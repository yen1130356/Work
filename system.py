# 引入網頁伺服器核心框架 (Flask)、樣本渲染 (render_template)、請求物件 (request)、JSON 回傳 (jsonify)、檔案送出 (send_file)
from flask import Flask, render_template, request, jsonify, send_file
# 引入 PostgreSQL 資料庫連線套件及其高級字典/游標工具 (DictCursor)
import psycopg2  
import psycopg2.extras
# 引入開源 Excel 新版擴展庫 (.xlsx)，主要用於動態生成與匯出統計報表
import openpyxl
# 引入高速數據分析矩陣庫 (Pandas)
import pandas as pd
import numpy as np
# 引入時間日期操作函數
from datetime import datetime, date, timedelta
# 引入系統作業路徑、記憶體二進位流處理、以及安全執行緒阻斷延遲功能
import os
import io
import time

# 💡 從專案內部的 download_report.py 導入 Selenium RPA 遠端下載模組與共享上傳資料夾路徑
from download_report import download_yesterday_report, DOWNLOAD_DIR
# 引入排程套件
from flask_apscheduler import APScheduler
from pytz import timezone

# 初始化 Flask 應用程式實例
app = Flask(__name__)

@app.context_processor
def inject_now():
    # 每次渲染 HTML 時，自動生成當前時間戳記字串
    return {'version': datetime.now().strftime('%Y%m%d%H%M%S')}

# 確保 Flask 配置中寫入時區
app.config['SCHEDULER_API_ENABLED'] = True
app.config['SCHEDULER_TIMEZONE'] = 'Asia/Taipei'

# 初始化與設定排程
scheduler = APScheduler()

# 啟動排程器
scheduler.init_app(app)
scheduler.start()

@scheduler.task('cron', id='auto_download_job',minute="*")
def scheduled_auto_sync():
    print(f"📢 [時區檢查] 伺服器本機標準時間(UTC/Local)目前為: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"⏰ [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 定時任務啟動：開始自動更新中榮昨日傳送數據...")
    
    with app.app_context():
        success = download_yesterday_report()
        if not success:
            print("❌ [定時任務中斷] 遠端 RPA 模擬登入下載失敗。")
            return

        time.sleep(1)
        
        files = [
            os.path.join(DOWNLOAD_DIR, f) 
            for f in os.listdir(DOWNLOAD_DIR) 
            if (f.endswith(".xlsx") or f.endswith(".xls")) and not f.startswith("~$")
        ]
        if not files: 
            print("❌ [定時任務中斷] 下載目錄未偵測到標準 Excel 檔案。")
            return
            
        latest_file = max(files, key=os.path.getctime)

        try:
            raw_df = pd.read_excel(latest_file, dtype=str)
        except Exception as e:
            print(f"❌ [定時任務中斷] Pandas 解析失敗: {str(e)}")
            return

        cleaned_df = run_r_preprocessing_engine(raw_df)
        
        # 💡 同步加上過濾邏輯：只保留任務時間或結束時間包含「昨天」的紀錄
        yesterday_obj = date.today() - timedelta(days=1)
        target_date_str = yesterday_obj.strftime("%Y-%m-%d")
        
        def is_in_target_range_scheduled(row):
            t_time = parse_datetime(row.get('任務時間'))
            e_time = parse_datetime(row.get('結束時間'))
            t_date = t_time[:10] if t_time else ""
            e_date = e_time[:10] if e_time else ""
            if t_date == target_date_str or e_date == target_date_str: return True
            if not t_date and not e_date: return True
            return False
            
        filtered_df = cleaned_df[cleaned_df.apply(is_in_target_range_scheduled, axis=1)]
        
        conn = get_db()
        c = conn.cursor()
        inserted = skipped = errors = 0
        affected_batches = set()

        INSERT_SQL = """INSERT INTO task_records (
            import_batch, 單號, 勤務中心, 申請者, 派工單位, 任務時間,
            起始地點, 結束地點, 任務, 特別指示, 病人姓名, 病歷號, 病房, 病床, 設備, 任務狀態,
            傳送人員, 建立時間, 預達時間, 派工時間, 派工時段,
            回應時間, 準備時間, 開始時間, 手圈時間, 結束時間, 執行時間,
            派工人員, 病人非病人, 緊急等級, 取消手動完成時間, 取消手動完成操作者, 取消手動完成原因, 手動狀態,
            等待操作時間, 等待操作者, 等待原因, 返回操作時間, 返回操作者, 返回原因,
            套餐母單號, 花費時間, 是否延遲, 派工紀錄, 藥袋,
            不需計算, 排程需排除, 不屬延遲, 班別
        ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        ON CONFLICT (單號, import_batch) DO UPDATE SET
            勤務中心 = EXCLUDED.勤務中心,
            申請者 = EXCLUDED.申請者,
            派工單位 = EXCLUDED.派工單位,
            任務時間 = EXCLUDED.任務時間,
            起始地點 = EXCLUDED.起始地點,
            結束地點 = EXCLUDED.結束地點,
            任務 = EXCLUDED.任務,
            特別指示 = EXCLUDED.特別指示,
            病人姓名 = EXCLUDED.病人姓名,
            病歷號 = EXCLUDED.病歷號,
            病房 = EXCLUDED.病房,
            病床 = EXCLUDED.病床,
            設備 = EXCLUDED.設備,
            任務狀態 = EXCLUDED.任務狀態,
            傳送人員 = EXCLUDED.傳送人員,
            建立時間 = EXCLUDED.建立時間,
            預達時間 = EXCLUDED.預達時間,
            派工時間 = EXCLUDED.派工時間,
            派工時段 = EXCLUDED.派工時段,
            回應時間 = EXCLUDED.回應時間,
            準備時間 = EXCLUDED.準備時間,
            開始時間 = EXCLUDED.開始時間,
            手圈時間 = EXCLUDED.手圈時間,
            結束時間 = EXCLUDED.結束時間,
            執行時間 = EXCLUDED.執行時間,
            派工人員 = EXCLUDED.派工人員,
            病人非病人 = EXCLUDED.病人非病人,
            緊急等級 = EXCLUDED.緊急等級,
            是否延遲 = EXCLUDED.是否延遲,
            派工紀錄 = EXCLUDED.派工紀錄,
            藥袋 = EXCLUDED.藥袋,
            排程需排除 = EXCLUDED.排程需排除,
            不屬延遲 = EXCLUDED.不屬延遲,
            班別 = EXCLUDED.班別,
            "不需計算" = CASE
                WHEN task_records."不需計算" LIKE '人工覆核判定%%' THEN task_records."不需計算"
                ELSE EXCLUDED.不需計算
            END,
            "延遲調整" = CASE
                WHEN task_records."不需計算" LIKE '人工覆核判定%%' THEN task_records."延遲調整"
                ELSE NULL
            END,
            "命中規則描述" = CASE
                WHEN task_records."不需計算" LIKE '人工覆核判定%%' THEN task_records."命中規則描述"
                ELSE NULL
            END"""

        for _, r in filtered_df.iterrows():
            c.execute("SAVEPOINT scheduled_row_savepoint")
            current_sn = r.get('單號', '未知單號')
            try:
                task_date_str = parse_datetime(r.get('任務時間'))
                row_batch = task_date_str[:10] if task_date_str else date.today().strftime("%Y-%m-%d")
                affected_batches.add(row_batch)
                
                c.execute(INSERT_SQL, (
                    row_batch,
                    parse_text(r.get('單號')), parse_text(r.get('勤務中心')), parse_text(r.get('申請者')), parse_text(r.get('派工單位')), parse_datetime(r.get('任務時間')),
                    parse_text(r.get('起始地點')), parse_text(r.get('結束地點')), parse_text(r.get('任務')), parse_text(r.get('特別指示')),
                    parse_text(r.get('病人姓名')), parse_text(r.get('病歷號')), parse_text(r.get('病房')), parse_text(r.get('病床')), parse_text(r.get('設備')), parse_text(r.get('任務狀態')),
                    parse_text(r.get('傳送人員')), parse_datetime(r.get('建立時間')), parse_datetime(r.get('預達時間')), parse_datetime(r.get('派工時間')), parse_int(r.get('派工時段')),
                    parse_datetime(r.get('回應時間')), parse_datetime(r.get('準備時間')), parse_datetime(r.get('開始時間')), parse_datetime(r.get('手圈時間')), parse_datetime(r.get('結束時間')), parse_int(r.get('執行時間')),
                    parse_text(r.get('派工人員')), parse_text(r.get('病人非病人')), parse_text(r.get('緊急等級')),
                    parse_datetime(r.get('取消手動完成時間')), parse_text(r.get('取消手動完成操作者')), parse_text(r.get('取消手動完成原因')), parse_text(r.get('手動狀態')),
                    parse_datetime(r.get('等待操作時間')), parse_text(r.get('等待操作者')), parse_text(r.get('等待原因')),
                    parse_datetime(r.get('返回操作時間')), parse_text(r.get('返回操作者')), parse_text(r.get('返回原因')),
                    parse_text(r.get('套餐母單號')), parse_text(r.get('花費時間')), parse_text(r['是否延遲']), parse_text(r.get('派工紀錄')), parse_text(r.get('藥袋')),
                    parse_text(r['不需計算']), parse_text(r['排程需排除']), parse_text(r['不屬延遲']), parse_text(r['班別'])
                ))
                if c.rowcount > 0: inserted += 1
                else: skipped += 1
                c.execute("RELEASE SAVEPOINT scheduled_row_savepoint")
            except Exception as ex:
                c.execute("ROLLBACK TO SAVEPOINT scheduled_row_savepoint")
                print(f"⚠️ [定時任務警告] 單號 {current_sn} 寫入資料庫失敗，已跳過。原因: {str(ex)}")
                errors += 1

        conn.commit()
        
        for b in affected_batches:
            try:
                execute_system_7_cleaning_rules(conn, b)
            except Exception as clean_ex:
                print(f"⚠️ 批次 {b} 運行清洗規則失敗: {str(clean_ex)}")
                
        conn.commit()
        
        if os.path.exists(latest_file): 
            os.remove(latest_file)
            
        print(f"✅ [定時任務成功] 匯入完成！新增: {inserted}筆, 跳過: {skipped}筆, 錯誤: {errors}筆")

        # === 🛠️ 核心修改：精準儲存各項卡片統計數據與時間戳記 ===
        yesterday_obj = date.today() - timedelta(days=1)
        target_date_str = yesterday_obj.strftime("%Y-%m-%d")
        total_rows = len(filtered_df) # 本次實際處理的總筆數

        try:
            # 建立一個共用的更新 SQL
            UPSERT_STATUS_SQL = """
                INSERT INTO porter_system_status (key, value, updated_at)
                VALUES (%s, %s, CURRENT_TIMESTAMP)
                ON CONFLICT (key) 
                DO UPDATE SET value = EXCLUDED.value, updated_at = CURRENT_TIMESTAMP
            """
            
            # 1. 記錄最新的資料日期 (資料日期批次)
            c.execute(UPSERT_STATUS_SQL, ('latest_sync_date', target_date_str))
            
            # 2. 將 4 個卡片數據分別寫入對應的 key，方便前端渲染
            c.execute(UPSERT_STATUS_SQL, ('sync_total', str(total_rows)))
            c.execute(UPSERT_STATUS_SQL, ('sync_success', str(inserted)))
            c.execute(UPSERT_STATUS_SQL, ('sync_skipped', str(skipped)))
            c.execute(UPSERT_STATUS_SQL, ('sync_failed', str(errors)))
            
            # 3. 另外保留一個方便顯示的文字備註
            stats_msg = f"自動定時更新成功(原始:{total_rows}, 新增:{inserted}, 跳過:{skipped}, 失敗:{errors})"
            c.execute(UPSERT_STATUS_SQL, ('latest_sync_stats', stats_msg))
            
            conn.commit()
            print(f"📌 [系統狀態更新] 定時更新卡片數據同步完畢！時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        except Exception as status_ex:
            print(f"❌ 寫入系統狀態失敗: {str(status_ex)}")
        finally:
            conn.close()

# ============================================================
# 本地連線設定區
# ============================================================
DB_CONFIG = {
    "host": "127.0.0.1",         # 資料庫本機 IP
    "database": "postgres",      # 資料庫名稱
    "user": "postgres_local",    # 資料庫帳號
    "password": "",              # 資料庫密碼
    "port": "5432"               # PostgreSQL 標準通訊埠
}

def get_db():
    """ 建立並返回一個全新的資料庫連線實例 """
    return psycopg2.connect(**DB_CONFIG)

def update_system_status(batch_name, msg, total=0, success=0, skipped=0, failed=0):
    """更新系統最新狀態與統計卡片數據"""
    conn = get_db()
    try:
        c = conn.cursor()
        # 先清除舊狀態，保持只有最新的一筆
        c.execute("DELETE FROM system_status")
        c.execute("""
            INSERT INTO system_status 
            (latest_import_batch, status_message, total_rows, success_rows, skipped_rows, failed_rows, updated_at)
            VALUES (%s, %s, %s, %s, %s, %s, NOW())
        """, (batch_name, msg, total, success, skipped, failed))
        conn.commit()
    except Exception as e:
        print(f"❌ 更新系統狀態失敗: {str(e)}")
        conn.rollback()

# ============================================================
# 工具與核心數據清洗函式 (文字過濾與資料轉型)
# ============================================================
def parse_datetime(value):
    """ 強固型時間解析器：將 Excel 內各種混雜的時間字串標準化為 PostgreSQL 接受的 YYYY-MM-DD HH:MM:SS """
    if value is None or str(value).strip() in ("", "-"): return None
    if isinstance(value, datetime): return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y/%m/%d %H:%M:%S"):
            try: return datetime.strptime(value.strip(), fmt).strftime("%Y-%m-%d %H:%M:%S")
            except ValueError: continue
    return None

def parse_int(value):
    """ 整數安全轉換器：排除中榮 Excel 中的 '-' 或空白，防止轉型失敗導致資料庫拒絕入庫 """
    try: return int(value) if value not in (None, "", "-") else None
    except (ValueError, TypeError): return None

def parse_text(value):
    """ 字串清洗過濾器：將文字前後的空白修剪乾淨，並將無效的系統占位符 '-' 自動降級為 NULL """
    if value is None: return None
    s = str(value).strip()
    return None if s in ("", "-") else s

def init_db():
    """ 🏥 資料庫維護大腦：建立核心任務明細表並自動動態增開延伸欄位 """
    conn = get_db()
    c = conn.cursor()
    
    # 建立基礎任務明細表
    c.execute("""
        CREATE TABLE IF NOT EXISTS task_records (
            id SERIAL PRIMARY KEY,
            import_batch TEXT NOT NULL,
            單號 TEXT,
            勤務中心 TEXT,
            申請者 TEXT,
            派工單位 TEXT,
            任務時間 TEXT,
            起始地點 TEXT,
            結束地點 TEXT,
            任務 TEXT,
            特別指示 TEXT,
            病人姓名 TEXT,
            病歷號 TEXT,
            病房 TEXT,
            病床 TEXT,
            設備 TEXT,
            任務狀態 TEXT,
            傳送人員 TEXT,
            建立時間 TEXT,
            預達時間 TEXT,
            派工時間 TEXT,
            派工時段 INTEGER,
            回應時間 TEXT,
            準備時間 TEXT,
            開始時間 TEXT,
            手圈時間 TEXT,
            結束時間 TEXT,
            執行時間 INTEGER,
            派工人員 TEXT,
            病人非病人 TEXT,
            緊急等級 TEXT,
            取消手動完成時間 TEXT,
            取消手動完成操作者 TEXT,
            取消手動完成原因 TEXT,
            手動狀態 TEXT,
            等待操作時間 TEXT,
            等待操作者 TEXT,
            等待原因 TEXT,
            返回操作時間 TEXT,
            返回操作者 TEXT,
            返回原因 TEXT,
            套餐母單號 TEXT,
            花費時間 TEXT,
            是否延遲 TEXT,
            派工紀錄 TEXT,
            藥袋 TEXT,
            UNIQUE (單號, import_batch)
        )
    """)

    # 建立系統狀態表（用來記錄自動更新到哪一天）
    c.execute("""
    CREATE TABLE IF NOT EXISTS system_status (
        id SERIAL PRIMARY KEY,
        latest_update TEXT,
        excel_total INTEGER DEFAULT 0,
        success_count INTEGER DEFAULT 0,
        skip_count INTEGER DEFAULT 0,
        error_count INTEGER DEFAULT 0,
        updated_at TEXT
        )
    """)
    
    # 初始化一筆「最新更新日期」的紀錄（如果不存在的話）
    c.execute("""
        INSERT INTO porter_system_status (key, value) 
        VALUES ('latest_sync_date', '尚無資料') 
        ON CONFLICT (key) DO NOTHING
    """)
    
    # 1. 建立「等待原因排除」規則表
    c.execute("""
        CREATE TABLE IF NOT EXISTS rule_wait_reasons (
            id SERIAL PRIMARY KEY,
            reason TEXT UNIQUE NOT NULL,
            action TEXT NOT NULL DEFAULT '未延遲', -- 未延遲 / 需檢查 / 延遲
            enabled BOOLEAN DEFAULT TRUE,
            note TEXT
        )
    """)
    
    # 預塞你目前的 5 個寫死條件
    default_reasons = ['等單位通知', '前置作業未完成', '病人個人因素', '病人不在', '病人檢查中']
    for reason in default_reasons:
        c.execute("""
            INSERT INTO rule_wait_reasons (reason, action, enabled, note)
            VALUES (%s, '未延遲', TRUE, '中榮原生不可歸責等待原因')
            ON CONFLICT (reason) DO NOTHING
        """, (reason,))

    # 2. 建立「關鍵字排除」規則表 (包含特別指示、等待原因等)
    c.execute("""
        CREATE TABLE IF NOT EXISTS rule_keywords (
            id SERIAL PRIMARY KEY,
            target_field TEXT NOT NULL,  -- 特別指示 / 等待原因 / 派工紀錄
            keyword TEXT NOT NULL,
            match_type TEXT NOT NULL DEFAULT 'contains', -- contains / equals
            action TEXT NOT NULL DEFAULT '需檢查',
            enabled BOOLEAN DEFAULT TRUE,
            note TEXT,
            UNIQUE(target_field, keyword)
        )
    """)
    
    # 預塞目前的規則 7 與 8 的關鍵字
    default_keywords = [
        ('特別指示', '備註時間', 'contains', '需檢查', '需人工確認時間描述'),
        ('特別指示', '同一位勤務', 'contains', '需檢查', '特指排除條件'),
        ('特別指示', '非勤務工作', 'contains', '需檢查', '可能不列入延遲'),
        ('特別指示', '故障', 'contains', '未延遲', '設備環境故障排除'),
        ('特別指示', '內勤工作', 'contains', '未延遲', '內勤排除'),
        ('特別指示', '現場狀況', 'contains', '未延遲', '突發狀況排除'),
        ('等待原因', '故障', 'contains', '未延遲', '等待原因含故障排除'),
        ('等待原因', '內勤工作', 'contains', '未延遲', '等待原因含內勤排除'),
        ('等待原因', '現場狀況', 'contains', '未延遲', '等待原因含現場狀況排除'),
    ]
    for field, kw, m_type, act, nt in default_keywords:
        c.execute("""
            INSERT INTO rule_keywords (target_field, keyword, match_type, action, enabled, note)
            VALUES (%s, %s, %s, %s, TRUE, %s)
            ON CONFLICT (target_field, keyword) DO NOTHING
        """, (field, kw, m_type, act, nt))

    # 3. 建立「排程地點時段」規則表
    c.execute("""
        CREATE TABLE IF NOT EXISTS rule_schedule_excludes (
            id SERIAL PRIMARY KEY,
            end_location TEXT NOT NULL,
            expected_time TEXT NOT NULL, -- 預達時段 (HH:MM:SS)
            label TEXT,
            enabled BOOLEAN DEFAULT TRUE,
            UNIQUE(end_location, expected_time)
        )
    """)

    try:
        c.execute('ALTER TABLE task_records ADD COLUMN IF NOT EXISTS "不需計算" TEXT DEFAULT \'\';')
        c.execute('ALTER TABLE task_records ADD COLUMN IF NOT EXISTS "排程需排除" TEXT DEFAULT \'\';')
        c.execute('ALTER TABLE task_records ADD COLUMN IF NOT EXISTS "不屬延遲" TEXT DEFAULT \'\';')
        c.execute('ALTER TABLE task_records ADD COLUMN IF NOT EXISTS "班別" TEXT DEFAULT \'\';')
        c.execute('ALTER TABLE task_records ADD COLUMN IF NOT EXISTS "延遲調整" TEXT DEFAULT NULL;')
        c.execute('ALTER TABLE task_records ADD COLUMN IF NOT EXISTS "命中規則描述" TEXT DEFAULT NULL;')
    except Exception as e:
        print(f"💡 欄位檢查提示 (可能已存在): {str(e)}")
        conn.rollback()

    # 針對頻繁進行條件檢索與報表統計的欄位建立優化索引
    c.execute("CREATE INDEX IF NOT EXISTS idx_task_time ON task_records (任務時間)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_import_batch ON task_records (import_batch)")
    
    conn.commit()
    conn.close()
    print("🏥 PostgreSQL 資料庫結構前處理與『延遲調整』欄位自我校正檢查完成！")

def run_r_preprocessing_engine(df):
    if df.empty:
        return df

    # 強制進行型態標準化修剪，防止空值 NaN 導致文字處理中斷
    df['申請者'] = df['申請者'].fillna('').astype(str).str.strip()
    df['緊急等級'] = df['緊急等級'].fillna('').astype(str).str.strip()
    df['單號'] = df['單號'].fillna('').astype(str).str.strip()
    df['套餐母單號'] = df['套餐母單號'].fillna('').astype(str).str.strip()
    df['起始地點'] = df['起始地點'].fillna('').astype(str).str.strip()
    df['結束地點'] = df['結束地點'].fillna('').astype(str).str.strip()
    df['預達時間'] = df['預達時間'].fillna('').astype(str).str.strip()
    df['開始時間'] = df['開始時間'].fillna('').astype(str).str.strip()
    df['手動狀態'] = df['手動狀態'].fillna('').astype(str).str.strip()
    df['任務'] = df['任務'].fillna('').astype(str).str.strip()
    df['是否延遲'] = df['是否延遲'].fillna('').astype(str).str.strip()
    df['勤務中心'] = df['勤務中心'].fillna('').astype(str).str.strip()

    # 1. 預建衍生衍生時間字串 
    df['預達時段'] = df['預達時間'].apply(lambda x: x[11:19] if len(x) >= 19 else '')

    # 2. 讀取專案目錄底下的「條件列表.xlsx」交叉對照組 Sheet
    con_df = pd.DataFrame(columns=['結束地點', '時間', '標籤'])
    task_late_list = []
    or_list = []
    
    cond_file = "條件列表.xlsx"
    if os.path.exists(cond_file):
        try:
            # Sheet 1: 組合排除條件 
            con_excel = pd.read_excel(cond_file, sheet_name=0)
            con_df = con_excel.copy()
            con_df.columns = ['結束地點', '時間', '標籤']
            con_df['時間'] = con_df['時間'].astype(str).apply(lambda x: x[11:19] if len(x) >= 19 else x)
            
            # Sheet 2: 接送病人任務清單
            task_late_df = pd.read_excel(cond_file, sheet_name=1, header=None)
            task_late_list = task_late_df[0].fillna('').astype(str).str.strip().tolist()
            
            # Sheet 3: OR 開刀房地點矩陣
            or_df = pd.read_excel(cond_file, sheet_name=2, header=None)
            or_list = or_df[0].fillna('').astype(str).str.strip().tolist()
        except Exception as e:
            print(f"⚠️ 讀取 條件列表.xlsx 失敗，將使用空矩陣運行: {str(e)}")

    # 3. 「不需計算」：自建任務(x1)+返回排程(x2)+子單號(x3)+CSR車(x4) 拼接
    df['不需計算'] = ''
    def calc_not_need(row):
        x1 = "自建任務" if row['申請者'] == "Auto" else ""
        x2 = "返回之排程" if (row['緊急等級'] == "排程" and len(row['單號']) > 10) else ""
        x3 = "子單號" if (row['套餐母單號'] != "" and row['單號'] != row['套餐母單號']) else ""
        x4 = "CSR衛材車" if row['起始地點'] == "第一醫療大樓 地下一樓 - CSR供應中心" else ""
        return f"{x1}{x2}{x3}{x4}"
    df['不需計算'] = df.apply(calc_not_need, axis=1)

    # 4. 「排程需排除」：比對結束地點與預達時段組合是否命中特定標籤，包含「剃雉」
    df['排程需排除'] = ''
    con_dict = {}
    for _, r in con_df.iterrows():
        key = f"{str(r['結束地點']).strip()}{str(r['時間']).strip()}"
        con_dict[key] = str(r['標籤']).strip()

    def calc_exclude(row):
        if row['任務'] == "剃雉":
            return "剃雉"
        if row['緊急等級'] == "排程":
            match_key = f"{row['結束地點']}{row['預達時段']}"
            if match_key in con_dict:
                return con_dict[match_key]
        return ""
    df['排程需排除'] = df.apply(calc_exclude, axis=1)

    # 5. 「洗腎室特判」：針對 HD 洗腎室排程任務，直接根據時間大小覆寫原始「是否延遲」欄位
    def fix_hd_dialysis(row):
        if row['起始地點'] == "第二醫療大樓 地下二樓 - 【HD】洗腎室" and row['緊急等級'] == "排程" and row['手動狀態'] != "取消任務":
            if row['開始時間'] > row['預達時間']:
                return "延遲"
            else:
                return "未延遲"
        return row['是否延遲']
    df['是否延遲'] = df.apply(fix_hd_dialysis, axis=1)

    # 6. 「不屬延遲」：非病人任務放寬校正與開刀房特急件校正
    df['不屬延遲'] = df['是否延遲']
    unique_tasks = df['任務'].unique().tolist()
    non_patient_tasks = [t for t in unique_tasks if t not in task_late_list]

    def calc_not_real_late(row):
        current_status = row['是否延遲']
        if current_status == "延遲" and row['不需計算'] == "":
            if row['緊急等級'] != "排程" and row['任務'] in non_patient_tasks:
                return "未延遲"
            if row['緊急等級'] == "特急" and row['任務'] == "病人送檢查/復健" and row['結束地點'] not in or_list:
                return "未延遲"
        return current_status
    df['不屬延遲'] = df.apply(calc_not_real_late, axis=1)

    # 7. 「班別劃分」：白班、小夜班、大夜班切分
    def calc_shift(row):
        if not row['預達時段']:
            return ""
        try:
            hour = int(row['預達時段'][:2])
        except:
            return ""
            
        if hour == 2:
            return "大夜班"
        elif 8 <= hour < 16:
            return f"{row['勤務中心']}勤白班"
        elif 16 <= hour < 24:
            return "小夜班"
        else:
            return "大夜班"
    df['班別'] = df.apply(calc_shift, axis=1)

    return df

def execute_system_7_cleaning_rules(conn, import_batch):
    """
    🎯 智慧動態洗滌引擎：從資料庫讀取規則，全自動判定並留存脈絡
    """
    c_clean = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
    
    # 1. 預先載入資料庫中所有「啟用中」的規則清單
    c_clean.execute("SELECT reason, action, note FROM rule_wait_reasons WHERE enabled = TRUE")
    db_wait_reasons = {r["reason"]: (r["action"], r["note"]) for r in c_clean.fetchall()}
    
    c_clean.execute("SELECT target_field, keyword, match_type, action, note FROM rule_keywords WHERE enabled = TRUE")
    db_keywords = [dict(r) for r in c_clean.fetchall()]
    
    # 2. 撈出當前批次尚未被過濾判定的任務明細
    c_clean.execute("""
        SELECT id, 任務, 是否延遲, 特別指示, 等待原因, 派工紀錄, 傳送人員, 緊急等級, "不屬延遲", "不需計算" 
        FROM task_records 
        WHERE import_batch = %s AND "延遲調整" IS NULL
    """, (import_batch,))
    
    to_clean_rows = [dict(row) for row in c_clean.fetchall()]
    results = {"原始未延遲": 0, "原始延遲": 0, "規則1": 0, "規則2": 0, "規則3": 0, "規則4": 0, "規則7": 0, "規則8": 0}
    
    # 💡【優化：宣告收集器】將原本逐筆更新的動作，改為先收集在記憶體中
    update_batch_data = []

    for r in to_clean_rows:
        rid = r["id"]
        task = r["任務"] or ""
        spec = r["特別指示"] or ""
        wait_reason = r["等待原因"] or ""
        dispatch_log = r["派工紀錄"] or ""
        person = r["傳送人員"] or ""
        current_r_status = r["不屬延遲"] or r["是否延遲"]
        current_not_need = r["不需計算"] or ""
        current_hit_msg = None

        # 基礎防禦：如果原本就是未延遲，直接認列
        if current_r_status == '未延遲':
            update_batch_data.append(('未延遲', current_not_need, current_hit_msg, rid))
            results["原始未延遲"] += 1
            continue

        hit_rule = False
        hit_message = ""

        # ==========================================
        # 🧪 動態規則檢測 A：比對「等待原因排除清單」
        # ==========================================
        if wait_reason in db_wait_reasons:
            action, note = db_wait_reasons[wait_reason]
            hit_message = f"命中等待原因[{wait_reason}] -> 調整為({action})"
            update_batch_data.append((action, f"{current_not_need}{wait_reason}", hit_message, rid))
            results["規則1"] += 1
            hit_rule = True

        # ==========================================
        # 🧪 動態規則檢測 B：比對「關鍵字排除清單」
        # ==========================================
        if not hit_rule:
            for kw_rule in db_keywords:
                field_value = ""
                if kw_rule["target_field"] == "特別指示": field_value = spec
                elif kw_rule["target_field"] == "等待原因": field_value = wait_reason
                elif kw_rule["target_field"] == "派工紀錄": field_value = dispatch_log
                
                # 實作包含（contains）判定
                if kw_rule["match_type"] == "contains" and kw_rule["keyword"] in field_value:
                    action = kw_rule["action"]
                    hit_message = f"命中[{kw_rule['target_field']}]關鍵字({kw_rule['keyword']}) -> 調整為({action})"
                    update_batch_data.append((action, f"{current_not_need}關鍵字過濾", hit_message, rid))
                    
                    if action == '需檢查': results["規則7"] += 1
                    else: results["規則8"] += 1
                    hit_rule = True
                    break

        # ==========================================
        # 🛡️ 系統底層保留規則：維持原寫死的 OR 房/急領藥等高度複雜判定
        # ==========================================
        if not hit_rule:
            if '重新建立' in dispatch_log:
                update_batch_data.append(('未延遲', current_not_need, '系統硬編碼：重新建立件', rid))
                results["規則2"] += 1
                hit_rule = True
            elif task == '送開刀房' and ('OR' in person or '刀房內' in person):
                update_batch_data.append(('未延遲', current_not_need, '系統硬編碼：開刀房內勤人員', rid))
                results["規則3"] += 1
                hit_rule = True
            elif task == '急領藥品' and ('內' in person or '其他人員' in person):
                update_batch_data.append(('需檢查', current_not_need, '系統硬編碼：急領藥品待審', rid))
                results["規則4"] += 1
                hit_rule = True

        # ==========================================
        # 最終認列：若上述動態與靜態規則皆未命中，判定為真正的「延遲」
        # ==========================================
        if not hit_rule:
            update_batch_data.append(('延遲', current_not_need, current_hit_msg, rid))
            results["原始延遲"] += 1
            
    # 💡【優化：批次大爆發】一次性將上萬筆的判定結果更新回資料庫
    if update_batch_data:
        psycopg2.extras.execute_batch(c_clean, """
            UPDATE task_records 
            SET "延遲調整" = %s, "不需計算" = %s, "命中規則描述" = %s
            WHERE id = %s
        """, update_batch_data, page_size=1000)
        
    return results

# ============================================================
# Web 路由接口 (與前端互動傳輸端)
# ============================================================
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/auto_sync", methods=["POST"])
def auto_sync():
    # 💡 判斷是網頁發送的 POST 請求還是定時任務直接調用
    # 如果 request.has_data 或 ctx 存在則解析，若無則代表是排程靜態調用
    try:
        req_data = request.json or {}
        date_from_str = req_data.get("date_from")  
        date_to_str = req_data.get("date_to")      
    except:
        req_data = {}
        date_from_str = None
        date_to_str = None
    
    if date_from_str:
        # A. 網頁人工手動觸發「更新指定區間」
        print(f"🚀 前端發送過來的原始同步請求: {date_from_str} 至 {date_to_str}")
        import_batch = date_from_str.replace('/', '-')
        actual_to = date_to_str if date_to_str else date_from_str
        success = download_yesterday_report(str(date_from_str), str(actual_to))
    else:
        # B. ⏰ 早上 07:00 定時自動更新昨日整天數據
        yesterday_obj = date.today() - timedelta(days=1)
        import_batch = yesterday_obj.strftime("%Y-%m-%d") # 實打實的昨天 (YYYY-MM-DD)
        print(f"⏰ 背景定時任務：啟動自動更新，目標鎖定昨日日期 [{import_batch}]")
        success = download_yesterday_report()

    if os.path.exists(DOWNLOAD_DIR):
        for f in os.listdir(DOWNLOAD_DIR):
            if f.endswith(".xlsx") or f.endswith(".xls") or f.startswith("~$"):
                try: os.remove(os.path.join(DOWNLOAD_DIR, f))
                except: pass

    if not success:
        return jsonify({"success": False, "message": "遠端報表下載失敗，請檢查 ePorter 模擬登入狀態。"})

    time.sleep(2)  
    
    files = [
        os.path.join(DOWNLOAD_DIR, f) 
        for f in os.listdir(DOWNLOAD_DIR) 
        if (f.endswith(".xlsx") or f.endswith(".xls")) and not f.startswith("~$")
    ]
    if not files: 
        return jsonify({"success": False, "message": "同步異常：未成功將 Excel 下載至暫存目錄！"})
        
    latest_file = max(files, key=os.path.getctime)

    try:
        raw_df = pd.read_excel(latest_file, dtype=str)
    except Exception as e:
        return jsonify({"success": False, "message": f"Pandas 解析下載 Excel 失敗: {str(e)}"})

    conn = get_db()
    c = conn.cursor()
    inserted = skipped = errors = 0  
    affected_batches = set()

    # 清空本次要覆蓋匯入的批次舊數據
    try:
        c.execute("DELETE FROM task_records WHERE import_batch = %s", (import_batch,))
    except Exception as e:
        conn.rollback()

    INSERT_SQL = """INSERT INTO task_records (
        import_batch, 單號, 勤務中心, 申請者, 派工單位, 任務時間,
        起始地點, 結束地點, 任務, 特別指示, 病人姓名, 病歷號, 病房, 病床, 設備, 任務狀態,
        傳送人員, 建立時間, 預達時間, 派工時間, 派工時段,
        回應時間, 準備時間, 開始時間, 手圈時間, 結束時間, 執行時間,
        派工人員, 病人非病人, 緊急等級, 取消手動完成時間, 取消手動完成操作者, 取消手動完成原因, 手動狀態,
        等待操作時間, 等待操作者, 等待原因, 返回操作時間, 返回操作者, 返回原因,
        套餐母單號, 花費時間, 是否延遲, 派工紀錄, 藥袋,
        不需計算, 排程需排除, 不屬延遲, 班別
    ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
    ON CONFLICT (單號, import_batch) DO NOTHING"""

    records_to_insert = []
    for _, r in raw_df.iterrows():
        task_time_raw = parse_text(r.get('任務時間'))
        end_time_raw = parse_text(r.get('結束時間'))
        
        has_task_match = task_time_raw and import_batch in task_time_raw
        has_end_match = end_time_raw and import_batch in end_time_raw
        
        # 如果既不是這個日期的建立單、也不是這個日期的結束單，才予以過濾
        if not (has_task_match or has_end_match):
            continue

        row_batch = import_batch
        affected_batches.add(row_batch)

        records_to_insert.append((
            row_batch,
            parse_text(r.get('單號')), parse_text(r.get('勤務中心')), parse_text(r.get('申請者')), parse_text(r.get('派工單位')), parse_datetime(r.get('任務時間')),
            parse_text(r.get('起始地點')), parse_text(r.get('結束地點')), parse_text(r.get('任務')), parse_text(r.get('特別指示')),
            parse_text(r.get('病人姓名')), parse_text(r.get('病歷號')), parse_text(r.get('病房')), parse_text(r.get('病床')), parse_text(r.get('設備')), parse_text(r.get('任務狀態')),
            parse_text(r.get('傳送人員')), parse_datetime(r.get('建立時間')), parse_datetime(r.get('預達時間')), parse_datetime(r.get('派工時間')), parse_int(r.get('派工時段')),
            parse_datetime(r.get('回應時間')), parse_datetime(r.get('準備時間')), parse_datetime(r.get('開始時間')), parse_datetime(r.get('手圈時間')), parse_datetime(r.get('結束時間')), parse_int(r.get('執行時間')),
            parse_text(r.get('派工人員')), parse_text(r.get('病人非病人')), parse_text(r.get('緊急等級')),
            parse_datetime(r.get('取消手動完成時間')), parse_text(r.get('取消手動完成操作者')), parse_text(r.get('取消手動完成原因')), parse_text(r.get('手動狀態')),
            parse_datetime(r.get('等待操作時間')), parse_text(r.get('等待操作者')), parse_text(r.get('等待原因')),
            parse_datetime(r.get('返回操作時間')), parse_text(r.get('返回操作者')), parse_text(r.get('返回原因')),
            parse_text(r.get('套餐母單號')), parse_text(r.get('花費時間')), parse_text(r.get('是否延遲')), parse_text(r.get('派工紀錄')), parse_text(r.get('藥袋')),
            parse_text(r.get('不需計算')), parse_text(r.get('排程需排除')), parse_text(r.get('不屬延遲')), parse_text(r.get('班別'))
        ))
            
    for rec in records_to_insert:
        try:
            c.execute(INSERT_SQL, rec)
            if c.rowcount > 0: inserted += 1
            else: skipped += 1
        except Exception as ex:
            errors += 1

    total_excel_rows = len(records_to_insert)
    conn.commit()

    results = {}
    for b in affected_batches:
        try:
            batch_res = execute_system_7_cleaning_rules(conn, b)
            for k, v in batch_res.items():
                results[k] = results.get(k, 0) + v
        except Exception as clean_ex:
            print(f"⚠️ 批次 {b} 運行清洗規則失敗: {str(clean_ex)}")

    # === 網頁同步成功後，更新數據匯入報告卡片 ===
    try:
        current_time_str = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        UPSERT_STATUS_SQL = """
            INSERT INTO porter_system_status (key, value, updated_at)
            VALUES (%s, %s, CURRENT_TIMESTAMP)
            ON CONFLICT (key) 
            DO UPDATE SET value = EXCLUDED.value, updated_at = CURRENT_TIMESTAMP
        """
        c.execute(UPSERT_STATUS_SQL, ('latest_sync_date', import_batch))
        c.execute(UPSERT_STATUS_SQL, ('sync_total', str(total_excel_rows)))
        c.execute(UPSERT_STATUS_SQL, ('sync_success', str(inserted)))
        c.execute(UPSERT_STATUS_SQL, ('sync_skipped', str(skipped)))
        c.execute(UPSERT_STATUS_SQL, ('sync_failed', str(errors)))
        print(f"📌 [系統狀態] 網頁介面同步卡片數據已寫入資料庫。時間：{current_time_str}")
    except Exception as status_ex:
        print(f"❌ 網頁同步更新卡片數據失敗: {str(status_ex)}")

    conn.commit()
    conn.close()
    
    if os.path.exists(latest_file): 
        try: os.remove(latest_file)
        except: pass
    
    summary_msg = f"Excel 原始符合 {total_excel_rows} 筆 / 成功新增 {inserted} 筆 / 已存在跳過 {skipped} 筆 / 錯誤失敗 {errors} 筆"
    print(f"📋 [同步完成報告] {summary_msg}")
    
    return jsonify({
        "success": True, 
        "imported": summary_msg,  
        "cleaned_details": results,
        "stats": { "total": total_excel_rows, "success": inserted, "skipped": skipped, "error": errors }
    })

#手動上傳
@app.route("/api/upload", methods=["POST"])
def upload_file():
    if 'file' not in request.files: return jsonify({"success": False, "message": "未偵測到上傳檔案"})
    f = request.files['file']
    if f.filename == '': return jsonify({"success": False, "message": "未選取檔案"})
    
    if f and (f.filename.endswith('.xlsx') or f.filename.endswith('.xls')):
        file_path = os.path.join(DOWNLOAD_DIR, f.filename)
        f.save(file_path)
        try:
            raw_df = pd.read_excel(file_path, dtype=str)
        except Exception as e:
            if os.path.exists(file_path): os.remove(file_path)
            return jsonify({"success": False, "message": f"解析 Excel 失敗: {str(e)}"})

        cleaned_df = run_r_preprocessing_engine(raw_df)
        total_excel_rows = len(cleaned_df)

        conn = get_db()
        c = conn.cursor()
        inserted = skipped = errors = 0  
        affected_batches = set()

        INSERT_SQL = """INSERT INTO task_records (
            import_batch, 單號, 勤務中心, 申請者, 派工單位, 任務時間,
            起始地點, 結束地點, 任務, 特別指示, 病人姓名, 病歷號, 病房, 病床, 設備, 任務狀態,
            傳送人員, 建立時間, 預達時間, 派工時間, 派工時段,
            回應時間, 準備時間, 開始時間, 手圈時間, 結束時間, 執行時間,
            派工人員, 病人非病人, 緊急等級, 取消手動完成時間, 取消手動完成操作者, 取消手動完成原因, 手動狀態,
            等待操作時間, 等待操作者, 等待原因, 返回操作時間, 返回操作者, 返回原因,
            套餐母單號, 花費時間, 是否延遲, 派工紀錄, 藥袋,
            不需計算, 排程需排除, 不屬延遲, 班別
        ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        ON CONFLICT (單號, import_batch) DO UPDATE SET
            勤務中心 = EXCLUDED.勤務中心,
            申請者 = EXCLUDED.申請者,
            派工單位 = EXCLUDED.派工單位,
            任務時間 = EXCLUDED.任務時間,
            起始地點 = EXCLUDED.起始地點,
            結束地點 = EXCLUDED.結束地點,
            任務 = EXCLUDED.任務,
            特別指示 = EXCLUDED.特別指示,
            病人姓名 = EXCLUDED.病人姓名,
            病歷號 = EXCLUDED.病歷號,
            病房 = EXCLUDED.病房,
            病床 = EXCLUDED.病床,
            設備 = EXCLUDED.設備,
            任務狀態 = EXCLUDED.任務狀態,
            傳送人員 = EXCLUDED.傳送人員,
            建立時間 = EXCLUDED.建立時間,
            預達時間 = EXCLUDED.預達時間,
            派工時間 = EXCLUDED.派工時間,
            派工時段 = EXCLUDED.派工時段,
            回應時間 = EXCLUDED.回應時間,
            準備時間 = EXCLUDED.準備時間,
            開始時間 = EXCLUDED.開始時間,
            手圈時間 = EXCLUDED.手圈時間,
            結束時間 = EXCLUDED.結束時間,
            執行時間 = EXCLUDED.執行時間,
            派工人員 = EXCLUDED.派工人員,
            病人非病人 = EXCLUDED.病人非病人,
            緊急等級 = EXCLUDED.緊急等級,
            是否延遲 = EXCLUDED.是否延遲,
            派工紀錄 = EXCLUDED.派工紀錄,
            藥袋 = EXCLUDED.藥袋,
            排程需排除 = EXCLUDED.排程需排除,
            不屬延遲 = EXCLUDED.不屬延遲,
            班別 = EXCLUDED.班別,
            -- 如果該筆紀錄已經被人工覆核過，則保留原有的審核結果，不再用 NULL 或新值覆蓋
            "不需計算" = CASE
                WHEN task_records."不需計算" LIKE '人工覆核判定%%' THEN task_records."不需計算"
                ELSE EXCLUDED.不需計算
            END,
            "延遲調整" = CASE
                WHEN task_records."不需計算" LIKE '人工覆核判定%%' THEN task_records."延遲調整"
                ELSE NULL
            END,
            "命中規則描述" = CASE
                WHEN task_records."不需計算" LIKE '人工覆核判定%%' THEN task_records."命中規則描述"
                ELSE NULL
            END"""

        records_to_insert = []
        for _, r in cleaned_df.iterrows():
            task_date_str = parse_datetime(r.get('任務時間'))
            row_batch = task_date_str[:10] if task_date_str else date.today().strftime("%Y-%m-%d")
            affected_batches.add(row_batch)
            
            records_to_insert.append((
                row_batch,
                parse_text(r.get('單號')), parse_text(r.get('勤務中心')), parse_text(r.get('申請者')), parse_text(r.get('派工單位')), parse_datetime(r.get('任務時間')),
                parse_text(r.get('起始地點')), parse_text(r.get('結束地點')), parse_text(r.get('任務')), parse_text(r.get('特別指示')),
                parse_text(r.get('病人姓名')), parse_text(r.get('病歷號')), parse_text(r.get('病房')), parse_text(r.get('病床')), parse_text(r.get('設備')), parse_text(r.get('任務狀態')),
                parse_text(r.get('傳送人員')), parse_datetime(r.get('建立時間')), parse_datetime(r.get('預達時間')), parse_datetime(r.get('派工時間')), parse_int(r.get('派工時段')),
                parse_datetime(r.get('回應時間')), parse_datetime(r.get('準備時間')), parse_datetime(r.get('開始時間')), parse_datetime(r.get('手圈時間')), parse_datetime(r.get('結束時間')), parse_int(r.get('執行時間')),
                parse_text(r.get('派工人員')), parse_text(r.get('病人非病人')), parse_text(r.get('緊急等級')),
                parse_datetime(r.get('取消手動完成時間')), parse_text(r.get('取消手動完成操作者')), parse_text(r.get('取消手動完成原因')), parse_text(r.get('手動狀態')),
                parse_datetime(r.get('等待操作時間')), parse_text(r.get('等待操作者')), parse_text(r.get('等待原因')),
                parse_datetime(r.get('返回操作時間')), parse_text(r.get('返回操作者')), parse_text(r.get('返回原因')),
                parse_text(r.get('套餐母單號')), parse_text(r.get('花費時間')), parse_text(r['是否延遲']), parse_text(r.get('派工紀錄')), parse_text(r.get('藥袋')),
                parse_text(r['不需計算']), parse_text(r['排程需排除']), parse_text(r['不屬延遲']), parse_text(r['班別'])
            ))
                
        try:
            psycopg2.extras.execute_batch(c, INSERT_SQL, records_to_insert, page_size=1000)
            inserted = len(records_to_insert)
        except Exception as batch_ex:
            print(f"⚠️ [手動上傳] 批次寫入遭遇異常格式，啟動逐筆安全降級模式... ({str(batch_ex)})")
            conn.rollback()
            inserted = skipped = errors = 0
            for rec in records_to_insert:
                c.execute("SAVEPOINT upload_row_savepoint")
                try:
                    c.execute(INSERT_SQL, rec)
                    if c.rowcount > 0: inserted += 1
                    else: skipped += 1
                    c.execute("RELEASE SAVEPOINT upload_row_savepoint")
                except Exception as ex:
                    c.execute("ROLLBACK TO SAVEPOINT upload_row_savepoint")
                    errors += 1

        conn.commit()
        
        # 運行後續清洗規則
        results = {}
        for b in affected_batches:
            try:
                batch_res = execute_system_7_cleaning_rules(conn, b)
                for k, v in batch_res.items():
                    results[k] = results.get(k, 0) + v
            except Exception as clean_ex:
                print(f"⚠️ 批次 {b} 運行清洗規則失敗: {str(clean_ex)}")

        # === 手動上傳成功後，更新數據匯入報告卡片 ===
        try:
            current_time_str = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            UPSERT_STATUS_SQL = """
                INSERT INTO porter_system_status (key, value, updated_at)
                VALUES (%s, %s, CURRENT_TIMESTAMP)
                ON CONFLICT (key) 
                DO UPDATE SET value = EXCLUDED.value, updated_at = CURRENT_TIMESTAMP
            """
            # 手動上傳時，最新日期抓取受影響批次的第一個日期，若無則用當天
            display_batch = list(affected_batches)[0] if affected_batches else date.today().strftime("%Y-%m-%d")
            
            c.execute(UPSERT_STATUS_SQL, ('latest_sync_date', display_batch))
            c.execute(UPSERT_STATUS_SQL, ('sync_total', str(total_excel_rows)))
            c.execute(UPSERT_STATUS_SQL, ('sync_success', str(inserted)))
            c.execute(UPSERT_STATUS_SQL, ('sync_skipped', str(skipped)))
            c.execute(UPSERT_STATUS_SQL, ('sync_failed', str(errors)))
            print(f"📌 [系統狀態] 手動上傳 Excel 卡片數據已寫入資料庫。時間：{current_time_str}")
        except Exception as status_ex:
            print(f"❌ 手動上傳更新卡片數據失敗: {str(status_ex)}")
                
        conn.commit()
        conn.close()
        
        if os.path.exists(file_path): 
            os.remove(file_path) 
        
        summary_msg = f"Excel 原始 {total_excel_rows} 筆\n成功新增 {inserted} 筆\n已存在跳過 {skipped} 筆\n錯誤失敗 {errors} 筆"
        print(f"📋 [上傳完成報告] {summary_msg}")
        
        return jsonify({
            "success": True, 
            "imported": summary_msg,  
            "cleaned_details": results,
            "stats": {
                "total": total_excel_rows,
                "success": inserted,
                "skipped": skipped,
                "error": errors
            }
        })
    
@app.route("/api/report")
def report():
    df = request.args.get("date_from", "")
    dt = request.args.get("date_to", "")
    
    conditions = ["\"延遲調整\" IS NOT NULL"]
    params = []
    
    if df: 
        conditions.append("任務時間 >= %s")
        params.append(df + " 00:00:00")
    if dt: 
        conditions.append("任務時間 <= %s")
        params.append(dt + " 23:59:59")
        
    where_clause = " AND ".join(conditions)

    conn = get_db()
    c = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
    
    # 使用 SUBSTRING 將 TEXT 型態的任務時間切出日期 (YYYY-MM-DD) 來做 GROUP BY
    sql = f"""
        SELECT SUBSTRING(任務時間 FROM 1 FOR 10) AS 日期, COUNT(*) AS 總任務數,
               SUM(CASE WHEN "延遲調整"='延遲' THEN 1 ELSE 0 END) AS 修正後延遲數
        FROM task_records 
        WHERE {where_clause} AND 任務時間 IS NOT NULL AND 任務時間 != ''
        GROUP BY SUBSTRING(任務時間 FROM 1 FOR 10)
        ORDER BY 日期 DESC
    """
    try:
        c.execute(sql, params)
        rows = [dict(r) for r in c.fetchall()]
    except Exception as e:
        print(f"❌ SQL 執行失敗: {str(e)}")
        rows = []
    finally:
        conn.close()
    
    for r in rows:
        r["日期"] = str(r["日期"])
        tot = r["總任務數"]
        r["修正後延遲率"] = round((r["修正後延遲數"] / tot * 100), 2) if tot > 0 else 0
    return jsonify(rows)

@app.route("/api/check")
def check():
    conn = get_db()
    c = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
    
    c.execute("""
        SELECT *, "延遲調整" as is_delayed_adjusted, "不需計算" as exclude_reason
        FROM task_records 
        WHERE "延遲調整"='需檢查' 
        ORDER BY 任務時間 DESC
    """)
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    
    # 在 Python 端安全地處理日期切取，防止資料庫轉型崩潰
    for r in rows: 
        t_time = r.get("任務時間") or ""
        r["日期"] = t_time[:10] if len(t_time) >= 10 else "無日期"
        
    return jsonify(rows)

@app.route('/api/task_package_details', methods=['GET'])
def get_task_package_details():
    """ 撈取同單號或同病人同一天的完整套餐任務明細 """
    task_id = request.args.get('task_id', '')      # 單號
    patient_name = request.args.get('patient', '') # 病人姓名
    task_time = request.args.get('date', '')       # 任務時間 (取前10碼日期)
    
    if not task_id and not patient_name:
        return jsonify([])
        
    conn = get_db()
    c = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
    
    # 擷取日期 YYYY-MM-DD
    date_str = task_time[:10] if len(task_time) >= 10 else ""
    
    # SQL 邏輯：同單號，或者是（同病人姓名且同一天）
    sql = """
        SELECT id, 單號, 任務, 傳送人員, 任務時間, 
               "是否延遲", "延遲調整", "不需計算", "備註-時間描述", "備註-開始結束說明"
        FROM task_records
        WHERE 單號 = %s 
           OR (病人姓名 = %s AND SUBSTRING(任務時間 FROM 1 FOR 10) = %s)
        ORDER BY 任務時間 ASC
    """
    
    try:
        c.execute(sql, (task_id, patient_name, date_str))
        rows = [dict(r) for r in c.fetchall()]
    except Exception as e:
        print(f"❌ 撈取套餐明細失敗: {str(e)}")
        rows = []
    finally:
        conn.close()
        
    return jsonify(rows)

@app.route("/api/update", methods=["POST"])
def update_record():
    """ 人工確認：更新覆寫最新的『延遲調整』覆核狀態 """
    data = request.json; rid, status = data.get("id"), data.get("is_delayed_adjusted")
    conn = get_db(); c = conn.cursor()
    c.execute("""
        UPDATE task_records 
        SET "延遲調整"=%s, "不需計算"=%s 
        WHERE id=%s
    """, (status, f"人工覆核判定：{status}", rid))
    conn.commit(); conn.close()
    return jsonify({"success": True})

@app.route("/api/search_options")
def search_options():
    conn = get_db(); c = conn.cursor()
    c.execute("SELECT DISTINCT 任務 FROM task_records WHERE 任務 IS NOT NULL AND 任務 != '' ORDER BY 任務 ASC")
    tasks = [r[0] for r in c.fetchall()]
    c.execute("SELECT DISTINCT 傳送人員 FROM task_records WHERE 傳送人員 IS NOT NULL AND 傳送人員 != '' ORDER BY 傳送人員 ASC")
    persons = [r[0] for r in c.fetchall()]
    conn.close()
    return jsonify({"tasks": tasks, "persons": persons})

@app.route("/api/search")
def search():
    """🔍 查詢資料分頁數據"""
    df, dt, tk, ps, st = request.args.get("date_from",""), request.args.get("date_to",""), request.args.get("task",""), request.args.get("person",""), request.args.get("status","")
    page = int(request.args.get("page", 1))
    per_page = int(request.args.get("per_page", 15))
    
    conditions = ["1=1"]
    params = []
    if df: conditions.append("任務時間 >= %s"); params.append(df + " 00:00:00")
    if dt: conditions.append("任務時間 <= %s"); params.append(dt + " 23:59:59")
    if tk: conditions.append("任務 LIKE %s"); params.append(f"%{tk}%")
    if ps: conditions.append("傳送人員 LIKE %s"); params.append(f"%{ps}%")
    if st: conditions.append("\"延遲調整\" = %s"); params.append(st)
    
    where = " AND ".join(conditions)
    conn = get_db(); c = conn.cursor()
    c.execute(f"SELECT COUNT(*) FROM task_records WHERE {where}", params)
    total = c.fetchone()[0]
    offset = (page - 1) * per_page
    
    c.execute(f"""
        SELECT id, 單號, 任務時間::
        date AS 日期, 任務, 傳送人員, 派工單位, 特別指示, 是否延遲, "延遲調整" as is_delayed_adjusted, "不需計算" as exclude_reason
        FROM task_records 
        WHERE {where} 
        ORDER BY 任務時間 
        DESC LIMIT %s
        OFFSET %s""", 
        params + [per_page, offset])
    colnames = [desc[0] for desc in c.description]
    rows = [dict(zip(colnames, r)) for r in c.fetchall()]; conn.close()
    for r in rows: r["日期"] = str(r["日期"])
    return jsonify({"total": total, "page": page, "per_page": per_page, "rows": rows})

@app.route("/api/stats")
def get_stats():
    conn = get_db()
    c = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
    
    # 建立一個安全的預設結構字典，防止前端抓不到資料而報錯
    status_data = {
        "latest_update": "無資料",
        "excel_total": 0,
        "success_count": 0,
        "skip_count": 0,
        "error_count": 0,
        "updated_at": "-"
    }
    
    try:
        # 1. 配合您資料庫的 key-value 設計，將所有統計狀態一次撈出
        c.execute("""
            SELECT key, value, to_char(updated_at, 'YYYY/MM/DD HH24:MI:SS') as t_str 
            FROM porter_system_status 
            WHERE key IN ('latest_sync_date', 'sync_total', 'sync_success', 'sync_skipped', 'sync_failed')
        """)
        rows = c.fetchall()
        
        # 將資料庫一筆一筆的 key-value 拆解對齊並塞入預設字典
        for row in rows:
            k = row['key']
            v = row['value']
            t = row['t_str']
            
            if k == 'latest_sync_date':
                status_data["latest_update"] = v or "無資料"
            elif k == 'sync_total':
                status_data["excel_total"] = int(v) if (v and v.isdigit()) else 0
            elif k == 'sync_success':
                status_data["success_count"] = int(v) if (v and v.isdigit()) else 0
            elif k == 'sync_skipped':
                status_data["skip_count"] = int(v) if (v and v.isdigit()) else 0
            elif k == 'sync_failed':
                status_data["error_count"] = int(v) if (v and v.isdigit()) else 0
            
            # 只要有更新時間就進行保留，記錄最後一筆更新時間
            if t:
                status_data["updated_at"] = t

    except Exception as e:
        print(f"⚠️ [警告] 撈取數據匯入報告狀態時出錯（可能資料表尚未建立或欄位有衝突）: {str(e)}")
        # 出錯時不中斷，維持預設 0 筆回傳，防止前端崩潰爆出 500
        pass

    # 2. 計算側邊欄紅色數位貼紙（尚未審核的人工確認項目筆數）
    # 使用符合您原本資料庫實際使用的中文名稱「延遲調整」進行篩選，加入安全防禦機制
    badge_count = 0
    try:
        c.execute("""
            SELECT COUNT(*) FROM task_records 
            WHERE "延遲調整" IS NULL 
              AND 是否延遲 = '是' 
              AND 不需計算 = '' 
              AND 排程需排除 = '' 
              AND 不屬延遲 = ''
        """)
        badge_count = c.fetchone()[0]
    except Exception as badge_e:
        print(f"⚠️ [警告] 計算未審核紅點筆數時出錯: {str(badge_e)}")
        # 若資料表名稱打架則 fallback 為 0 筆，保障全網頁暢通
        try:
            conn.rollback() # 發生異常時回滾事物，避免連線死鎖
            c.execute("""
                SELECT COUNT(*) FROM task_records 
                WHERE is_delayed_adjusted IS NULL 
                  AND 是否延遲 = '是' 
                  AND 不需計算 = '' 
                  AND 排程需排除 = '' 
                  AND 不屬延遲 = ''
            """)
            badge_count = c.fetchone()[0]
        except:
            badge_count = 0

    conn.close()
    
    # 最終打包成 JSON 格式吐給前端 main.js
    return jsonify({
        "success": True,
        "status": status_data,
        "badge_count": badge_count
    })

@app.route("/api/export/report")
def export_report():
    """📥 下載 Excel 統計報表"""
    df = request.args.get("date_from", "")
    dt = request.args.get("date_to", "")
    
    conditions = ["\"延遲調整\" IS NOT NULL"]
    params = []
    if df: 
        conditions.append("任務時間 >= %s")
        params.append(df + " 00:00:00")
    if dt: 
        conditions.append("任務時間 <= %s")
        params.append(dt + " 23:59:59")
        
    where_clause = " AND ".join(conditions)

    conn = get_db()
    c = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
    
    sql = f"""
        SELECT SUBSTRING(任務時間 FROM 1 FOR 10) AS 日期, COUNT(*) AS 總任務數, 
               SUM(CASE WHEN "延遲調整"='延遲' THEN 1 ELSE 0 END) AS 修正後延遲數 
        FROM task_records 
        WHERE {where_clause} AND 任務時間 IS NOT NULL AND 任務時間 != ''
        GROUP BY SUBSTRING(任務時間 FROM 1 FOR 10) 
        ORDER BY 日期 DESC
    """
    try:
        c.execute(sql, params)
        rows = [dict(r) for r in c.fetchall()]
    except Exception as e:
        print(f"❌ 匯出 SQL 執行失敗: {str(e)}")
        rows = []
    finally:
        conn.close()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "統計報表"
    ws.append(["日期", "總任務數", "修正後延遲數", "修正後延遲率"])
    
    for r in rows:
        tot = r["總任務數"]
        ad_r = round((r["修正後延遲數"]/tot*100), 2) if tot > 0 else 0
        ws.append([str(r["日期"]), tot, r["修正後延遲數"], f"{ad_r}%"])
        
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    
    filename = "中榮傳送改善統計報表"
    if df or dt: 
        filename += f"_{df.replace('-', '')}-{dt.replace('-', '')}"
    filename += ".xlsx"
    
    return send_file(stream, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name=filename)

# ============================================================
# ⚙️ 規則管理 API 接口區
# ============================================================

@app.route("/api/rules/<rule_type>", methods=["GET"])
def get_rules(rule_type):
    """取得特定類型的規則清單"""
    table_map = {
        "wait_reasons": "rule_wait_reasons ORDER BY id ASC",
        "keywords": "rule_keywords ORDER BY id ASC",
        "schedule_excludes": "rule_schedule_excludes ORDER BY id ASC"
    }
    if rule_type not in table_map:
        return jsonify({"success": False, "message": "未知的規則類型"})
        
    conn = get_db()
    c = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
    c.execute(f"SELECT * FROM {table_map[rule_type]}")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify(rows)

@app.route("/api/rules/<rule_type>", methods=["POST"])
def add_rule(rule_type):
    """新增規則"""
    data = request.json or {}
    conn = get_db()
    c = conn.cursor()
    try:
        if rule_type == "wait_reasons":
            c.execute("""INSERT INTO rule_wait_reasons (reason, action, note) VALUES (%s, %s, %s)""",
                      (data.get("reason"), data.get("action", "未延遲"), data.get("note")))
        elif rule_type == "keywords":
            c.execute("""INSERT INTO rule_keywords (target_field, keyword, match_type, action, note) VALUES (%s, %s, %s, %s, %s)""",
                      (data.get("target_field"), data.get("keyword"), data.get("match_type", "contains"), data.get("action", "需檢查"), data.get("note")))
        elif rule_type == "schedule_excludes":
            c.execute("""INSERT INTO rule_schedule_excludes (end_location, expected_time, label) VALUES (%s, %s, %s)""",
                      (data.get("end_location"), data.get("expected_time"), data.get("label")))
        else:
            return jsonify({"success": False, "message": "未知的規則類型"})
        conn.commit()
        return jsonify({"success": True, "message": "規則新增成功"})
    except Exception as e:
        conn.rollback()
        return jsonify({"success": False, "message": f"新增失敗，可能資料重複。原因: {str(e)}"})
    finally:
        conn.close()

@app.route("/api/rules/<rule_type>/<int:rule_id>", methods=["PUT"])
def update_rule_status(rule_type, rule_id):
    """更新規則的啟用狀態 (True/False)"""
    table_map = {
        "wait_reasons": "rule_wait_reasons",
        "keywords": "rule_keywords",
        "schedule_excludes": "rule_schedule_excludes"
    }
    if rule_type not in table_map:
        return jsonify({"success": False, "message": "未知的規則類型"})
        
    data = request.json or {}
    enabled = data.get("enabled")
    
    conn = get_db()
    c = conn.cursor()
    c.execute(f"UPDATE {table_map[rule_type]} SET enabled = %s WHERE id = %s", (enabled, rule_id))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "規則狀態已更新"})

@app.route("/api/rules/<rule_type>/<int:rule_id>", methods=["DELETE"])
def delete_rule(rule_type, rule_id):
    """刪除規則 (備用)"""
    table_map = {
        "wait_reasons": "rule_wait_reasons",
        "keywords": "rule_keywords",
        "schedule_excludes": "rule_schedule_excludes"
    }
    if rule_type not in table_map:
        return jsonify({"success": False, "message": "未知的規則類型"})
        
    conn = get_db()
    c = conn.cursor()
    c.execute(f"DELETE FROM {table_map[rule_type]} WHERE id = %s", (rule_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "規則已刪除"})


if __name__ == "__main__":
    # 使用 Flask 上下文確保資料庫與排程安全初始化
    with app.app_context():
        init_db()
    
    app.run(
        debug=True, 
        port=5000, 
        host="0.0.0.0",
        # 💡 核心修正：關閉 reloader，防止 Flask 在開發模式下啟動兩個行程，造成排程重複執行
        use_reloader=False,
        extra_files=[
            './templates/index.html',
            './static/style.css',
            './static/main.js'
        ]
    )