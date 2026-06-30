/* ==========================================================================
   1. 網頁初始化與全域狀態管理 (Initialization & Globals)
   ========================================================================== */
let currentPage = 1;      // 紀錄目前查詢資料分頁的位置
let reportChart = null;     // 儲存 Chart.js 實例
let selectedFile = null;    // 紀錄手動上傳檔案
let currentCheckList = [];  // 暫存目前整份需確認項目的陣列

document.addEventListener("DOMContentLoaded", async function() {
    // 🎯 優先載入狀態與統計卡片
    try { await loadStats(); } catch(e) { console.error("Stats 載入失敗:", e); }
    try { loadSearchOptions(); } catch(e) { console.error("SearchOptions 載入失敗:", e); }
    try { initSyncDates(); } catch(e) { console.error("SyncDates 載入失敗:", e); }
    try { initReportDates(); } catch(e) { console.error("ReportDates 載入失敗:", e); }
    try { await loadReport(); } catch(e) { console.error("Report 載入失敗:", e); }
    try { await loadCheck(); } catch(e) { console.error("Check 載入失敗:", e); }

    // 綁定篩選佈局中的「🔍 開始查詢」按鈕 (加上阻止預設提交防閃退)
    const btnSearch = document.getElementById("btnSearch");
    if (btnSearch) {
        btnSearch.addEventListener("click", function(e) {
            e.preventDefault(); 
            doSearch(1);    
        });
    }

    // 強固化首頁引導：安全切換至資料管理
    showPage('import');
});

function initReportDates() {
    const reportFrom = document.getElementById('reportDateFrom');
    const reportTo = document.getElementById('reportDateTo');
    if (reportFrom && reportTo && (!reportFrom.value || !reportTo.value)) {
        const today = new Date();
        const past30 = new Date();
        past30.setDate(today.getDate() - 30);
        const formatDate = (d) => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
        reportFrom.value = formatDate(past30);
        reportTo.value = formatDate(today);
    }
}

function initSyncDates() {
    const syncFrom = document.getElementById('syncDateFrom');
    const syncTo = document.getElementById('syncDateTo');
    if (syncFrom && syncTo) {
        const yesterday = new Date();
        yesterday.setDate(yesterday.getDate() - 1); 
        const yesterdayStr = `${yesterday.getFullYear()}-${String(yesterday.getMonth() + 1).padStart(2, '0')}-${String(yesterday.getDate()).padStart(2, '0')}`; 
        syncFrom.value = yesterdayStr;
        syncTo.value = yesterdayStr;
    }
}

/* ==========================================================================
   2. 頁面控制與 UI 提示組件 (Routing & Base UI)
   ========================================================================== */
function showPage(pageId) {
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
    
    const targetPage = document.getElementById(`page-${pageId}`);
    if (targetPage) targetPage.classList.add('active');

    if (window.event && window.event.currentTarget) {
        window.event.currentTarget.classList.add('active');
    } else {
        const defaultBtn = document.querySelector(`button[onclick*="showPage('${pageId}')"]`);
        if (defaultBtn) defaultBtn.classList.add('active');
    }
    
    if (pageId === 'rules') {
        if (typeof switchRuleTab === 'function') switchRuleTab('wait_reasons'); 
    }
}

function showToast(msg, type='success') {
    const t = document.getElementById('toast');
    if(!t) return;
    t.textContent = msg;
    t.className = 'toast show ' + type;
    setTimeout(() => t.className = 'toast', 3000);
}

/* ==========================================================================
   2.5 🎯 核心封裝：不論來源，100% 統一渲染大標題與指標卡片
   ========================================================================== */
function renderStatsPanel(stats, status) {
    // 1. 處理大標題旁的「🍏 系統最新數據已更新至」外框
    const alertBox = document.getElementById('update-status-alert');
    const dateText = document.getElementById('latest-date-text');
    
    if (alertBox && dateText && status) {
        if (status.latest_batch && status.latest_batch !== "尚無資料" && status.latest_batch !== "-") {
            alertBox.style.display = 'inline-block';
            let displayDate = status.latest_batch.replace(/-/g, '/');
            let shortTime = "";
            if (status.updated_at && status.updated_at.length >= 16) {
                shortTime = status.updated_at.substring(11, 19);
            }
            // 完美符合你想要的更新時間提示字樣
            dateText.textContent = `🍏 系統最新數據已更新至：${displayDate} (自動定時更新於 ${shortTime})`;
        } else {
            alertBox.style.display = 'none';
        }
    }

    // 2. 處理四張指標數據膠囊卡片
    const statsCard = document.getElementById('upload-stats-card');
    if (statsCard && stats) {
        // 兼容後端多欄位命名 (統計、排除、失敗)
        const total = parseInt(stats.total || 0);
        const success = parseInt(stats.success || 0);
        const skipped = parseInt(stats.skipped || stats.skip || 0);
        const failed = parseInt(stats.failed || stats.error || 0);

        if (total > 0 || success > 0) {
            statsCard.style.display = 'block';
            document.getElementById('stat-total').textContent = total.toLocaleString();
            document.getElementById('stat-success').textContent = success.toLocaleString();
            document.getElementById('stat-skipped').textContent = skipped.toLocaleString();
            document.getElementById('stat-error').textContent = failed.toLocaleString();
        }
    }
}

/* ==========================================================================
   3. 資料管理頁面控制：遠端 RPA 同步與備用手動上傳 (Data Management)
   ========================================================================== */
/**
 * 情境 A：[定時自動更新] 或網頁初始化載入最新數據
 */
async function loadStats() {
    try {
        const res = await fetch('/api/stats');
        if (!res.ok) return;
        const data = await res.json();
        
        // 1. 更新側邊欄需確認項目的紅色數位貼紙
        const badge = document.getElementById('badge-count');
        if (badge) {
            if (data.badge_count > 0) {
                badge.textContent = data.badge_count;
                badge.style.display = 'inline-block';
            } else {
                badge.style.display = 'none';
            }
        }
        
        // 2. 渲染右上角最新數據更新提示與匯入報告卡片
        if (data.status) {
            const st = data.status;
            
            // 填入右上角狀態提示
            const lblLatestDate = document.getElementById('lblLatestDate');
            const lblUpdateTime = document.getElementById('lblUpdateTime');
            if (lblLatestDate) lblLatestDate.textContent = st.latest_update;
            if (lblUpdateTime) lblUpdateTime.textContent = "自動定時更新於 " + st.updated_at.split(' ')[1] || st.updated_at;
            
            // 填入最近一次數據匯入報告的標籤膠囊數字
            const lblReportDetailTime = document.getElementById('lblReportDetailTime');
            if (lblReportDetailTime) lblReportDetailTime.textContent = `(更新完成時間：${st.updated_at})`;
            
            if (document.getElementById('valExcelTotal')) {
                document.getElementById('valExcelTotal').textContent = Number(st.excel_total).toLocaleString();
                document.getElementById('valSuccessCount').textContent = Number(st.success_count).toLocaleString();
                document.getElementById('valSkipCount').textContent = Number(st.skip_count).toLocaleString();
                document.getElementById('valErrorCount').textContent = Number(st.error_count).toLocaleString();
            }
        }
    } catch (e) {
        console.error("載入統計狀態失敗:", e);
    }
}

/**
 * 情境 B：[RPA 線上同步] 一鍵觸發指定區間更新
 */
async function doAutoSync() {
    const rawFrom = document.getElementById('syncDateFrom').value;
    const rawTo = document.getElementById('syncDateTo').value;
    
    if (!rawFrom || !rawTo) {
        showToast('請先選取日期區間！', 'error');
        return;
    }

    const dateFrom = rawFrom.replace(/-/g, '/');
    const dateTo = rawTo.replace(/-/g, '/');
    const btn = document.getElementById('btn-auto-sync');
    
    if (btn) { btn.disabled = true; btn.textContent = '⏳ 同步中...'; }
    
    try {
        const res = await fetch('/api/auto_sync', { 
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ date_from: dateFrom, date_to: dateTo })
        });
        const d = await res.json();
        if (d.success) {
            showToast('更新完成！', 'success');

            // 🎯 使用統一渲染：傳入更新後的 stats 與 status
            await loadStats();
            renderStatsPanel(d.stats, d.status || d.stats); 
        } else {
            showToast(d.message, 'error');
            alert(`❌ 更新失敗：\n${d.message}`);
        }
    } catch (e) {
        alert(`❌ 伺服器連線錯誤，無法完成更新！\n原因: ${e.message}`);
    }
    if (btn) { btn.disabled = false; btn.textContent = '🔄 更新指定區間數據'; }
}

/* 📁 手動上傳區事件 */
const uploadZone = document.getElementById('uploadZone');
if (uploadZone) {
    uploadZone.addEventListener('dragover', e => { e.preventDefault(); uploadZone.classList.add('dragover'); });
    uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('dragover'));
    uploadZone.addEventListener('drop', e => {
        e.preventDefault(); uploadZone.classList.remove('dragover');
        const f = e.dataTransfer.files[0];
        if (f) setFile(f);
    });
}

const fileInput = document.getElementById('excelFile');
if (fileInput) {
    fileInput.addEventListener('change', e => {
        if (e.target.files[0]) setFile(e.target.files[0]);
    });
}

function setFile(f) {
    selectedFile = f;
    if (uploadZone) {
        uploadZone.innerHTML = `<div class="upload-icon">📄</div><p style="color:var(--green)">${f.name}</p><small>${(f.size/1024/1024).toFixed(2)} MB</small>`;
    }
    const btnUpload = document.getElementById('btn-upload');
    if (btnUpload) btnUpload.disabled = false;
}

/**
 * 情境 C：[手動上傳 Excel] 備用檔案匯入
 */
async function doUpload() {
    if (!selectedFile) return;
    const btn = document.getElementById('btn-upload');
    if (btn) { btn.disabled = true; btn.textContent = '上傳中...'; }

    const form = new FormData();
    form.append('file', selectedFile);
    try {
        const res = await fetch('/api/upload', { method: 'POST', body: form });
        const d = await res.json();
        if (d.success) {
            showToast(`匯入成功！`, 'success');

            loadSearchOptions(); 
            await loadStats();
            // 🎯 使用統一渲染：傳入手動上傳成功後的 stats 物件
            renderStatsPanel(d.stats, d.status || d.stats);
        } else {
            showToast(d.message, 'error');
            alert(`❌ 匯入失敗：\n${d.message}`);
        }
    } catch(e) {
        alert(`❌ 檔案上傳期間發生預期外連線錯誤！\n原因: ${e.message}`);
    }
    if (btn) { btn.disabled = false; btn.textContent = '⬆️ 上傳並匯入'; }
}

/* ==========================================================================
   4. 延遲率報表功能區：趨勢圖渲染與 Excel 下載 (Reports & Charts)
   ========================================================================== */
async function loadReport() {
    const tbody = document.getElementById('report-table');
    if (!tbody) return; 
    
    let dateFrom = document.getElementById('reportDateFrom').value;
    let dateTo = document.getElementById('reportDateTo').value;
    
    if (!dateFrom || !dateTo) {
        initReportDates();
        dateFrom = document.getElementById('reportDateFrom').value;
        dateTo = document.getElementById('reportDateTo').value;
    }
    
    const params = new URLSearchParams();
    if (dateFrom) params.append('date_from', dateFrom);
    if (dateTo) params.append('date_to', dateTo);

    try {
        const res = await fetch('/api/report?' + params);
        const data = await res.json();
        
        if (!data || !data.length) {
            tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:var(--text3);padding:30px">該區間尚無報表資料，請先更新數據 </td></tr>';
            if (reportChart) { reportChart.destroy(); reportChart = null; }
            return;
        }
        
        tbody.innerHTML = data.map(r => `
            <tr>
                <td class="mono">${r.日期}</td>
                <td>${r.總任務數}</td>
                <td><span class="badge badge-yellow">${r.修正後延遲數}</span></td>
                <td><strong style="color:var(--green)">${r.修正後延遲率}%</strong></td>
            </tr>`
        ).join('');

        const canvas = document.getElementById('reportChart');
        if (!canvas) return;
        
        const labels = data.map(r => r.日期 || '').reverse();
        const adj = data.map(r => r.修正後延遲率 || 0).reverse();
        
        if (reportChart) { reportChart.destroy(); }
        
        try {
            reportChart = new Chart(canvas, {
                type: 'line',
                data: {
                    labels: labels,
                    datasets: [{ 
                        label: '修正後延遲率', 
                        data: adj, 
                        borderColor: '#34d399', 
                        backgroundColor: 'rgba(52,211,153,0.1)', 
                        tension: 0.3, 
                        fill: true,
                        pointBackgroundColor: '#34d399',
                        pointRadius: 4
                    }]
                },
                options: {
                    responsive: true, 
                    maintainAspectRatio: false,
                    plugins: { legend: { labels: { color: '#94a3b8', font: { size: 12, family: 'Noto Sans TC' } } } },
                    scales: {
                        x: { ticks: { color: '#64748b', font: { family: 'JetBrains Mono' } }, grid: { color: '#1e2235' } },
                        y: { ticks: { color: '#64748b', callback: v => v + '%' }, grid: { color: '#1e2235' }, min: 0 }
                    }
                }
            });
        } catch (chartErr) {
            console.error("❌ Chart.js 渲染引擎發生異常:", chartErr);
        }
    } catch (err) {
        console.error("❌ 報表 API 連線或資料解析失敗:", err);
    }
}

function exportReportWithDate() {
    const dateFrom = document.getElementById('reportDateFrom').value;
    const dateTo = document.getElementById('reportDateTo').value;
    const params = new URLSearchParams();
    if (dateFrom) params.append('date_from', dateFrom);
    if (dateTo) params.append('date_to', dateTo);
    window.location.href = '/api/export/report?' + params;
}

/* ==========================================================================
   5. 需確認項目功能區：人工二次覆核彈窗 (Manual Review & Modal)
   ========================================================================== */
async function loadCheck() {
    const res = await fetch('/api/check');
    const data = await res.json();
    currentCheckList = data; 
    const tbody = document.getElementById('check-table');
    if (!tbody) return;
    if (!data.length) {
        tbody.innerHTML = '<tr><td colspan="7"><div class="empty-state"><div class="icon">✅</div><p>目前沒有需要人工確認的項目</p></div></td></tr>';
        return;
    }
    
    tbody.innerHTML = data.map((r, index) => {
        return `
        <tr>
            <td class="mono" style="font-size:11px">${r.單號}</td>
            <td class="mono">${r.日期}</td>
            <td>${r.任務 || '—'}</td>
            <td>${r.傳送人員 || '—'}</td>
            <td><div style="font-size:12px;max-width:260px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${r.特別指示||''}">${r.特別指示 || '—'}</div></td>
            <td><span class="badge badge-yellow" style="font-size:10px">${r.exclude_reason||'原清洗規則保留審核'}</span></td>
            <td><button class="btn btn-primary btn-sm" onclick="openEdit(${index})">確認</button></td>
        </tr>`;
    }).join('');
}

function openEdit(index) {
    if (index < 0 || index >= currentCheckList.length) return;
    const r = currentCheckList[index];

    document.getElementById('edit-id').value = r.id;
    document.getElementById('edit-special').value = r.特別指示 || '—';
    document.getElementById('edit-note-time').value = '';
    document.getElementById('edit-note-open-close').value = '';
    document.getElementById('edit-status').value = '未延遲'; 

    const modalHeader = document.querySelector('#editModal .modal h3');
    if (modalHeader) {
        let navBar = document.getElementById('modal-nav-bar');
        if (!navBar) {
            navBar = document.createElement('div');
            navBar.id = 'modal-nav-bar';
            navBar.style = 'display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; background:var(--surface2); padding:8px 12px; border-radius:8px; font-size:13px; color:var(--text2);';
            modalHeader.parentNode.insertBefore(navBar, modalHeader.nextSibling);
        }
        
        const hasPrev = index > 0;
        const hasNext = index < currentCheckList.length - 1;
        
        navBar.innerHTML = `
            <button class="btn btn-ghost btn-sm" ${hasPrev ? '' : 'disabled'} onclick="openEdit(${index - 1})" style="padding: 4px 10px;">◀ 上一筆</button>
            <span style="font-family:'JetBrains Mono'; font-weight:600; color:var(--accent)">第 ${index + 1} / ${currentCheckList.length} 筆 (單號: ${r.單號})</span>
            <button class="btn btn-ghost btn-sm" ${hasNext ? '' : 'disabled'} onclick="openEdit(${index + 1})" style="padding: 4px 10px;">下一筆 ▶</button>
        `;
    }

    const detailsContainer = document.getElementById('edit-full-details');
    if (detailsContainer) {
        const excludeKeys = ['id', 'import_batch', 'is_delayed_adjusted', 'exclude_reason', '日期'];
        let htmlContent = '<div style="display:grid; grid-template-columns: 110px 1fr; gap: 6px 10px; border-top: 1px solid var(--border); padding-top:10px;">';
        for (const [key, value] of Object.entries(r)) {
            if (!excludeKeys.includes(key) && value !== null && value !== undefined && String(value).trim() !== '') {
                htmlContent += `<div style="color: var(--text3); font-weight: 600; text-align: right;">${key}：</div><div style="word-break: break-all; color: var(--text); font-family: sans-serif;">${value}</div>`;
            }
        }
        htmlContent += '</div>';
        detailsContainer.innerHTML = htmlContent;
    }

    const pkgTbody = document.getElementById('package-table-body');
    if (pkgTbody) {
        pkgTbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:15px; color:var(--text3);">⏳ 正在深度檢索套餐關聯資料...</td></tr>';
        fetch(`/api/task_package_details?task_id=${encodeURIComponent(r.單號 || '')}&patient=${encodeURIComponent(r.病人姓名 || '')}&date=${encodeURIComponent(r.任務時間 || '')}`)
            .then(res => res.json())
            .then(packageTasks => {
                if (!packageTasks || packageTasks.length === 0) {
                    pkgTbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:15px; color:var(--text3);">無其他關聯套餐任務</td></tr>';
                    return;
                }
                let html = '';
                packageTasks.forEach(t => {
                    const isCurrent = (t.id == r.id) ? 'background: rgba(79, 142, 247, 0.15); font-weight: bold; color: var(--accent);' : '';
                    const oriStatus = t.是否延遲 === '是' ? `<span style="color:var(--red)">延遲</span>` : '正常';
                    let adjStatus = t.延延迟調整 || '';
                    if (adjStatus === '延遲') adjStatus = `<span style="color:var(--red)">延遲</span>`;
                    if (adjStatus === '未延遲') adjStatus = `<span style="color:var(--green)">未延遲</span>`;
                    if (t.不需計算 === '是') adjStatus = `<span style="color:var(--text3); text-decoration: line-through;">不需計算</span>`;
                    if (!adjStatus) adjStatus = `<span style="color:var(--yellow)">待審核</span>`;
                    const desc = [t['備註-時間描述'], t['備註-開始結束說明']].filter(Boolean).join(' | ') || '(無)';
                    
                    html += `
                        <tr style="border-bottom: 1px solid var(--border); ${isCurrent}">
                            <td style="padding: 8px 6px;">${t.任務 || ''}</td>
                            <td style="padding: 8px 6px;">${t.傳送人員 || ''}</td>
                            <td style="padding: 8px 6px; font-family: monospace;">${(t.任務時間 || '').substring(11, 16)}</td>
                            <td style="padding: 8px 6px;">${oriStatus}</td>
                            <td style="padding: 8px 6px;">${adjStatus}</td>
                            <td style="padding: 8px 6px; color: var(--text3); max-width: 180px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${desc}">${desc}</td>
                        </tr>`;
                });
                pkgTbody.innerHTML = html;
            })
            .catch(err => {
                console.error("撈取套餐任務失敗:", err);
                pkgTbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:15px; color:var(--red);">❌ 關聯資料載入失敗</td></tr>';
            });
    }
    document.getElementById('editModal').classList.add('open');
}

function closeModal() {
    document.getElementById('editModal').classList.remove('open');
}

async function saveEdit() {
    const id = document.getElementById('edit-id').value;
    const status = document.getElementById('edit-status').value;
    const currentIndex = currentCheckList.findIndex(item => item.id == id);
    
    const res = await fetch('/api/update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id, is_delayed_adjusted: status })
    });
    const d = await res.json();
    if (d.success) {
        showToast('已儲存修正結果');
        await loadCheck(); 
        await loadStats();         
        loadSearchOptions(); 

        if (currentIndex !== -1 && currentIndex < currentCheckList.length) {
            openEdit(currentIndex);
        } else {
            closeModal();
        }
    }
}

/* ==========================================================================
   6. 查詢資料功能區：多條件組合篩選與資料分頁控制 (Advanced Query & Search)
   ========================================================================== */
function loadSearchOptions() {
    fetch('/api/search_options')
        .then(res => res.json())
        .then(data => {
            const taskSelect = document.getElementById('searchTask');
            const personSelect = document.getElementById('searchPerson');
            if (!taskSelect || !personSelect) return;
            taskSelect.innerHTML = '<option value="">全部</option>';
            data.tasks.forEach(task => { taskSelect.innerHTML += `<option value="${task}">${task}</option>`; });
            personSelect.innerHTML = '<option value="">全部</option>';
            data.persons.forEach(person => { personSelect.innerHTML += `<option value="${person}">${person}</option>`; });
        })
        .catch(err => console.error("載入篩選選單失敗:", err));
}

async function doSearch(page) {
    currentPage = page;
    const perPageEl = document.getElementById('searchPerPage');
    const perPage = perPageEl ? perPageEl.value : 15;
    
    const params = new URLSearchParams({
        date_from: document.getElementById('searchDateFrom').value,
        date_to: document.getElementById('searchDateTo').value,
        task: document.getElementById('searchTask').value, 
        person: document.getElementById('searchPerson').value,
        status: document.getElementById('searchStatus').value,
        per_page: perPage, 
        page: page,
    });
    
    const res = await fetch('/api/search?' + params);
    const d = await res.json();
    
    const searchInfo = document.getElementById('search-info');
    if (searchInfo) searchInfo.textContent = `共 ${d.total.toLocaleString()} 筆資料`;
    
    const tbody = document.getElementById('search-table');
    if (!tbody) return;
    if (!d.rows.length) {
        tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--text3);padding:30px">查無資料</td></tr>';
        const pagCtrl = document.getElementById('pagination');
        if (pagCtrl) pagCtrl.innerHTML = '';
        return;
    }
    
    const statusBadge = s => {
        if (s === '延遲') return '<span class="badge badge-red">延遲</span>';
        if (s === '未延遲') return '<span class="badge badge-green">未延遲</span>';
        if (s === '需檢查') return '<span class="badge badge-yellow">需檢查</span>';
        return s || '—';
    };
    
    tbody.innerHTML = d.rows.map(r => `
        <tr>
            <td class="mono" style="font-size:11px">${r.單號}</td>
            <td class="mono">${r.日期}</td>
            <td>${r.任務||'—'}</td>
            <td>${r.傳送人員||'—'}</td>
            <td style="font-size:12px">${r.派工單位||'—'}</td>
            <td>${statusBadge(r.是否延遲)}</td>
            <td>${statusBadge(r.is_delayed_adjusted)}</td>
            <td style="font-size:11px;color:var(--text3)">${r.exclude_reason||'—'}</td>
        </tr>`).join('');

    const totalPages = Math.ceil(d.total / d.per_page);
    let pageSelectHtml = `<select onchange="doSearch(parseInt(this.value))" style="padding: 4px 8px; background: var(--surface); border: 1px solid var(--border); color: var(--accent); border-radius: 6px; height: 30px; font-family: 'JetBrains Mono', monospace; font-weight: 600; cursor: pointer; margin: 0 4px;">`;
    let startPage = Math.max(1, page - 50);
    let endPage = Math.min(totalPages, page + 50);
    
    if (startPage > 1) pageSelectHtml += `<option value="1">1 ...</option>`;
    for (let i = startPage; i <= endPage; i++) {
        pageSelectHtml += `<option value="${i}" ${i === page ? 'selected' : ''}>${i}</option>`;
    }
    if (endPage < totalPages) pageSelectHtml += `<option value="${totalPages}">... ${totalPages}</option>`;
    pageSelectHtml += `</select>`;

    let pag = '';
    if (page > 1) pag += `<button class="page-btn" onclick="doSearch(${page-1})">‹ 上一頁</button>`;
    pag += `<span class="page-info">${pageSelectHtml} / ${totalPages} 頁</span>`;
    if (page < totalPages) pag += `<button class="page-btn" onclick="doSearch(${page+1})">下一頁 ›</button>`;
    
    const paginationCtrl = document.getElementById('pagination');
    if (paginationCtrl) paginationCtrl.innerHTML = pag;
}

function clearSearch() {
    document.getElementById('searchDateFrom').value = '';
    document.getElementById('searchDateTo').value = '';
    document.getElementById('searchTask').value = '';   
    document.getElementById('searchPerson').value = ''; 
    document.getElementById('searchStatus').value = '';
    document.getElementById('searchPerPage').value = '15'; 
    
    const tbody = document.getElementById('search-table');
    if (tbody) tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:var(--text3);padding:30px">請輸入條件查詢</td></tr>';
    
    const searchInfo = document.getElementById('search-info');
    if (searchInfo) searchInfo.textContent = '';
    
    const paginationCtrl = document.getElementById('pagination');
    if (paginationCtrl) paginationCtrl.innerHTML = '';
}

/* ==========================================================================
   7. ⚙️ 規則設定頁面全自動控制模組 (Dynamic Rule Management)
   ========================================================================== */
let currentRuleTab = 'wait_reasons'; 

async function switchRuleTab(tabName) {
    currentRuleTab = tabName;
    document.getElementById('tab-wait').classList.toggle('active', tabName === 'wait_reasons');
    document.getElementById('tab-keyword').classList.toggle('active', tabName === 'keywords');
    
    const formTitle = document.getElementById('rule-form-title');
    const inputsFlex = document.getElementById('rule-inputs-flex');
    
    if (tabName === 'wait_reasons') {
        formTitle.innerHTML = '➕ 新增等待原因排除規則 (規則 1)';
        inputsFlex.innerHTML = `
            <div style="flex: 2; min-width: 200px;">
                <label style="display:block; margin-bottom:5px; font-size:12px; color:var(--text2);">中榮原生等待原因名稱</label>
                <input type="text" id="new-wait-reason" placeholder="例如：等單位通知" style="width:100%; padding:8px; background:var(--bg); border:1px solid var(--border); border-radius:6px; color:white; height:38px;">
            </div>
            <div style="flex: 1; min-width: 120px;">
                <label style="display:block; margin-bottom:5px; font-size:12px; color:var(--text2);">動作判定</label>
                <select id="new-wait-action" style="width:100%; padding:8px; background:var(--bg); border:1px solid var(--border); border-radius:6px; color:white; height:38px;">
                    <option value="未延遲">✅ 未延遲</option>
                    <option value="需檢查">⚠️ 需檢查</option>
                    <option value="延遲">❌ 延遲</option>
                </select>
            </div>
            <div style="flex: 2; min-width: 200px;">
                <label style="display:block; margin-bottom:5px; font-size:12px; color:var(--text2);">規則備註/說明</label>
                <input type="text" id="new-wait-note" placeholder="例：不可歸責傳送人員" style="width:100%; padding:8px; background:var(--bg); border:1px solid var(--border); border-radius:6px; color:white; height:38px;">
            </div>
            <div><button class="btn btn-success" onclick="submitNewRule()" style="height:38px;">💾 儲存規則</button></div>`;
    } else if (tabName === 'keywords') {
        formTitle.innerHTML = '➕ 新增特別指示 / 等待原因 關鍵字過濾 (規則 7 & 8)';
        inputsFlex.innerHTML = `
            <div style="flex: 1; min-width: 130px;">
                <label style="display:block; margin-bottom:5px; font-size:12px; color:var(--text2);">比對目標欄位</label>
                <select id="new-kw-field" style="width:100%; padding:8px; background:var(--bg); border:1px solid var(--border); border-radius:6px; color:white; height:38px;">
                    <option value="特別指示">特別指示</option>
                    <option value="等待原因">等待原因</option>
                    <option value="派工紀錄">派工紀錄</option>
                </select>
            </div>
            <div style="flex: 1.5; min-width: 160px;">
                <label style="display:block; margin-bottom:5px; font-size:12px; color:var(--text2);">關鍵字</label>
                <input type="text" id="new-kw-keyword" placeholder="例如：故障" style="width:100%; padding:8px; background:var(--bg); border:1px solid var(--border); border-radius:6px; color:white; height:38px;">
            </div>
            <div style="flex: 1; min-width: 120px;">
                <label style="display:block; margin-bottom:5px; font-size:12px; color:var(--text2);">動作判定</label>
                <select id="new-kw-action" style="width:100%; padding:8px; background:var(--bg); border:1px solid var(--border); border-radius:6px; color:white; height:38px;">
                    <option value="未延遲">✅ 未延遲</option>
                    <option value="需檢查">⚠️ 需檢查</option>
                    <option value="延遲">❌ 延遲</option>
                </select>
            </div>
            <div style="flex: 2; min-width: 200px;">
                <label style="display:block; margin-bottom:5px; font-size:12px; color:var(--text2);">規則備註/說明</label>
                <input type="text" id="new-kw-note" placeholder="說明此關鍵字用途" style="width:100%; padding:8px; background:var(--bg); border:1px solid var(--border); border-radius:6px; color:white; height:38px;">
            </div>
            <div><button class="btn btn-success" onclick="submitNewRule()" style="height:38px;">💾 儲存規則</button></div>`;
    }
    await loadRulesList();
}

async function loadRulesList() {
    const tbody = document.getElementById('rule-table-body');
    const thead = document.getElementById('rule-table-header');
    if (!tbody || !thead) return;
    
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:15px; color:var(--text3);">⏳ 正在向資料庫深度檢索規則...</td></tr>';
    
    try {
        const res = await fetch(`/api/rules/${currentRuleTab}`);
        const rules = await res.json();
        
        if (currentRuleTab === 'wait_reasons') {
            thead.innerHTML = `
                <th style="padding:12px 14px; text-align:left;">ID</th>
                <th style="padding:12px 14px; text-align:left;">排除的等待原因</th>
                <th style="padding:12px 14px; text-align:left;">判定結果</th>
                <th style="padding:12px 14px; text-align:left;">備註脈絡</th>
                <th style="padding:12px 14px; text-align:left;">啟用狀態</th>
                <th style="padding:12px 14px; text-align:center;">操作</th>`;
            
            if (rules.length === 0) {
                tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:15px; color:var(--text3);">目前無設定任何等待原因規則</td></tr>';
                return;
            }
            
            tbody.innerHTML = rules.map(r => {
                const statusBadge = r.action === '未延遲' ? 'badge-green' : (r.action === '需檢查' ? 'badge-yellow' : 'badge-red');
                return `
                <tr style="border-bottom: 1px solid var(--border);">
                    <td style="padding:12px 14px;" class="mono">${r.id}</td>
                    <td style="padding:12px 14px; font-weight:600; color:var(--text);">${r.reason}</td>
                    <td style="padding:12px 14px;"><span class="badge ${statusBadge}">${r.action}</span></td>
                    <td style="padding:12px 14px; color:var(--text3); font-size:12px;">${r.note || '—'}</td>
                    <td style="padding:12px 14px;"><input type="checkbox" ${r.enabled ? 'checked' : ''} onchange="toggleRuleEnable('${currentRuleTab}', ${r.id}, this.checked)"></td>
                    <td style="padding:12px 14px; text-align:center;"><button class="btn btn-danger btn-sm" onclick="deleteRuleItem('${currentRuleTab}', ${r.id})">🗑️ 刪除</button></td>
                </tr>`;
            }).join('');
            
        } else if (currentRuleTab === 'keywords') {
            thead.innerHTML = `
                <th style="padding:12px 14px; text-align:left;">ID</th>
                <th style="padding:12px 14px; text-align:left;">比對目標欄位</th>
                <th style="padding:12px 14px; text-align:left;">關鍵字</th>
                <th style="padding:12px 14px; text-align:left;">動作判定</th>
                <th style="padding:12px 14px; text-align:left;">規則說明</th>
                <th style="padding:12px 14px; text-align:left;">啟用狀態</th>
                <th style="padding:12px 14px; text-align:center;">操作</th>`;
            
            if (rules.length === 0) {
                tbody.innerHTML = '<tr><td colspan="7" style="text-align:center; padding:15px; color:var(--text3);">目前無設定任何關鍵字規則</td></tr>';
                return;
            }
            
            tbody.innerHTML = rules.map(r => {
                const statusBadge = r.action === '未延遲' ? 'badge-green' : (r.action === '需檢查' ? 'badge-yellow' : 'badge-red');
                return `
                <tr style="border-bottom: 1px solid var(--border);">
                    <td style="padding:12px 14px;" class="mono">${r.id}</td>
                    <td style="padding:12px 14px;"><span class="badge badge-blue">${r.target_field}</span></td>
                    <td style="padding:12px 14px;" class="mono"><strong style="color:var(--yellow)">"${r.keyword}"</strong></td>
                    <td style="padding:12px 14px;"><span class="badge ${statusBadge}">${r.action}</span></td>
                    <td style="padding:12px 14px; font-size:12px; color:var(--text3);">${r.note || '—'}</td>
                    <td style="padding:12px 14px;"><input type="checkbox" ${r.enabled ? 'checked' : ''} onchange="toggleRuleEnable('${currentRuleTab}', ${r.id}, this.checked)"></td>
                    <td style="padding:12px 14px; text-align:center;"><button class="btn btn-danger btn-sm" onclick="deleteRuleItem('${currentRuleTab}', ${r.id})">🗑️ 刪除</button></td>
                </tr>`;
            }).join('');
        }
    } catch (e) {
        console.error("載入規則失敗:", e);
        tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding:15px; color:var(--red);">⚠️ 載入後端資料失敗</td></tr>';
    }
}

async function submitNewRule() {
    let payload = {};
    if (currentRuleTab === 'wait_reasons') {
        const reason = document.getElementById('new-wait-reason').value.trim();
        if(!reason) { alert("請輸入等待原因！"); return; }
        payload = { reason, action: document.getElementById('new-wait-action').value, note: document.getElementById('new-wait-note').value.trim() };
    } else if (currentRuleTab === 'keywords') {
        const keyword = document.getElementById('new-kw-keyword').value.trim();
        if(!keyword) { alert("請輸入關鍵字！"); return; }
        payload = { target_field: document.getElementById('new-kw-field').value, keyword, match_type: 'contains', action: document.getElementById('new-kw-action').value, note: document.getElementById('new-kw-note').value.trim() };
    }

    const res = await fetch(`/api/rules/${currentRuleTab}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
    });
    const d = await res.json();
    if (d.success) {
        showToast("規則已成功新增！");
        switchRuleTab(currentRuleTab); 
    } else {
        alert(d.message);
    }
}

async function toggleRuleEnable(tabName, id, isEnabled) {
    await fetch(`/api/rules/${tabName}/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ enabled: isEnabled })
    });
    showToast("規則啟用狀態已變更");
}

async function deleteRuleItem(tabName, id) {
    if (!confirm("確定要徹底刪除這條過濾規則嗎？")) return;
    const res = await fetch(`/api/rules/${tabName}/${id}`, { method: 'DELETE' });
    const d = await res.json();
    if (d.success) {
        showToast("規則已刪除");
        switchRuleTab(tabName);
    }
}

/* ==========================================================================
   8. 🕒 定時自動偵測後端更新並全自動刷新網頁數據 (Auto Refresh Engine)
   ========================================================================== */
let lastKnownUpdateString = null; 

async function checkBackendUpdateAndRefresh() {
    try {
        const res = await fetch('/api/stats');
        if (!res.ok) return;
        const data = await res.json();
        
        if (data && data.status && data.status.updated_at) {
            const currentUpdateTime = data.status.updated_at;
            
            if (lastKnownUpdateString === null) {
                lastKnownUpdateString = currentUpdateTime;
                return;
            }
            
            if (currentUpdateTime !== lastKnownUpdateString) {
                lastKnownUpdateString = currentUpdateTime;
                showToast('⏰ 系統已完成每日定時自動更新，正在為您同步最新數據！', 'success');
                
                try { await loadStats(); } catch(e) {}
                try { await loadCheck(); } catch(e) {}
                try { await loadReport(); } catch(e) {}
                try { if (typeof currentPage !== 'undefined') { doSearch(currentPage); } } catch(e) {}
            }
        }
    } catch (e) {
        console.warn("📌 自動偵測後端狀態時連線異常，將於下個週期重新嘗試。", e);
    }
}

// 啟動監聽器移至最外層，移除重複綁定與語法衝突
document.addEventListener('DOMContentLoaded', () => {
    const hasRuleTable = document.getElementById('rule-table-body');
    if (hasRuleTable && typeof switchRuleTab === 'function') {
        switchRuleTab('wait_reasons');
    }
    setInterval(loadStats, 300000); 
    setTimeout(checkBackendUpdateAndRefresh, 2000);
});