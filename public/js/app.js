// 政策红利追踪工具 - 主应用脚本

let currentData = null;
let charts = {};

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    initForm();
    initFileUpload();
    refreshData();
});

// 初始化标签页
function initTabs() {
    const tabs = document.querySelectorAll('.tab');
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            tabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            
            const tabId = tab.dataset.tab;
            document.querySelectorAll('.tab-content').forEach(content => {
                content.classList.remove('active');
            });
            document.getElementById(tabId).classList.add('active');
            
            // 如果切换到图表标签，重新渲染图表
            if (tabId === 'chart' && currentData) {
                renderCharts();
            }
        });
    });
}

// 初始化表单
function initForm() {
    const form = document.getElementById('dataForm');
    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        await saveFormData();
    });
}

// 初始化文件上传
function initFileUpload() {
    const fileInput = document.getElementById('jsonFile');
    const preview = document.getElementById('filePreview');
    
    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = JSON.parse(e.target.result);
                    preview.innerHTML = '<pre>' + JSON.stringify(data, null, 2).substring(0, 500) + '...</pre>';
                    preview.dataset.content = e.target.result;
                } catch (err) {
                    preview.innerHTML = '<span style="color: red;">JSON格式错误</span>';
                }
            };
            reader.readAsText(file);
        }
    });
}

// 刷新数据
async function refreshData() {
    try {
        const response = await fetch('/api/data');
        currentData = await response.json();
        
        // 更新时间显示
        const updateTime = new Date(currentData.lastUpdated).toLocaleString('zh-CN');
        document.getElementById('lastUpdated').textContent = '最后更新: ' + updateTime;
        
        // 渲染仪表盘
        renderDashboard();
    } catch (err) {
        showToast('数据加载失败', 'error');
    }
}

// 渲染仪表盘
function renderDashboard() {
    // 财政预算
    const fiscalHtml = `
        <div class="metric-item">
            <span class="metric-label">全国财政赤字</span>
            <span class="metric-value">${currentData.fiscalBudget.deficit}<span class="metric-change positive">+2300亿</span></span>
        </div>
        <div class="metric-item">
            <span class="metric-label">赤字率</span>
            <span class="metric-value">${currentData.fiscalBudget.deficitRate}%</span>
        </div>
        <div class="metric-item">
            <span class="metric-label">超长期特别国债</span>
            <span class="metric-value">${currentData.fiscalBudget.longTermBond}亿</span>
        </div>
        <div class="metric-item">
            <span class="metric-label">地方政府专项债</span>
            <span class="metric-value">${currentData.fiscalBudget.specialBond}亿</span>
        </div>
    `;
    document.getElementById('fiscalBudget').innerHTML = fiscalHtml;
    
    // 两新政策
    const twoNewHtml = `
        <div class="metric-item">
            <span class="metric-label">全年目标</span>
            <span class="metric-value">${currentData.twoNewPolicy.totalBudget}亿</span>
        </div>
        <div class="metric-item">
            <span class="metric-label">累计下达</span>
            <span class="metric-value">${currentData.twoNewPolicy.cumulative}亿</span>
        </div>
        <div class="metric-item">
            <span class="metric-label">完成进度</span>
            <span class="metric-value">${currentData.twoNewPolicy.progressRate}%</span>
        </div>
    `;
    document.getElementById('twoNewPolicy').innerHTML = twoNewHtml;
    
    // 融资数据
    const invHtml = `
        <div class="metric-item">
            <span class="metric-label">Q1融资事件</span>
            <span class="metric-value">${currentData.investment.q1_2026_events}起</span>
        </div>
        <div class="metric-item">
            <span class="metric-label">Q1融资总额</span>
            <span class="metric-value">${currentData.investment.q1_2026_amount}亿</span>
        </div>
        <div class="metric-item">
            <span class="metric-label">同比增长</span>
            <span class="metric-value">+${currentData.investment.yoy_events}%</span>
        </div>
        <div class="metric-item">
            <span class="metric-label">全球AI投资</span>
            <span class="metric-value">${currentData.investment.aiGlobalQ1}亿美元</span>
        </div>
    `;
    document.getElementById('investment').innerHTML = invHtml;
    
    // 低空经济
    const lowAltHtml = `
        <div class="metric-item">
            <span class="metric-label">市场规模</span>
            <span class="metric-value">${(currentData.lowAltitude.marketSize/10000).toFixed(1)}万亿</span>
        </div>
        <div class="metric-item">
            <span class="metric-label">eVTOL市场</span>
            <span class="metric-value">${currentData.lowAltitude.evtvolMarket}亿</span>
        </div>
        <div class="metric-item">
            <span class="metric-label">无人机登记</span>
            <span class="metric-value">${currentData.lowAltitude.uavRegistration}万架</span>
        </div>
        <div class="metric-item">
            <span class="metric-label">相关企业</span>
            <span class="metric-value">${(currentData.lowAltitude.companies/10000).toFixed(1)}万家</span>
        </div>
    `;
    document.getElementById('lowAltitude').innerHTML = lowAltHtml;
    
    // 投资建议表格
    const sectorsHtml = Object.entries(currentData.sectors).map(([key, val]) => {
        const ratingClass = val.rating === '超配' ? 'high' : 'medium';
        return `
            <tr>
                <td>${getSectorName(key)}</td>
                <td><span class="rating-badge ${ratingClass}">${val.rating}</span></td>
                <td><span class="signal-stars">${'★'.repeat(val.signal)}</span></td>
                <td>${val.coreLogic}</td>
            </tr>
        `;
    }).join('');
    document.getElementById('sectorsBody').innerHTML = sectorsHtml;
    
    // 填充表单
    fillForm();
}

// 获取领域名称
function getSectorName(key) {
    const names = {
        ai: 'AI/算力',
        lowAltitude: '低空经济',
        consumption: '消费/汽车',
        robot: '机器人',
        semiconductor: '半导体设备',
        newEnergy: '新能源',
        zeroCarbon: '零碳园区'
    };
    return names[key] || key;
}

// 填充表单
function fillForm() {
    if (!currentData) return;
    
    // 财政预算
    document.getElementById('deficit').value = currentData.fiscalBudget.deficit;
    document.getElementById('deficitRate').value = currentData.fiscalBudget.deficitRate;
    document.getElementById('longTermBond').value = currentData.fiscalBudget.longTermBond;
    document.getElementById('specialBond').value = currentData.fiscalBudget.specialBond;
    document.getElementById('centralBudget').value = currentData.fiscalBudget.centralBudget;
    document.getElementById('bankCapitalBond').value = currentData.fiscalBudget.bankCapitalBond || 0;
    
    // 两新政策
    document.getElementById('totalBudget').value = currentData.twoNewPolicy.totalBudget;
    document.getElementById('cumulative').value = currentData.twoNewPolicy.cumulative;
    document.getElementById('progressRate').value = currentData.twoNewPolicy.progressRate;
    
    // 融资数据
    document.getElementById('q1_events').value = currentData.investment.q1_2026_events;
    document.getElementById('q1_amount').value = currentData.investment.q1_2026_amount;
    document.getElementById('yoy').value = currentData.investment.yoy_events;
    document.getElementById('aiGlobal').value = currentData.investment.aiGlobalQ1;
    
    // 低空经济
    document.getElementById('marketSize').value = currentData.lowAltitude.marketSize;
    document.getElementById('evtvolMarket').value = currentData.lowAltitude.evtvolMarket;
    document.getElementById('uavRegistration').value = currentData.lowAltitude.uavRegistration;
    document.getElementById('companies').value = (currentData.lowAltitude.companies / 10000).toFixed(1);
}

// 保存表单数据
async function saveFormData() {
    const data = {
        fiscalBudget: {
            deficit: parseInt(document.getElementById('deficit').value),
            deficitRate: parseFloat(document.getElementById('deficitRate').value),
            longTermBond: parseInt(document.getElementById('longTermBond').value),
            specialBond: parseInt(document.getElementById('specialBond').value),
            centralBudget: parseInt(document.getElementById('centralBudget').value),
            bankCapitalBond: parseInt(document.getElementById('bankCapitalBond').value) || 0
        },
        twoNewPolicy: {
            totalBudget: parseInt(document.getElementById('totalBudget').value),
            cumulative: parseInt(document.getElementById('cumulative').value),
            progressRate: parseInt(document.getElementById('progressRate').value)
        },
        investment: {
            q1_2026_events: parseInt(document.getElementById('q1_events').value),
            q1_2026_amount: parseInt(document.getElementById('q1_amount').value),
            yoy_events: parseInt(document.getElementById('yoy').value),
            aiGlobalQ1: parseInt(document.getElementById('aiGlobal').value)
        },
        lowAltitude: {
            marketSize: parseInt(document.getElementById('marketSize').value),
            evtvolMarket: parseInt(document.getElementById('evtvolMarket').value),
            uavRegistration: parseInt(document.getElementById('uavRegistration').value),
            companies: Math.round(parseFloat(document.getElementById('companies').value) * 10000)
        }
    };
    
    try {
        const response = await fetch('/api/data', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        const result = await response.json();
        if (result.success) {
            showToast('数据保存成功！', 'success');
            await refreshData();
        } else {
            showToast('保存失败: ' + result.message, 'error');
        }
    } catch (err) {
        showToast('保存失败', 'error');
    }
}

// 导入文件
async function importFile() {
    const preview = document.getElementById('filePreview');
    const content = preview.dataset.content;
    
    if (!content) {
        showToast('请先选择JSON文件', 'error');
        return;
    }
    
    try {
        const data = JSON.parse(content);
        const response = await fetch('/api/import', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        const result = await response.json();
        if (result.success) {
            showToast('数据导入成功！', 'success');
            await refreshData();
        } else {
            showToast('导入失败: ' + result.message, 'error');
        }
    } catch (err) {
        showToast('JSON格式错误', 'error');
    }
}

// 下载模板
function downloadTemplate() {
    const template = {
        version: "1.0.0",
        fiscalBudget: {
            deficit: 58900,
            deficitRate: 4,
            specialBond: 44000,
            longTermBond: 13000,
            centralBudget: 7550,
            bankCapitalBond: 3000
        },
        twoNewPolicy: {
            totalBudget: 2000,
            cumulative: 1851,
            progressRate: 92
        },
        investment: {
            q1_2026_events: 2865,
            q1_2026_amount: 2560,
            yoy_events: 52,
            aiGlobalQ1: 2555
        },
        lowAltitude: {
            marketSize: 10600,
            evtvolMarket: 95,
            uavRegistration: 1200,
            companies: 163000
        },
        sectors: {
            ai: { rating: "超配", signal: 5, coreLogic: "Q1投资超2025全年，资本开支爆发" },
            lowAltitude: { rating: "超配", signal: 5, coreLogic: "商业化元年，市场规模破万亿" },
            consumption: { rating: "超配", signal: 5, coreLogic: "两新资金92%已下达，补贴力度大" },
            robot: { rating: "超配", signal: 5, coreLogic: "具身智能元年，融资爆发" },
            semiconductor: { rating: "标配", signal: 4, coreLogic: "国产替代加速，但需警惕周期" },
            newEnergy: { rating: "标配", signal: 4, coreLogic: "市场化转型，关注结构性机会" },
            zeroCarbon: { rating: "关注", signal: 4, coreLogic: "十五五将建100个，政策刚启动" }
        }
    };
    
    downloadFile(JSON.stringify(template, null, 2), 'policy_data_template.json', 'application/json');
}

// 导出数据
async function exportData() {
    if (!currentData) {
        showToast('数据加载中...', 'error');
        return;
    }
    downloadFile(JSON.stringify(currentData, null, 2), 'policy_data_export.json', 'application/json');
}

// 下载文件
function downloadFile(content, filename, type) {
    const blob = new Blob([content], { type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
}

// 渲染图表
function renderCharts() {
    if (!currentData) return;
    
    // 销毁旧图表
    Object.values(charts).forEach(chart => chart.destroy());
    charts = {};
    
    // 两新政策进度图表
    const twoNewCtx = document.getElementById('twoNewChart').getContext('2d');
    charts.twoNew = new Chart(twoNewCtx, {
        type: 'doughnut',
        data: {
            labels: ['已下达', '待下达'],
            datasets: [{
                data: [
                    currentData.twoNewPolicy.cumulative,
                    currentData.twoNewPolicy.totalBudget - currentData.twoNewPolicy.cumulative
                ],
                backgroundColor: ['#2563eb', '#e2e8f0']
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { position: 'bottom' },
                title: {
                    display: true,
                    text: `完成进度: ${currentData.twoNewPolicy.progressRate}%`
                }
            }
        }
    });
    
    // 融资数据图表
    const invCtx = document.getElementById('investmentChart').getContext('2d');
    charts.investment = new Chart(invCtx, {
        type: 'bar',
        data: {
            labels: ['Q1融资事件', 'Q1融资总额(亿)', '全球AI投资(亿美元)'],
            datasets: [{
                label: '2026年数据',
                data: [
                    currentData.investment.q1_2026_events,
                    currentData.investment.q1_2026_amount,
                    currentData.investment.aiGlobalQ1
                ],
                backgroundColor: ['#2563eb', '#16a34a', '#f59e0b']
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: { beginAtZero: true }
            }
        }
    });
    
    // 低空经济图表
    const lowAltCtx = document.getElementById('lowAltitudeChart').getContext('2d');
    charts.lowAltitude = new Chart(lowAltCtx, {
        type: 'bar',
        data: {
            labels: ['市场规模(亿)', 'eVTOL市场(亿)', '无人机登记(万)', '相关企业(万)'],
            datasets: [{
                label: '低空经济数据',
                data: [
                    currentData.lowAltitude.marketSize,
                    currentData.lowAltitude.evtvolMarket,
                    currentData.lowAltitude.uavRegistration,
                    currentData.lowAltitude.companies / 10000
                ],
                backgroundColor: ['#2563eb', '#16a34a', '#f59e0b', '#dc2626']
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: { beginAtZero: true }
            }
        }
    });
    
    // 行业信号图表
    const sectorsCtx = document.getElementById('sectorsChart').getContext('2d');
    const sectorNames = Object.keys(currentData.sectors);
    const sectorSignals = Object.values(currentData.sectors).map(s => s.signal);
    const colors = sectorSignals.map(s => s >= 5 ? '#16a34a' : '#f59e0b');
    
    charts.sectors = new Chart(sectorsCtx, {
        type: 'radar',
        data: {
            labels: sectorNames.map(getSectorName),
            datasets: [{
                label: '信号强度',
                data: sectorSignals,
                backgroundColor: 'rgba(37, 99, 235, 0.2)',
                borderColor: '#2563eb',
                pointBackgroundColor: colors
            }]
        },
        options: {
            responsive: true,
            scales: {
                r: {
                    beginAtZero: true,
                    max: 5
                }
            }
        }
    });
}

// 生成报告
async function generateReport() {
    if (!currentData) {
        showToast('数据加载中...', 'error');
        return;
    }
    
    const preview = document.getElementById('reportPreview');
    const updateTime = new Date(currentData.lastUpdated).toLocaleString('zh-CN');
    
    let html = `
        <h1>政策红利抓取方法论调研报告</h1>
        <p style="text-align:center;color:#666;">基于"钱"与"试点"信号的调研分析报告</p>
        <p style="text-align:center;font-size:12px;color:#999;">更新时间: ${updateTime}</p>
        
        <h2>一、核心方法论框架</h2>
        <p>政策红利抓取的核心在于建立系统化的信息追踪机制，紧盯"钱"和"试点"两大信号。</p>
        
        <h2>二、看钱：资金流向追踪体系</h2>
        <table>
            <tr><th>指标</th><th>数值</th><th>变化</th></tr>
            <tr><td>全国财政赤字</td><td>${currentData.fiscalBudget.deficit}亿元</td><td>+2300亿元</td></tr>
            <tr><td>赤字率</td><td>${currentData.fiscalBudget.deficitRate}%</td><td>持平</td></tr>
            <tr><td>超长期特别国债</td><td>${currentData.fiscalBudget.longTermBond}亿元</td><td>持平</td></tr>
            <tr><td>地方政府专项债</td><td>${currentData.fiscalBudget.specialBond}亿元</td><td>略降</td></tr>
            <tr><td>中央预算内投资</td><td>${currentData.fiscalBudget.centralBudget}亿元</td><td>增加</td></tr>
        </table>
        
        <h3>两新政策资金下达情况</h3>
        <table>
            <tr><th>指标</th><th>数值</th></tr>
            <tr><td>全年目标</td><td>${currentData.twoNewPolicy.totalBudget}亿元</td></tr>
            <tr><td>累计下达</td><td>${currentData.twoNewPolicy.cumulative}亿元</td></tr>
            <tr><td>完成进度</td><td>${currentData.twoNewPolicy.progressRate}%</td></tr>
        </table>
        
        <h2>三、一级市场融资信号</h2>
        <table>
            <tr><th>指标</th><th>数值</th><th>同比</th></tr>
            <tr><td>Q1融资事件</td><td>${currentData.investment.q1_2026_events}起</td><td>+${currentData.investment.yoy_events}%</td></tr>
            <tr><td>Q1融资总额</td><td>${currentData.investment.q1_2026_amount}亿元</td><td>+${currentData.investment.yoy_amount}%</td></tr>
            <tr><td>全球AI投资</td><td>${currentData.investment.aiGlobalQ1}亿美元</td><td>-</td></tr>
        </table>
        
        <h2>四、低空经济（商业化元年）</h2>
        <table>
            <tr><th>指标</th><th>数值</th></tr>
            <tr><td>市场规模</td><td>${(currentData.lowAltitude.marketSize/10000).toFixed(1)}万亿元</td></tr>
            <tr><td>eVTOL市场</td><td>${currentData.lowAltitude.evtvolMarket}亿元</td></tr>
            <tr><td>无人机登记</td><td>${currentData.lowAltitude.uavRegistration}万架</td></tr>
            <tr><td>相关企业</td><td>${(currentData.lowAltitude.companies/10000).toFixed(1)}万家</td></tr>
        </table>
        
        <h2>五、投资建议</h2>
        <table>
            <tr><th>领域</th><th>评级</th><th>信号</th><th>核心逻辑</th></tr>
            ${Object.entries(currentData.sectors).map(([key, val]) => `
                <tr>
                    <td>${getSectorName(key)}</td>
                    <td>${val.rating}</td>
                    <td>${'★'.repeat(val.signal)}</td>
                    <td>${val.coreLogic}</td>
                </tr>
            `).join('')}
        </table>
    `;
    
    preview.innerHTML = html;
    showToast('报告已生成！', 'success');
}

// 下载报告（DOCX完整版）
async function downloadReport() {
    showToast('正在生成DOCX报告...', 'info');
    try {
        const response = await fetch('/api/report/docx');
        if (!response.ok) throw new Error('生成失败');
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const now = new Date();
        const dateStr = now.getFullYear() + '' + String(now.getMonth()+1).padStart(2,'0') + String(now.getDate()).padStart(2,'0');
        a.download = `政策红利调研报告_${dateStr}.docx`;
        a.click();
        URL.revokeObjectURL(url);
        showToast('报告下载成功！', 'success');
    } catch (err) {
        showToast('下载失败: ' + err.message, 'error');
    }
}

// 显示提示
function showToast(message, type = 'info') {
    const toast = document.getElementById('toast');
    toast.textContent = message;
    toast.className = 'toast ' + type;
    toast.classList.add('show');
    
    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
}
