const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');
const generateDocx = require('./generate-report');

const PORT = 3000;

// 数据存储
let policyData = {
    lastUpdated: new Date().toISOString(),
    version: '1.0.0',
    fiscalBudget: {
        year: 2026,
        deficit: 58900,
        deficitRate: 4,
        specialBond: 44000,
        specialBondPurpose: '重大项目、置换隐性债务、消化拖欠账款',
        longTermBond: 13000,
        centralBudget: 7550,
        bankCapitalBond: 3000
    },
    twoNewPolicy: {
        totalBudget: 2000,
        firstBatch: 625,
        secondBatch: 915,
        cumulative: 1851,
        progressRate: 92,
        subsidyScope: ['老旧小区加装电梯', '养老机构', '消防救援', '检验检测', '商业综合体']
    },
    investment: {
        q1_2026_events: 2865,
        q1_2026_amount: 2560,
        yoy_events: 52,
        yoy_amount: 48,
        aiGlobalQ1: 2555,
        semiconductorQ1: 80
    },
    lowAltitude: {
        marketSize: 10600,
        marketSize2035: 35000,
        evtvolMarket: 95,
        uavRegistration: 1200,
        companies: 163000,
        cities: ['合肥', '杭州', '深圳', '苏州', '成都', '重庆']
    },
    trials: {
        carbonPeak: {
            secondBatch: 27,
            zeroCarbonParks: 100
        },
        carbonMarket: {
            coverage: 70,
            securitiesCompanies: 26
        }
    },
    sectors: {
        ai: { rating: '超配', signal: 5, coreLogic: 'Q1投资超2025全年，资本开支爆发' },
        lowAltitude: { rating: '超配', signal: 5, coreLogic: '商业化元年，市场规模破万亿' },
        consumption: { rating: '超配', signal: 5, coreLogic: '两新资金92%已下达，补贴力度大' },
        robot: { rating: '超配', signal: 5, coreLogic: '具身智能元年，融资爆发' },
        semiconductor: { rating: '标配', signal: 4, coreLogic: '国产替代加速，但需警惕周期' },
        newEnergy: { rating: '标配', signal: 4, coreLogic: '市场化转型，关注结构性机会' },
        zeroCarbon: { rating: '关注', signal: 4, coreLogic: '十五五将建100个，政策刚启动' }
    },
    history: []
};

// 加载保存的数据
const dataFilePath = path.join(__dirname, 'data.json');
if (fs.existsSync(dataFilePath)) {
    try {
        const savedData = JSON.parse(fs.readFileSync(dataFilePath, 'utf8'));
        policyData = { ...policyData, ...savedData };
    } catch (e) {
        console.log('使用默认数据');
    }
}

// 辅助函数：读取静态文件
function serveStaticFile(res, filePath, contentType) {
    fs.readFile(filePath, (err, data) => {
        if (err) {
            res.writeHead(404, { 'Content-Type': 'text/plain' });
            res.end('File not found');
            return;
        }
        res.writeHead(200, { 'Content-Type': contentType });
        res.end(data);
    });
}

// 辅助函数：发送JSON响应
function sendJson(res, data, statusCode = 200) {
    res.writeHead(statusCode, {
        'Content-Type': 'application/json',
        'Access-Control-Allow-Origin': '*'
    });
    res.end(JSON.stringify(data));
}

// 保存数据到文件
function saveData() {
    fs.writeFileSync(dataFilePath, JSON.stringify(policyData, null, 2));
}

// 创建HTTP服务器
const server = http.createServer((req, res) => {
    const parsedUrl = url.parse(req.url, true);
    const pathname = parsedUrl.pathname;
    const method = req.method;

    // CORS预检
    if (method === 'OPTIONS') {
        res.writeHead(200, {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type'
        });
        res.end();
        return;
    }

    // API路由
    if (pathname === '/api/data' && method === 'GET') {
        sendJson(res, policyData);
        return;
    }

    if (pathname === '/api/data' && method === 'POST') {
        let body = '';
        req.on('data', chunk => { body += chunk; });
        req.on('end', () => {
            try {
                const newData = JSON.parse(body);
                // 记录历史
                policyData.history.push({
                    timestamp: policyData.lastUpdated,
                    data: JSON.parse(JSON.stringify(policyData))
                });
                // 限制历史记录数量
                if (policyData.history.length > 50) {
                    policyData.history = policyData.history.slice(-50);
                }
                // 更新数据
                policyData = { ...policyData, ...newData, lastUpdated: new Date().toISOString() };
                saveData();
                sendJson(res, { success: true, message: '数据已更新', lastUpdated: policyData.lastUpdated });
            } catch (e) {
                sendJson(res, { success: false, message: '数据格式错误: ' + e.message }, 400);
            }
        });
        return;
    }

    if (pathname === '/api/import' && method === 'POST') {
        let body = '';
        req.on('data', chunk => { body += chunk; });
        req.on('end', () => {
            try {
                const importedData = JSON.parse(body);
                policyData.history.push({
                    timestamp: policyData.lastUpdated,
                    data: JSON.parse(JSON.stringify(policyData))
                });
                policyData = { ...policyData, ...importedData, lastUpdated: new Date().toISOString() };
                saveData();
                sendJson(res, { success: true, message: '数据导入成功', lastUpdated: policyData.lastUpdated });
            } catch (e) {
                sendJson(res, { success: false, message: '导入数据格式错误' }, 400);
            }
        });
        return;
    }

    if (pathname === '/api/export' && method === 'GET') {
        sendJson(res, policyData);
        return;
    }

    if (pathname === '/api/report' && method === 'GET') {
        const reportData = generateReportData(policyData);
        sendJson(res, reportData);
        return;
    }

    if (pathname === '/api/history' && method === 'GET') {
        sendJson(res, policyData.history || []);
        return;
    }

    // DOCX报告下载
    if (pathname === '/api/report/docx' && method === 'GET') {
        res.writeHead(200, {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': 'attachment; filename=policy_report.docx',
            'Access-Control-Allow-Origin': '*'
        });
        generateDocx(policyData).then(buffer => {
            res.end(buffer);
        }).catch(err => {
            res.writeHead(500, { 'Content-Type': 'text/plain' });
            res.end('Report generation failed: ' + err.message);
        });
        return;
    }

    // 静态文件服务
    if (pathname === '/' || pathname === '/index.html') {
        serveStaticFile(res, path.join(__dirname, 'public', 'index.html'), 'text/html');
        return;
    }

    if (pathname.startsWith('/css/')) {
        serveStaticFile(res, path.join(__dirname, 'public', pathname), 'text/css');
        return;
    }

    if (pathname.startsWith('/js/')) {
        serveStaticFile(res, path.join(__dirname, 'public', pathname), 'application/javascript');
        return;
    }

    // 404
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('Not found');
});

// 生成报告数据
function generateReportData(data) {
    return {
        title: '政策红利抓取方法论调研报告',
        subtitle: '——基于"钱"与"试点"信号的调研分析报告',
        version: data.version,
        lastUpdated: data.lastUpdated,
        sections: {
            fiscalBudget: {
                title: '二、看钱：资金流向追踪体系',
                data: [
                    { label: '全国财政赤字', value: data.fiscalBudget.deficit + '亿元', change: '+2300亿元' },
                    { label: '赤字率', value: data.fiscalBudget.deficitRate + '%', change: '持平' },
                    { label: '超长期特别国债', value: data.fiscalBudget.longTermBond + '亿元', change: '持平' },
                    { label: '地方政府专项债', value: data.fiscalBudget.specialBond + '亿元', change: '略降' },
                    { label: '中央预算内投资', value: data.fiscalBudget.centralBudget + '亿元', change: '增加' }
                ]
            },
            twoNewPolicy: {
                title: '两新政策资金下达情况',
                data: [
                    { label: '全年目标', value: data.twoNewPolicy.totalBudget + '亿元' },
                    { label: '累计下达', value: data.twoNewPolicy.cumulative + '亿元' },
                    { label: '完成进度', value: data.twoNewPolicy.progressRate + '%' }
                ]
            },
            investment: {
                title: '一级市场融资信号',
                data: [
                    { label: 'Q1融资事件', value: data.investment.q1_2026_events + '起', change: '+' + data.investment.yoy_events + '%' },
                    { label: 'Q1融资总额', value: data.investment.q1_2026_amount + '亿元', change: '+' + data.investment.yoy_amount + '%' },
                    { label: '全球AI投资', value: data.investment.aiGlobalQ1 + '亿美元' }
                ]
            },
            lowAltitude: {
                title: '低空经济（商业化元年）',
                data: [
                    { label: '市场规模', value: (data.lowAltitude.marketSize / 10000).toFixed(1) + '万亿元' },
                    { label: 'eVTOL市场规模', value: data.lowAltitude.evtvolMarket + '亿元' },
                    { label: '无人机实名登记', value: data.lowAltitude.uavRegistration + '万架' },
                    { label: '相关企业', value: (data.lowAltitude.companies / 10000).toFixed(1) + '万家' }
                ]
            },
            sectors: {
                title: '四、重点领域政策红利分析',
                data: Object.entries(data.sectors).map(([key, val]) => ({
                    name: getSectorName(key),
                    rating: val.rating,
                    signal: '★'.repeat(val.signal),
                    logic: val.coreLogic
                }))
            }
        }
    };
}

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

// 启动服务器
server.listen(PORT, () => {
    console.log('政策红利追踪工具已启动!');
    console.log('访问地址: http://localhost:' + PORT);
});
