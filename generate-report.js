const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat,
        HeadingLevel, BorderStyle, WidthType, ShadingType,
        PageNumber, PageBreak } = require('docx');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function C(text, width, opts = {}) {
    const { bold = false, fill = null } = opts;
    return new TableCell({
        borders, width: { size: width, type: WidthType.DXA },
        shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text, bold, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } })] })]
    });
}

function P(text, opts = {}) {
    const { bold = false, size = 24, color = undefined, spacing = {}, heading = undefined, numbering = undefined, alignment = undefined } = opts;
    const para = { children: [new TextRun({ text, bold, size, color, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } })] };
    if (heading) para.heading = heading;
    if (numbering) para.numbering = numbering;
    if (alignment) para.alignment = alignment;
    if (spacing.before || spacing.after) para.spacing = spacing;
    return new Paragraph(para);
}

function PB() { return new Paragraph({ children: [new PageBreak()] }); }

function sectorName(k) {
    return { ai: 'AI/算力', lowAltitude: '低空经济', consumption: '消费/汽车', robot: '机器人', semiconductor: '半导体设备', newEnergy: '新能源', zeroCarbon: '零碳园区' }[k] || k;
}

function stars(n) { return '\u2605'.repeat(n); }

function bullet(ref, text) {
    return new Paragraph({ numbering: { reference: ref, level: 0 }, children: [new TextRun({ text, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } })] });
}

function num(ref, text) {
    return new Paragraph({ numbering: { reference: ref, level: 0 }, children: [new TextRun({ text, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } })] });
}

module.exports = async function generateDocx(d) {
    const ut = new Date(d.lastUpdated).toLocaleString('zh-CN');

    const doc = new Document({
        styles: {
            default: { document: { run: { font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" }, size: 24 } } },
            paragraphStyles: [
                { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
                  run: { size: 36, bold: true, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } },
                  paragraph: { spacing: { before: 400, after: 240 }, outlineLevel: 0, keepNext: false, keepLines: false } },
                { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
                  run: { size: 30, bold: true, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } },
                  paragraph: { spacing: { before: 300, after: 180 }, outlineLevel: 1, keepNext: false, keepLines: false } },
                { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
                  run: { size: 26, bold: true, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } },
                  paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 2, keepNext: false, keepLines: false } },
            ]
        },
        numbering: {
            config: [
                { reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
                { reference: "n", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
            ]
        },
        sections: [{
            properties: {
                page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
            },
            headers: {
                default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "\u653F\u7B56\u7EA2\u5229\u6293\u53D6\u65B9\u6CD5\u8BBA\u8C03\u7814\u62A5\u544A", size: 20, color: "666666", font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } })] })] })
            },
            footers: {
                default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "\u7B2C ", size: 20, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } }), new TextRun({ children: [PageNumber.CURRENT], size: 20, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } }), new TextRun({ text: " \u9875", size: 20, font: { ascii: "Arial", hAnsi: "Arial", eastAsia: "Microsoft YaHei" } })] })] })
            },
            children: [
                // ====== 封面 ======
                new Paragraph({ spacing: { before: 2000 }, children: [] }),
                P("\u653F\u7B56\u7EA2\u5229\u6293\u53D6\u65B9\u6CD5\u8BBA", { bold: true, size: 52, alignment: AlignmentType.CENTER }),
                P("\u2014\u2014\u57FA\u4E8E\u201C\u94B1\u201D\u4E0E\u201C\u8BD5\u70B9\u201D\u4FE1\u53F7\u7684\u8C03\u7814\u5206\u6790\u62A5\u544A", { size: 32, alignment: AlignmentType.CENTER, spacing: { before: 400 } }),
                new Paragraph({ spacing: { before: 1500 }, children: [] }),
                P("\u6295\u8D44\u51B3\u7B56\u53C2\u8003", { size: 28, alignment: AlignmentType.CENTER }),
                new Paragraph({ spacing: { before: 400 }, children: [] }),
                P("\u3010" + ut + " \u66F4\u65B0\u3011", { size: 26, bold: true, color: "FF6600", alignment: AlignmentType.CENTER }),
                PB(),

                // ====== 更新说明 ======
                P("\u66F4\u65B0\u8BF4\u660E", { heading: HeadingLevel.HEADING_1 }),
                P("\u672C\u62A5\u544A\u57FA\u4E8E\u6700\u65B0\u653F\u7B56\u6570\u636E\u5168\u9762\u66F4\u65B0\uFF0C\u4E3B\u8981\u66F4\u65B0\u5185\u5BB9\u5305\u62EC\uFF1A", { spacing: { before: 200, after: 200 } }),
                num("n", "\u8D22\u653F\u8D64\u5B57" + d.fiscalBudget.deficit + "\u4EBF\u5143\u3001\u8D64\u5B57\u7387" + d.fiscalBudget.deficitRate + "%\u3001\u8D85\u957F\u671F\u7279\u522B\u56FD\u503A" + d.fiscalBudget.longTermBond + "\u4EBF\u5143"),
                num("n", "\u4E13\u9879\u503A" + d.fiscalBudget.specialBond + "\u4EBF\u5143\uFF0C\u5B8C\u5584\u8D1F\u9762\u6E05\u5355\u7BA1\u7406\u548C\u81EA\u5BA1\u81EA\u53D1\u8BD5\u70B9"),
                num("n", "\u4E24\u65B0\u653F\u7B56\u7D2F\u8BA1" + d.twoNewPolicy.cumulative + "\u4EBF\u5143\u5DF2\u4E0B\u8FBE\uFF08\u5360\u5168\u5E74" + d.twoNewPolicy.totalBudget + "\u4EBF\u5143\u7684" + d.twoNewPolicy.progressRate + "%\uFF09"),
                num("n", "\u4F4E\u7A7A\u7ECF\u6D4E\u5E02\u573A\u89C4\u6A21\u7A81\u7834" + (d.lowAltitude.marketSize/10000).toFixed(1) + "\u4E07\u4EBF\u5143\uFF0CeVTOL\u8FDB\u5165\u5546\u4E1A\u5316\u5143\u5E74"),
                num("n", "AI\u6295\u8D44\u7206\u53D1\uFF1AQ1\u5168\u7403" + d.investment.aiGlobalQ1 + "\u4EBF\u7F8E\u5143\uFF0C\u8D85\u8FC72025\u5E74\u5168\u5E74"),
                num("n", "\u78B3\u5E02\u573A\u8FDB\u51652.0\u65F6\u4EE3\uFF0C\u5341\u4E94\u4E94\u5C06\u5EFA" + d.trials.carbonPeak.zeroCarbonParks + "\u4E2A\u96F6\u78B3\u56ED\u533A"),
                PB(),

                // ====== 一、核心方法论 ======
                P("\u4E00\u3001\u6838\u5FC3\u65B9\u6CD5\u8BBA\u6846\u67B6", { heading: HeadingLevel.HEADING_1 }),
                P("\u653F\u7B56\u7EA2\u5229\u6293\u53D6\u7684\u6838\u5FC3\u5728\u4E8E\u5EFA\u7ACB\u7CFB\u7EDF\u5316\u7684\u4FE1\u606F\u8FFD\u8E2A\u673A\u5236\uFF0C\u7D27\u76EF\u201C\u94B1\u201D\u548C\u201C\u8BD5\u70B9\u201D\u4E24\u5927\u4FE1\u53F7\u3002\u8FD9\u4E00\u65B9\u6CD5\u8BBA\u57FA\u4E8E\u4EE5\u4E0B\u6838\u5FC3\u903B\u8F91\uFF1A\u653F\u7B56\u843D\u5730\u524D\u5FC5\u7136\u4F34\u968F\u8D44\u91D1\u5B89\u6392\u548C\u8BD5\u70B9\u63A2\u7D22\uFF0C\u63D0\u524D\u8BC6\u522B\u8FD9\u4E9B\u4FE1\u53F7\u80FD\u591F\u6293\u4F4F\u653F\u7B56\u843D\u5730\u524D\u7684\u65F6\u95F4\u7A97\u53E3\uFF0C\u83B7\u53D6\u8D85\u989D\u6536\u76CA\u3002", { spacing: { before: 200, after: 200 } }),

                P("1.1 \u65B9\u6CD5\u8BBA\u6838\u5FC3\u903B\u8F91", { heading: HeadingLevel.HEADING_2 }),
                P("\u653F\u7B56\u7EA2\u5229\u6293\u53D6\u65B9\u6CD5\u8BBA\u5EFA\u7ACB\u5728\u4E09\u4E2A\u57FA\u672C\u5047\u8BBE\u4E4B\u4E0A\uFF1A", { spacing: { before: 200, after: 200 } }),
                num("n", "\u653F\u7B56\u5148\u884C\uFF1A\u4EFB\u4F55\u91CD\u5927\u4EA7\u4E1A\u653F\u7B56\u7684\u63A8\u51FA\u90FD\u9700\u8981\u524D\u671F\u8D44\u91D1\u51C6\u5907\u548C\u8BD5\u70B9\u9A8C\u8BC1"),
                num("n", "\u8D44\u91D1\u5BFC\u5411\uFF1A\u8D22\u653F\u8D44\u91D1\u3001\u4E13\u9879\u503A\u3001\u793E\u4F1A\u8D44\u672C\u6D41\u5411\u76F4\u63A5\u53CD\u6620\u653F\u7B56\u6267\u884C\u529B\u5EA6"),
                num("n", "\u8BD5\u70B9\u9A8C\u8BC1\uFF1A\u53D1\u6539\u59D4\u8BD5\u70B9\u9879\u76EE\u7684KPI\u8FBE\u6210\u60C5\u51B5\u662F\u884C\u4E1A\u62D0\u70B9\u7684\u5148\u884C\u6307\u6807"),
                P("\u57FA\u4E8E\u4E0A\u8FF0\u903B\u8F91\uFF0C\u6211\u4EEC\u6784\u5EFA\u4E86\u53CC\u8F6E\u9A71\u52A8\u7684\u8FFD\u8E2A\u4F53\u7CFB\uFF1A", { spacing: { before: 300 } }),

                new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: [2340, 3510, 3510], rows: [
                    new TableRow({ cantSplit: true, children: [C("\u7EF4\u5EA6", 2340, { bold: true, fill: "D5E8F0" }), C("\u770B\u94B1", 3510, { bold: true, fill: "D5E8F0" }), C("\u770B\u8BD5\u70B9", 3510, { bold: true, fill: "D5E8F0" })] }),
                    new TableRow({ cantSplit: true, children: [C("\u6838\u5FC3\u6307\u6807", 2340, { bold: true }), C("\u9884\u7B97\u62E8\u6B3E\u3001\u4E13\u9879\u503A\u3001\u878D\u8D44\u89C4\u6A21", 3510), C("\u8BD5\u70B9\u540D\u5355\u3001KPI\u8FBE\u6210\u7387", 3510)] }),
                    new TableRow({ cantSplit: true, children: [C("\u6570\u636E\u6765\u6E90", 2340, { bold: true }), C("\u8D22\u653F\u90E8\u5B98\u7F51\u3001\u53D1\u6539\u59D4\u3001IT\u6854\u5B50\u7B49", 3510), C("\u53D1\u6539\u59D4\u5B98\u7F51\u3001\u5730\u65B9\u653F\u5E9C\u516C\u544A", 3510)] }),
                    new TableRow({ cantSplit: true, children: [C("\u8DDF\u8E2A\u9891\u7387", 2340, { bold: true }), C("\u6BCF\u670830\u5206\u949F\uFF08\u4E13\u9879\u503A\uFF09\u3001\u6BCF\u5B63\u5EA6\uFF08\u9884\u7B97\uFF09", 3510), C("\u6BCF\u5B63\u5EA6\uFF08\u8BD5\u70B9\u8FDB\u5C55\uFF09", 3510)] }),
                    new TableRow({ cantSplit: true, children: [C("\u4FE1\u53F7\u610F\u4E49", 2340, { bold: true }), C("\u589E\u957F\u91CF=\u6267\u884C\u529B\u5EA6\uFF0C\u8FDE\u7EED\u589E\u957F=\u673A\u4F1A\u4FE1\u53F7", 3510), C("\u8D85\u9884\u671F=\u884C\u4E1A\u62D0\u70B9\u4FE1\u53F7", 3510)] }),
                ] }),
                PB(),

                // ====== 二、看钱 ======
                P("\u4E8C\u3001\u770B\u94B1\uFF1A\u8D44\u91D1\u6D41\u5411\u8FFD\u8E2A\u4F53\u7CFB", { heading: HeadingLevel.HEADING_1 }),
                P("\u8D44\u91D1\u662F\u653F\u7B56\u6267\u884C\u7684\u6838\u5FC3\u8F7D\u4F53\u3002\u901A\u8FC7\u8FFD\u8E2A\u4E09\u7C7B\u8D44\u91D1\u2014\u2014\u90E8\u59D4\u9884\u7B97\u62E8\u6B3E\u3001\u5730\u65B9\u4E13\u9879\u503A\u3001\u4E00\u7EA7\u5E02\u573A\u878D\u8D44\u2014\u2014\u53EF\u4EE5\u7CBE\u51C6\u628A\u63E1\u653F\u7B56\u6267\u884C\u529B\u5EA6\u548C\u4EA7\u4E1A\u666F\u6C14\u5EA6\u3002", { spacing: { before: 200, after: 200 } }),

                // 2.1
                P("2.1 \u8D22\u653F\u9884\u7B97\u4E0E\u8D85\u957F\u671F\u7279\u522B\u56FD\u503A", { heading: HeadingLevel.HEADING_2 }),
                P("2026\u5E74\u653F\u5E9C\u5DE5\u4F5C\u62A5\u544A\u660E\u786E\u7EE7\u7EED\u5B9E\u65BD\u66F4\u52A0\u79EF\u6781\u7684\u8D22\u653F\u653F\u7B56\uFF0C\u8D64\u5B57\u7387\u6309" + d.fiscalBudget.deficitRate + "%\u5DE6\u53F3\u5B89\u6392\uFF0C\u5168\u56FD\u8D22\u653F\u8D64\u5B57" + d.fiscalBudget.deficit + "\u4EBF\u5143\u3002", { spacing: { before: 200, after: 200 } }),

                new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: [3120, 2340, 2340, 2340], rows: [
                    new TableRow({ cantSplit: true, children: [C("\u8D44\u91D1\u9879\u76EE", 3120, { bold: true, fill: "D5E8F0" }), C("2025\u5E74", 2340, { bold: true, fill: "D5E8F0" }), C("2026\u5E74", 2340, { bold: true, fill: "D5E8F0" }), C("\u53D8\u5316", 2340, { bold: true, fill: "D5E8F0" })] }),
                    new TableRow({ cantSplit: true, children: [C("\u5168\u56FD\u8D22\u653F\u8D64\u5B57", 3120), C("56600\u4EBF\u5143", 2340), C(d.fiscalBudget.deficit + "\u4EBF\u5143", 2340), C("+2300\u4EBF\u5143", 2340)] }),
                    new TableRow({ cantSplit: true, children: [C("\u8D64\u5B57\u7387", 3120), C("\u7EA64%", 2340), C(d.fiscalBudget.deficitRate + "%\u5DE6\u53F3", 2340), C("\u6301\u5E73", 2340)] }),
                    new TableRow({ cantSplit: true, children: [C("\u8D85\u957F\u671F\u7279\u522B\u56FD\u503A", 3120), C("1.3\u4E07\u4EBF\u5143", 2340), C(d.fiscalBudget.longTermBond + "\u4EBF\u5143", 2340), C("\u6301\u5E73", 2340)] }),
                    new TableRow({ cantSplit: true, children: [C("\u5176\u4E2D\uFF1A\u4E24\u91CD\u5EFA\u8BBE", 3120), C("8000\u4EBF\u5143", 2340), C("8000\u4EBF\u5143", 2340), C("\u6301\u5E73", 2340)] }),
                    new TableRow({ cantSplit: true, children: [C("\u5176\u4E2D\uFF1A\u4E24\u65B0\u5DE5\u4F5C", 3120), C("\u7EA65000\u4EBF\u5143", 2340), C("5000\u4EBF\u5143", 2340), C("\u6301\u5E73", 2340)] }),
                    new TableRow({ cantSplit: true, children: [C("\u94F6\u884C\u6CE8\u8D44\u7279\u522B\u56FD\u503A", 3120), C("\u65E0", 2340), C((d.fiscalBudget.bankCapitalBond || 3000) + "\u4EBF\u5143", 2340), C("\u65B0\u589E", 2340)] }),
                    new TableRow({ cantSplit: true, children: [C("\u4E2D\u592E\u9884\u7B97\u5185\u6295\u8D44", 3120), C("\u7EA67000\u4EBF\u5143", 2340), C(d.fiscalBudget.centralBudget + "\u4EBF\u5143", 2340), C("\u589E\u52A0", 2340)] }),
                ] }),
                P("\u6570\u636E\u6765\u6E90\uFF1A2026\u5E74\u653F\u5E9C\u5DE5\u4F5C\u62A5\u544A\u3001\u8D22\u653F\u90E8\u9884\u7B97\u8349\u6848\u62A5\u544A", { size: 20, italics: true, color: "666666", spacing: { before: 200 } }),
                P("2026\u5E74\u5173\u952E\u53D1\u73B0\uFF1A", { bold: true, spacing: { before: 200 } }),
                bullet("b", "\u8D22\u653F\u8D64\u5B57\u589E\u52A02300\u4EBF\u5143\uFF0C\u4F53\u73B0\u66F4\u52A0\u79EF\u6781\u7684\u8D22\u653F\u653F\u7B56\u53D6\u5411"),
                bullet("b", "\u65B0\u589E3000\u4EBF\u5143\u7279\u522B\u56FD\u503A\u652F\u6301\u56FD\u6709\u5927\u578B\u5546\u4E1A\u94F6\u884C\u8865\u5145\u8D44\u672C"),
                bullet("b", "\u4E2D\u592E\u9884\u7B97\u5185\u6295\u8D44\u589E\u81F3" + d.fiscalBudget.centralBudget + "\u4EBF\u5143\uFF0C\u91CD\u70B9\u652F\u6301\u65B0\u8D28\u751F\u4EA7\u529B\u3001\u65B0\u578B\u57CE\u9547\u5316"),
                bullet("b", "\u4E24\u65B0\u8D44\u91D1\u7B2C\u4E00\u6279\uFF1A625\u4EBF\u5143\uFF0C\u7B2C\u4E8C\u6279\uFF1A915\u4EBF\u5143\uFF0C\u7D2F\u8BA1" + d.twoNewPolicy.cumulative + "\u4EBF\u5143\uFF08\u5360\u5168\u5E74" + d.twoNewPolicy.progressRate + "%\uFF09"),

                // 2.2
                P("2.2 \u5730\u65B9\u4E13\u9879\u503A\u6D41\u5411\u5206\u6790", { heading: HeadingLevel.HEADING_2 }),
                P("2026\u5E74\u5730\u65B9\u653F\u5E9C\u4E13\u9879\u503A\u5B89\u6392" + d.fiscalBudget.specialBond + "\u4EBF\u5143\uFF0C\u5B8C\u5584\u4E13\u9879\u503A\u5238\u9879\u76EE\u8D1F\u9762\u6E05\u5355\u7BA1\u7406\u548C\u81EA\u5BA1\u81EA\u53D1\u8BD5\u70B9\uFF0C\u91CD\u70B9\u652F\u6301\u5EFA\u8BBE\u91CD\u5927\u9879\u76EE\u3001\u7F6E\u6362\u9690\u6027\u503A\u52A1\u3001\u6D88\u5316\u653F\u5E9C\u62D6\u6B20\u8D26\u6B3E\u7B49\u3002", { spacing: { before: 200, after: 200 } }),

                P("2.2.1 \u4E13\u9879\u503A\u65B0\u653F\u4EAE\u70B9", { heading: HeadingLevel.HEADING_3 }),
                num("n", "\u6295\u5411\u9886\u57DF\u8D1F\u9762\u6E05\u5355\u7BA1\u7406\uFF1A\u672A\u7EB3\u5165\u8D1F\u9762\u6E05\u5355\u7684\u9879\u76EE\u5747\u53EF\u7533\u8BF7\u4E13\u9879\u503A\u8D44\u91D1\uFF0C\u5927\u5E45\u62D3\u5BBD\u6295\u5411\u8303\u56F4"),
                num("n", "\u8D44\u672C\u91D1\u8303\u56F4\u6B63\u9762\u6E05\u5355\u6269\u5BB9\uFF1A\u5C06\u4FE1\u606F\u6280\u672F\u3001\u65B0\u6750\u6599\u3001\u751F\u7269\u5236\u9020\u3001\u6570\u5B57\u7ECF\u6D4E\u3001\u4F4E\u7A7A\u7ECF\u6D4E\u3001\u91CF\u5B50\u79D1\u6280\u3001\u751F\u547D\u79D1\u5B66\u3001\u5546\u4E1A\u822A\u5929\u3001\u5317\u6597\u7B49\u65B0\u5174\u4EA7\u4E1A\u57FA\u7840\u8BBE\u65BD\u7EB3\u5165"),
                num("n", "\u8D44\u672C\u91D1\u6BD4\u4F8B\u4E0A\u9650\u63D0\u9AD8\uFF1A\u753125%\u63D0\u9AD8\u81F330%"),
                num("n", "\u81EA\u5BA1\u81EA\u53D1\u8BD5\u70B9\uFF1A\u6C5F\u82CF\u3001\u5C71\u4E1C\u3001\u5E7F\u4E1C\u3001\u56DB\u5DDD\u3001\u5317\u4EAC\u7B4910\u7701\u4EFD\u53CA\u96C4\u5B89\u65B0\u533A\u8BD5\u70B9"),

                // 2.3
                P("2.3 \u4E00\u7EA7\u5E02\u573A\u878D\u8D44\u4FE1\u53F7", { heading: HeadingLevel.HEADING_2 }),
                P("2026\u5E74\u4E00\u7EA7\u5E02\u573A\u878D\u8D44\u5448\u73B0\u7206\u53D1\u5F0F\u589E\u957F\uFF0CAI\u9886\u57DF\u6295\u8D44\u521B\u5386\u53F2\u65B0\u9AD8\uFF0C\u786C\u79D1\u6280\u8D5B\u9053\u6301\u7EED\u706B\u70ED\u3002", { spacing: { before: 200, after: 200 } }),

                P("2.3.1 2026\u5E74Q1\u4E2D\u56FD\u521B\u6295\u5E02\u573A", { heading: HeadingLevel.HEADING_3 }),
                new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: [3120, 2340, 2340, 2340], rows: [
                    new TableRow({ cantSplit: true, children: [C("\u6307\u6807", 3120, { bold: true, fill: "D5E8F0" }), C("2026\u5E74Q1", 2340, { bold: true, fill: "D5E8F0" }), C("\u73AF\u6BD4\u589E\u957F", 2340, { bold: true, fill: "D5E8F0" }), C("\u540C\u6BD4\u589E\u957F", 2340, { bold: true, fill: "D5E8F0" })] }),
                    new TableRow({ cantSplit: true, children: [C("\u878D\u8D44\u4E8B\u4EF6\u6570", 3120), C(d.investment.q1_2026_events + "\u8D77", 2340), C("+2.5%", 2340), C("+" + d.investment.yoy_events + "%", 2340)] }),
                    new TableRow({ cantSplit: true, children: [C("\u62AB\u9732\u878D\u8D44\u603B\u989D", 3120), C(d.investment.q1_2026_amount + "\u4EBF\u5143", 2340), C("+11.4%", 2340), C("+" + d.investment.yoy_amount + "%", 2340)] }),
                    new TableRow({ cantSplit: true, children: [C("\u5355\u7B14\u5E73\u5747\u91D1\u989D", 3120), C("0.89\u4EBF\u5143", 2340), C("+0.15\u4EBF\u5143", 2340), C("-", 2340)] }),
                ] }),

                P("2.3.2 \u5168\u7403AI\u6295\u8D44\u7206\u53D1", { heading: HeadingLevel.HEADING_3 }),
                P("\u6839\u636E PitchBook \u6570\u636E\uFF0C2026\u5E74Q1\u5168\u7403AI\u6295\u8D44\u8FBE\u5230\u5386\u53F2\u6027\u65B0\u9AD8\uFF1A", { spacing: { before: 200 } }),
                bullet("b", "\u5168\u7403AI\u6295\u8D44\uFF1A" + d.investment.aiGlobalQ1 + "\u4EBF\u7F8E\u5143\uFF08\u4EC5Q1\uFF09\uFF0C\u8D85\u8FC72025\u5E74\u5168\u5E742544\u4EBF\u7F8E\u5143"),
                bullet("b", "\u6C34\u5E73\u5E73\u53F0\uFF1A1970\u4EBF\u7F8E\u5143\uFF08396\u7B14\u4EA4\u6613\uFF09\uFF0C\u8D85\u8FC72025\u5E74\u5168\u5E741100\u4EBF\u7F8E\u5143"),
                bullet("b", "\u534A\u5BFC\u4F53\u6295\u8D44\uFF1A50\u4EBF\u7F8E\u5143\uFF0884\u7B14\u4EA4\u6613\uFF09\uFF0C\u5386\u53F2\u7B2C\u4E8C\u9AD8\u6C34\u5E73"),
                bullet("b", "\u81EA\u4E3B\u673A\u5668\uFF1A290\u4EBF\u7F8E\u5143\uFF08118\u7B14\u4EA4\u6613\uFF09\uFF0C\u662F2025\u5E74Q4\u76843\u500D\u4EE5\u4E0A"),

                P("2.3.3 \u786C\u79D1\u6280\u6295\u8D44\u70ED\u70B9\u8FC1\u79FB", { heading: HeadingLevel.HEADING_3 }),
                num("n", "2023\u5E74\uFF1AAI\u82AF\u7247\uFF08\u786C\u4EF6\u57FA\u7840\uFF09"),
                num("n", "2024\u5E74\uFF1A\u5DE5\u4E1AAI\u4E0E\u5927\u6A21\u578B\u7206\u53D1"),
                num("n", "2025\u5E74\uFF1A\u673A\u5668\u4EBA\u548C\u5DE5\u4E1AAI\u6210\u4E3A\u7EDD\u5BF9\u70ED\u70B9"),
                num("n", "2026\u5E74\uFF1A\u5177\u8EAB\u667A\u80FD\u3001AI Agent\u3001\u884C\u4E1A\u5927\u6A21\u578B\u5E94\u7528\u5168\u9762\u7206\u53D1"),
                PB(),

                // ====== 三、看试点 ======
                P("\u4E09\u3001\u770B\u8BD5\u70B9\uFF1A\u653F\u7B56\u5148\u884C\u533A\u4E0EKPI\u8FFD\u8E2A", { heading: HeadingLevel.HEADING_1 }),
                P("\u53D1\u6539\u59D4\u8BD5\u70B9\u9879\u76EE\u662F\u653F\u7B56\u843D\u5730\u524D\u7684\u201C\u8BD5\u9A8C\u7530\u201D\uFF0C\u8BD5\u70B9\u9879\u76EE\u7684KPI\u8FBE\u6210\u60C5\u51B5\u662F\u5224\u65AD\u884C\u4E1A\u662F\u5426\u8FDB\u5165\u62D0\u70B9\u7684\u5173\u952E\u6307\u6807\u3002", { spacing: { before: 200, after: 200 } }),

                P("3.1 \u4F4E\u7A7A\u7ECF\u6D4E\u8BD5\u70B9\u8FDB\u5C55", { heading: HeadingLevel.HEADING_2 }),
                P("2026\u5E74\u4F4E\u7A7A\u7ECF\u6D4E\u8FDB\u5165\u5546\u4E1A\u5316\u5143\u5E74\uFF0C\u591A\u9879\u91CC\u7A0B\u7891\u4E8B\u4EF6\u8FBE\u6210\uFF1A", { spacing: { before: 200 } }),
                bullet("b", "\u5E02\u573A\u89C4\u6A21\uFF1A2026\u5E74\u7A81\u7834" + (d.lowAltitude.marketSize/10000).toFixed(1) + "\u4E07\u4EBF\u5143\uFF0C2035\u5E74\u9884\u8BA1\u8FBE" + (d.lowAltitude.marketSize2035/10000).toFixed(1) + "\u4E07\u4EBF\u5143"),
                bullet("b", "eVTOL\u89C4\u6A21\uFF1A\u4ECE2023\u5E749.8\u4EBF\u5143\u6FC0\u589E\u81F3" + d.lowAltitude.evtvolMarket + "\u4EBF\u5143"),
                bullet("b", "\u5546\u4E1A\u5316\u7A81\u7834\uFF1A\u4EBF\u822AEEH216-S\u83B7\u5168\u7403\u9996\u5F20eVTOL\u9002\u822A\u8BC1\uFF0C\u5E7F\u5DDE\u3001\u5408\u80A5\u5F00\u901A\u5546\u4E1A\u5316\u552E\u7968\u822A\u7EBF\uFF0C\u7968\u4EF7299\u5143\u8D77"),
                bullet("b", "\u8BA2\u5355\u7206\u53D1\uFF1A\u5C0F\u9E4F\u6C47\u5929\u98DE\u884C\u6C7D\u8F66\u8BA2\u5355\u8D857000\u53F0\uFF0C\u56DB\u5B63\u5EA6\u542F\u52A8\u4EA4\u4ED8"),
                bullet("b", "\u65E0\u4EBA\u673A\u5B9E\u540D\u767B\u8BB0\uFF1A\u7A81\u7834" + d.lowAltitude.uavRegistration + "\u4E07\u67B6\uFF08\u8D85\u9884\u671F\uFF09"),
                bullet("b", "\u4F01\u4E1A\u6570\u91CF\uFF1A" + (d.lowAltitude.companies/10000).toFixed(1) + "\u4E07\u5BB6\uFF0C2025\u5E74\u65B0\u589E4.9\u4E07\u5BB6\uFF08+142%\uFF09"),
                bullet("b", "\u8BD5\u70B9\u57CE\u5E02\uFF1A" + (d.lowAltitude.cities || []).join("\u3001")),

                P("3.2 \u78B3\u8FBE\u5CF0\u4E0E\u78B3\u5E02\u573A\u8BD5\u70B9", { heading: HeadingLevel.HEADING_2 }),
                P("2026\u5E74\u4E2D\u56FD\u78B3\u5E02\u573A\u8FDB\u51652.0\u65F6\u4EE3\uFF0C\u6269\u56F4+\u91D1\u878D\u5316\u53CC\u8F6E\u9A71\u52A8\uFF1A", { spacing: { before: 200 } }),
                bullet("b", "\u7B2C\u4E8C\u6279\u78B3\u8FBE\u5CF0\u8BD5\u70B9\uFF1A" + d.trials.carbonPeak.secondBatch + "\u4E2A\u57CE\u5E02\u548C\u56ED\u533A"),
                bullet("b", "\u96F6\u78B3\u56ED\u533A\uFF1A\u5341\u4E94\u4E94\u5C06\u5EFA" + d.trials.carbonPeak.zeroCarbonParks + "\u4E2A\u5DE6\u53F3\u56FD\u5BB6\u7EA7\u96F6\u78B3\u56ED\u533A"),
                bullet("b", "\u8986\u76D6\u6269\u56F4\uFF1A\u4ECE\u7535\u529B\u884C\u4E1A\u6269\u81F3\u94A2\u94C1\u3001\u6C34\u6CE5\u3001\u94DD\u51B6\u70BC\u7B49\u516B\u5927\u884C\u4E1A\uFF0C2027\u5E74\u57FA\u672C\u8986\u76D6\u5DE5\u4E1A\u9886\u57DF\u4E3B\u8981\u6392\u653E\u884C\u4E1A"),
                bullet("b", "\u91D1\u878D\u673A\u6784\u5165\u573A\uFF1A" + (d.trials.carbonMarket ? d.trials.carbonMarket.securitiesCompanies : 26) + "\u5BB6\u5238\u5546\u83B7\u6279\u53C2\u4E0E\u78B3\u6392\u653E\u6743\u4EA4\u6613"),
                bullet("b", "\u78B3\u671F\u8D27\u7B79\u5907\uFF1A\u5E7F\u5DDE\u671F\u8D27\u4EA4\u6613\u6240\u57FA\u672C\u5B8C\u6210\u78B3\u6392\u653E\u6743\u671F\u8D27\u5408\u7EA6\u5236\u5EA6\u8BBE\u8BA1"),
                bullet("b", "\u8003\u6838\u673A\u5236\uFF1A2026\u5E744\u6708\u5370\u53D1\u300A\u78B3\u8FBE\u5CF0\u78B3\u4E2D\u548C\u7EFC\u5408\u8BC4\u4EF7\u8003\u6838\u529E\u6CD5\u300B"),

                P("3.3 KPI\u8FBE\u6210\u60C5\u51B5\u4E0E\u884C\u4E1A\u62D0\u70B9", { heading: HeadingLevel.HEADING_2 }),
                P("\u8BD5\u70B9\u9879\u76EE\u7684KPI\u8FBE\u6210\u60C5\u51B5\u662F\u5224\u65AD\u884C\u4E1A\u662F\u5426\u8FDB\u5165\u62D0\u70B9\u7684\u5173\u952E\u6307\u6807\u3002\u8D85\u9884\u671F\u5B8C\u6210\u5F80\u5F80\u9884\u793A\u7740\u884C\u4E1A\u5373\u5C06\u8FDB\u5165\u5FEB\u901F\u6210\u957F\u671F\u3002", { spacing: { before: 200, after: 200 } }),
                new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: [2340, 2340, 2340, 2340], rows: [
                    new TableRow({ cantSplit: true, children: [C("\u6307\u6807", 2340, { bold: true, fill: "D5E8F0" }), C("2025\u5E74", 2340, { bold: true, fill: "D5E8F0" }), C("2026\u5E74\u76EE\u6807", 2340, { bold: true, fill: "D5E8F0" }), C("\u8FBE\u6210\u60C5\u51B5", 2340, { bold: true, fill: "D5E8F0" })] }),
                    new TableRow({ cantSplit: true, children: [C("\u4F4E\u7A7A\u5E02\u573A\u89C4\u6A21", 2340), C("\u7EA67000\u4EBF\u5143", 2340), C("\u7A81\u78341\u4E07\u4EBF\u5143", 2340), C("\u5DF2\u8FBE\u6210", 2340)] }),
                    new TableRow({ cantSplit: true, children: [C("\u65E0\u4EBA\u673A\u767B\u8BB0", 2340), C("\u7EA6800\u4E07\u67B6", 2340), C("1000\u4E07\u67B6", 2340), C("\u7A81\u78341200\u4E07\u67B6\uFF08\u8D85\u9884\u671F\uFF09", 2340)] }),
                    new TableRow({ cantSplit: true, children: [C("\u4E24\u65B0\u8D44\u91D1\u4E0B\u8FBE", 2340), C("-", 2340), C(d.twoNewPolicy.totalBudget + "\u4EBF\u5143", 2340), C(d.twoNewPolicy.cumulative + "\u4EBF\u5143\uFF08" + d.twoNewPolicy.progressRate + "%\uFF09", 2340)] }),
                    new TableRow({ cantSplit: true, children: [C("\u8BBE\u5907\u6295\u8D44\u589E\u901F", 2340), C("+15.7%", 2340), C("\u6301\u7EED\u589E\u957F", 2340), C("Q1\u540C\u6BD4+13.9%", 2340)] }),
                ] }),
                PB(),

                // ====== 四、重点领域 ======
                P("\u56DB\u3001\u91CD\u70B9\u9886\u57DF\u653F\u7B56\u7EA2\u5229\u6DF1\u5EA6\u5206\u6790", { heading: HeadingLevel.HEADING_1 }),
                P("\u57FA\u4E8E\u770B\u94B1\u4E0E\u770B\u8BD5\u70B9\u65B9\u6CD5\u8BBA\uFF0C\u6211\u4EEC\u5BF9\u56DB\u5927\u91CD\u70B9\u9886\u57DF\u8FDB\u884C\u6DF1\u5EA6\u5206\u6790\uFF0C\u8BC6\u522B\u653F\u7B56\u7EA2\u5229\u673A\u4F1A\u3002", { spacing: { before: 200, after: 200 } }),

                // 4.1
                P("4.1 \u79D1\u6280\u521B\u65B0/AI/\u534A\u5BFC\u4F53", { heading: HeadingLevel.HEADING_2 }),
                P("\u8D44\u91D1\u652F\u6301\uFF1A", { bold: true, spacing: { before: 200 } }),
                bullet("b", "\u4E2D\u592E\u9884\u7B97\u5185\u6295\u8D44" + d.fiscalBudget.centralBudget + "\u4EBF\u5143\uFF0C\u805A\u7126\u65B0\u8D28\u751F\u4EA7\u529B\u3001\u65B0\u578B\u57CE\u9547\u5316"),
                bullet("b", "\u8D85\u957F\u671F\u7279\u522B\u56FD\u503A8000\u4EBF\u5143\u652F\u6301\u4E24\u91CD\u5EFA\u8BBE"),
                bullet("b", "\u79D1\u6280\u5DE8\u5934\u8D44\u672C\u5F00\u652F\uFF1A2026\u5E74\u5168\u7403\u79D1\u6280\u5DE8\u5934AI\u6295\u5165\u9884\u8BA1\u8FBE8000\u4EBF\u7F8E\u5143\uFF08\u540C\u6BD4+80%\uFF09"),
                P("\u4EA7\u4E1A\u8FDB\u5C55\uFF1A", { bold: true, spacing: { before: 200 } }),
                bullet("b", "\u56FD\u4EA7\u5927\u6A21\u578B\uFF1A\u9636\u8DC3\u661F\u8FB0\u5B8C\u621050\u4EBF\u5143B+\u8F6E\u878D\u8D44\uFF0C\u521B\u8FC7\u53BB12\u4E2A\u6708\u56FD\u5185\u5927\u6A21\u578B\u8D5B\u9053\u5355\u7B14\u878D\u8D44\u6700\u9AD8\u7EAA\u5F55"),
                bullet("b", "OpenAI\uFF1A\u5B8C\u62101100\u4EBF\u7F8E\u5143\u878D\u8D44\uFF0C\u4F30\u503C7300\u4EBF\u7F8E\u5143"),
                bullet("b", "\u534A\u5BFC\u4F53\uFF1AQ1\u903E80\u5BB6\u534A\u5BFC\u4F53\u521D\u521B\u4F01\u4E1A\u5408\u8BA1\u878D\u8D44\u8D8580\u4EBF\u7F8E\u5143"),
                P("\u4FE1\u53F7\u5F3A\u5EA6\uFF1A\u2605\u2605\u2605\u2605\u2605\uFF08\u5F3A\u70C8\u63A8\u8350\uFF09", { bold: true, color: "FF0000", spacing: { before: 200 } }),
                P("AI\u6295\u8D44\u521B\u5386\u53F2\u65B0\u9AD8\uFF0C\u653F\u7B56\u4E0E\u8D44\u672C\u53CC\u8F6E\u9A71\u52A8\uFF0C\u5EFA\u8BAE\u91CD\u70B9\u5173\u6CE8\uFF1A", { spacing: { before: 100 } }),
                num("n", "AI\u82AF\u7247\u53CA\u7B97\u529B\u57FA\u7840\u8BBE\u65BD"),
                num("n", "AI Agent\u4E0E\u884C\u4E1A\u5927\u6A21\u578B\u5E94\u7528"),
                num("n", "\u5177\u8EAB\u667A\u80FD/\u4EBA\u5F62\u673A\u5668\u4EBA\u4EA7\u4E1A\u94FE"),
                num("n", "\u534A\u5BFC\u4F53\u8BBE\u5907\u4E0EChiplet\u6280\u672F"),

                // 4.2
                P("4.2 \u65B0\u80FD\u6E90/\u53CC\u78B3/\u73AF\u4FDD", { heading: HeadingLevel.HEADING_2 }),
                P("\u5341\u4E94\u4E94\u76EE\u6807\uFF1A", { bold: true, spacing: { before: 200 } }),
                bullet("b", "\u78B3\u6392\u653E\u5F3A\u5EA6\u8F832005\u5E74\u4E0B\u964D65%\u4EE5\u4E0A"),
                bullet("b", "\u975E\u5316\u77F3\u80FD\u6E90\u6D88\u8D39\u5360\u6BD4\u8FBE25%"),
                bullet("b", "\u7164\u70AD\u3001\u77F3\u6CB9\u6D88\u8D392030\u5E74\u524D\u8FBE\u5CF0"),
                bullet("b", "\u5EFA\u6210" + d.trials.carbonPeak.zeroCarbonParks + "\u4E2A\u5DE6\u53F3\u56FD\u5BB6\u7EA7\u96F6\u78B3\u56ED\u533A"),
                P("\u4FE1\u53F7\u5F3A\u5EA6\uFF1A\u2605\u2605\u2605\u2605", { bold: true, color: "FF6600", spacing: { before: 200 } }),
                P("\u53CC\u78B3\u76EE\u6807\u8FDB\u5165\u653B\u575A\u671F\uFF0C\u5EFA\u8BAE\u5173\u6CE8\u7ED3\u6784\u6027\u673A\u4F1A\uFF1A", { spacing: { before: 100 } }),
                num("n", "\u96F6\u78B3\u56ED\u533A\u5EFA\u8BBE\u76F8\u5173\uFF08\u5206\u5E03\u5F0F\u80FD\u6E90\u3001\u5FAE\u7535\u7F51\u3001\u50A8\u80FD\uFF09"),
                num("n", "\u78B3\u6355\u96C6\u4E0E\u5C01\u5B58\uFF08CCUS\uFF09"),
                num("n", "\u78B3\u8D44\u4EA7\u7BA1\u7406\u4E0E\u78B3\u91D1\u878D\u670D\u52A1"),
                num("n", "\u98CE\u5149\u8BBE\u5907\u5FAA\u73AF\u5229\u7528\uFF08\u9000\u5F79\u6F6E\u6765\u4E34\uFF09"),

                // 4.3
                P("4.3 \u57FA\u5EFA/\u65B0\u57FA\u5EFA", { heading: HeadingLevel.HEADING_2 }),
                P("\u8D44\u91D1\u652F\u6301\uFF1A", { bold: true, spacing: { before: 200 } }),
                bullet("b", "\u4E13\u9879\u503A" + d.fiscalBudget.specialBond + "\u4EBF\u5143\uFF0C\u652F\u6301\u91CD\u5927\u9879\u76EE\u3001\u7F6E\u6362\u9690\u6027\u503A\u52A1"),
                bullet("b", "\u8D85\u957F\u671F\u7279\u522B\u56FD\u503A8000\u4EBF\u5143\u652F\u6301\u4E24\u91CD\u5EFA\u8BBE"),
                bullet("b", "\u4F4E\u7A7A\u7ECF\u6D4E\u57FA\u7840\u8BBE\u65BD\u7EB3\u5165\u4E13\u9879\u503A\u8D44\u672C\u91D1\u8303\u56F4"),
                P("\u4FE1\u53F7\u5F3A\u5EA6\uFF1A\u2605\u2605\u2605\u2605\u2605\uFF08\u5F3A\u70C8\u63A8\u8350\uFF09", { bold: true, color: "FF0000", spacing: { before: 200 } }),
                P("\u4F4E\u7A7A\u7ECF\u6D4E\u8FDB\u5165\u5546\u4E1A\u5316\u5143\u5E74\uFF0C\u57FA\u5EFA\u5148\u884C\uFF1A", { spacing: { before: 100 } }),
                num("n", "\u4F4E\u7A7A\u8D77\u964D\u70B9/vertiport\u5EFA\u8BBE"),
                num("n", "\u4F4E\u7A7A\u667A\u8054\u7F51/\u7A7A\u7BA1\u7CFB\u7EDF"),
                num("n", "\u7B97\u529B\u57FA\u7840\u8BBE\u65BD\uFF08\u667A\u7B97\u4E2D\u5FC3\uFF09"),
                num("n", "\u57CE\u5E02\u66F4\u65B0\u4E0E\u5730\u4E0B\u7BA1\u7F51"),

                // 4.4
                P("4.4 \u6D88\u8D39/\u533B\u7597/\u6559\u80B2", { heading: HeadingLevel.HEADING_2 }),
                P("2026\u5E74\u4E24\u65B0\u653F\u7B56\u52A0\u529B\u6269\u56F4\uFF1A", { bold: true, spacing: { before: 200 } }),
                bullet("b", "\u8D44\u91D1\u4E0B\u8FBE\uFF1A\u7D2F\u8BA1" + d.twoNewPolicy.cumulative + "\u4EBF\u5143\u5DF2\u4E0B\u8FBE\uFF0C\u5360\u5168\u5E74" + d.twoNewPolicy.totalBudget + "\u4EBF\u5143\u7684" + d.twoNewPolicy.progressRate + "%"),
                bullet("b", "\u6C7D\u8F66\u8865\u8D34\uFF1A\u65B0\u80FD\u6E90\u8865\u8D34\u8F66\u4EF712%\uFF08\u6700\u9AD82\u4E07\u5143\uFF09\uFF0C\u71C3\u6CB9\u8865\u8D3410%\uFF08\u6700\u9AD81.5\u4E07\u5143\uFF09"),
                bullet("b", "\u5BB6\u7535\u8865\u8D34\uFF1A1\u7EA7\u80FD\u6548\u4EA7\u54C1\u552E\u4EF7\u768415%\uFF0C\u5355\u4EF6\u6700\u9AD81500\u5143"),
                bullet("b", "\u6570\u7801\u8865\u8D34\uFF1A\u624B\u673A\u3001\u5E73\u677F\u3001\u667A\u80FD\u624B\u8868\u3001\u667A\u80FD\u773C\u955C\u3001\u667A\u80FD\u5BB6\u5C45\uFF0C\u8865\u8D3415%"),
                bullet("b", "\u8BBE\u5907\u66F4\u65B0\u6269\u56F4\uFF1A\u65B0\u589E\u8001\u65E7\u5C0F\u533A\u52A0\u88C5\u7535\u68AF\u3001\u517B\u8001\u673A\u6784\u3001\u6D88\u9632\u6551\u63F4\u3001\u68C0\u9A8C\u68C0\u6D4B\u3001\u5546\u4E1A\u7EFC\u5408\u4F53"),
                P("\u4FE1\u53F7\u5F3A\u5EA6\uFF1A\u2605\u2605\u2605\u2605\u2605\uFF08\u5F3A\u70C8\u63A8\u8350\uFF09", { bold: true, color: "FF0000", spacing: { before: 200 } }),
                P("\u4E24\u65B0\u653F\u7B56\u6548\u679C\u663E\u8457\uFF0C2026\u5E74\u52A0\u529B\u6269\u56F4\uFF0C\u5EFA\u8BAE\u91CD\u70B9\u5173\u6CE8\uFF1A", { spacing: { before: 100 } }),
                num("n", "\u65B0\u80FD\u6E90\u6C7D\u8F66\u4EA7\u4E1A\u94FE\uFF08\u6E17\u900F\u7387\u8D8550%\uFF09"),
                num("n", "\u667A\u80FD\u5BB6\u7535\u4E0E\u6570\u7801\u4EA7\u54C1"),
                num("n", "\u533B\u7597\u8BBE\u5907\u66F4\u65B0\uFF08\u5F71\u50CF\u3001\u68C0\u9A8C\u8BBE\u5907\uFF09"),
                num("n", "\u6559\u80B2\u4FE1\u606F\u5316\u4E0E\u517B\u8001\u8BBE\u5907"),
                PB(),

                // ====== 五、实操指南 ======
                P("\u4E94\u3001\u5B9E\u64CD\u6307\u5357\uFF1A30\u5206\u949F\u6708\u5EA6\u8DDF\u8E2A\u6CD5", { heading: HeadingLevel.HEADING_1 }),
                P("\u5EFA\u7ACB\u7CFB\u7EDF\u5316\u7684\u4FE1\u606F\u8FFD\u8E2A\u673A\u5236\u662F\u6293\u53D6\u653F\u7B56\u7EA2\u5229\u7684\u5173\u952E\u3002\u4EE5\u4E0B\u662F\u7ECF\u8FC7\u9A8C\u8BC1\u7684\u5B9E\u64CD\u65B9\u6CD5\uFF1A", { spacing: { before: 200, after: 200 } }),

                P("5.1 \u6708\u5EA6\u8DDF\u8E2A\u6E05\u5355\uFF0830\u5206\u949F\uFF09", { heading: HeadingLevel.HEADING_2 }),
                new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: [1560, 3120, 2340, 2340], rows: [
                    new TableRow({ cantSplit: true, children: [C("\u65F6\u95F4", 1560, { bold: true, fill: "D5E8F0" }), C("\u4EFB\u52A1", 3120, { bold: true, fill: "D5E8F0" }), C("\u6570\u636E\u6765\u6E90", 2340, { bold: true, fill: "D5E8F0" }), C("\u5173\u6CE8\u8981\u70B9", 2340, { bold: true, fill: "D5E8F0" })] }),
                    new TableRow({ cantSplit: true, children: [C("10\u5206\u949F", 1560), C("\u4E13\u9879\u503A\u6D41\u5411\u8FFD\u8E2A", 3120), C("\u4F01\u4E1A\u9884\u8B66\u901A\u3001Wind", 2340), C("\u6295\u5411\u9886\u57DF\u53D8\u5316\u3001\u89C4\u6A21\u589E\u957F", 2340)] }),
                    new TableRow({ cantSplit: true, children: [C("10\u5206\u949F", 1560), C("\u8BD5\u70B9\u9879\u76EE\u8FDB\u5C55", 3120), C("\u53D1\u6539\u59D4\u5B98\u7F51\u3001\u5730\u65B9\u516C\u544A", 2340), C("KPI\u8FBE\u6210\u60C5\u51B5\u3001\u8D85\u9884\u671F\u4FE1\u53F7", 2340)] }),
                    new TableRow({ cantSplit: true, children: [C("10\u5206\u949F", 1560), C("\u878D\u8D44\u6570\u636E\u626B\u63CF", 3120), C("IT\u6854\u5B50\u3001PitchBook", 2340), C("\u8FDE\u7EED\u5B63\u5EA6\u589E\u957F\u3001\u70ED\u70B9\u8FC1\u79FB", 2340)] }),
                ] }),

                P("5.2 2026\u5E74\u91CD\u70B9\u8DDF\u8E2A\u6307\u6807", { heading: HeadingLevel.HEADING_2 }),
                num("n", "\u4E24\u65B0\u8D44\u91D1\u4E0B\u8FBE\u8FDB\u5EA6\uFF1A\u5168\u5E74" + d.twoNewPolicy.totalBudget + "\u4EBF\u5143\uFF0C\u5DF2\u4E0B\u8FBE" + d.twoNewPolicy.cumulative + "\u4EBF\u5143\uFF08" + d.twoNewPolicy.progressRate + "%\uFF09"),
                num("n", "\u4F4E\u7A7A\u7ECF\u6D4E\u5546\u4E1A\u5316\u8FDB\u5C55\uFF1AeVTOL\u9002\u822A\u53D6\u8BC1\u3001\u822A\u7EBF\u5F00\u901A\u3001\u7968\u4EF7\u53D8\u5316"),
                num("n", "AI\u6295\u8D44\u70ED\u5EA6\uFF1A\u5B63\u5EA6\u878D\u8D44\u89C4\u6A21\u3001\u72EC\u89D2\u517D\u751F\u6570\u91CF"),
                num("n", "\u78B3\u5E02\u573A\u6269\u56F4\uFF1A\u94A2\u94C1\u884C\u4E1A\u7EB3\u5165\u65F6\u95F4\u3001\u78B3\u4EF7\u8D70\u52BF"),
                num("n", "\u96F6\u78B3\u56ED\u533A\u5EFA\u8BBE\uFF1A\u9996\u6279\u96F6\u78B3\u56ED\u533A\u540D\u5355\u3001\u6295\u8D44\u89C4\u6A21"),

                P("5.3 \u5173\u952E\u4FE1\u53F7\u8BC6\u522B", { heading: HeadingLevel.HEADING_2 }),
                P("\u4EE5\u4E0B\u4FE1\u53F7\u51FA\u73B0\u65F6\uFF0C\u901A\u5E38\u9884\u793A\u7740\u653F\u7B56\u7EA2\u5229\u7A97\u53E3\u671F\uFF1A", { spacing: { before: 200 } }),
                new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: [2340, 3510, 3510], rows: [
                    new TableRow({ cantSplit: true, children: [C("\u4FE1\u53F7\u7C7B\u578B", 2340, { bold: true, fill: "D5E8F0" }), C("\u5177\u4F53\u8868\u73B0", 3510, { bold: true, fill: "D5E8F0" }), C("\u6295\u8D44\u542B\u4E49", 3510, { bold: true, fill: "D5E8F0" })] }),
                    new TableRow({ cantSplit: true, children: [C("\u8D44\u91D1\u589E\u957F", 2340), C("\u8FDE\u7EED\u4E24\u5B63\u5EA6\u9884\u7B97/\u4E13\u9879\u503A\u589E\u957F\u8D8515%", 3510), C("\u653F\u7B56\u6267\u884C\u529B\u5EA6\u52A0\u5927\uFF0C\u884C\u4E1A\u666F\u6C14\u5EA6\u4E0A\u5347", 3510)] }),
                    new TableRow({ cantSplit: true, children: [C("\u8BD5\u70B9\u8D85\u9884\u671F", 2340), C("\u8BD5\u70B9\u9879\u76EEKPI\u8FBE\u6210\u7387\u8D85120%", 3510), C("\u6280\u672F/\u6A21\u5F0F\u9A8C\u8BC1\u6210\u529F\uFF0C\u5373\u5C06\u89C4\u6A21\u5316\u63A8\u5E7F", 3510)] }),
                    new TableRow({ cantSplit: true, children: [C("\u878D\u8D44\u6D3B\u8DC3", 2340), C("\u4E00\u7EA7\u5E02\u573A\u878D\u8D44\u8FDE\u7EED\u4E24\u5B63\u5EA6\u589E\u957F\u8D8530%", 3510), C("\u8D44\u672C\u8BA4\u53EF\u5EA6\u63D0\u5347\uFF0C\u4F30\u503C\u4E2D\u67A2\u4E0A\u79FB", 3510)] }),
                    new TableRow({ cantSplit: true, children: [C("\u5546\u4E1A\u5316\u7A81\u7834", 2340), C("\u8BD5\u70B9\u57CE\u5E02\u5F00\u901A\u5546\u4E1A\u5316\u8FD0\u8425\u3001\u5B9E\u73B0\u76C8\u5229", 3510), C("\u4ECE\u653F\u7B56\u9A71\u52A8\u8F6C\u5411\u4E1A\u7EE9\u5151\u73B0", 3510)] }),
                ] }),
                PB(),

                // ====== 六、风险提示 ======
                P("\u516D\u3001\u98CE\u9669\u63D0\u793A\u4E0E\u6295\u8D44\u5EFA\u8BAE", { heading: HeadingLevel.HEADING_1 }),

                P("6.1 \u4E3B\u8981\u98CE\u9669\u63D0\u793A", { heading: HeadingLevel.HEADING_2 }),
                P("1. \u653F\u7B56\u6267\u884C\u98CE\u9669", { bold: true, spacing: { before: 200 } }),
                P("\u9884\u7B97\u62E8\u6B3E\u548C\u4E13\u9879\u503A\u89C4\u6A21\u589E\u957F\u4E0D\u7B49\u4E8E\u5B9E\u9645\u6267\u884C\u6548\u679C\uFF0C\u9700\u5173\u6CE8\u8D44\u91D1\u5230\u4F4D\u7387\u548C\u9879\u76EE\u843D\u5730\u8FDB\u5EA6\u3002\u90E8\u5206\u5730\u533A\u5B58\u5728\u94B1\u7B49\u9879\u76EE\u6216\u9879\u76EE\u7B49\u94B1\u7684\u60C5\u51B5\u3002", { spacing: { before: 100 } }),
                P("2. \u6280\u672F\u8FED\u4EE3\u98CE\u9669", { bold: true, spacing: { before: 200 } }),
                P("AI\u3001\u4F4E\u7A7A\u7ECF\u6D4E\u7B49\u9886\u57DF\u6280\u672F\u8FED\u4EE3\u8FC5\u901F\uFF0C\u8BD5\u70B9\u9879\u76EE\u7684\u6280\u672F\u8DEF\u7EBF\u53EF\u80FD\u88AB\u98A0\u8986\u3002\u9700\u5173\u6CE8\u6280\u672F\u6210\u719F\u5EA6\u548C\u5546\u4E1A\u5316\u53EF\u884C\u6027\u3002", { spacing: { before: 100 } }),
                P("3. \u4F30\u503C\u6CE1\u6CAB\u98CE\u9669", { bold: true, spacing: { before: 200 } }),
                P("AI\u9886\u57DF\u6295\u8D44\u8FC7\u70ED\uFF0C2026\u5E74Q1\u5168\u7403AI\u6295\u8D44" + d.investment.aiGlobalQ1 + "\u4EBF\u7F8E\u5143\u5DF2\u8D85\u8FC72025\u5E74\u5168\u5E74\uFF0C\u9700\u8B66\u60D5\u4F30\u503C\u6CE1\u6CAB\u3002", { spacing: { before: 100 } }),
                P("4. \u5B8F\u89C2\u73AF\u5883\u98CE\u9669", { bold: true, spacing: { before: 200 } }),
                P("\u5730\u7F18\u653F\u6CBB\u3001\u8D38\u6613\u6469\u64E6\u3001\u7F8E\u8054\u50A8\u653F\u7B56\u7B49\u5916\u90E8\u56E0\u7D20\u53EF\u80FD\u5F71\u54CD\u653F\u7B56\u6267\u884C\u8282\u594F\u548C\u4EA7\u4E1A\u4F9B\u5E94\u94FE\u7A33\u5B9A\u6027\u3002", { spacing: { before: 100 } }),

                P("6.2 \u6295\u8D44\u5EFA\u8BAE", { heading: HeadingLevel.HEADING_2 }),
                P("\u57FA\u4E8E\u770B\u94B1\u4E0E\u770B\u8BD5\u70B9\u65B9\u6CD5\u8BBA\uFF0C\u7ED3\u5408" + (new Date().getFullYear()) + "\u5E74\u6700\u65B0\u6570\u636E\uFF0C\u5BF9\u5404\u9886\u57DF\u6295\u8D44\u5EFA\u8BAE\u5982\u4E0B\uFF1A", { spacing: { before: 200, after: 200 } }),

                new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: [2340, 1560, 2340, 3120], rows: [
                    new TableRow({ cantSplit: true, children: [C("\u9886\u57DF", 2340, { bold: true, fill: "D5E8F0" }), C("\u8BC4\u7EA7", 1560, { bold: true, fill: "D5E8F0" }), C("\u4FE1\u53F7\u5F3A\u5EA6", 2340, { bold: true, fill: "D5E8F0" }), C("\u6838\u5FC3\u903B\u8F91", 3120, { bold: true, fill: "D5E8F0" })] }),
                    ...Object.entries(d.sectors).map(([k, v]) =>
                        new TableRow({ cantSplit: true, children: [C(sectorName(k), 2340), C(v.rating, 1560), C(stars(v.signal), 2340), C(v.coreLogic, 3120)] })
                    )
                ] }),

                P("\u514D\u8D23\u58F0\u660E\uFF1A", { bold: true, spacing: { before: 400 } }),
                P("\u672C\u62A5\u544A\u4EC5\u4F9B\u53C2\u8003\uFF0C\u4E0D\u6784\u6210\u6295\u8D44\u5EFA\u8BAE\u3002\u6295\u8D44\u8005\u5E94\u6839\u636E\u81EA\u8EAB\u60C5\u51B5\u72EC\u7ACB\u5224\u65AD\uFF0C\u81EA\u884C\u627F\u62C5\u6295\u8D44\u98CE\u9669\u3002", { size: 20, color: "666666", spacing: { before: 100 } }),

                P("\u6570\u636E\u6765\u6E90\u8BF4\u660E\uFF1A", { bold: true, spacing: { before: 400 } }),
                bullet("b", "2026\u5E74\u653F\u5E9C\u5DE5\u4F5C\u62A5\u544A\u3001\u8D22\u653F\u90E8\u9884\u7B97\u8349\u6848\u62A5\u544A"),
                bullet("b", "\u56FD\u5BB6\u53D1\u6539\u59D4\u300A\u5173\u4E8E2026\u5E74\u5B9E\u65BD\u5927\u89C4\u6A21\u8BBE\u5907\u66F4\u65B0\u548C\u6D88\u8D39\u54C1\u4EE5\u65E7\u6362\u65B0\u653F\u7B56\u7684\u901A\u77E5\u300B"),
                bullet("b", "IT\u6854\u5B50\u300A2026\u5E74Q1\u521B\u6295\u62A5\u544A\u300B"),
                bullet("b", "PitchBook AI\u6295\u8D44\u6570\u636E"),
                bullet("b", "\u8D5B\u8FEA\u7814\u7A76\u9662\u4F4E\u7A7A\u7ECF\u6D4E\u6570\u636E"),
                bullet("b", "\u56FD\u52A1\u9662\u3001\u56FD\u5BB6\u53D1\u6539\u59D4\u5B98\u7F51\u516C\u5F00\u653F\u7B56\u6587\u4EF6"),
            ]
        }]
    });

    return await Packer.toBuffer(doc);
};
