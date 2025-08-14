document.addEventListener('DOMContentLoaded', () => {
    // 确保所有依赖库都已加载
    if (typeof docx === 'undefined' || typeof marked === 'undefined' || typeof saveAs === 'undefined') {
        alert("错误：一个或多个必需的库未能加载。请检查您的网络连接和文件路径。");
        return;
    }

    const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, convertInchesToTwip, convertMillimetersToTwip, LineRuleType, Numbering, Indent } = docx;

    // --- 1. 获取所有 DOM 元素 ---
    const mdInput = document.getElementById('md-input');
    const convertBtn = document.getElementById('convert-btn');
    const mdUpload = document.getElementById('md-upload');
    const autoNumberingCheckbox = document.getElementById('auto-numbering');
    const downloadFontsBtn = document.getElementById('download-fonts-btn');

    // --- 2. 加载默认/示例 Markdown ---
    // 恢复加载 README.md 以匹配旧版行为
    fetch('README.md')
        .then(res => {
            if (!res.ok) throw new Error(`无法加载默认文件 README.md, 状态: ${res.status}`);
            return res.text();
        })
        .then(text => {
            mdInput.value = text;
        })
        .catch(err => {
            console.warn(err);
            mdInput.placeholder = "无法加载 README.md。请确保文件存在，或直接粘贴内容/上传文件。";
        });

    // --- 3. 设置事件监听器 ---

    // 文件上传功能
    mdUpload.addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (file) {
            mdInput.value = await file.text();
        }
    });

    // 字体下载功能
    downloadFontsBtn.addEventListener('click', () => {
        // 使用旧脚本中的 URL
        const fontZipUrl = 'https://github.com/AngelSnow1129/md2docx/releases/download/V1.0.0/default.zip';
        window.open(fontZipUrl, '_blank');
        console.log(`正在尝试从以下地址下载字体包: ${fontZipUrl}`);
    });

    // 核心转换功能触发
    convertBtn.addEventListener('click', () => {
        const markdownText = mdInput.value;
        if (!markdownText.trim()) {
            alert('内容不能为空，请输入或上传 Markdown 文本！');
            return;
        }
        // 将复选框的状态传递给生成函数
        generateDocx(markdownText, autoNumberingCheckbox.checked);
    });

    // --- 4. 核心 DOCX 生成函数 (融合版) ---
    function generateDocx(mdText, enableAutoNumbering) {
        // 定义字体和字号 (pt)
        const FONT_FANGSONG_GB2312 = "仿宋_GB2312";
        const FONT_XIAOBIAOSONG = "方正小标宋简体";
        const FONT_HEITI = "黑体";
        const FONT_KAITI_GB2312 = "楷体_GB2312";

        const SIZE_H1 = 22 * 2; // 二号 (22pt)
        const SIZE_H_COMMON = 16 * 2; // 三号 (16pt)
        const SIZE_BODY = 16 * 2; // 三号 (16pt)

        // 定义自动编号规则
        const numbering = new Numbering({
            config: [{
                reference: "gongwen-numbering",
                levels: [
                    { level: 0, format: "chineseCounting", text: "%1、", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { firstLine: convertInchesToTwip(0.44) } }, run: { font: FONT_HEITI, size: SIZE_H_COMMON } } },
                    { level: 1, format: "chineseCounting", text: "（%2）", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { firstLine: convertInchesToTwip(0.44) } }, run: { font: FONT_KAITI_GB2312, size: SIZE_H_COMMON } } },
                    { level: 2, format: "decimal", text: "%3.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { firstLine: convertInchesToTwip(0.44) } }, run: { font: FONT_FANGSONG_GB2312, size: SIZE_H_COMMON, bold: true } } },
                    { level: 3, format: "decimal", text: "（%4）", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { firstLine: convertInchesToTwip(0.44) } }, run: { font: FONT_FANGSONG_GB2312, size: SIZE_H_COMMON } } },
                ],
            }],
        });

        const tokens = marked.lexer(mdText);
        const docChildren = [];

        tokens.forEach(token => {
            let paragraph;
            switch (token.type) {
                case 'heading':
                    const textWithoutFormatting = token.text;
                    let headingProperties = { text: textWithoutFormatting };

                    // 根据复选框状态决定是否应用编号
                    if (enableAutoNumbering) {
                        const levelMap = { 2: 0, 3: 1, 4: 2, 5: 3 };
                        if (levelMap[token.depth] !== undefined) {
                            headingProperties.numbering = { reference: "gongwen-numbering", level: levelMap[token.depth] };
                        }
                    }
                    
                    // 为标题应用基础样式
                    const styleMap = { 2: "h2", 3: "h3", 4: "h4", 5: "h5" };
                    if (styleMap[token.depth]) {
                        headingProperties.style = styleMap[token.depth];
                    }

                    if (token.depth === 1) { // 主标题特殊处理
                        paragraph = new Paragraph({
                            children: [new TextRun({ text: textWithoutFormatting, font: FONT_XIAOBIAOSONG, size: SIZE_H1 })],
                            alignment: AlignmentType.CENTER,
                            spacing: { line: 30 * 20, lineRule: LineRuleType.EXACT, after: 16 * 20 },
                        });
                    } else {
                        paragraph = new Paragraph(headingProperties);
                    }
                    break;

                // "扁平化"处理逻辑保持不变
                case 'paragraph':
                case 'list':
                case 'blockquote':
                case 'code':
                case 'hr':
                    const plainText = token.text || (token.tokens ? token.tokens.map(t => t.text).join(' ') : '');
                    if (plainText.trim()) {
                        paragraph = new Paragraph({
                            children: [new TextRun({ text: plainText, font: FONT_FANGSONG_GB2312, size: SIZE_BODY })],
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 640 }, // 32pt = 640 twips
                            spacing: { line: 28 * 20, lineRule: LineRuleType.EXACT },
                        });
                    }
                    break;

                default:
                    paragraph = null;
                    break;
            }

            if (paragraph) {
                docChildren.push(paragraph);
            }
        });

        // 创建文档
        const doc = new Document({
            numbering: numbering,
            styles: {
                paragraphStyles: [
                    { id: "h2", name: "Heading 2", run: { font: FONT_HEITI, size: SIZE_H_COMMON }, paragraph: { spacing: { line: 28 * 20, lineRule: LineRuleType.EXACT }, indent: { firstLine: convertInchesToTwip(0.44) } } },
                    { id: "h3", name: "Heading 3", run: { font: FONT_KAITI_GB2312, size: SIZE_H_COMMON }, paragraph: { spacing: { line: 28 * 20, lineRule: LineRuleType.EXACT }, indent: { firstLine: convertInchesToTwip(0.44) } } },
                    { id: "h4", name: "Heading 4", run: { font: FONT_FANGSONG_GB2312, size: SIZE_H_COMMON, bold: true }, paragraph: { spacing: { line: 28 * 20, lineRule: LineRuleType.EXACT }, indent: { firstLine: convertInchesToTwip(0.44) } } },
                    { id: "h5", name: "Heading 5", run: { font: FONT_FANGSONG_GB2312, size: SIZE_H_COMMON }, paragraph: { spacing: { line: 28 * 20, lineRule: LineRuleType.EXACT }, indent: { firstLine: convertInchesToTwip(0.44) } } },
                ]
            },
            sections: [{
                properties: {
                    page: {
                        size: { width: convertMillimetersToTwip(210), height: convertMillimetersToTwip(297) }, // A4
                        margin: {
                            top: convertMillimetersToTwip(37), bottom: convertMillimetersToTwip(35),
                            left: convertMillimetersToTwip(28), right: convertMillimetersToTwip(26),
                        },
                        headers: { default: { properties: { header: { margin: { top: convertMillimetersToTwip(15) } } } } },
                        footers: { default: { properties: { footer: { margin: { top: convertMillimetersToTwip(28) } } } } },
                    },
                },
                children: docChildren,
            }],
        });

        // 生成并下载
        Packer.toBlob(doc).then(blob => {
            const timestamp = new Date().toISOString().replace(/[-:.]/g, "").replace("T", "_").slice(0, 15);
            saveAs(blob, `公文_${timestamp}.docx`);
        }).catch(err => {
            console.error("生成 Word 文档时出错:", err);
            alert("生成失败，详情请查看浏览器控制台。");
        });
    }
});