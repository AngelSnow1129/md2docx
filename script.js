document.addEventListener('DOMContentLoaded', () => {
    // 确保所有依赖库都已加载
    if (typeof docx === 'undefined' || typeof marked === 'undefined' || typeof saveAs === 'undefined') {
        alert("错误：一个或多个必需的库未能加载。请检查您的网络连接和文件路径。");
        return;
    }

    const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, convertInchesToTwip, convertMillimetersToTwip, LineRuleType, Numbering, Indent } = docx;

    const mdInput = document.getElementById('md-input');
    const convertBtn = document.getElementById('convert-btn');

    // 加载示例文本
    fetch('sample.md')
        .then(res => {
            if (!res.ok) throw new Error(`无法加载示例文件 sample.md, 状态: ${res.status}`);
            return res.text();
        })
        .then(text => {
            mdInput.value = text;
        })
        .catch(err => {
            console.warn(err);
            mdInput.placeholder = "无法加载 sample.md。请确保文件存在，或直接在此处粘贴内容。";
        });

    convertBtn.addEventListener('click', () => {
        const markdownText = mdInput.value;
        if (!markdownText.trim()) {
            alert('内容不能为空，请输入或上传 Markdown 文本！');
            return;
        }
        generateDocx(markdownText);
    });

    function generateDocx(mdText) {
        // --- 1. 定义字体和字号 (pt) ---
        const FONT_FANGSONG_GB2312 = "仿宋_GB2312";
        const FONT_XIAOBIAOSONG = "方正小标宋简体";
        const FONT_HEITI = "黑体";
        const FONT_KAITI_GB2312 = "楷体_GB2312";

        const SIZE_H1 = 22 * 2; // 二号 (22pt)
        const SIZE_H_COMMON = 16 * 2; // 三号 (16pt)
        const SIZE_BODY = 16 * 2; // 三号 (16pt)

        // --- 2. 定义自动编号 ---
        const numbering = new Numbering({
            config: [{
                reference: "gongwen-numbering",
                levels: [
                    { // 一级标题: 一、
                        level: 0,
                        format: "chineseCounting",
                        text: "%1、",
                        alignment: AlignmentType.LEFT,
                        style: {
                            paragraph: {
                                indent: { left: convertInchesToTwip(0), firstLine: convertInchesToTwip(0.44) }, // 2字符缩进
                            },
                            run: {
                                font: FONT_HEITI,
                                size: SIZE_H_COMMON,
                            }
                        }
                    },
                    { // 二级标题: (一)
                        level: 1,
                        format: "chineseCounting",
                        text: "（%2）",
                        alignment: AlignmentType.LEFT,
                        style: {
                            paragraph: {
                                indent: { left: convertInchesToTwip(0), firstLine: convertInchesToTwip(0.44) }, // 2字符缩进
                            },
                            run: {
                                font: FONT_KAITI_GB2312,
                                size: SIZE_H_COMMON,
                            }
                        }
                    },
                    { // 三级标题: 1.
                        level: 2,
                        format: "decimal",
                        text: "%3.",
                        alignment: AlignmentType.LEFT,
                        style: {
                            paragraph: {
                                indent: { left: convertInchesToTwip(0), firstLine: convertInchesToTwip(0.44) }, // 2字符缩进
                            },
                            run: {
                                font: FONT_FANGSONG_GB2312,
                                size: SIZE_H_COMMON,
                                bold: true,
                            }
                        }
                    },
                    { // 四级标题: (1)
                        level: 3,
                        format: "decimal",
                        text: "（%4）",
                        alignment: AlignmentType.LEFT,
                        style: {
                            paragraph: {
                                indent: { left: convertInchesToTwip(0), firstLine: convertInchesToTwip(0.44) }, // 2字符缩进
                            },
                            run: {
                                font: FONT_FANGSONG_GB2312,
                                size: SIZE_H_COMMON,
                            }
                        }
                    },
                ],
            }],
        });

        // --- 3. 解析 Markdown ---
        const tokens = marked.lexer(mdText);
        const docChildren = [];

        tokens.forEach(token => {
            let paragraph;
            switch (token.type) {
                case 'heading':
                    const textWithoutFormatting = token.text;
                    switch (token.depth) {
                        case 1: // 主标题
                            paragraph = new Paragraph({
                                children: [new TextRun({ text: textWithoutFormatting, font: FONT_XIAOBIAOSONG, size: SIZE_H1 })],
                                alignment: AlignmentType.CENTER,
                                spacing: { line: 30 * 20, lineRule: LineRuleType.EXACT, after: 16 * 20 }, // 30磅行距, 段后空一行(16pt)
                            });
                            break;
                        case 2: // 一级标题
                            paragraph = new Paragraph({
                                text: textWithoutFormatting,
                                numbering: { reference: "gongwen-numbering", level: 0 },
                                style: "h2", // 使用预定义样式
                            });
                            break;
                        case 3: // 二级标题
                            paragraph = new Paragraph({
                                text: textWithoutFormatting,
                                numbering: { reference: "gongwen-numbering", level: 1 },
                                style: "h3",
                            });
                            break;
                        case 4: // 三级标题
                            paragraph = new Paragraph({
                                text: textWithoutFormatting,
                                numbering: { reference: "gongwen-numbering", level: 2 },
                                style: "h4",
                            });
                            break;
                        case 5: // 四级标题
                            paragraph = new Paragraph({
                                text: textWithoutFormatting,
                                numbering: { reference: "gongwen-numbering", level: 3 },
                                style: "h5",
                            });
                            break;
                    }
                    break;

                // "扁平化"处理：将所有其他类型的块级元素（段落、列表、引用等）统一视为正文
                case 'paragraph':
                case 'list':
                case 'blockquote':
                case 'code':
                case 'hr':
                    // 提取纯文本，对于列表等复杂结构，marked会提供一个.text属性
                    const plainText = token.text || (token.tokens ? token.tokens.map(t => t.text).join(' ') : '');
                    if (plainText.trim()) {
                        paragraph = new Paragraph({
                            children: [new TextRun({ text: plainText, font: FONT_FANGSONG_GB2312, size: SIZE_BODY })],
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 640 }, // 2个字符，32pt = 640 twips
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

        // --- 4. 创建文档 ---
        const doc = new Document({
            numbering: numbering,
            styles: {
                paragraphStyles: [
                    // 预定义样式以供编号系统引用
                    { id: "h2", name: "Heading 2", run: { font: FONT_HEITI, size: SIZE_H_COMMON }, paragraph: { spacing: { line: 28 * 20, lineRule: LineRuleType.EXACT } } },
                    { id: "h3", name: "Heading 3", run: { font: FONT_KAITI_GB2312, size: SIZE_H_COMMON }, paragraph: { spacing: { line: 28 * 20, lineRule: LineRuleType.EXACT } } },
                    { id: "h4", name: "Heading 4", run: { font: FONT_FANGSONG_GB2312, size: SIZE_H_COMMON, bold: true }, paragraph: { spacing: { line: 28 * 20, lineRule: LineRuleType.EXACT } } },
                    { id: "h5", name: "Heading 5", run: { font: FONT_FANGSONG_GB2312, size: SIZE_H_COMMON }, paragraph: { spacing: { line: 28 * 20, lineRule: LineRuleType.EXACT } } },
                ]
            },
            sections: [{
                properties: {
                    page: {
                        size: { width: convertMillimetersToTwip(210), height: convertMillimetersToTwip(297) }, // A4
                        margin: {
                            top: convertMillimetersToTwip(37),
                            bottom: convertMillimetersToTwip(35),
                            left: convertMillimetersToTwip(28),
                            right: convertMillimetersToTwip(26),
                        },
                        pageBorders: {},
                        headers: { default: { properties: { header: { margin: { top: convertMillimetersToTwip(15) } } } } },
                        footers: { default: { properties: { footer: { margin: { top: convertMillimetersToTwip(28) } } } } },
                    },
                },
                children: docChildren,
            }],
        });

        // --- 5. 生成并下载 ---
        Packer.toBlob(doc).then(blob => {
            const timestamp = new Date().toISOString().replace(/[-:.]/g, "").replace("T", "_").slice(0, 15);
            saveAs(blob, `公文_${timestamp}.docx`);
        }).catch(err => {
            console.error("生成 Word 文档时出错:", err);
            alert("生成失败，详情请查看浏览器控制台。");
        });
    }
});
