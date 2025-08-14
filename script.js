document.addEventListener('DOMContentLoaded', () => {
    // 确保所有依赖库都已加载
    if (typeof docx === 'undefined' || typeof marked === 'undefined' || typeof saveAs === 'undefined') {
        alert("错误：一个或多个必需的库未能加载。请检查您的网络连接和文件路径。");
        return;
    }

    const { Document, Packer, Paragraph, TextRun, AlignmentType, convertMillimetersToTwip, LineRuleType } = docx;

    // --- 1. 获取所有 DOM 元素 ---
    const mdInput = document.getElementById('md-input');
    const convertBtn = document.getElementById('convert-btn');
    const mdUpload = document.getElementById('md-upload');
    // 移除了 autoNumberingCheckbox 的引用
    const downloadFontsBtn = document.getElementById('download-fonts-btn');

    // --- 2. 加载默认/示例 Markdown ---
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
        // 不再需要传递复选框状态
        generateDocx(markdownText);
    });

    // --- 4. 核心 DOCX 生成函数 (修改版) ---
    function generateDocx(mdText) {
        // 定义字体和字号 (pt)
        const FONT_FANGSONG_GB2312 = "仿宋_GB2312";
        const FONT_XIAOBIAOSONG = "方正小标宋简体";
        const FONT_HEITI = "黑体";
        const FONT_KAITI_GB2312 = "楷体_GB2312";

        const SIZE_H1 = 22 * 2; // 二号 (22pt)
        const SIZE_H_COMMON = 16 * 2; // 三号 (16pt)
        const SIZE_BODY = 16 * 2; // 三号 (16pt)

        // 移除了自动编号 (Numbering) 的定义

        const tokens = marked.lexer(mdText);
        const docChildren = [];

        tokens.forEach(token => {
            let paragraph;
            switch (token.type) {
                case 'heading':
                    const textWithoutFormatting = token.text;
                    let textRun;

                    switch (token.depth) {
                        case 1: // 主标题
                            paragraph = new Paragraph({
                                children: [new TextRun({ text: textWithoutFormatting, font: FONT_XIAOBIAOSONG, size: SIZE_H1 })],
                                alignment: AlignmentType.CENTER,
                                // 移除了 spacing.after，行距为 30pt 固定值
                                spacing: { line: 30 * 20, lineRule: LineRuleType.EXACT },
                            });
                            break;
                        case 2: // 一级标题
                            textRun = new TextRun({ text: textWithoutFormatting, font: FONT_HEITI, size: SIZE_H_COMMON });
                            break;
                        case 3: // 二级标题
                            textRun = new TextRun({ text: textWithoutFormatting, font: FONT_KAITI_GB2312, size: SIZE_H_COMMON });
                            break;
                        case 4: // 三级标题
                            textRun = new TextRun({ text: textWithoutFormatting, font: FONT_FANGSONG_GB2312, size: SIZE_H_COMMON, bold: true });
                            break;
                        case 5: // 四级标题
                            textRun = new TextRun({ text: textWithoutFormatting, font: FONT_FANGSONG_GB2312, size: SIZE_H_COMMON });
                            break;
                    }
                    
                    // 为 H2-H5 创建段落
                    if (token.depth > 1 && textRun) {
                        paragraph = new Paragraph({
                            children: [textRun],
                            // 所有内文标题与正文共享相同的缩进和行距设置
                            alignment: AlignmentType.JUSTIFIED,
                            indent: { firstLine: 640 }, // 2字符缩进
                            spacing: { line: 28 * 20, lineRule: LineRuleType.EXACT },
                        });
                    }
                    break;

                // 正确处理多行文本（如表格、代码块、引用）
                case 'paragraph':
                case 'list':
                case 'blockquote':
                case 'code':
                case 'table': // Explicitly handle tables
                case 'hr':
                    // For any non-heading content, we get its raw text, which preserves line breaks.
                    const rawText = token.raw || '';
                    if (rawText.trim()) {
                        // Split the raw text into individual lines.
                        const lines = rawText.split('\n');
                        // Create a new paragraph for each line to ensure formatting is correct.
                        lines.forEach(line => {
                            const lineParagraph = new Paragraph({
                                children: [new TextRun({ text: line, font: FONT_FANGSONG_GB2312, size: SIZE_BODY })],
                                alignment: AlignmentType.JUSTIFIED,
                                indent: { firstLine: 640 }, // 2-character indent
                                spacing: { line: 28 * 20, lineRule: LineRuleType.EXACT },
                            });
                            docChildren.push(lineParagraph);
                        });
                    }
                    // Set paragraph to null because we've already added children to the doc.
                    paragraph = null; 
                    break;

                default:
                    // Any other unhandled token types will be ignored.
                    paragraph = null;
                    break;
            }

            // This will now only handle heading paragraphs, as others are pushed directly.
            if (paragraph) {
                docChildren.push(paragraph);
            }
        });

        // 创建文档
        const doc = new Document({
            // 移除了 numbering 和 styles.paragraphStyles
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
