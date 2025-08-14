document.addEventListener('DOMContentLoaded', () => {
    if (typeof docx === 'undefined' || typeof marked === 'undefined' || typeof saveAs === 'undefined') {
        alert("错误：依赖库未能成功加载，请检查网络或文件路径。");
        return;
    }

    const { Document, Packer, Paragraph, TextRun, AlignmentType, convertMillimetersToTwip, LineRuleType } = docx;

    const mdInput = document.getElementById('md-input');
    const convertBtn = document.getElementById('convert-btn');
    const mdUpload = document.getElementById('md-upload');
    const autoNumberingCheckbox = document.getElementById('auto-numbering');
    const downloadFontsBtn = document.getElementById('download-fonts-btn'); // 新增：获取字体下载按钮
    let markdownText = "";

    fetch('sample.md').then(res => res.ok ? res.text() : "").then(text => {
        mdInput.value = text;
        markdownText = text;
    }).catch(console.warn);

    mdUpload.addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (file) {
            markdownText = await file.text();
            mdInput.value = markdownText;
        }
    });
    
    mdInput.addEventListener('input', () => markdownText = mdInput.value);

    convertBtn.addEventListener('click', () => {
        if (!markdownText.trim()) {
            alert('内容为空，请输入或上传 Markdown 文件！');
            return;
        }
        generateDocx(markdownText);
    });

    // --- 新增：字体下载按钮点击事件 ---
    downloadFontsBtn.addEventListener('click', () => {
        const fontZipUrl = 'https://github.com/AngelSnow1129/md2docx/releases/download/V1.0.0/default.zip';
        window.open(fontZipUrl, '_blank');
        console.log(`正在尝试从以下地址下载字体包: ${fontZipUrl}`);
    });


    function generateDocx(mdText) {
        const enableAutoNumbering = autoNumberingCheckbox.checked;

        const STYLES = {
            title: { font: "方正小标宋简体", size: 44, bold: true },
            docNumber: { font: "楷体_GB2312", size: 32 },
            body: { font: "仿宋_GB2312", size: 32 },
            h1: { font: "黑体_GB2312", size: 32 },
            h2: { font: "楷体_GB2312", size: 32 },
            h3: { font: "仿宋_GB2312", size: 32, bold: true },
            times: { font: "Times New Roman", size: 32 }
        };
        const LINE_SPACING = 28 * 20;

        const tokens = marked.lexer(mdText);
        let docChildren = [];
        let counters = [0, 0, 0];

        const createMixedFontRuns = (text, style) => {
            const runs = [];
            const regex = /([a-zA-Z0-9\s.,-]+)/g;
            const parts = text.split(regex).filter(p => p);
            for (const part of parts) {
                const isLatin = regex.test(part);
                const currentStyle = isLatin ? STYLES.times : style;
                runs.push(new TextRun({
                    text: part,
                    font: currentStyle.font,
                    size: currentStyle.size,
                    bold: style.bold || false,
                }));
            }
            return runs;
        };
        
        const titleToken = tokens.find(t => t.type === 'heading' && t.depth === 1);
        if (titleToken) {
            docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, children: createMixedFontRuns(titleToken.text, STYLES.title), spacing: { after: 400 } }));
            tokens.splice(tokens.indexOf(titleToken), 1);
        }

        const docNumToken = tokens.find(t => t.type === 'paragraph');
        if (docNumToken) {
             docChildren.push(new Paragraph({ alignment: AlignmentType.CENTER, children: createMixedFontRuns(docNumToken.text, STYLES.docNumber), spacing: { after: 600 } }));
             tokens.splice(tokens.indexOf(docNumToken), 1);
        }

        tokens.forEach(token => {
            let para;
            switch (token.type) {
                case 'heading':
                    let text = token.text, style;
                    const chineseNumerals = "一二三四五六七八九十";
                    
                    if (token.depth === 2) { // ## -> 一、
                        style = STYLES.h1;
                        if (enableAutoNumbering) {
                            counters[0]++; counters[1] = 0; counters[2] = 0;
                            if (!/^第?[一二三四五六七八九十]+[、章]/.test(text)) {
                                text = `第${chineseNumerals[counters[0]-1]}章 ${token.text}`;
                            }
                        }
                    } else if (token.depth === 3) { // ### -> (一)
                        style = STYLES.h2;
                        if (enableAutoNumbering) {
                            counters[1]++; counters[2] = 0;
                            if (!/^[（(][一二三四五六七八九十]+[)）]/.test(text)) {
                                text = `（${chineseNumerals[counters[1]-1]}）${token.text}`;
                            }
                        }
                    } else if (token.depth === 4) { // #### -> 1.
                        style = STYLES.h3;
                        if (enableAutoNumbering) {
                           counters[2]++;
                           if (!/^\d+\./.test(text)) {
                               text = `${counters[2]}. ${token.text}`;
                           }
                        }
                    } else { return; }
                    
                    para = new Paragraph({ children: createMixedFontRuns(text, style), spacing: { before: 240, after: 240, line: LINE_SPACING, lineRule: LineRuleType.EXACT } });
                    break;
                case 'paragraph':
                    para = new Paragraph({ children: createMixedFontRuns(token.text, STYLES.body), spacing: { line: LINE_SPACING, lineRule: LineRuleType.EXACT }, indent: { firstLine: convertMillimetersToTwip(10.5) } });
                    break;
                default: return;
            }
            docChildren.push(para);
        });

        const doc = new Document({
            sections: [{
                properties: { page: { margin: { top: convertMillimetersToTwip(37), bottom: convertMillimetersToTwip(35), left: convertMillimetersToTwip(28), right: convertMillimetersToTwip(26) } } },
                children: docChildren,
            }],
        });

        // --- 恢复：文件名时间戳功能 ---
        const now = new Date();
        const pad = (n) => n.toString().padStart(2, '0');
        const timestamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}_`;

        Packer.toBlob(doc).then(blob => {
            saveAs(blob, `${timestamp}_公文格式文档.docx`);
        }).catch(err => console.error("生成文档失败:", err));
    }
});