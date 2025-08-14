document.addEventListener('DOMContentLoaded', () => {
    if (typeof docx === 'undefined' || typeof marked === 'undefined' || typeof saveAs === 'undefined') {
        alert("错误：依赖库未能成功加载，请检查网络或文件路径。");
        return;
    }

    const { Document, Packer, Paragraph, TextRun, AlignmentType, convertMillimetersToTwip, LineRuleType } = docx;

    const mdInput = document.getElementById('md-input');
    const convertBtn = document.getElementById('convert-btn');
    const mdUpload = document.getElementById('md-upload');
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

    function generateDocx(mdText) {
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
        
        // Main Title
        const titleToken = tokens.find(t => t.type === 'heading' && t.depth === 1);
        if (titleToken) {
            docChildren.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                children: createMixedFontRuns(titleToken.text, STYLES.title),
                spacing: { after: 400 }
            }));
            tokens.splice(tokens.indexOf(titleToken), 1);
        }

        // Document Number
        const docNumToken = tokens.find(t => t.type === 'paragraph');
        if (docNumToken) {
             docChildren.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                children: createMixedFontRuns(docNumToken.text, STYLES.docNumber),
                spacing: { after: 600 }
            }));
            tokens.splice(tokens.indexOf(docNumToken), 1);
        }

        // Process rest of the tokens
        tokens.forEach(token => {
            let para;
            switch (token.type) {
                case 'heading':
                    let text, style;
                    if (token.depth === 2) { // ## -> 一、
                        counters[0]++; counters[1] = 0; counters[2] = 0;
                        text = `第${"一二三四五六七八九十"[counters[0]-1]}章 ${token.text}`;
                        style = STYLES.h1;
                    } else if (token.depth === 3) { // ### -> (一)
                        counters[1]++; counters[2] = 0;
                        text = `（${"一二三四五六七八九十"[counters[1]-1]}）${token.text}`;
                        style = STYLES.h2;
                    } else if (token.depth === 4) { // #### -> 1.
                        counters[2]++;
                        text = `${counters[2]}. ${token.text}`;
                        style = STYLES.h3;
                    } else { return; }
                    
                    para = new Paragraph({
                        children: createMixedFontRuns(text, style),
                        spacing: { before: 240, after: 240, line: LINE_SPACING, lineRule: LineRuleType.EXACT },
                    });
                    break;
                case 'paragraph':
                    para = new Paragraph({
                        children: createMixedFontRuns(token.text, STYLES.body),
                        spacing: { line: LINE_SPACING, lineRule: LineRuleType.EXACT },
                        indent: { firstLine: convertMillimetersToTwip(10.5) }
                    });
                    break;
                default: return;
            }
            docChildren.push(para);
        });

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: convertMillimetersToTwip(37),
                            bottom: convertMillimetersToTwip(35),
                            left: convertMillimetersToTwip(28),
                            right: convertMillimetersToTwip(26),
                        },
                    },
                },
                children: docChildren,
            }],
        });

        Packer.toBlob(doc).then(blob => {
            saveAs(blob, "公文格式文档.docx");
        }).catch(err => console.error("生成文档失败:", err));
    }
});
