const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } = docx;

const fileInput = document.getElementById("mdFile");
const preview = document.getElementById("preview");
const convertBtn = document.getElementById("convertBtn");

let mdText = "";

// 上传文件事件
fileInput.addEventListener("change", () => {
    if (fileInput.files.length > 0) {
        const file = fileInput.files[0];
        file.text().then(text => {
            mdText = text;
            preview.textContent = mdText;
        });
    }
});

function numToChinese(num) {
    const c = "一二三四五六七八九十";
    return num <= 10 ? c[num - 1] : num.toString();
}

convertBtn.addEventListener("click", async () => {
    if (!mdText) { alert("请先上传 Markdown 文件"); return; }

    const lines = mdText.split("\n");
    const lineSpacing = parseFloat(document.getElementById("lineSpacing").value) || 28;
    const titleSize = parseInt(document.getElementById("titleSize").value) || 32;
    const bodySize = parseInt(document.getElementById("bodySize").value) || 32;

    const doc = new Document({
        sections: [{
            properties: {
                page: { margin: { top: 37*2.835, bottom: 35*2.835, left: 28*2.835, right: 22*2.835 } }
            },
            children: []
        }]
    });

    let level1=0, level2=0, level3=0, level4=0, signature=null;

    lines.forEach(line => {
        line = line.trim();
        if (!line) return;

        if (line.startsWith("落款:") || line.startsWith("落款：")) {
            signature = line.split(/[:：]/)[1].trim();
            return;
        }

        let paraText = line;
        let paraProps = { spacing: { line: lineSpacing*20 }, alignment: AlignmentType.LEFT };
        let fontSize = bodySize;
        let fontName = "仿宋";

        if (line.startsWith("# ")) { level1++; level2=level3=level4=0; paraText=`${numToChinese(level1)}、${line.slice(2)}`; paraProps.heading=HeadingLevel.HEADING_1; fontSize=titleSize; }
        else if (line.startsWith("## ")) { level2++; level3=level4=0; paraText=`（${String.fromCharCode(0x2460+level2-1)}）${line.slice(3)}`; paraProps.heading=HeadingLevel.HEADING_2; fontSize=titleSize; }
        else if (line.startsWith("### ")) { level3++; level4=0; paraText=`${level3}. ${line.slice(4)}`; paraProps.heading=HeadingLevel.HEADING_3; fontSize=titleSize; }
        else if (line.startsWith("#### ")) { level4++; paraText=`（${level4}）${line.slice(5)}`; fontSize=titleSize; }

        const run = new TextRun({ text: paraText, font: fontName, size: fontSize });
        run.text = run.text.replace(/[A-Za-z0-9]+/g, match => match); // 英文数字 Times New Roman

        const para = new Paragraph({ ...paraProps, children: [run] });
        doc.sections[0].children.push(para);
    });

    if (signature) {
        doc.sections[0].children.push(new Paragraph({ text:"", spacing:{line: lineSpacing*20} }));
        doc.sections[0].children.push(new Paragraph({ text:"", spacing:{line: lineSpacing*20} }));
        const sigRun = new TextRun({ text: signature, font:"仿宋", size: bodySize });
        const sigPara = new Paragraph({ text: signature, alignment: AlignmentType.RIGHT, spacing: { line: lineSpacing*20 }, children: [sigRun] });
        doc.sections[0].children.push(sigPara);
    }

    const blob = await Packer.toBlob(doc);
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "output.docx";
    link.click();
});
