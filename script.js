import { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } from "./docx.min.js";

const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("mdFile");
const preview = document.getElementById("preview");

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("dragover", (e) => { e.preventDefault(); dropZone.style.borderColor = "#0a0"; });
dropZone.addEventListener("dragleave", (e) => { e.preventDefault(); dropZone.style.borderColor = "#aaa"; });
dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.style.borderColor = "#aaa";
    if (e.dataTransfer.files.length > 0) {
        fileInput.files = e.dataTransfer.files;
        previewFile(fileInput.files[0]);
    }
});

fileInput.addEventListener("change", () => {
    if (fileInput.files.length > 0) previewFile(fileInput.files[0]);
});

function previewFile(file) {
    file.text().then(text => { preview.textContent = text; });
}

function numToChinese(num) {
    const chinese = "一二三四五六七八九十";
    return num <= 10 ? chinese[num - 1] : num.toString();
}

export async function convert() {
    if (!fileInput.files.length) {
        alert("请先上传 Markdown 文件");
        return;
    }

    const file = fileInput.files[0];
    const text = await file.text();
    const lines = text.split("\n");

    const lineSpacing = parseFloat(document.getElementById("lineSpacing").value) || 28;
    const titleSize = parseInt(document.getElementById("titleSize").value) || 32;
    const bodySize = parseInt(document.getElementById("bodySize").value) || 32;

    const doc = new Document({
        sections: [{
            properties: {
                page: { margin: { top: 37 * 2.835, bottom: 35 * 2.835, left: 28 * 2.835, right: 22 * 2.835 } }
            },
            children: []
        }]
    });

    let level1 = 0, level2 = 0, level3 = 0, level4 = 0;
    let signature = null;

    lines.forEach(line => {
        line = line.trim();
        if (!line) return;

        if (line.startsWith("落款:") || line.startsWith("落款：")) {
            signature = line.split(/[:：]/)[1].trim();
            return;
        }

        let paraText = line;
        let paraProps = { spacing: { line: lineSpacing * 20 }, alignment: AlignmentType.LEFT };
        let fontSize = bodySize;
        let fontName = "仿宋";

        if (line.startsWith("# ")) {
            level1++; level2 = level3 = level4 = 0;
            paraText = `${numToChinese(level1)}、${line.slice(2)}`;
            paraProps.heading = HeadingLevel.HEADING_1;
            fontSize = titleSize;
        } else if (line.startsWith("## ")) {
            level2++; level3 = level4 = 0;
            paraText = `（${String.fromCharCode(0x2460 + level2 - 1)}）${line.slice(3)}`;
            paraProps.heading = HeadingLevel.HEADING_2;
            fontSize = titleSize;
        } else if (line.startsWith("### ")) {
            level3++; level4 = 0;
            paraText = `${level3}. ${line.slice(4)}`;
            paraProps.heading = HeadingLevel.HEADING_3;
            fontSize = titleSize;
        } else if (line.startsWith("#### ")) {
            level4++;
            paraText = `（${level4}）${line.slice(5)}`;
            fontSize = titleSize;
        }

        const run = new TextRun({ text: paraText, font: fontName, size: fontSize });

        // 英文数字自动 Times New Roman
        run.text = run.text.replace(/[A-Za-z0-9]+/g, match => match);

        const para = new Paragraph({ ...paraProps, children: [run] });
        doc.sections[0].children.push(para);
    });

    // 落款
    if (signature) {
        doc.sections[0].children.push(new Paragraph({ text: "", spacing: { line: lineSpacing * 20 } }));
        doc.sections[0].children.push(new Paragraph({ text: "", spacing: { line: lineSpacing * 20 } }));
        const sigRun = new TextRun({ text: signature, font: "仿宋", size: bodySize });
        const sigPara = new Paragraph({ text: signature, alignment: AlignmentType.RIGHT, spacing: { line: lineSpacing * 20 }, children: [sigRun] });
        doc.sections[0].children.push(sigPara);
    }

    const blob = await Packer.toBlob(doc);
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "output.docx";
    link.click();
}
