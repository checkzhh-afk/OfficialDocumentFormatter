// 将 Markdown 建议文稿转换为公文格式 Word 文档
// 格式规则与 公文格式化WPS宏.js 保持一致：
//   标题：宋体 22pt 加粗居中；一级标题（一、）：黑体 16pt；二级标题：楷体 16pt 加粗；
//   正文：仿宋 16pt；西文 Times New Roman；行距固定 28 磅；首行缩进 2 字符；页码居中 TNR 14pt
// 用法: node md2docx_gongwen.js <输入.md> <输出.docx>
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Footer, AlignmentType, PageNumber,
} = require("docx");

const [, , inPath, outPath] = process.argv;
if (!inPath || !outPath) { console.error("用法: node md2docx_gongwen.js 输入.md 输出.docx"); process.exit(1); }

const FONT_BODY = "仿宋", FONT_TITLE = "宋体", FONT_L1 = "黑体", FONT_L2 = "楷体", FONT_WEST = "Times New Roman";
const SIZE_BODY = 32;   // 16pt 三号（半磅单位）
const SIZE_TITLE = 44;  // 22pt 二号
const LINE = { line: 560, lineRule: "exact", before: 0, after: 0 }; // 固定值 28 磅
const INDENT2 = { firstLine: 640 }; // 2 字符 × 16pt = 32pt = 640 twips

function runs(text, cnFont, bold) {
  // 中西文分段：西文/数字用 Times New Roman，中文用指定字体
  const parts = text.split(/([\x00-\x7f]+)/).filter(Boolean);
  return parts.map(p =>
    new TextRun({
      text: p,
      bold: !!bold,
      size: arguments[3] || SIZE_BODY,
      font: /^[\x00-\x7f]+$/.test(p)
        ? { ascii: FONT_WEST, hAnsi: FONT_WEST, eastAsia: cnFont }
        : { ascii: FONT_WEST, hAnsi: FONT_WEST, eastAsia: cnFont },
    })
  );
}

function para(text, kind) {
  if (kind === "title")
    return new Paragraph({ alignment: AlignmentType.CENTER, spacing: LINE,
      children: runs(text, FONT_TITLE, true, SIZE_TITLE) });
  if (kind === "level1")
    return new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: LINE, indent: INDENT2,
      children: runs(text, FONT_L1, false) });
  if (kind === "level2")
    return new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: LINE, indent: INDENT2,
      children: runs(text, FONT_L2, true) });
  return new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: LINE, indent: INDENT2,
    children: runs(text, FONT_BODY, false) });
}

const blank = () => new Paragraph({ spacing: LINE, children: [] });

const md = fs.readFileSync(inPath, "utf-8");
const lines = md.split(/\r?\n/);
const children = [];

for (const raw of lines) {
  const line = raw.trim();
  if (!line || line === "---") continue;
  const plain = line.replace(/\*\*/g, ""); // 公文正文不保留加粗标记
  if (line.startsWith("# ") && !line.startsWith("## ")) {
    children.push(blank(), blank());               // 标题前空两行
    children.push(para(plain.replace(/^#\s*/, ""), "title"));
    children.push(blank());                        // 标题后空一行
  } else if (line.startsWith("## ")) {
    children.push(para(plain.replace(/^##\s*/, ""), "level1"));
  } else if (line.startsWith("### ")) {
    children.push(para(plain.replace(/^###\s*/, ""), "level2"));
  } else if (/^-\s+/.test(line)) {
    children.push(para(plain.replace(/^-\s+/, ""), "body"));
  } else {
    children.push(para(plain, "body"));
  }
}

const doc = new Document({
  styles: { default: { document: { run: { font: { ascii: FONT_WEST, eastAsia: FONT_BODY }, size: SIZE_BODY } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 2098, bottom: 1984, left: 1588, right: 1474 }, // 3.7/3.5/2.8/2.6 cm
      },
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({
            children: ["— ", PageNumber.CURRENT, " —"],
            size: 28, // 14pt 四号
            font: { ascii: FONT_WEST, hAnsi: FONT_WEST, eastAsia: FONT_TITLE },
          })],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outPath, buf);
  console.log("已生成: " + outPath);
});
