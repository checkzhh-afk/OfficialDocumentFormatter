const fs = require("fs");
const path = require("path");

function requireProjectModule(name) {
  try {
    return require(name);
  } catch (globalError) {
    return require(path.join(__dirname, "..", "..", "OPENCODE测试", "node_modules", name));
  }
}

const {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Packer,
  PageNumber,
  Paragraph,
  Run,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
  XmlAttributeComponent,
  XmlComponent,
} = requireProjectModule("docx");
const JSZip = requireProjectModule("jszip");

const macroPath = path.join(__dirname, "公文格式化控制台WPS宏.js");
const outputPath = path.join(__dirname, "公文格式控制台_居中页码_GB2312.docm");
const pageWidth = 11906;
const pageHeight = 16838;
const margin = 850;
const contentWidth = pageWidth - margin * 2;

const colors = {
  red: "B42318",
  redDark: "7A271A",
  charcoal: "242424",
  gray: "667085",
  border: "C7CDD4",
  light: "F2F4F7",
  buttonSecondary: "E7ECF2",
  buttonPrimaryHighlight: "F5B7B1",
  buttonPrimaryShadow: "7A271A",
  buttonSecondaryHighlight: "FFFFFF",
  buttonSecondaryShadow: "98A2B3",
  redLight: "FEF3F2",
  greenLight: "ECFDF3",
  white: "FFFFFF",
};

const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: colors.border };
const cellBorders = {
  top: thinBorder,
  bottom: thinBorder,
  left: thinBorder,
  right: thinBorder,
};

function buttonBorders(primary) {
  const highlight = primary ? colors.buttonPrimaryHighlight : colors.buttonSecondaryHighlight;
  const shadow = primary ? colors.buttonPrimaryShadow : colors.buttonSecondaryShadow;
  return {
    top: { style: BorderStyle.SINGLE, size: 8, color: highlight },
    left: { style: BorderStyle.SINGLE, size: 8, color: highlight },
    bottom: { style: BorderStyle.SINGLE, size: 14, color: shadow },
    right: { style: BorderStyle.SINGLE, size: 14, color: shadow },
  };
}

function escapeXml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;")
    .replace(/\r\n|\r|\n/g, "&#x0A;");
}

function jdeProject(source) {
  return [
    '<?xml version="1.0" encoding="UTF-8" ?>',
    '<document version="2.0">',
    "    <name>Project</name>",
    '    <property desc="" lock="false" password="" />',
    "    <activemodule>1</activemodule>",
    '    <codemodule name="Module1" id="1">',
    '        <window cursorpos="0" actived="true" visible="true" />',
    "        <codetext>" + escapeXml(source) + "</codetext>",
    "    </codemodule>",
    "    <functionsdata />",
    "</document>",
    "",
  ].join("\r\n");
}

function addRelationship(xml) {
  const type = "http://www.wps.cn/officeDocument/2018/jdeExtension";
  if (xml.includes(type)) return xml;
  const relationship =
    '<Relationship Id="rIdJde" Type="' + type + '" Target="JDEData.bin"/>';
  return xml.replace("</Relationships>", relationship + "</Relationships>");
}

function addContentType(xml) {
  let result = xml.replace(
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "application/vnd.ms-word.document.macroEnabled.main+xml",
  );
  if (!result.includes('PartName="/word/JDEData.bin"')) {
    result = result.replace(
      "</Types>",
      '<Override PartName="/word/JDEData.bin" ContentType="application/octet-stream"/></Types>',
    );
  }
  return result;
}

function textParagraph(text, options) {
  const opts = options || {};
  return new Paragraph({
    alignment: opts.alignment || AlignmentType.LEFT,
    spacing: {
      before: opts.before || 0,
      after: opts.after || 0,
      line: opts.line || 320,
    },
    border: opts.border,
    children: [
      new TextRun({
        text,
        bold: !!opts.bold,
        color: opts.color || colors.charcoal,
        size: opts.size || 20,
        font: opts.font || "Microsoft YaHei",
      }),
    ],
  });
}

function sectionTitle(text) {
  return textParagraph(text, {
    bold: true,
    size: 23,
    color: colors.redDark,
    before: 150,
    after: 70,
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 8, color: colors.red, space: 3 },
    },
  });
}

class FieldCharAttributes extends XmlAttributeComponent {
  constructor(type) {
    super({ type });
    this.xmlKeys = { type: "w:fldCharType" };
  }
}

class FieldChar extends XmlComponent {
  constructor(type) {
    super("w:fldChar");
    this.root.push(new FieldCharAttributes(type));
  }
}

class PreserveSpaceAttributes extends XmlAttributeComponent {
  constructor() {
    super({ space: "preserve" });
    this.xmlKeys = { space: "xml:space" };
  }
}

class MacroButtonInstruction extends XmlComponent {
  constructor(macro, label) {
    super("w:instrText");
    this.root.push(new PreserveSpaceAttributes());
    this.root.push(" MACROBUTTON " + macro + " " + label + " ");
  }
}

function macroButtonRuns(command, primary) {
  const visibleLabel = "双击：" + command.label;
  return [
    new Run({ children: [new FieldChar("begin")] }),
    new Run({ children: [new MacroButtonInstruction(command.macro, visibleLabel)] }),
    new Run({ children: [new FieldChar("separate")] }),
    new TextRun({
      text: visibleLabel,
      bold: true,
      color: primary ? colors.white : colors.charcoal,
      size: primary ? 22 : 20,
      font: "Microsoft YaHei",
    }),
    new Run({ children: [new FieldChar("end")] }),
  ];
}

function commandCell(command, width) {
  const primary = !!command.primary;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    borders: buttonBorders(primary),
    shading: {
      fill: primary ? colors.red : colors.buttonSecondary,
      type: ShadingType.CLEAR,
    },
    margins: { top: 125, bottom: 115, left: 105, right: 105 },
    children: [
      new Paragraph({
        style: primary ? "PrimaryButton" : "SecondaryButton",
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0, line: 250 },
        children: macroButtonRuns(command, primary),
      }),
    ],
  });
}

function emptyCell(width) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: cellBorders,
    shading: { fill: colors.white, type: ShadingType.CLEAR },
    children: [new Paragraph("")],
  });
}

function commandTable(commands, columns) {
  const baseWidth = Math.floor(contentWidth / columns);
  const widths = [];
  for (let i = 0; i < columns; i++) {
    widths.push(i === columns - 1 ? contentWidth - baseWidth * (columns - 1) : baseWidth);
  }
  const rows = [];
  for (let i = 0; i < commands.length; i += columns) {
    const cells = [];
    for (let column = 0; column < columns; column++) {
      const command = commands[i + column];
      cells.push(command ? commandCell(command, widths[column]) : emptyCell(widths[column]));
    }
    rows.push(new TableRow({ children: cells }));
  }
  return new Table({
    width: { size: contentWidth, type: WidthType.DXA },
    columnWidths: widths,
    rows,
  });
}

const documentCommands = [
  { label: "同目录一键格式化", macro: "ConsoleFormatFolderDocument", primary: true },
  { label: "操作菜单", macro: "ConsoleRunMenu" },
  { label: "选择同目录目标", macro: "ConsoleSelectTargetDocument" },
  { label: "打开已选目标", macro: "ConsoleOpenTargetDocument" },
  { label: "格式化已选目标", macro: "ConsoleFormatTargetDocument" },
  { label: "选择已打开文档", macro: "ConsoleChooseOpenDocument" },
  { label: "查看当前目标", macro: "ConsoleShowTargetDocument" },
  { label: "脚本自检", macro: "ConsoleSyntaxCheck" },
  { label: "表格字体字号", macro: "ConsoleApplyTableTextFormat" },
  { label: "脚注字体", macro: "ConsoleApplyFootnoteTextFormat" },
  { label: "居中页码", macro: "ConsoleApplyPageNumberFormat" },
];

const selectionCommands = [
  { label: "大标题", macro: "ConsoleApplyTitleFormat" },
  { label: "主送单位", macro: "ConsoleApplyMainRecipientFormat" },
  { label: "正文", macro: "ConsoleApplyBodyFormat" },
  { label: "一级标题", macro: "ConsoleApplyLevel1HeadingFormat" },
  { label: "二级标题", macro: "ConsoleApplyLevel2HeadingFormat" },
  { label: "三级标题", macro: "ConsoleApplyLevel3HeadingFormat" },
  { label: "四级标题", macro: "ConsoleApplyLevel4HeadingFormat" },
  { label: "一是/二是加粗", macro: "ConsoleApplyYiShiFormat" },
  { label: "成文日期", macro: "ConsoleApplyDateFormat" },
  { label: "落款单位", macro: "ConsoleApplySignatureFormat" },
  { label: "附件说明", macro: "ConsoleApplyAttachmentNoteFormat" },
  { label: "附件说明续行", macro: "ConsoleApplyAttachmentNoteContinuationFormat" },
  { label: "附件序号", macro: "ConsoleApplyAttachmentSequenceFormat" },
  { label: "附件标题", macro: "ConsoleApplyAttachmentTitleFormat" },
];

function createDocument() {
  return new Document({
    creator: "Codex",
    title: "公文格式控制台 - 居中页码 GB2312 字体版",
    description: "嵌入 WPS JSA 宏和可执行 MacroButton 字段的居中页码 GB2312 字体版控制台",
    styles: {
      default: {
        document: {
          run: { font: "Microsoft YaHei", size: 20, color: colors.charcoal },
          paragraph: { spacing: { line: 320 } },
        },
      },
      paragraphStyles: [
        {
          id: "PrimaryButton",
          name: "Primary Button",
          basedOn: "Normal",
          next: "Normal",
          run: { font: "Microsoft YaHei", size: 22, bold: true, color: colors.white },
          paragraph: { alignment: AlignmentType.CENTER },
        },
        {
          id: "SecondaryButton",
          name: "Secondary Button",
          basedOn: "Normal",
          next: "Normal",
          run: { font: "Microsoft YaHei", size: 20, bold: true, color: colors.charcoal },
          paragraph: { alignment: AlignmentType.CENTER },
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: pageWidth, height: pageHeight },
            margin: { top: margin, right: margin, bottom: margin, left: margin },
          },
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: "居中页码 GB2312 字体版  |  ",
                    size: 15,
                    color: colors.gray,
                    font: "Microsoft YaHei",
                  }),
                  new TextRun({
                    children: [PageNumber.CURRENT],
                    size: 15,
                    color: colors.gray,
                    font: "Times New Roman",
                  }),
                ],
              }),
            ],
          }),
        },
        children: [
          textParagraph("公文格式控制台·GB2312", {
            bold: true,
            size: 36,
            color: colors.redDark,
            alignment: AlignmentType.CENTER,
            after: 45,
          }),
          textParagraph("居中页码｜方正小标宋｜仿宋/楷体_GB2312", {
            size: 19,
            color: colors.gray,
            alignment: AlignmentType.CENTER,
            after: 100,
          }),
          new Table({
            width: { size: contentWidth, type: WidthType.DXA },
            columnWidths: [contentWidth],
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    width: { size: contentWidth, type: WidthType.DXA },
                    borders: cellBorders,
                    shading: { fill: colors.greenLight, type: ShadingType.CLEAR },
                    margins: { top: 95, bottom: 95, left: 130, right: 130 },
                    children: [
                      textParagraph(
                        "全文：双击“同目录一键格式化”。局部：在已打开公文中选中文字，再回到本控制台双击对应格式。",
                        { size: 17, color: "166534", line: 280 },
                      ),
                    ],
                  }),
                ],
              }),
            ],
          }),
          sectionTitle("文档操作"),
          commandTable(documentCommands, 4),
          sectionTitle("段落选区格式"),
          commandTable(selectionCommands, 3),
          textParagraph(
            "全文格式化会同时处理表格、脚注/尾注和居中页码。多文件时使用上下方向键选择并按 Enter；WPS 的 MacroButton 通常需要双击。",
            {
              size: 16,
              color: colors.gray,
              before: 90,
              alignment: AlignmentType.CENTER,
            },
          ),
        ],
      },
    ],
  });
}

function validateMacroSource(source) {
  if (/[^\x00-\x7F]/.test(source)) {
    throw new Error("正式方案 B 宏必须保持纯 ASCII。");
  }
  if (source.replace(/\r\n/g, "").includes("\n") || source.replace(/\r\n/g, "").includes("\r")) {
    throw new Error("正式方案 B 宏必须统一使用 CRLF 换行。");
  }
  for (const command of documentCommands.concat(selectionCommands)) {
    const pattern = new RegExp("function\\s+" + command.macro + "\\s*\\(");
    if (!pattern.test(source)) throw new Error("宏入口不存在：" + command.macro);
  }
}

async function buildConsoleDocument() {
  const macroSource = fs.readFileSync(macroPath, "ascii");
  validateMacroSource(macroSource);

  const buffer = await Packer.toBuffer(createDocument());
  const zip = await JSZip.loadAsync(buffer);
  const relsPath = "word/_rels/document.xml.rels";
  const rels = await zip.file(relsPath).async("string");
  const contentTypes = await zip.file("[Content_Types].xml").async("string");

  zip.file("word/JDEData.bin", jdeProject(macroSource));
  zip.file(relsPath, addRelationship(rels));
  zip.file("[Content_Types].xml", addContentType(contentTypes));

  const result = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
  fs.writeFileSync(outputPath, result);
  console.log(outputPath);
}

buildConsoleDocument().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
