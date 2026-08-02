/* 本地识别逻辑测试：node 公文格式化识别逻辑测试.js */
"use strict";

var fs = require("fs");
var path = require("path");
var vm = require("vm");

// 方案A：唯一主本为插件目录的 main.js（本文件与其保持路径引用，不再维护根目录副本）
var macroPath = path.join(__dirname, "..", "方案A_ribbonUI插件", "OfficialDocumentFormat_0.01", "main.js");
var macroCode = fs.readFileSync(macroPath, "utf8");
var sandbox = {};
vm.createContext(sandbox);
vm.runInContext(macroCode, sandbox, { filename: macroPath });
var formatter = sandbox.GetOfficialDocumentFormatter();

var inspectPath = path.join(__dirname, "_inspect.json");
var raw = JSON.parse(fs.readFileSync(inspectPath, "utf8"));

function parseAlign(value) {
  var text = String(value || "");
  if (text.indexOf("CENTER") >= 0) return formatter.CONST.WD_ALIGN_CENTER;
  if (text.indexOf("RIGHT") >= 0) return formatter.CONST.WD_ALIGN_RIGHT;
  if (text.indexOf("JUSTIFY") >= 0) return formatter.CONST.WD_ALIGN_JUSTIFY;
  return formatter.CONST.WD_ALIGN_LEFT;
}

var items = raw.map(function (row, i) {
  return {
    index: i,
    text: row.text || "",
    align: parseAlign(row.align),
    hasBreak: row.text === "" && (i === 21 || i === 27)
  };
});

function sameArray(actual, expected) {
  if (actual.length !== expected.length) return false;
  for (var i = 0; i < actual.length; i++) {
    if (actual[i] !== expected[i]) return false;
  }
  return true;
}

function assertArray(name, actual, expected) {
  if (!sameArray(actual, expected)) {
    throw new Error(name + " expected [" + expected.join(",") + "], got [" + actual.join(",") + "]");
  }
  console.log("PASS " + name + ": [" + actual.join(",") + "]");
}

function assertEqual(name, actual, expected) {
  if (actual !== expected) {
    throw new Error(name + " expected " + expected + ", got " + actual);
  }
  console.log("PASS " + name + ": " + actual);
}

function headingIndexes(result, level) {
  var output = [];
  Object.keys(result.headings).forEach(function (key) {
    if (result.headings[key] === level) output.push(Number(key));
  });
  output.sort(function (a, b) { return a - b; });
  return output;
}

var result = formatter.classifyDocument(items);
var clearedItems = items.map(function (item) {
  return {
    index: item.index,
    text: item.text,
    align: formatter.CONST.WD_ALIGN_LEFT,
    hasBreak: item.hasBreak
  };
});
var clearedResult = formatter.classifyDocument(clearedItems);

assertArray("大标题", result.titleIndexes, [0, 1]);
assertArray("清除格式后大标题兜底", clearedResult.titleIndexes, [0, 1]);
assertEqual("主送单位", result.mainRecipientIndex, 2);
assertEqual("成文日期", result.dateIndex, 19);
assertEqual("署名", result.signatureIndex, 18);
assertArray("附件说明行", result.attachmentNoteIndexes, [16]);
assertArray("附件说明续行", result.attachmentNoteContinuationIndexes, [17]);
assertArray("附件顺序号行", result.attachmentSequenceIndexes, [22, 28]);
assertArray("附件标题", result.attachmentTitleIndexes, [23, 30]);
assertArray("一级标题", headingIndexes(result, "level1"), [4, 24]);
assertArray("二级标题", headingIndexes(result, "level2"), [5, 32]);
assertArray("三级标题", headingIndexes(result, "level3"), [11, 33]);
assertArray("四级标题", headingIndexes(result, "level4"), [12, 34]);
assertArray("一是/二是句式", result.yiShiIndexes, [6, 7, 10, 26]);

assertEqual("带句号的'一、'列举视为正文", result.headings[3], undefined);
assertEqual("超长'1.'列举视为正文", result.headings[13], undefined);
assertEqual("'一是否'不匹配一是句式", formatter.isYiShi(items[9].text), false);
assertEqual("段内'一是/二是'匹配", formatter.isYiShi("资产质量方面，一是资产共用。二是租约集中到期。"), true);
assertEqual("段内'一是/二是'数量", formatter.yiShiMatches("资产质量方面，一是资产共用。二是租约集中到期。").length, 2);
assertEqual("附件详见后页不匹配附件说明", formatter.isAttachmentNote("附件详见后页"), false);
assertEqual("附件说明续行支持点号", formatter.isAttachmentNoteContinuation("2.深交所无异议函"), true);
assertEqual("附件说明续行缩进", formatter.attachmentNoteContinuationIndent(), 80);
assertEqual("署名字符宽度", formatter.charWidth("债券监督管理司"), 7);
assertEqual("日期字符宽度", formatter.charWidth("2024年3月5日"), 6);
assertEqual("占位符成文日期", formatter.isDateText("2026年【】月【】日"), true);
assertEqual("正文内日期不作成文日期", formatter.isDateText("该REIT采用100%收益法，以2025年12月31日为估值基准日。"), false);

var placeholderDateItems = [
  { text: "关于准予北京国资商业不动产REIT注册的请示", align: formatter.CONST.WD_ALIGN_CENTER, hasBreak: false },
  { text: "中国证监会：", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "该REIT采用100%收益法，以2025年12月31日为估值基准日。", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "（三）估值合理性", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "附件：1.注册会会议纪要", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "2.深交所无异议函", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "债券司", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "2026年【】月【】日", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false }
];
var placeholderResult = formatter.classifyDocument(placeholderDateItems);
assertEqual("占位符日期定位", placeholderResult.dateIndex, 8);
assertEqual("占位符日期署名定位", placeholderResult.signatureIndex, 7);
assertArray("点号附件续行定位", placeholderResult.attachmentNoteContinuationIndexes, [5]);
assertEqual("估值合理性保持二级标题", placeholderResult.headings[3], "level2");

var noRecipientItems = [
  { text: "信创系统建设进展情况报告", align: formatter.CONST.WD_ALIGN_CENTER, hasBreak: false },
  { text: "", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "按照工作安排，现将有关情况报告如下。", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "一、基本情况", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "项目总体推进平稳。", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "附件：1.系统建设进展明细", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "信息技术部", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false },
  { text: "2026年7月23日", align: formatter.CONST.WD_ALIGN_LEFT, hasBreak: false }
];
var noRecipientResult = formatter.classifyDocument(noRecipientItems);
assertEqual("无主送单位文档不误判称呼", noRecipientResult.mainRecipientIndex, -1);
assertArray("无主送单位标题", noRecipientResult.titleIndexes, [0]);
assertArray("无主送单位正文", noRecipientResult.bodyIndexes, [2, 4]);
assertArray("无主送单位一级标题", headingIndexes(noRecipientResult, "level1"), [3]);
assertArray("无主送单位附件说明", noRecipientResult.attachmentNoteIndexes, [5]);
assertEqual("无主送单位成文日期", noRecipientResult.dateIndex, 8);
assertEqual("无主送单位署名", noRecipientResult.signatureIndex, 7);

console.log("全部识别逻辑测试通过。");
