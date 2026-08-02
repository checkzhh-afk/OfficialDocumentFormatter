/* 公文格式化 WPS JS 宏
 * 运行环境：WPS Office JS 宏。菜单按钮通过 customUI.xml 绑定本文件中的全局函数。
 */
  var CONST;
  var OfficialDocumentFormatter;

  function ensureOfficialDocumentFormatter() {
    if (CONST) return;
    CONST = {};
    CONST.FONT_BODY = "仿宋";
    CONST.FONT_TITLE = "宋体";
    CONST.FONT_LEVEL1 = "黑体";
    CONST.FONT_LEVEL2 = "楷体";
    CONST.FONT_WEST = "Times New Roman";
    CONST.SIZE_BODY = 16;
    CONST.SIZE_TITLE = 22;
    CONST.LINE_SPACING = 28;
    CONST.CHAR_PT = 16;
    CONST.WD_ALIGN_LEFT = 0;
    CONST.WD_ALIGN_CENTER = 1;
    CONST.WD_ALIGN_RIGHT = 2;
    CONST.WD_ALIGN_JUSTIFY = 3;
    CONST.WD_LINE_SPACE_EXACTLY = 4;
    CONST.WD_WITHIN_TABLE = 12;
    CONST.WD_HEADER_FOOTER_PRIMARY = 1;
    CONST.WD_HEADER_FOOTER_FIRST_PAGE = 2;
    CONST.WD_HEADER_FOOTER_EVEN_PAGES = 3;
    CONST.WD_FIELD_PAGE = 33;
    CONST.WD_ALIGN_PAGE_NUMBER_CENTER = 1;
    OfficialDocumentFormatter = {};
    OfficialDocumentFormatter.CONST = CONST;
    OfficialDocumentFormatter.trimText = trimText;
    OfficialDocumentFormatter.charWidth = charWidth;
    OfficialDocumentFormatter.isHeading = isHeading;
    OfficialDocumentFormatter.isYiShi = isYiShi;
    OfficialDocumentFormatter.yiShiPrefixLength = yiShiPrefixLength;
    OfficialDocumentFormatter.yiShiMatches = yiShiMatches;
    OfficialDocumentFormatter.isDateText = isDateText;
    OfficialDocumentFormatter.findDateText = findDateText;
    OfficialDocumentFormatter.isAttachmentNote = isAttachmentNote;
    OfficialDocumentFormatter.isAttachmentNoteContinuation = isAttachmentNoteContinuation;
    OfficialDocumentFormatter.isAttachmentSequence = isAttachmentSequence;
    OfficialDocumentFormatter.attachmentNoteContinuationIndent = attachmentNoteContinuationIndent;
    OfficialDocumentFormatter.signatureLeftIndent = signatureLeftIndent;
    OfficialDocumentFormatter.classifyDocument = classifyDocument;
  }

  function trimText(text) {
    if (text === null || text === undefined) return "";
    return String(text)
      .replace(/[\r\n\u0007]/g, "")
      .replace(/\f/g, "")
      .replace(/^\s+|\s+$/g, "");
  }

  function visibleLength(text) {
    return trimText(text).length;
  }

  function charWidth(text) {
    var s = trimText(text);
    var total = 0;
    for (var i = 0; i < s.length; i++) {
      total += /[\x00-\x7f]/.test(s.charAt(i)) ? 0.5 : 1;
    }
    return total;
  }

  function endsWithColon(text) {
    return /[:：]$/.test(trimText(text));
  }

  function endsWithPeriod(text) {
    return /[。.]$/.test(trimText(text));
  }

  function hasSerialPrefix(text) {
    return /^([一二三四五六七八九十百]+、|（[一二三四五六七八九十百]+）|\d+\.|（\d+）)/.test(trimText(text));
  }

  function isCenteredParagraph(meta) {
    return Number(meta && meta.align) === CONST.WD_ALIGN_CENTER;
  }

  function isTitleCandidate(meta, allowShortContinuation, allowUncentered) {
    var text = trimText(meta && meta.text);
    var len = visibleLength(text);
    if (!text) return false;
    if (!allowUncentered && !isCenteredParagraph(meta)) return false;
    if (hasSerialPrefix(text)) return false;
    if (endsWithColon(text)) return false;
    if (len > 50) return false;
    return allowShortContinuation ? len >= 2 : len >= 5;
  }

  function isHeading(text) {
    var s = trimText(text);
    var len = visibleLength(s);
    if (len < 2 || len > 40 || endsWithPeriod(s)) return null;
    if (/^[一二三四五六七八九十百]+、/.test(s)) return "level1";
    if (/^（[一二三四五六七八九十百]+）/.test(s)) return "level2";
    if (/^\d+\./.test(s)) return "level3";
    if (/^（\d+）/.test(s)) return "level4";
    return null;
  }

  function isYiShi(text) {
    return yiShiMatches(text).length > 0;
  }

  function yiShiPrefixLength(text) {
    var matches = yiShiMatches(text);
    return matches.length && matches[0].start === 0 ? matches[0].length : 0;
  }

  function yiShiMatches(text) {
    var s = String(text === null || text === undefined ? "" : text).replace(/[\r\n\u0007\f]+$/g, "");
    var matches = [];
    var re = /(^|[，。；;、\s])([一二三四五六七八九十]+是)(?!否)/g;
    var match;
    while ((match = re.exec(s)) !== null) {
      matches.push({ start: match.index + match[1].length, length: match[2].length, text: match[2] });
      if (match.index === re.lastIndex) re.lastIndex++;
    }
    return matches;
  }

  function dateLinePattern() {
    return "\\d{4}年(?:\\d{1,2}|[【\\[]\\s*[】\\]]|_+|＿+)月(?:\\d{1,2}|[【\\[]\\s*[】\\]]|_+|＿+)日";
  }

  function isDateText(text) {
    return new RegExp("^" + dateLinePattern() + "$").test(trimText(text));
  }

  function findDateText(text) {
    var matches = trimText(text).match(new RegExp(dateLinePattern(), "g"));
    return matches && matches.length ? matches[matches.length - 1] : "";
  }

  function isAttachmentNote(text) {
    return /^附件(?:\d+|\s|[:：])/.test(trimText(text));
  }

  function isAttachmentNoteContinuation(text) {
    return /^\d+(?:\s+|[.．、])\S+/.test(trimText(text));
  }

  function isAttachmentSequence(text) {
    return /^附件\s*\d*$/.test(trimText(text));
  }

  function attachmentNoteHangingIndent(text) {
    var s = trimText(text);
    var match = s.match(/^附件(?:[:：]|\s*)/);
    if (!match) return 0;
    return charWidth(match[0]) * CONST.CHAR_PT;
  }

  function attachmentNoteContinuationIndent() {
    return CONST.CHAR_PT * 2 + charWidth("附件：") * CONST.CHAR_PT;
  }

  function paragraphMeta(paragraph, index) {
    var text = "";
    var align = CONST.WD_ALIGN_LEFT;
    var hasBreak = false;
    try {
      text = paragraph.Range.Text;
      hasBreak = String(text).indexOf("\f") >= 0;
    } catch (e1) {}
    try {
      align = paragraph.Range.ParagraphFormat.Alignment;
    } catch (e2) {
      try {
        align = paragraph.Alignment;
      } catch (e3) {}
    }
    return { paragraph: paragraph, index: index, text: trimText(text), rawText: text, align: align, hasBreak: hasBreak };
  }

  function collectionToArray(collection) {
    var result = [];
    if (!collection) return result;
    var count = 0;
    try { count = collection.Count; } catch (e1) {}
    if (count) {
      for (var i = 1; i <= count; i++) {
        try { result.push(collection.Item(i)); } catch (e2) {}
      }
      return result;
    }
    try {
      for (var j = 0; j < collection.length; j++) result.push(collection[j]);
    } catch (e3) {}
    return result;
  }

  function findTitleIndexes(items) {
    var indexes = findTitleIndexesWithMode(items, false);
    if (!indexes.length) indexes = findTitleIndexesWithMode(items, true);
    return indexes;
  }

  function findTitleIndexesWithMode(items, allowUncentered) {
    var indexes = [];
    var nonEmptySeen = 0;
    var started = false;
    for (var i = 0; i < items.length && nonEmptySeen < 10; i++) {
      var text = trimText(items[i].text);
      if (!text) continue;
      nonEmptySeen++;
      if (!started) {
        if (isTitleCandidate(items[i], false, allowUncentered)) {
          indexes.push(i);
          started = true;
        }
      } else if (isTitleCandidate(items[i], true, allowUncentered)) {
        indexes.push(i);
      } else {
        break;
      }
    }
    return indexes;
  }

  function findMainRecipientIndex(items, titleIndexes) {
    if (!titleIndexes.length) return -1;
    var start = titleIndexes[titleIndexes.length - 1] + 1;
    for (var i = start; i < items.length; i++) {
      var text = trimText(items[i].text);
      if (!text) continue;
      return endsWithColon(text) ? i : -1;
    }
    return -1;
  }

  function findDateIndex(items) {
    var idx = -1;
    for (var i = 0; i < items.length; i++) {
      if (isDateText(items[i].text)) idx = i;
    }
    return idx;
  }

  function findSignatureIndex(items, dateIndex) {
    if (dateIndex <= 0) return -1;
    for (var i = dateIndex - 1; i >= 0; i--) {
      if (trimText(items[i].text)) return i;
    }
    return -1;
  }

  function findAttachmentNoteIndexes(items, dateIndex) {
    var indexes = [];
    var end = dateIndex >= 0 ? dateIndex : items.length;
    for (var i = 0; i < end; i++) {
      if (isAttachmentNote(items[i].text)) indexes.push(i);
    }
    return indexes;
  }

  function findAttachmentNoteContinuationIndexes(items, dateIndex, attachmentNoteIndexes) {
    var indexes = [];
    var noteMap = {};
    for (var n = 0; n < attachmentNoteIndexes.length; n++) noteMap[attachmentNoteIndexes[n]] = true;
    var end = dateIndex >= 0 ? dateIndex : items.length;
    var inAttachmentNotes = false;
    for (var i = 0; i < end; i++) {
      var text = trimText(items[i].text);
      if (!text) {
        inAttachmentNotes = false;
        continue;
      }
      if (noteMap[i]) {
        inAttachmentNotes = true;
        continue;
      }
      if (inAttachmentNotes && isAttachmentNoteContinuation(text)) {
        indexes.push(i);
        continue;
      }
      inAttachmentNotes = false;
    }
    return indexes;
  }

  function findAttachmentBody(items, dateIndex) {
    var sequenceIndexes = [];
    var titleIndexes = [];
    var afterBreak = false;
    for (var i = 0; i < items.length; i++) {
      if (items[i].hasBreak) afterBreak = true;
      if (dateIndex >= 0 && i <= dateIndex) continue;
      var text = trimText(items[i].text);
      if (!text) {
        if (i > dateIndex) afterBreak = true;
        continue;
      }
      if ((afterBreak || dateIndex >= 0) && isAttachmentSequence(text)) {
        sequenceIndexes.push(i);
        for (var j = i + 1; j < items.length; j++) {
          if (trimText(items[j].text)) {
            titleIndexes.push(j);
            break;
          }
        }
        afterBreak = false;
      }
    }
    return { sequenceIndexes: sequenceIndexes, titleIndexes: titleIndexes };
  }

  function classifyDocument(items) {
    ensureOfficialDocumentFormatter();
    var titleIndexes = findTitleIndexes(items);
    var mainRecipientIndex = findMainRecipientIndex(items, titleIndexes);
    var dateIndex = findDateIndex(items);
    var signatureIndex = findSignatureIndex(items, dateIndex);
    var attachmentNoteIndexes = findAttachmentNoteIndexes(items, dateIndex);
    var attachmentNoteContinuationIndexes = findAttachmentNoteContinuationIndexes(items, dateIndex, attachmentNoteIndexes);
    var attachmentBody = findAttachmentBody(items, dateIndex);
    var start = mainRecipientIndex >= 0 ? mainRecipientIndex + 1 : (titleIndexes.length ? titleIndexes[titleIndexes.length - 1] + 1 : 0);
    var end = dateIndex >= 0 ? dateIndex : items.length;
    var headings = {};
    var yiShiIndexes = [];
    var bodyIndexes = [];
    for (var i = start; i < items.length; i++) {
      if (i >= end && attachmentBody.sequenceIndexes.indexOf(i) < 0 && attachmentBody.titleIndexes.indexOf(i) < 0) {
        if (dateIndex >= 0 && i <= dateIndex) continue;
      }
      var text = trimText(items[i].text);
      if (!text) continue;
      if (titleIndexes.indexOf(i) >= 0 || i === mainRecipientIndex || i === dateIndex || i === signatureIndex) continue;
      if (attachmentNoteIndexes.indexOf(i) >= 0 || attachmentNoteContinuationIndexes.indexOf(i) >= 0 || attachmentBody.sequenceIndexes.indexOf(i) >= 0 || attachmentBody.titleIndexes.indexOf(i) >= 0) continue;
      var heading = isHeading(text);
      if (heading) {
        headings[i] = heading;
      } else {
        bodyIndexes.push(i);
      }
      if (isYiShi(text)) yiShiIndexes.push(i);
    }
    return {
      titleIndexes: titleIndexes,
      mainRecipientIndex: mainRecipientIndex,
      dateIndex: dateIndex,
      signatureIndex: signatureIndex,
      attachmentNoteIndexes: attachmentNoteIndexes,
      attachmentNoteContinuationIndexes: attachmentNoteContinuationIndexes,
      attachmentSequenceIndexes: attachmentBody.sequenceIndexes,
      attachmentTitleIndexes: attachmentBody.titleIndexes,
      headings: headings,
      yiShiIndexes: yiShiIndexes,
      bodyIndexes: bodyIndexes
    };
  }

  function canFormatParagraph(paragraph) {
    try {
      if (paragraph.Range.Information(CONST.WD_WITHIN_TABLE)) return false;
    } catch (e1) {}
    try {
      if (paragraph.Range.InlineShapes && paragraph.Range.InlineShapes.Count > 0) return false;
    } catch (e2) {}
    try {
      if (paragraph.Range.ShapeRange && paragraph.Range.ShapeRange.Count > 0) return false;
    } catch (e3) {}
    try {
      if (String(paragraph.Range.Text).indexOf("\f") >= 0) return false;
    } catch (e4) {}
    return true;
  }

  function setFont(range, cnFont, westFont, size, bold) {
    var font = range.Font;
    try { font.Name = cnFont; } catch (e1) {}
    try { font.NameFarEast = cnFont; } catch (e2) {}
    try { font.NameAscii = westFont; } catch (e3) {}
    try { font.NameOther = westFont; } catch (e4) {}
    try { font.NameBi = westFont; } catch (e5) {}
    try { font.Size = size; } catch (e6) {}
    try { font.Bold = bold ? true : false; } catch (e7) {}
  }

  function setParagraph(paragraph, align, leftIndent, firstLineIndent, rightIndent) {
    var pf = paragraph.Range.ParagraphFormat;
    try { pf.Alignment = align; } catch (e1) {}
    try { pf.LeftIndent = leftIndent || 0; } catch (e2) {}
    try { pf.FirstLineIndent = firstLineIndent || 0; } catch (e3) {}
    try { pf.RightIndent = rightIndent || 0; } catch (e4) {}
    try { pf.LineSpacingRule = CONST.WD_LINE_SPACE_EXACTLY; } catch (e5) {}
    try { pf.LineSpacing = CONST.LINE_SPACING; } catch (e6) {}
    try { pf.SpaceBefore = 0; } catch (e7) {}
    try { pf.SpaceAfter = 0; } catch (e8) {}
  }

  function setWesternFont(range, fontName) {
    var font = range.Font;
    try { font.Name = fontName; } catch (e1) {}
    try { font.NameFarEast = fontName; } catch (e2) {}
    try { font.NameAscii = fontName; } catch (e3) {}
    try { font.NameOther = fontName; } catch (e4) {}
    try { font.NameBi = fontName; } catch (e5) {}
  }

  function normalizeParagraphText(paragraph) {
    if (!paragraph || !canFormatParagraph(paragraph)) return "";
    var text = "";
    try { text = trimText(paragraph.Range.Text); } catch (e1) {}
    if (!text) return "";
    try {
      var range = paragraph.Range.Duplicate;
      range.End = range.End - 1;
      if (trimText(range.Text) !== text) range.Text = text;
    } catch (e2) {}
    return text;
  }

  function applyNamedFormat(paragraph, name) {
    if (!paragraph || !canFormatParagraph(paragraph)) return;
    if (name === "title" || name === "attachmentTitle") {
      setFont(paragraph.Range, CONST.FONT_TITLE, CONST.FONT_WEST, CONST.SIZE_TITLE, true);
      setParagraph(paragraph, CONST.WD_ALIGN_CENTER, 0, 0, 0);
      return;
    }
    if (name === "mainRecipient" || name === "attachmentNote") {
      var noteIndent = name === "attachmentNote" ? attachmentNoteHangingIndent(paragraph.Range.Text) : 0;
      setFont(paragraph.Range, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
      if (name === "attachmentNote") {
        setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, CONST.CHAR_PT * 2 + noteIndent, -noteIndent, 0);
      } else {
        setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, noteIndent, -noteIndent, 0);
      }
      return;
    }
    if (name === "attachmentNoteContinuation") {
      setFont(paragraph.Range, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
      setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, attachmentNoteContinuationIndent(), 0, 0);
      return;
    }
    if (name === "level1") {
      setFont(paragraph.Range, CONST.FONT_LEVEL1, CONST.FONT_WEST, CONST.SIZE_BODY, false);
      setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, 0, CONST.CHAR_PT * 2, 0);
      return;
    }
    if (name === "level2") {
      setFont(paragraph.Range, CONST.FONT_LEVEL2, CONST.FONT_WEST, CONST.SIZE_BODY, true);
      setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, 0, CONST.CHAR_PT * 2, 0);
      return;
    }
    if (name === "level3") {
      setFont(paragraph.Range, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, true);
      setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, 0, CONST.CHAR_PT * 2, 0);
      return;
    }
    if (name === "level4" || name === "body") {
      setFont(paragraph.Range, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
      setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, 0, CONST.CHAR_PT * 2, 0);
      return;
    }
    if (name === "date") {
      normalizeParagraphText(paragraph);
      setFont(paragraph.Range, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
      setParagraph(paragraph, CONST.WD_ALIGN_RIGHT, 0, 0, CONST.CHAR_PT * 4);
      return;
    }
    if (name === "attachmentSequence") {
      setFont(paragraph.Range, CONST.FONT_LEVEL1, CONST.FONT_WEST, CONST.SIZE_BODY, false);
      setParagraph(paragraph, CONST.WD_ALIGN_LEFT, 0, 0, 0);
    }
  }

  function pageContentWidth(doc) {
    try {
      return doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin;
    } catch (e1) {
      return 451.3;
    }
  }

  function signatureLeftIndent(signatureText, dateText, usableWidth) {
    var rightBoundary = usableWidth - CONST.CHAR_PT * 4;
    var dateWidth = charWidth(dateText) * CONST.CHAR_PT;
    var dateLeft = rightBoundary - dateWidth;
    var dateCenter = (dateLeft + rightBoundary) / 2;
    var signatureWidth = charWidth(signatureText) * CONST.CHAR_PT;
    var leftIndent = dateCenter - signatureWidth / 2;
    return leftIndent < 0 ? 0 : leftIndent;
  }

  function applySignatureFormat(paragraph, dateText, doc) {
    if (!paragraph || !canFormatParagraph(paragraph)) return;
    var signature = normalizeParagraphText(paragraph);
    var date = findDateText(dateText) || trimText(dateText);
    var usable = pageContentWidth(doc);
    var leftIndent = signatureLeftIndent(signature, date, usable);
    setFont(paragraph.Range, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
    setParagraph(paragraph, CONST.WD_ALIGN_LEFT, leftIndent, 0, 0);
  }

  function recognizedDateText(doc) {
    if (!doc) return "";
    var items = documentItems(doc);
    var result = classifyDocument(items);
    if (result.dateIndex >= 0 && items[result.dateIndex]) return items[result.dateIndex].text;
    return "";
  }

  function clearParagraph(paragraph) {
    if (!paragraph || !canFormatParagraph(paragraph)) return;
    try { paragraph.Range.Font.Reset(); } catch (e1) {}
    try { paragraph.Range.ParagraphFormat.Reset(); } catch (e2) {}
    try { paragraph.Range.HighlightColorIndex = 0; } catch (e3) {}
  }

  function applyYiShi(paragraph) {
    if (!paragraph || !canFormatParagraph(paragraph)) return;
    var rawText = "";
    try { rawText = String(paragraph.Range.Text || ""); } catch (e1) {}
    rawText = rawText.replace(/[\r\n\u0007\f]+$/g, "");
    var matches = yiShiMatches(rawText);
    if (!matches.length) return;
    applyNamedFormat(paragraph, "body");
    try {
      var start = paragraph.Range.Start;
      for (var i = 0; i < matches.length; i++) {
        var marker = paragraph.Range.Duplicate;
        marker.SetRange(start + matches[i].start, start + matches[i].start + matches[i].length);
        setFont(marker, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, true);
      }
    } catch (e2) {}
  }

  function hostApplication() {
    try {
      if (typeof wps !== "undefined" && wps.WpsApplication) return wps.WpsApplication();
    } catch (e1) {}
    try {
      if (typeof Application !== "undefined" && Application.ActiveDocument) return Application;
    } catch (e2) {}
    try {
      if (typeof Application !== "undefined" && Application.WpsApplication) return Application.WpsApplication();
    } catch (e3) {}
    return Application;
  }

  function currentDocument() {
    return hostApplication().ActiveDocument;
  }

  function documentItems(doc) {
    var paragraphs = collectionToArray(doc.Paragraphs);
    var items = [];
    for (var i = 0; i < paragraphs.length; i++) {
      items.push(paragraphMeta(paragraphs[i], i));
    }
    return items;
  }

  function nextNonEmptyIndex(items, start) {
    for (var i = start; i < items.length; i++) {
      if (trimText(items[i].text)) return i;
    }
    return -1;
  }

  function previousNonEmptyIndex(items, start) {
    for (var i = start; i >= 0; i--) {
      if (trimText(items[i].text)) return i;
    }
    return -1;
  }

  function deleteParagraph(paragraph) {
    try {
      paragraph.Range.Delete();
      return true;
    } catch (e1) {}
    return false;
  }

  function insertParagraphAfter(paragraph) {
    try {
      paragraph.Range.InsertParagraphAfter();
      return true;
    } catch (e1) {}
    return false;
  }

  function insertParagraphBefore(paragraph) {
    try {
      paragraph.Range.InsertParagraphBefore();
      return true;
    } catch (e1) {}
    return false;
  }

  function normalizeBlankLinesBeforeTitle(items, result) {
    if (!result.titleIndexes.length) return false;
    var titleIndex = result.titleIndexes[0];
    var changed = false;
    for (var i = titleIndex - 1; i >= 0; i--) {
      if (!trimText(items[i].text) && !items[i].hasBreak) {
        if (deleteParagraph(items[i].paragraph)) changed = true;
        continue;
      }
      break;
    }
    if (insertParagraphBefore(items[titleIndex].paragraph)) changed = true;
    if (insertParagraphBefore(items[titleIndex].paragraph)) changed = true;
    return changed;
  }

  function formatBlankLinesBeforeTitle(items, result) {
    if (!result.titleIndexes.length) return;
    var firstTitleIndex = result.titleIndexes[0];
    for (var i = firstTitleIndex - 1; i >= 0 && i >= firstTitleIndex - 2; i--) {
      if (trimText(items[i].text)) continue;
      if (!canFormatParagraph(items[i].paragraph)) continue;
      setFont(items[i].paragraph.Range, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
      setParagraph(items[i].paragraph, CONST.WD_ALIGN_LEFT, 0, 0, 0);
    }
  }

  function applyOpeningWithoutMainRecipient(items, result) {
    if (result.mainRecipientIndex >= 0 || !result.titleIndexes.length) return;
    for (var t = 0; t < result.titleIndexes.length; t++) {
      applyNamedFormat(items[result.titleIndexes[t]].paragraph, "title");
    }
    var start = result.titleIndexes[result.titleIndexes.length - 1] + 1;
    var end = result.dateIndex >= 0 ? result.dateIndex : items.length;
    for (var i = start; i < end; i++) {
      if (!items[i] || !trimText(items[i].text)) continue;
      if (i === result.signatureIndex) continue;
      if (result.attachmentNoteIndexes.indexOf(i) >= 0 || result.attachmentNoteContinuationIndexes.indexOf(i) >= 0) continue;
      if (result.attachmentSequenceIndexes.indexOf(i) >= 0 || result.attachmentTitleIndexes.indexOf(i) >= 0) continue;
      if (result.headings.hasOwnProperty(i)) {
        applyNamedFormat(items[i].paragraph, result.headings[i]);
      } else if (result.yiShiIndexes.indexOf(i) >= 0) {
        applyYiShi(items[i].paragraph);
      } else if (result.bodyIndexes.indexOf(i) >= 0) {
        applyNamedFormat(items[i].paragraph, "body");
      }
    }
  }

  function normalizeBlankLineAfterTitle(items, result) {
    if (!result.titleIndexes.length) return false;
    var titleIndex = result.titleIndexes[result.titleIndexes.length - 1];
    var targetIndex = result.mainRecipientIndex >= 0 ? result.mainRecipientIndex : nextNonEmptyIndex(items, titleIndex + 1);
    if (targetIndex < 0 || targetIndex <= titleIndex) return false;
    var changed = false;
    for (var i = targetIndex - 1; i > titleIndex; i--) {
      if (!trimText(items[i].text) && !items[i].hasBreak) {
        if (deleteParagraph(items[i].paragraph)) changed = true;
      }
    }
    if (insertParagraphAfter(items[titleIndex].paragraph)) changed = true;
    return changed;
  }

  function formatBlankLineAfterTitle(items, result) {
    if (!result.titleIndexes.length) return;
    var blankIndex = result.titleIndexes[result.titleIndexes.length - 1] + 1;
    if (blankIndex >= items.length || trimText(items[blankIndex].text)) return;
    if (!canFormatParagraph(items[blankIndex].paragraph)) return;
    setFont(items[blankIndex].paragraph.Range, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
    setParagraph(items[blankIndex].paragraph, CONST.WD_ALIGN_LEFT, 0, 0, 0);
  }

  function normalizeBlankLineBeforeSignature(items, result) {
    if (result.signatureIndex < 0 || result.dateIndex < 0) return false;
    var previousIndex = previousNonEmptyIndex(items, result.signatureIndex - 1);
    if (previousIndex < 0) return false;
    var changed = false;
    for (var i = result.signatureIndex - 1; i > previousIndex; i--) {
      if (!trimText(items[i].text) && !items[i].hasBreak) {
        if (deleteParagraph(items[i].paragraph)) changed = true;
      }
    }
    if (insertParagraphAfter(items[previousIndex].paragraph)) changed = true;
    if (insertParagraphAfter(items[previousIndex].paragraph)) changed = true;
    return changed;
  }

  function formatBlankLineBeforeSignature(items, result) {
    if (result.signatureIndex < 0 || result.dateIndex < 0) return;
    for (var i = result.signatureIndex - 1; i >= 0 && i >= result.signatureIndex - 2; i--) {
      if (trimText(items[i].text)) continue;
      if (!canFormatParagraph(items[i].paragraph)) continue;
      setFont(items[i].paragraph.Range, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
      setParagraph(items[i].paragraph, CONST.WD_ALIGN_LEFT, 0, 0, 0);
    }
  }

  function footerByIndex(section, index) {
    try { return section.Footers.Item(index); } catch (e1) {}
    try { return section.Footers(index); } catch (e2) {}
    return null;
  }

  function footerIsEnabled(section, index) {
    if (index === CONST.WD_HEADER_FOOTER_PRIMARY) return true;
    try {
      if (index === CONST.WD_HEADER_FOOTER_FIRST_PAGE) {
        return !!section.PageSetup.DifferentFirstPageHeaderFooter;
      }
      if (index === CONST.WD_HEADER_FOOTER_EVEN_PAGES) {
        if (section.PageSetup.OddAndEvenPagesHeaderFooter) return true;
      }
    } catch (e1) {}
    try {
      if (index === CONST.WD_HEADER_FOOTER_EVEN_PAGES) {
        return !!hostApplication().Options.OddAndEvenPagesHeaderFooter;
      }
    } catch (e2) {}
    return false;
  }

  function isPageField(field) {
    try {
      if (Number(field.Type) === CONST.WD_FIELD_PAGE) return true;
    } catch (e1) {}
    try {
      return /\bPAGE\b/i.test(String(field.Code.Text || ""));
    } catch (e2) {}
    return false;
  }

  function removeExistingPageFields(footer) {
    var fields = [];
    try { fields = collectionToArray(footer.Range.Fields); } catch (e1) {}
    for (var i = fields.length - 1; i >= 0; i--) {
      if (!isPageField(fields[i])) continue;
      try { fields[i].Delete(); } catch (e2) {}
    }
  }

  function formatPageField(field) {
    if (!field) return;
    try {
      var pageRange = field.Result;
      setWesternFont(pageRange, CONST.FONT_WEST);
      pageRange.Font.Bold = false;
      pageRange.ParagraphFormat.Alignment = CONST.WD_ALIGN_CENTER;
    } catch (e1) {}
  }

  function applyCenteredPageNumbers(doc) {
    var sections = collectionToArray(doc.Sections);
    var footerIndexes = [
      CONST.WD_HEADER_FOOTER_PRIMARY,
      CONST.WD_HEADER_FOOTER_FIRST_PAGE,
      CONST.WD_HEADER_FOOTER_EVEN_PAGES
    ];
    for (var i = 0; i < sections.length; i++) {
      for (var j = 0; j < footerIndexes.length; j++) {
        var footerIndex = footerIndexes[j];
        if (!footerIsEnabled(sections[i], footerIndex)) continue;
        var footer = footerByIndex(sections[i], footerIndex);
        if (!footer) continue;
        removeExistingPageFields(footer);
        try {
          var pageNumbers = footer.PageNumbers;
          try { pageNumbers.RestartNumberingAtSection = false; } catch (e1) {}
          pageNumbers.Add(CONST.WD_ALIGN_PAGE_NUMBER_CENTER, true);
        } catch (e2) {}
        try {
          var fields = collectionToArray(footer.Range.Fields);
          for (var f = 0; f < fields.length; f++) {
            if (isPageField(fields[f])) formatPageField(fields[f]);
          }
        } catch (e3) {}
      }
    }
  }

  function askYesNo(message, title) {
    try {
      var answer = MsgBox(message, 4, title || "公文格式");
      return answer === 6 || answer === "Yes" || answer === true;
    } catch (e1) {}
    try {
      return hostApplication().Confirm(message);
    } catch (e2) {}
    try {
      if (typeof confirm !== "undefined") return confirm(message);
    } catch (e3) {}
    return false;
  }

  function alertMessage(message) {
    try { MsgBox(message, 0, "公文格式"); return; } catch (e1) {}
    try { hostApplication().Alert(message); return; } catch (e2) {}
    try { if (typeof alert !== "undefined") alert(message); return; } catch (e3) {}
  }

  function formatWholeDocument() {
    ensureOfficialDocumentFormatter();
    if (!askYesNo("将调整标题、空行、正文、落款和页码，是否继续？", "公文格式")) return;
    var doc = currentDocument();
    var items = documentItems(doc);
    var result = classifyDocument(items);
    if (normalizeBlankLinesBeforeTitle(items, result)) {
      items = documentItems(doc);
      result = classifyDocument(items);
    }
    if (normalizeBlankLineAfterTitle(items, result)) {
      items = documentItems(doc);
      result = classifyDocument(items);
    }
    if (normalizeBlankLineBeforeSignature(items, result)) {
      items = documentItems(doc);
      result = classifyDocument(items);
    }
    if (askYesNo("是否先清除全文格式？", "公文格式")) {
      for (var c = 0; c < items.length; c++) clearParagraph(items[c].paragraph);
      items = documentItems(doc);
      result = classifyDocument(items);
    }
    items = documentItems(doc);
    result = classifyDocument(items);

    for (var i = 0; i < result.titleIndexes.length; i++) applyNamedFormat(items[result.titleIndexes[i]].paragraph, "title");
    formatBlankLinesBeforeTitle(items, result);
    formatBlankLineAfterTitle(items, result);
    if (result.mainRecipientIndex >= 0) applyNamedFormat(items[result.mainRecipientIndex].paragraph, "mainRecipient");
    if (result.dateIndex >= 0) applyNamedFormat(items[result.dateIndex].paragraph, "date");
    if (result.signatureIndex >= 0 && result.dateIndex >= 0) {
      applySignatureFormat(items[result.signatureIndex].paragraph, items[result.dateIndex].text, doc);
    }
    formatBlankLineBeforeSignature(items, result);
    for (var n = 0; n < result.attachmentNoteIndexes.length; n++) {
      applyNamedFormat(items[result.attachmentNoteIndexes[n]].paragraph, "attachmentNote");
    }
    for (var an = 0; an < result.attachmentNoteContinuationIndexes.length; an++) {
      applyNamedFormat(items[result.attachmentNoteContinuationIndexes[an]].paragraph, "attachmentNoteContinuation");
    }
    for (var s = 0; s < result.attachmentSequenceIndexes.length; s++) {
      applyNamedFormat(items[result.attachmentSequenceIndexes[s]].paragraph, "attachmentSequence");
    }
    for (var t = 0; t < result.attachmentTitleIndexes.length; t++) {
      applyNamedFormat(items[result.attachmentTitleIndexes[t]].paragraph, "attachmentTitle");
    }
    for (var key in result.headings) {
      if (result.headings.hasOwnProperty(key)) applyNamedFormat(items[Number(key)].paragraph, result.headings[key]);
    }
    for (var b = 0; b < result.bodyIndexes.length; b++) {
      applyNamedFormat(items[result.bodyIndexes[b]].paragraph, "body");
    }
    for (var y = 0; y < result.yiShiIndexes.length; y++) {
      applyYiShi(items[result.yiShiIndexes[y]].paragraph);
    }
    applyOpeningWithoutMainRecipient(items, result);
    applyCenteredPageNumbers(doc);
    alertMessage("格式化完成");
  }

  function targetParagraphs() {
    try {
      var selection = hostApplication().Selection;
      var range = selection.Range;
      if (range && range.Start !== range.End) return collectionToArray(range.Paragraphs);
    } catch (e1) {}
    return [];
  }

  function applyToSelection(formatName) {
    ensureOfficialDocumentFormatter();
    var paragraphs = targetParagraphs();
    if (!paragraphs.length) {
      alertMessage("请先选中要设置格式的文字或段落。");
      return;
    }
    applyToParagraphs(paragraphs, formatName, currentDocument());
  }

  function applyToParagraphs(paragraphs, formatName, doc) {
    var dateText = formatName === "signature" ? recognizedDateText(doc) : "";
    for (var i = 0; i < paragraphs.length; i++) {
      if (formatName === "yiShi") applyYiShi(paragraphs[i]);
      else if (formatName === "signature") {
        applySignatureFormat(paragraphs[i], dateText, doc);
      } else {
        applyNamedFormat(paragraphs[i], formatName);
      }
    }
  }

  function FormatOfficialDocument() { formatWholeDocument(); }
  function ApplyTitleFormat() { applyToSelection("title"); }
  function ApplyMainRecipientFormat() { applyToSelection("mainRecipient"); }
  function ApplyBodyFormat() { applyToSelection("body"); }
  function ApplyLevel1HeadingFormat() { applyToSelection("level1"); }
  function ApplyLevel2HeadingFormat() { applyToSelection("level2"); }
  function ApplyLevel3HeadingFormat() { applyToSelection("level3"); }
  function ApplyLevel4HeadingFormat() { applyToSelection("level4"); }
  function ApplyYiShiFormat() { applyToSelection("yiShi"); }
  function ApplyDateFormat() { applyToSelection("date"); }
  function ApplySignatureFormat() { applyToSelection("signature"); }
  function ApplyAttachmentNoteFormat() { applyToSelection("attachmentNote"); }
  function ApplyAttachmentNoteContinuationFormat() { applyToSelection("attachmentNoteContinuation"); }
  function ApplyAttachmentSequenceFormat() { applyToSelection("attachmentSequence"); }
  function ApplyAttachmentTitleFormat() { applyToSelection("attachmentTitle"); }
  function ApplyPageNumberFormat() {
    ensureOfficialDocumentFormatter();
    applyCenteredPageNumbers(currentDocument());
    alertMessage("页码格式化完成");
  }

  function GetOfficialDocumentFormatter() {
    ensureOfficialDocumentFormatter();
    return OfficialDocumentFormatter;
  }

  var OfficialDocumentConsole = {
    targetPath: "",
    targetDoc: null,
    selectionStart: -1,
    selectionEnd: -1,
    selectionDocPath: "",
    lastInputAvailable: false
  };

  function consoleInput(message, defaultValue) {
    OfficialDocumentConsole.lastInputAvailable = false;
    try {
      var globalInputValue = InputBox(message, "公文格式控制台", defaultValue || "");
      OfficialDocumentConsole.lastInputAvailable = true;
      return globalInputValue;
    } catch (e1) {}
    try {
      var applicationInputValue = hostApplication().InputBox(message, "公文格式控制台", defaultValue || "");
      OfficialDocumentConsole.lastInputAvailable = true;
      return applicationInputValue;
    } catch (e2) {}
    return "";
  }

  function consoleSelectedFile(dialog) {
    try { return String(dialog.SelectedItems.Item(1)); } catch (e1) {}
    try { return String(dialog.SelectedItems(1)); } catch (e2) {}
    try { return String(dialog.SelectedItems[0]); } catch (e3) {}
    return "";
  }

  function currentFolderPath() {
    try {
      var doc = currentDocument();
      if (doc && doc.Path) return String(doc.Path);
    } catch (e1) {}
    return "";
  }

  function pickTargetDocumentPath() {
    var app = hostApplication();
    var fileDialogAvailable = false;
    try {
      var dialog = app.FileDialog(3);
      if (!dialog) throw new Error("FileDialog unavailable");
      fileDialogAvailable = true;
      dialog.Title = "选择需要格式化的公文";
      dialog.AllowMultiSelect = false;
      try {
        dialog.Filters.Clear();
        dialog.Filters.Add("Word/WPS 文档", "*.docx;*.doc;*.wps;*.wpt");
      } catch (filterError) {}
      try {
        var folder = currentFolderPath();
        if (folder) dialog.InitialFileName = folder;
      } catch (folderError) {}
      var result = dialog.Show();
      if (result === -1 || result === true || result === 1) return consoleSelectedFile(dialog);
    } catch (e1) {}
    var path = consoleInput("请输入目标文档完整路径：", OfficialDocumentConsole.targetPath || "");
    if (!path && !fileDialogAvailable && !OfficialDocumentConsole.lastInputAvailable) {
      alertMessage("当前 WPS 不支持文件选择或路径输入。请先手动打开目标文档，再运行“当前文档设为目标”。");
    }
    return path;
  }

  function normalizePath(path) {
    return String(path || "").replace(/^"|"$/g, "").replace(/\\/g, "/").replace(/^\s+|\s+$/g, "");
  }

  function samePath(a, b) {
    var left = normalizePath(a);
    var right = normalizePath(b);
    if (/^[A-Za-z]:\//.test(left) && /^[A-Za-z]:\//.test(right)) {
      return left.toLowerCase() === right.toLowerCase();
    }
    return left === right;
  }

  function documentPath(doc) {
    try { return normalizePath(doc.FullName); } catch (e1) {}
    return "";
  }

  function documentName(doc) {
    try { return String(doc.Name || doc.FullName || "未命名文档"); } catch (e1) {}
    return "未命名文档";
  }

  function sameDocument(left, right) {
    if (!left || !right) return false;
    if (left === right) return true;
    var leftPath = documentPath(left);
    var rightPath = documentPath(right);
    return !!leftPath && !!rightPath && samePath(leftPath, rightPath);
  }

  function clearSavedSelection() {
    OfficialDocumentConsole.selectionStart = -1;
    OfficialDocumentConsole.selectionEnd = -1;
    OfficialDocumentConsole.selectionDocPath = "";
  }

  function activeSelectionRange() {
    try { return hostApplication().Selection.Range; } catch (e1) {}
    return null;
  }

  function rangeBelongsToDocument(range, doc) {
    if (!range || !doc) return false;
    try {
      if (range.Document) return sameDocument(range.Document, doc);
    } catch (e1) {}
    return sameDocument(currentDocument(), doc);
  }

  function captureSelectionForDocument(doc, showMessage) {
    if (!doc) return false;
    try { doc.Activate(); } catch (e1) {}
    var range = activeSelectionRange();
    if (!range || !rangeBelongsToDocument(range, doc)) {
      alertMessage("当前选区不属于目标文档。请切换到目标文档后重新选择段落。");
      return false;
    }
    var start = -1;
    var end = -1;
    try {
      start = Number(range.Start);
      end = Number(range.End);
    } catch (e2) {}
    if (start < 0 || end <= start) {
      alertMessage("请先选中至少一个字符或完整段落，再记录选区。");
      return false;
    }
    OfficialDocumentConsole.selectionStart = start;
    OfficialDocumentConsole.selectionEnd = end;
    OfficialDocumentConsole.selectionDocPath = documentPath(doc);
    if (showMessage) alertMessage("已记录目标文档当前选区。现在可切换回控制台执行局部格式按钮。");
    return true;
  }

  function savedSelectionParagraphs(doc) {
    if (OfficialDocumentConsole.selectionStart < 0 ||
        OfficialDocumentConsole.selectionEnd <= OfficialDocumentConsole.selectionStart) return [];
    var savedPath = OfficialDocumentConsole.selectionDocPath;
    if (savedPath && !samePath(savedPath, documentPath(doc))) return [];
    try {
      var range = doc.Range(OfficialDocumentConsole.selectionStart, OfficialDocumentConsole.selectionEnd);
      return collectionToArray(range.Paragraphs);
    } catch (e1) {
      clearSavedSelection();
      return [];
    }
  }

  function findOpenDocumentByPath(path) {
    var docs;
    try { docs = hostApplication().Documents; } catch (e1) { return null; }
    var count = 0;
    try { count = docs.Count; } catch (e2) { return null; }
    for (var i = 1; i <= count; i++) {
      var doc = null;
      try { doc = docs.Item(i); } catch (itemError) {}
      if (!doc) continue;
      try {
        if (samePath(doc.FullName, path)) return doc;
      } catch (fullNameError) {}
    }
    return null;
  }

  function openTargetDocument() {
    var path = normalizePath(OfficialDocumentConsole.targetPath);
    if (!path) {
      path = normalizePath(pickTargetDocumentPath());
      OfficialDocumentConsole.targetPath = path;
    }
    if (!path) {
      alertMessage("未选择目标文档。");
      return null;
    }
    var doc = findOpenDocumentByPath(path);
    if (!doc) {
      try {
        doc = hostApplication().Documents.Open(path);
      } catch (openError) {
        alertMessage("目标文档打开失败：" + openError.message);
        return null;
      }
    }
    OfficialDocumentConsole.targetDoc = doc;
    try { doc.Activate(); } catch (activateError) {}
    return doc;
  }

  function targetDocument() {
    if (OfficialDocumentConsole.targetDoc) {
      try {
        OfficialDocumentConsole.targetDoc.Activate();
        return OfficialDocumentConsole.targetDoc;
      } catch (e1) {
        OfficialDocumentConsole.targetDoc = null;
      }
    }
    return openTargetDocument();
  }

  function applyTargetSelection(formatName) {
    var doc = targetDocument();
    if (!doc) return;
    var paragraphs = savedSelectionParagraphs(doc);
    if (!paragraphs.length) {
      var range = activeSelectionRange();
      if (range && rangeBelongsToDocument(range, doc)) {
        paragraphs = collectionToArray(range.Paragraphs);
        captureSelectionForDocument(doc, false);
      }
    }
    if (!paragraphs.length) {
      alertMessage("未记录目标文档选区。请在目标文档选中段落后，运行“记录当前选区”，再点击格式按钮。");
      return;
    }
    applyToParagraphs(paragraphs, formatName, doc);
  }

  function ConsoleSelectTargetDocument() {
    var path = normalizePath(pickTargetDocumentPath());
    if (!path) {
      alertMessage("未选择目标文档。");
      return;
    }
    OfficialDocumentConsole.targetPath = path;
    OfficialDocumentConsole.targetDoc = null;
    clearSavedSelection();
    alertMessage("已选择目标文档：\n" + path);
  }

  function ConsoleOpenTargetDocument() {
    if (openTargetDocument()) alertMessage("目标文档已打开并切换到前台。");
  }

  function ConsoleUseActiveDocument() {
    var doc = currentDocument();
    if (!doc) {
      alertMessage("未找到当前活动文档。");
      return;
    }
    var name = documentName(doc);
    if (/公文格式化控制台/i.test(name)) {
      alertMessage("当前活动文档疑似控制台模板，不能设为目标。请切换到需要格式化的公文，或使用“选择目标文档”。");
      return;
    }
    if (!askYesNo("确认将以下当前活动文档设为目标？\n" + name, "公文格式控制台")) return;
    OfficialDocumentConsole.targetDoc = doc;
    OfficialDocumentConsole.targetPath = documentPath(doc);
    clearSavedSelection();
    var captured = captureSelectionForDocument(doc, false);
    alertMessage(captured ?
      "已将当前活动文档设为目标文档，并记录了当前选区。" :
      "已将当前活动文档设为目标文档。需要局部格式化时，请先选中段落并记录当前选区。");
  }

  function ConsoleCaptureTargetSelection() {
    var doc = OfficialDocumentConsole.targetDoc;
    if (!doc) {
      alertMessage("请先选择目标文档，或在目标文档中运行“当前文档设为目标”。");
      return;
    }
    captureSelectionForDocument(doc, true);
  }

  function ConsoleFormatTargetDocument() {
    var doc = targetDocument();
    if (!doc) return;
    formatWholeDocument();
  }

  function ConsoleApplyTitleFormat() { applyTargetSelection("title"); }
  function ConsoleApplyMainRecipientFormat() { applyTargetSelection("mainRecipient"); }
  function ConsoleApplyBodyFormat() { applyTargetSelection("body"); }
  function ConsoleApplyLevel1HeadingFormat() { applyTargetSelection("level1"); }
  function ConsoleApplyLevel2HeadingFormat() { applyTargetSelection("level2"); }
  function ConsoleApplyLevel3HeadingFormat() { applyTargetSelection("level3"); }
  function ConsoleApplyLevel4HeadingFormat() { applyTargetSelection("level4"); }
  function ConsoleApplyYiShiFormat() { applyTargetSelection("yiShi"); }
  function ConsoleApplyDateFormat() { applyTargetSelection("date"); }
  function ConsoleApplySignatureFormat() { applyTargetSelection("signature"); }
  function ConsoleApplyAttachmentNoteFormat() { applyTargetSelection("attachmentNote"); }
  function ConsoleApplyAttachmentNoteContinuationFormat() { applyTargetSelection("attachmentNoteContinuation"); }
  function ConsoleApplyAttachmentSequenceFormat() { applyTargetSelection("attachmentSequence"); }
  function ConsoleApplyAttachmentTitleFormat() { applyTargetSelection("attachmentTitle"); }
  function ConsoleApplyPageNumberFormat() {
    var doc = targetDocument();
    if (!doc) return;
    applyCenteredPageNumbers(doc);
    alertMessage("目标文档页码格式化完成");
  }

  function ConsoleRunMenu() {
    var menu =
      "请输入操作序号：\n" +
      "0  将当前活动文档设为目标\n" +
      "1  选择目标文档\n" +
      "2  打开目标文档\n" +
      "3  全文格式化\n" +
      "4  大标题\n" +
      "5  主送单位\n" +
      "6  正文\n" +
      "7  一级标题\n" +
      "8  二级标题\n" +
      "9  三级标题\n" +
      "10 四级标题\n" +
      "11 一是/二是加粗\n" +
      "12 成文日期\n" +
      "13 落款单位\n" +
      "14 附件说明\n" +
      "15 附件说明续行\n" +
      "16 附件序号\n" +
      "17 附件标题\n" +
      "18 居中页码\n" +
      "19 记录当前选区";
    var choice = trimText(consoleInput(menu, ""));
    if (!choice) return;
    if (choice === "0") ConsoleUseActiveDocument();
    else if (choice === "1") ConsoleSelectTargetDocument();
    else if (choice === "2") ConsoleOpenTargetDocument();
    else if (choice === "3") ConsoleFormatTargetDocument();
    else if (choice === "4") ConsoleApplyTitleFormat();
    else if (choice === "5") ConsoleApplyMainRecipientFormat();
    else if (choice === "6") ConsoleApplyBodyFormat();
    else if (choice === "7") ConsoleApplyLevel1HeadingFormat();
    else if (choice === "8") ConsoleApplyLevel2HeadingFormat();
    else if (choice === "9") ConsoleApplyLevel3HeadingFormat();
    else if (choice === "10") ConsoleApplyLevel4HeadingFormat();
    else if (choice === "11") ConsoleApplyYiShiFormat();
    else if (choice === "12") ConsoleApplyDateFormat();
    else if (choice === "13") ConsoleApplySignatureFormat();
    else if (choice === "14") ConsoleApplyAttachmentNoteFormat();
    else if (choice === "15") ConsoleApplyAttachmentNoteContinuationFormat();
    else if (choice === "16") ConsoleApplyAttachmentSequenceFormat();
    else if (choice === "17") ConsoleApplyAttachmentTitleFormat();
    else if (choice === "18") ConsoleApplyPageNumberFormat();
    else if (choice === "19") ConsoleCaptureTargetSelection();
    else alertMessage("未知操作序号：" + choice);
  }
