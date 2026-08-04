/* \u516c\u6587\u683c\u5f0f\u63a7\u5236\u53f0 WPS JS \u5b8f
 * \u72ec\u7acb\u63a7\u5236\u53f0\u65b9\u6848\uff1a\u6267\u884c\u65f6\u9009\u62e9\u76ee\u6807\u6587\u6863\uff0c\u6216\u628a\u5f53\u524d\u6d3b\u52a8\u6587\u6863\u8bbe\u4e3a\u76ee\u6807\u3002
 * \u539f Ribbon \u65b9\u6848\u4ecd\u4fdd\u7559\u5728\u201c\u65b9\u6848A_ribbonUI\u63d2\u4ef6\u201d\uff0c\u4e24\u5957\u811a\u672c\u4e92\u4e0d\u8986\u76d6\u3002
 */
  /* SHARED_CORE_START */
  var CONST;
  var OfficialDocumentFormatter;

  function ensureOfficialDocumentFormatter() {
    if (CONST) return;
    CONST = {};
    CONST.FONT_BODY = "\u4eff\u5b8b";
    CONST.FONT_TITLE = "\u5b8b\u4f53";
    CONST.FONT_LEVEL1 = "\u9ed1\u4f53";
    CONST.FONT_LEVEL2 = "\u6977\u4f53";
    CONST.FONT_WEST = "Times New Roman";
    CONST.FONT_PAGE_NUMBER = "\u5b8b\u4f53";
    CONST.SIZE_BODY = 16;
    CONST.SIZE_TITLE = 22;
    CONST.SIZE_PAGE_NUMBER = 14;
    CONST.PAGE_NUMBER_GAP_MM = 7;
    CONST.PAGE_NUMBER_TOP_OFFSET_PT = 11;
    CONST.PAGE_NUMBER_INDENT_PT = 14;
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
    CONST.WD_ALIGN_PAGE_NUMBER_LEFT = 0;
    CONST.WD_ALIGN_PAGE_NUMBER_RIGHT = 2;
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
    return /[:\uff1a]$/.test(trimText(text));
  }

  function endsWithPeriod(text) {
    return /[\u3002.]$/.test(trimText(text));
  }

  function hasSerialPrefix(text) {
    return /^([\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e]+\u3001|[\uff08(][\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e]+[\uff09)]|\d+[.\uff0e]|[\uff08(]\d+[\uff09)])/.test(trimText(text));
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
    if (/^[\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e]+\u3001/.test(s)) return "level1";
    if (/^[\uff08(][\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341\u767e]+[\uff09)]/.test(s)) return "level2";
    if (/^\d+[.\uff0e]/.test(s)) return "level3";
    if (/^[\uff08(]\d+[\uff09)]/.test(s)) return "level4";
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
    var re = /(^|[\uff0c\u3002\uff1b;\u3001\uff1a:\s])([\u4e00\u4e8c\u4e09\u56db\u4e94\u516d\u4e03\u516b\u4e5d\u5341]+\u662f)(?!\u5426)/g;
    var match;
    while ((match = re.exec(s)) !== null) {
      matches.push({ start: match.index + match[1].length, length: match[2].length, text: match[2] });
      if (match.index === re.lastIndex) re.lastIndex++;
    }
    return matches;
  }

  function dateLinePattern() {
    return "\\d{4}\u5e74(?:\\d{1,2}|[\u3010\\[]\\s*[\u3011\\]]|_+|\uff3f+)\u6708(?:\\d{1,2}|[\u3010\\[]\\s*[\u3011\\]]|_+|\uff3f+)\u65e5";
  }

  function isDateText(text) {
    return new RegExp("^" + dateLinePattern() + "$").test(trimText(text));
  }

  function findDateText(text) {
    var matches = trimText(text).match(new RegExp(dateLinePattern(), "g"));
    return matches && matches.length ? matches[matches.length - 1] : "";
  }

  function isAttachmentNote(text) {
    return /^\u9644\u4ef6(?:\d+|\s|[:\uff1a])/.test(trimText(text));
  }

  function isAttachmentNoteContinuation(text) {
    return /^\d+(?:\s+|[.\uff0e\u3001])\S+/.test(trimText(text));
  }

  function isAttachmentSequence(text) {
    return /^\u9644\u4ef6\s*\d*$/.test(trimText(text));
  }

  function attachmentNoteHangingIndent(text) {
    var s = trimText(text);
    var match = s.match(/^\u9644\u4ef6(?:[:\uff1a]|\s*)/);
    if (!match) return 0;
    return charWidth(match[0]) * CONST.CHAR_PT;
  }

  function attachmentNoteContinuationIndent() {
    return CONST.CHAR_PT * 2 + charWidth("\u9644\u4ef6\uff1a") * CONST.CHAR_PT;
  }

  function paragraphMeta(paragraph, index) {
    var text = "";
    var align = CONST.WD_ALIGN_LEFT;
    var leftIndent = 0;
    var hasBreak = false;
    try {
      text = paragraph.Range.Text;
      hasBreak = String(text).indexOf("\f") >= 0;
    } catch (e1) {}
    try {
      align = paragraph.Range.ParagraphFormat.Alignment;
      leftIndent = paragraph.Range.ParagraphFormat.LeftIndent;
    } catch (e2) {
      try {
        align = paragraph.Alignment;
      } catch (e3) {}
    }
    return {
      paragraph: paragraph,
      index: index,
      text: trimText(text),
      rawText: text,
      align: align,
      leftIndent: leftIndent,
      hasBreak: hasBreak
    };
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
    if (indexes.length) return indexes;
    indexes = findTitleIndexesWithMode(items, true);
    if (!indexes.length || indexes[0] !== firstNonEmptyItemIndex(items)) return [];
    if (!hasMainRecipientImmediatelyAfter(items, indexes)) return [];
    return indexes;
  }

  function firstNonEmptyItemIndex(items) {
    for (var i = 0; i < items.length; i++) {
      if (trimText(items[i].text)) return i;
    }
    return -1;
  }

  function hasMainRecipientImmediatelyAfter(items, titleIndexes) {
    var start = titleIndexes[titleIndexes.length - 1] + 1;
    for (var i = start; i < items.length; i++) {
      var text = trimText(items[i].text);
      if (!text) continue;
      return endsWithColon(text);
    }
    return false;
  }

  function findTitleIndexesWithMode(items, allowUncentered) {
    var indexes = [];
    var nonEmptySeen = 0;
    var started = false;
    var firstNonEmpty = firstNonEmptyItemIndex(items);
    if (firstNonEmpty < 0) return indexes;
    for (var i = 0; i < items.length && nonEmptySeen < 10; i++) {
      var text = trimText(items[i].text);
      if (!text) continue;
      nonEmptySeen++;
      if (!started) {
        if (i !== firstNonEmpty) return [];
        if (isTitleCandidate(items[i], false, allowUncentered)) {
          indexes.push(i);
          started = true;
        } else return [];
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

  function findDateIndex(items, endIndex) {
    var idx = -1;
    var end = typeof endIndex === "number" && endIndex >= 0 ? endIndex : items.length;
    for (var i = 0; i < end; i++) {
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

  function isTopLevelAttachmentSequence(meta) {
    return isAttachmentSequence(meta && meta.text);
  }

  function isAttachmentTitleCandidate(meta) {
    var text = trimText(meta && meta.text);
    var len = visibleLength(text);
    if (!text || len > 50) return false;
    if (isAttachmentSequence(text) || isDateText(text) || hasSerialPrefix(text)) return false;
    if (endsWithColon(text) || endsWithPeriod(text)) return false;
    return isCenteredParagraph(meta) || len >= 2;
  }

  function findAttachmentBody(items) {
    var sequenceIndexes = [];
    var titleIndexes = [];
    var afterBreak = false;
    for (var i = 0; i < items.length; i++) {
      if (items[i].hasBreak) afterBreak = true;
      var text = trimText(items[i].text);
      if (!text) continue;
      if (afterBreak && isTopLevelAttachmentSequence(items[i])) {
        sequenceIndexes.push(i);
        for (var j = i + 1; j < items.length; j++) {
          if (trimText(items[j].text)) {
            if (isAttachmentTitleCandidate(items[j])) titleIndexes.push(j);
            break;
          }
        }
        afterBreak = false;
        continue;
      }
      afterBreak = false;
    }
    return { sequenceIndexes: sequenceIndexes, titleIndexes: titleIndexes };
  }

  function classifyDocument(items) {
    ensureOfficialDocumentFormatter();
    var titleIndexes = findTitleIndexes(items);
    var mainRecipientIndex = findMainRecipientIndex(items, titleIndexes);
    var attachmentBody = findAttachmentBody(items);
    var attachmentStartIndex = attachmentBody.sequenceIndexes.length ? attachmentBody.sequenceIndexes[0] : items.length;
    var dateIndex = findDateIndex(items, attachmentStartIndex);
    var signatureIndex = findSignatureIndex(items, dateIndex);
    var attachmentNoteIndexes = findAttachmentNoteIndexes(items, dateIndex);
    var attachmentNoteContinuationIndexes = findAttachmentNoteContinuationIndexes(items, dateIndex, attachmentNoteIndexes);
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
      var rawText = String(paragraph.Range.Text || "");
      if (rawText.indexOf("\f") >= 0 && !trimText(rawText)) return false;
    } catch (e4) {}
    return true;
  }

  function paragraphTextRange(paragraph) {
    var baseRange = paragraph.Range;
    var rawText = "";
    try { rawText = String(baseRange.Text || ""); } catch (e1) { return baseRange; }
    if (rawText.indexOf("\f") < 0) return baseRange;

    var first = 0;
    var last = rawText.length;
    while (first < last && /[\r\n\u0007\f]/.test(rawText.charAt(first))) first++;
    while (last > first && /[\r\n\u0007\f]/.test(rawText.charAt(last - 1))) last--;
    if (first >= last) return null;

    try {
      var range = baseRange.Duplicate;
      var start = Number(baseRange.Start);
      if (isNaN(start)) start = 0;
      if (range.SetRange) range.SetRange(start + first, start + last);
      else {
        range.Start = start + first;
        range.End = start + last;
      }
      return range;
    } catch (e2) {
      return baseRange;
    }
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
    var rawText = "";
    var text = "";
    try {
      rawText = String(paragraph.Range.Text || "");
      text = trimText(rawText);
    } catch (e1) {}
    if (!text) return "";
    if (rawText.indexOf("\f") >= 0) return text;
    try {
      var range = paragraph.Range.Duplicate;
      range.End = range.End - 1;
      if (trimText(range.Text) !== text) range.Text = text;
    } catch (e2) {}
    return text;
  }

  function applyNamedFormat(paragraph, name) {
    if (!paragraph || !canFormatParagraph(paragraph)) return;
    var textRange = paragraphTextRange(paragraph);
    if (!textRange) return;
    if (name === "title" || name === "attachmentTitle") {
      setFont(textRange, CONST.FONT_TITLE, CONST.FONT_WEST, CONST.SIZE_TITLE, true);
      setParagraph(paragraph, CONST.WD_ALIGN_CENTER, 0, 0, 0);
      return;
    }
    if (name === "mainRecipient" || name === "attachmentNote") {
      var noteIndent = name === "attachmentNote" ? attachmentNoteHangingIndent(paragraph.Range.Text) : 0;
      setFont(textRange, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
      if (name === "attachmentNote") {
        setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, CONST.CHAR_PT * 2 + noteIndent, -noteIndent, 0);
      } else {
        setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, noteIndent, -noteIndent, 0);
      }
      return;
    }
    if (name === "attachmentNoteContinuation") {
      setFont(textRange, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
      setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, attachmentNoteContinuationIndent(), 0, 0);
      return;
    }
    if (name === "level1") {
      setFont(textRange, CONST.FONT_LEVEL1, CONST.FONT_WEST, CONST.SIZE_BODY, false);
      setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, 0, CONST.CHAR_PT * 2, 0);
      return;
    }
    if (name === "level2") {
      setFont(textRange, CONST.FONT_LEVEL2, CONST.FONT_WEST, CONST.SIZE_BODY, true);
      setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, 0, CONST.CHAR_PT * 2, 0);
      return;
    }
    if (name === "level3") {
      setFont(textRange, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, true);
      setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, 0, CONST.CHAR_PT * 2, 0);
      return;
    }
    if (name === "level4" || name === "body") {
      setFont(textRange, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
      setParagraph(paragraph, CONST.WD_ALIGN_JUSTIFY, 0, CONST.CHAR_PT * 2, 0);
      return;
    }
    if (name === "date") {
      normalizeParagraphText(paragraph);
      setFont(textRange, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
      setParagraph(paragraph, CONST.WD_ALIGN_RIGHT, 0, 0, CONST.CHAR_PT * 4);
      return;
    }
    if (name === "attachmentSequence") {
      setFont(textRange, CONST.FONT_LEVEL1, CONST.FONT_WEST, CONST.SIZE_BODY, false);
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
    var textRange = paragraphTextRange(paragraph);
    if (!textRange) return;
    setFont(textRange, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
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

  function applyYiShi(paragraph, preserveBaseFormat) {
    if (!paragraph || !canFormatParagraph(paragraph)) return;
    var rawText = "";
    try { rawText = String(paragraph.Range.Text || ""); } catch (e1) {}
    rawText = rawText.replace(/[\r\n\u0007\f]+$/g, "");
    var matches = yiShiMatches(rawText);
    if (!matches.length) return;
    if (!preserveBaseFormat) applyNamedFormat(paragraph, "body");
    try {
      var start = paragraph.Range.Start;
      for (var i = 0; i < matches.length; i++) {
        var marker = paragraph.Range.Duplicate;
        marker.SetRange(start + matches[i].start, start + matches[i].start + matches[i].length);
        if (preserveBaseFormat) marker.Font.Bold = true;
        else setFont(marker, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, true);
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

  function millimetersToPoints(value) {
    try {
      var converted = Number(hostApplication().MillimetersToPoints(value));
      if (isFinite(converted) && converted > 0) return converted;
    } catch (e1) {}
    return Number(value) * 72 / 25.4;
  }

  function configurePageNumberSection(section) {
    if (!section) return;
    var setup = null;
    try { setup = section.PageSetup; } catch (e1) {}
    if (!setup) return;
    try { setup.OddAndEvenPagesHeaderFooter = true; } catch (e2) {}
    var bottomMargin = millimetersToPoints(35);
    try {
      var currentBottomMargin = Number(setup.BottomMargin);
      if (isFinite(currentBottomMargin) && currentBottomMargin > 0) {
        bottomMargin = currentBottomMargin;
      }
    } catch (e3) {}
    var footerDistance = bottomMargin -
      millimetersToPoints(CONST.PAGE_NUMBER_GAP_MM) -
      CONST.PAGE_NUMBER_TOP_OFFSET_PT;
    if (footerDistance < 0) footerDistance = 0;
    try { setup.FooterDistance = footerDistance; } catch (e4) {}
  }

  function pageNumberAlignment(footerIndex) {
    return footerIndex === CONST.WD_HEADER_FOOTER_EVEN_PAGES ?
      CONST.WD_ALIGN_PAGE_NUMBER_LEFT :
      CONST.WD_ALIGN_PAGE_NUMBER_RIGHT;
  }

  function formatPageField(field, alignment) {
    if (!field) return;
    try {
      var pageRange = field.Result;
      pageRange.InsertBefore("\u2014 ");
      pageRange.InsertAfter(" \u2014");
      setFont(pageRange, CONST.FONT_PAGE_NUMBER, CONST.FONT_PAGE_NUMBER,
        CONST.SIZE_PAGE_NUMBER, false);
      var paragraphFormat = pageRange.ParagraphFormat;
      paragraphFormat.Alignment = alignment;
      paragraphFormat.LeftIndent = alignment === CONST.WD_ALIGN_PAGE_NUMBER_LEFT ?
        CONST.PAGE_NUMBER_INDENT_PT : 0;
      paragraphFormat.RightIndent = alignment === CONST.WD_ALIGN_PAGE_NUMBER_RIGHT ?
        CONST.PAGE_NUMBER_INDENT_PT : 0;
      paragraphFormat.FirstLineIndent = 0;
      paragraphFormat.SpaceBefore = 0;
      paragraphFormat.SpaceAfter = 0;
    } catch (e1) {}
  }

  function applyOfficialPageNumbers(doc) {
    ensureOfficialDocumentFormatter();
    var sections = collectionToArray(doc.Sections);
    var footerIndexes = [
      CONST.WD_HEADER_FOOTER_PRIMARY,
      CONST.WD_HEADER_FOOTER_FIRST_PAGE,
      CONST.WD_HEADER_FOOTER_EVEN_PAGES
    ];
    for (var i = 0; i < sections.length; i++) {
      configurePageNumberSection(sections[i]);
      for (var j = 0; j < footerIndexes.length; j++) {
        var footerIndex = footerIndexes[j];
        if (!footerIsEnabled(sections[i], footerIndex)) continue;
        var footer = footerByIndex(sections[i], footerIndex);
        if (!footer) continue;
        try { footer.LinkToPrevious = false; } catch (e1) {}
        removeExistingPageFields(footer);
        var alignment = pageNumberAlignment(footerIndex);
        try {
          var pageNumbers = footer.PageNumbers;
          try { pageNumbers.RestartNumberingAtSection = false; } catch (e2) {}
          pageNumbers.Add(alignment, true);
        } catch (e3) {}
        try {
          var fields = collectionToArray(footer.Range.Fields);
          for (var f = 0; f < fields.length; f++) {
            if (isPageField(fields[f])) formatPageField(fields[f], alignment);
          }
        } catch (e4) {}
      }
    }
  }

  function applyCenteredPageNumbers(doc) {
    applyOfficialPageNumbers(doc);
  }

  function askYesNo(message, title) {
    try {
      var answer = MsgBox(message, 4, title || "\u516c\u6587\u683c\u5f0f");
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
    try { MsgBox(message, 0, "\u516c\u6587\u683c\u5f0f"); return; } catch (e1) {}
    try { hostApplication().Alert(message); return; } catch (e2) {}
    try { if (typeof alert !== "undefined") alert(message); return; } catch (e3) {}
  }

  function formatWholeDocument() {
    ensureOfficialDocumentFormatter();
    if (!askYesNo("\u5c06\u8c03\u6574\u6807\u9898\u3001\u7a7a\u884c\u3001\u6b63\u6587\u3001\u843d\u6b3e\u548c\u9875\u7801\uff0c\u662f\u5426\u7ee7\u7eed\uff1f", "\u516c\u6587\u683c\u5f0f")) return;
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
    if (askYesNo("\u662f\u5426\u5148\u6e05\u9664\u5168\u6587\u683c\u5f0f\uff1f", "\u516c\u6587\u683c\u5f0f")) {
      for (var c = 0; c < items.length; c++) clearParagraph(items[c].paragraph);
    }

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
      applyYiShi(items[result.yiShiIndexes[y]].paragraph, true);
    }
    applyOfficialPageNumbers(doc);
    alertMessage("\u683c\u5f0f\u5316\u5b8c\u6210");
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
      alertMessage("\u8bf7\u5148\u9009\u4e2d\u8981\u8bbe\u7f6e\u683c\u5f0f\u7684\u6587\u5b57\u6216\u6bb5\u843d\u3002");
      return;
    }
    applyToParagraphs(paragraphs, formatName, currentDocument());
  }

  function applyToParagraphs(paragraphs, formatName, doc) {
    ensureOfficialDocumentFormatter();
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
    applyOfficialPageNumbers(currentDocument());
    alertMessage("\u9875\u7801\u683c\u5f0f\u5316\u5b8c\u6210");
  }

  function GetOfficialDocumentFormatter() {
    ensureOfficialDocumentFormatter();
    return OfficialDocumentFormatter;
  }
  /* SHARED_CORE_END */

  var OfficialDocumentConsole;

  function ensureOfficialDocumentConsole() {
    if (OfficialDocumentConsole) return;
    OfficialDocumentConsole = {};
    OfficialDocumentConsole.controllerDoc = null;
    OfficialDocumentConsole.targetPath = "";
    OfficialDocumentConsole.targetDoc = null;
    OfficialDocumentConsole.selectionStart = -1;
    OfficialDocumentConsole.selectionEnd = -1;
    OfficialDocumentConsole.selectionDocPath = "";
    OfficialDocumentConsole.lastInputAvailable = false;
    OfficialDocumentConsole.lastFileDialogAvailable = false;
  }

  function consoleInput(message, defaultValue) {
    ensureOfficialDocumentConsole();
    OfficialDocumentConsole.lastInputAvailable = false;
    try {
      var globalInputValue = InputBox(message, "\u516c\u6587\u683c\u5f0f\u63a7\u5236\u53f0", defaultValue || "");
      OfficialDocumentConsole.lastInputAvailable = true;
      return globalInputValue;
    } catch (e1) {}
    try {
      var applicationInputValue = hostApplication().InputBox(message, "\u516c\u6587\u683c\u5f0f\u63a7\u5236\u53f0", defaultValue || "");
      OfficialDocumentConsole.lastInputAvailable = true;
      return applicationInputValue;
    } catch (e2) {}
    return "";
  }

  function consoleCollectionItem(collection, index) {
    try { return collection.Item(index); } catch (e1) {}
    try { return collection(index); } catch (e2) {}
    try { return collection[index - 1]; } catch (e3) {}
    return null;
  }

  function consoleSelectedFile(dialog) {
    try { return String(dialog.SelectedItems.Item(1)); } catch (e1) {}
    try { return String(dialog.SelectedItems(1)); } catch (e2) {}
    try { return String(dialog.SelectedItems[0]); } catch (e3) {}
    return "";
  }

  function normalizePath(path) {
    var value = String(path || "").replace(/^"|"$/g, "").replace(/^\s+|\s+$/g, "");
    var normalized = "";
    for (var i = 0; i < value.length; i++) {
      normalized += value.charAt(i) === String.fromCharCode(92) ? "/" : value.charAt(i);
    }
    return normalized;
  }

  function isWindowsAbsolutePath(path) {
    return path.length >= 3 &&
      /[A-Za-z]/.test(path.charAt(0)) &&
      path.charAt(1) === ":" &&
      path.charAt(2) === "/";
  }

  function samePath(a, b) {
    var left = normalizePath(a);
    var right = normalizePath(b);
    if (isWindowsAbsolutePath(left) && isWindowsAbsolutePath(right)) {
      return left.toLowerCase() === right.toLowerCase();
    }
    return left === right;
  }

  function documentPath(doc) {
    try { return normalizePath(doc.FullName); } catch (e1) {}
    return "";
  }

  function documentName(doc) {
    try { return String(doc.Name || doc.FullName || "\u672a\u547d\u540d\u6587\u6863"); } catch (e1) {}
    return "\u672a\u547d\u540d\u6587\u6863";
  }

  function pathFileName(path) {
    var normalized = normalizePath(path);
    var index = normalized.lastIndexOf("/");
    return index >= 0 ? normalized.substring(index + 1) : normalized;
  }

  function pathFolder(path) {
    var normalized = normalizePath(path);
    var index = normalized.lastIndexOf("/");
    return index > 0 ? normalized.substring(0, index) : "";
  }

  function pathExtension(path) {
    var name = pathFileName(path);
    var index = name.lastIndexOf(".");
    return index >= 0 ? name.substring(index).toLowerCase() : "";
  }

  function isSupportedDocumentPath(path) {
    var extension = pathExtension(path);
    return extension === ".doc" ||
      extension === ".docx" ||
      extension === ".docm" ||
      extension === ".wps" ||
      extension === ".wpt";
  }

  function isTemporaryDocumentPath(path) {
    return pathFileName(path).indexOf("~$") === 0;
  }

  function isConsoleDocumentName(name) {
    return String(name || "").indexOf("\u516c\u6587\u683c\u5f0f\u63a7\u5236\u53f0") >= 0;
  }

  function sameDocument(left, right) {
    if (!left || !right) return false;
    if (left === right) return true;
    var leftPath = documentPath(left);
    var rightPath = documentPath(right);
    return !!leftPath && !!rightPath && samePath(leftPath, rightPath);
  }

  function documentHasConsoleMacroButton(doc) {
    if (!doc) return false;
    var fields;
    var count = 0;
    try {
      fields = doc.Fields;
      count = Number(fields.Count);
    } catch (e1) {
      return false;
    }
    for (var i = 1; i <= count; i++) {
      var field = consoleCollectionItem(fields, i);
      if (!field) continue;
      try {
        if (/MACROBUTTON\s+Console/i.test(String(field.Code.Text || ""))) return true;
      } catch (e2) {}
    }
    return false;
  }

  function isControllerDocumentPath(path) {
    var normalized = normalizePath(path);
    if (!normalized) return false;
    if (isConsoleDocumentName(pathFileName(normalized))) return true;
    try {
      if (OfficialDocumentConsole.controllerDoc &&
          samePath(documentPath(OfficialDocumentConsole.controllerDoc), normalized)) {
        return true;
      }
    } catch (e1) {}
    var docs;
    var count = 0;
    try {
      docs = hostApplication().Documents;
      count = Number(docs.Count);
    } catch (e2) {
      return false;
    }
    for (var i = 1; i <= count; i++) {
      var doc = consoleCollectionItem(docs, i);
      if (!doc || !samePath(documentPath(doc), normalized)) continue;
      if (documentHasConsoleMacroButton(doc)) return true;
    }
    return false;
  }

  function isDocumentOpen(doc) {
    if (!doc) return false;
    try {
      var ignored = doc.Name;
      return ignored !== null && ignored !== undefined;
    } catch (e1) {}
    return false;
  }

  function controllerDocument() {
    ensureOfficialDocumentConsole();
    if (isDocumentOpen(OfficialDocumentConsole.controllerDoc)) {
      return OfficialDocumentConsole.controllerDoc;
    }
    OfficialDocumentConsole.controllerDoc = null;
    var active = null;
    try { active = currentDocument(); } catch (e1) {}
    if (active && (isConsoleDocumentName(documentName(active)) || documentHasConsoleMacroButton(active))) {
      OfficialDocumentConsole.controllerDoc = active;
      return active;
    }
    var docs;
    var count = 0;
    try {
      docs = hostApplication().Documents;
      count = Number(docs.Count);
    } catch (e2) {
      return null;
    }
    for (var i = 1; i <= count; i++) {
      var doc = consoleCollectionItem(docs, i);
      if (!doc) continue;
      if (isConsoleDocumentName(documentName(doc)) || documentHasConsoleMacroButton(doc)) {
        OfficialDocumentConsole.controllerDoc = doc;
        return doc;
      }
    }
    return null;
  }

  function controllerFolderPath(showError) {
    var controller = controllerDocument();
    if (!controller) {
      if (showError) {
        alertMessage("\u672a\u627e\u5230\u201c\u516c\u6587\u683c\u5f0f\u63a7\u5236\u53f0\u201d\u6587\u6863\u3002\u8bf7\u5148\u6253\u5f00\u6b63\u5f0f DOCM \u63a7\u5236\u53f0\u3002");
      }
      return "";
    }
    var folder = "";
    try { folder = normalizePath(controller.Path); } catch (e1) {}
    if (!folder && showError) {
      alertMessage("\u63a7\u5236\u53f0\u6587\u6863\u5c1a\u672a\u4fdd\u5b58\u5230\u6587\u4ef6\u5939\u3002\u8bf7\u5148\u4fdd\u5b58 DOCM\uff0c\u518d\u6267\u884c\u540c\u76ee\u5f55\u683c\u5f0f\u5316\u3002");
    }
    return folder;
  }

  function validateTargetPath(path) {
    if (!path) return false;
    if (isTemporaryDocumentPath(path)) {
      alertMessage("\u4e0d\u80fd\u5904\u7406 WPS \u751f\u6210\u7684\u4e34\u65f6\u6587\u4ef6\uff1a\n" + pathFileName(path));
      return false;
    }
    if (!isSupportedDocumentPath(path)) {
      alertMessage("\u4ec5\u652f\u6301 .doc\u3001.docx\u3001.docm\u3001.wps \u548c .wpt \u6587\u6863\u3002");
      return false;
    }
    if (isControllerDocumentPath(path)) {
      alertMessage("\u4e0d\u80fd\u628a\u516c\u6587\u683c\u5f0f\u63a7\u5236\u53f0\u6587\u6863\u8bbe\u4e3a\u76ee\u6807\u3002");
      return false;
    }
    return true;
  }

  function pathAlreadyListed(paths, path) {
    for (var i = 0; i < paths.length; i++) {
      if (samePath(paths[i], path)) return true;
    }
    return false;
  }

  function discoverSiblingDocumentPaths() {
    var result = { available: false, folder: "", paths: [], error: "" };
    var folder = controllerFolderPath(true);
    if (!folder) return result;
    result.folder = folder;
    var patterns = ["*.doc", "*.docx", "*.docm", "*.wps", "*.wpt"];
    var expectedExtensions = [".doc", ".docx", ".docm", ".wps", ".wpt"];
    for (var p = 0; p < patterns.length; p++) {
      try {
        var search = hostApplication().FileSearch;
        search.NewSearch();
        search.LookIn = folder;
        search.SearchSubFolders = false;
        search.FileName = patterns[p];
        search.Execute();
        result.available = true;
        var count = Number(search.FoundFiles.Count);
        for (var i = 1; i <= count; i++) {
          var found = normalizePath(consoleCollectionItem(search.FoundFiles, i));
          if (!found || pathExtension(found) !== expectedExtensions[p]) continue;
          if (isControllerDocumentPath(found) ||
              isTemporaryDocumentPath(found) ||
              pathAlreadyListed(result.paths, found)) continue;
          result.paths.push(found);
        }
      } catch (searchError) {
        result.error = String(searchError.message || searchError);
      }
    }
    result.paths.sort(function (left, right) {
      var a = pathFileName(left).toLowerCase();
      var b = pathFileName(right).toLowerCase();
      if (a < b) return -1;
      if (a > b) return 1;
      return 0;
    });
    return result;
  }

  function pickTargetDocumentPath(folder, requireSameFolder) {
    ensureOfficialDocumentConsole();
    OfficialDocumentConsole.lastFileDialogAvailable = false;
    var app = hostApplication();
    try {
      var dialog = app.FileDialog(3);
      if (!dialog) throw new Error("FileDialog unavailable");
      OfficialDocumentConsole.lastFileDialogAvailable = true;
      dialog.Title = "\u9009\u62e9\u9700\u8981\u683c\u5f0f\u5316\u7684\u516c\u6587";
      dialog.AllowMultiSelect = false;
      try {
        dialog.Filters.Clear();
        dialog.Filters.Add("Word/WPS \u6587\u6863", "*.doc;*.docx;*.docm;*.wps;*.wpt");
      } catch (filterError) {}
      try {
        if (folder) dialog.InitialFileName = folder.replace(/\//g, String.fromCharCode(92)) + String.fromCharCode(92);
      } catch (folderError) {}
      while (true) {
        var result = dialog.Show();
        if (!(result === -1 || result === true || result === 1)) return "";
        var selected = normalizePath(consoleSelectedFile(dialog));
        if (!validateTargetPath(selected)) continue;
        if (requireSameFolder && folder && !samePath(pathFolder(selected), folder)) {
          alertMessage("\u8bf7\u9009\u62e9\u4e0e\u63a7\u5236\u53f0\u4f4d\u4e8e\u540c\u4e00\u6587\u4ef6\u5939\u7684\u6587\u6863\u3002");
          continue;
        }
        return selected;
      }
    } catch (e1) {
      OfficialDocumentConsole.lastFileDialogAvailable = false;
    }
    return "";
  }

  function chooseSiblingDocumentPath() {
    var discovery = discoverSiblingDocumentPaths();
    if (!discovery.folder) return "";
    if (!discovery.available) {
      var fallback = pickTargetDocumentPath(discovery.folder, true);
      if (!fallback) {
        alertMessage("\u5f53\u524d WPS \u4e0d\u652f\u6301\u81ea\u52a8\u626b\u63cf\u6587\u4ef6\u5939\uff0c\u4e14\u672a\u4ece\u6587\u4ef6\u9009\u62e9\u6846\u9009\u4e2d\u76ee\u6807\u6587\u6863\u3002");
      }
      return fallback;
    }
    if (!discovery.paths.length) {
      alertMessage("\u63a7\u5236\u53f0\u6240\u5728\u6587\u4ef6\u5939\u4e2d\u6ca1\u6709\u53ef\u5904\u7406\u7684 .doc\u3001.docx\u3001.docm\u3001.wps \u6216 .wpt \u6587\u6863\u3002");
      return "";
    }
    if (discovery.paths.length === 1) return discovery.paths[0];
    var selected = pickTargetDocumentPath(discovery.folder, true);
    if (selected || OfficialDocumentConsole.lastFileDialogAvailable) return selected;
    var menu = "\u540c\u76ee\u5f55\u68c0\u6d4b\u5230\u591a\u4e2a\u6587\u6863\uff1a\n";
    for (var i = 0; i < discovery.paths.length; i++) {
      menu += (i + 1) + "  " + pathFileName(discovery.paths[i]) + "\n";
    }
    menu += "\n\u5f53\u524d WPS \u65e0\u6cd5\u6253\u5f00\u539f\u751f\u6587\u4ef6\u9009\u62e9\u6846\uff0c\u8bf7\u8f93\u5165\u5e8f\u53f7\u3002";
    while (true) {
      var choice = trimText(consoleInput(menu, ""));
      if (!choice) {
        if (!OfficialDocumentConsole.lastInputAvailable) {
          return pickTargetDocumentPath(discovery.folder, true);
        }
        return "";
      }
      if (/^\d+$/.test(choice)) {
        var index = Number(choice) - 1;
        if (index >= 0 && index < discovery.paths.length) return discovery.paths[index];
      }
      alertMessage("\u8bf7\u8f93\u5165\u5217\u8868\u4e2d\u7684\u6709\u6548\u5e8f\u53f7\u3002");
    }
  }

  function clearSavedSelection() {
    ensureOfficialDocumentConsole();
    OfficialDocumentConsole.selectionStart = -1;
    OfficialDocumentConsole.selectionEnd = -1;
    OfficialDocumentConsole.selectionDocPath = "";
  }

  function rangeBelongsToDocument(range, doc) {
    if (!range || !doc) return false;
    try {
      if (range.Document) return sameDocument(range.Document, doc);
    } catch (e1) {}
    return sameDocument(currentDocument(), doc);
  }

  function documentSelectionRange(doc) {
    if (!doc) return null;
    var fallback = null;
    try {
      var windows = doc.Windows;
      var count = windows ? Number(windows.Count) : 0;
      for (var i = 1; i <= count; i++) {
        var window = consoleCollectionItem(windows, i);
        if (!window || !window.Selection) continue;
        var range = window.Selection.Range;
        if (!fallback) fallback = range;
        if (rangeHasSelection(range)) return range;
      }
    } catch (e1) {}
    if (fallback) return fallback;
    try {
      if (sameDocument(currentDocument(), doc)) return hostApplication().Selection.Range;
    } catch (e2) {}
    return null;
  }

  function rangeHasSelection(range) {
    if (!range) return false;
    try { return Number(range.End) > Number(range.Start); } catch (e1) {}
    return false;
  }

  function captureSelectionForDocument(doc, showMessage) {
    ensureOfficialDocumentConsole();
    if (!doc) return false;
    var range = documentSelectionRange(doc);
    if (!range || !rangeBelongsToDocument(range, doc)) {
      if (showMessage) {
        alertMessage("\u5f53\u524d\u9009\u533a\u4e0d\u5c5e\u4e8e\u76ee\u6807\u6587\u6863\u3002\u8bf7\u5207\u6362\u5230\u76ee\u6807\u6587\u6863\u540e\u91cd\u65b0\u9009\u62e9\u6bb5\u843d\u3002");
      }
      return false;
    }
    var start = -1;
    var end = -1;
    try {
      start = Number(range.Start);
      end = Number(range.End);
    } catch (e2) {}
    if (start < 0 || end <= start) {
      if (showMessage) {
        alertMessage("\u8bf7\u5148\u9009\u4e2d\u81f3\u5c11\u4e00\u4e2a\u5b57\u7b26\u6216\u5b8c\u6574\u6bb5\u843d\u3002");
      }
      return false;
    }
    OfficialDocumentConsole.selectionStart = start;
    OfficialDocumentConsole.selectionEnd = end;
    OfficialDocumentConsole.selectionDocPath = documentPath(doc);
    if (showMessage) alertMessage("\u5df2\u8bc6\u522b\u5e76\u8bb0\u5f55\u6587\u6863\u9009\u533a\uff1a\n" + documentName(doc));
    return true;
  }

  function savedSelectionParagraphs(doc) {
    ensureOfficialDocumentConsole();
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

  function openTargetDocuments() {
    var result = [];
    var docs;
    var count = 0;
    var controller = controllerDocument();
    try {
      docs = hostApplication().Documents;
      count = Number(docs.Count);
    } catch (e1) {
      return result;
    }
    for (var i = 1; i <= count; i++) {
      var doc = consoleCollectionItem(docs, i);
      if (!doc || sameDocument(doc, controller)) continue;
      if (isControllerDocumentPath(documentPath(doc))) continue;
      result.push(doc);
    }
    return result;
  }

  function targetDocumentIndex(documents) {
    ensureOfficialDocumentConsole();
    for (var i = 0; i < documents.length; i++) {
      if (sameDocument(documents[i], OfficialDocumentConsole.targetDoc)) return i;
      if (OfficialDocumentConsole.targetPath &&
          samePath(documentPath(documents[i]), OfficialDocumentConsole.targetPath)) return i;
    }
    return -1;
  }

  function chooseOpenTargetDocument(requireSelection, actionName) {
    var documents = openTargetDocuments();
    if (!documents.length) {
      alertMessage("\u672a\u627e\u5230\u5df2\u6253\u5f00\u7684\u5f85\u5904\u7406\u6587\u6863\u3002\u8bf7\u5148\u6253\u5f00\u516c\u6587\uff0c\u9700\u8981\u5c40\u90e8\u8c03\u6574\u65f6\u8bf7\u5148\u9009\u4e2d\u6587\u5b57\u3002");
      return null;
    }
    var selectedIndex = 0;
    if (documents.length > 1) {
      var menu = "\u5df2\u6253\u5f00\u591a\u4e2a\u6587\u6863\uff0c\u8bf7\u9009\u62e9\u8981\u6267\u884c\u201c" +
        actionName + "\u201d\u7684\u6587\u6863\uff1a\n";
      for (var i = 0; i < documents.length; i++) {
        var range = documentSelectionRange(documents[i]);
        var status = rangeHasSelection(range) ? "\u5df2\u9009\u4e2d\u6587\u5b57" : "\u672a\u9009\u4e2d\u6587\u5b57";
        menu += (i + 1) + "  " + documentName(documents[i]) + "  [" + status + "]\n";
      }
      var rememberedIndex = targetDocumentIndex(documents);
      var defaultChoice = rememberedIndex >= 0 ? String(rememberedIndex + 1) : "";
      while (true) {
        var choice = trimText(consoleInput(menu, defaultChoice));
        if (!choice) return null;
        if (/^\d+$/.test(choice)) {
          selectedIndex = Number(choice) - 1;
          if (selectedIndex >= 0 && selectedIndex < documents.length) break;
        }
        alertMessage("\u8bf7\u8f93\u5165\u5df2\u6253\u5f00\u6587\u6863\u5217\u8868\u4e2d\u7684\u6709\u6548\u5e8f\u53f7\u3002");
      }
    }
    var doc = documents[selectedIndex];
    var selectedRange = documentSelectionRange(doc);
    if (requireSelection && !rangeHasSelection(selectedRange)) {
      alertMessage("\u6240\u9009\u6587\u6863\u4e2d\u6ca1\u6709\u9009\u4e2d\u6587\u5b57\u3002\u8bf7\u5207\u6362\u5230\u8be5\u6587\u6863\u9009\u4e2d\u6587\u5b57\uff0c\u518d\u56de\u5230\u63a7\u5236\u53f0\u70b9\u51fb\u683c\u5f0f\u6309\u94ae\u3002");
      return null;
    }
    ensureOfficialDocumentConsole();
    OfficialDocumentConsole.targetDoc = doc;
    OfficialDocumentConsole.targetPath = documentPath(doc);
    if (rangeHasSelection(selectedRange)) captureSelectionForDocument(doc, false);
    else clearSavedSelection();
    return { doc: doc, range: selectedRange };
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

  function openDocumentAtPath(path, activateDocument) {
    path = normalizePath(path);
    if (!validateTargetPath(path)) return null;
    var doc = findOpenDocumentByPath(path);
    if (!doc) {
      try {
        doc = hostApplication().Documents.Open(path);
      } catch (openError) {
        alertMessage("\u76ee\u6807\u6587\u6863\u6253\u5f00\u5931\u8d25\uff1a" + String(openError.message || openError));
        return null;
      }
    }
    ensureOfficialDocumentConsole();
    OfficialDocumentConsole.targetDoc = doc;
    OfficialDocumentConsole.targetPath = path;
    if (activateDocument) {
      try { doc.Activate(); } catch (activateError) {}
    }
    return doc;
  }

  function openTargetDocument() {
    ensureOfficialDocumentConsole();
    var path = normalizePath(OfficialDocumentConsole.targetPath);
    if (!path) {
      path = normalizePath(chooseSiblingDocumentPath());
      OfficialDocumentConsole.targetPath = path;
    }
    if (!path) {
      alertMessage("\u672a\u9009\u62e9\u76ee\u6807\u6587\u6863\u3002");
      return null;
    }
    return openDocumentAtPath(path, true);
  }

  function targetDocument() {
    ensureOfficialDocumentConsole();
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

  function runWholeFormatOnDocument(doc) {
    if (!doc) return;
    try { doc.Activate(); } catch (e1) {}
    formatWholeDocument();
    try { doc.Activate(); } catch (e2) {}
  }

  function applyTargetSelection(formatName) {
    var target = chooseOpenTargetDocument(true, "\u5c40\u90e8\u683c\u5f0f\u8c03\u6574");
    if (!target) return;
    var paragraphs = [];
    try { paragraphs = collectionToArray(target.range.Paragraphs); } catch (e1) {}
    if (!paragraphs.length) {
      alertMessage("\u65e0\u6cd5\u8bfb\u53d6\u6240\u9009\u6587\u5b57\u6240\u5728\u6bb5\u843d\u3002");
      return;
    }
    applyToParagraphs(paragraphs, formatName, target.doc);
    alertMessage("\u5df2\u5bf9\u4ee5\u4e0b\u6587\u6863\u7684\u9009\u533a\u5e94\u7528\u683c\u5f0f\uff1a\n" + documentName(target.doc));
  }

  function ConsoleFormatFolderDocument() {
    ConsoleSelectAndFormatDocument();
  }

  function ConsoleSelectTargetDocument() {
    ensureOfficialDocumentConsole();
    var path = normalizePath(chooseSiblingDocumentPath());
    if (!path) {
      alertMessage("\u672a\u9009\u62e9\u76ee\u6807\u6587\u6863\u3002");
      return;
    }
    if (!validateTargetPath(path)) return;
    OfficialDocumentConsole.targetPath = path;
    OfficialDocumentConsole.targetDoc = null;
    clearSavedSelection();
    alertMessage("\u5df2\u9009\u62e9\u76ee\u6807\u6587\u6863\uff1a\n" + path);
  }

  function ConsoleSelectAndFormatDocument() {
    ensureOfficialDocumentConsole();
    var path = normalizePath(chooseSiblingDocumentPath());
    if (!path) return;
    if (!validateTargetPath(path)) return;
    OfficialDocumentConsole.targetPath = path;
    OfficialDocumentConsole.targetDoc = null;
    clearSavedSelection();
    var doc = openDocumentAtPath(path, false);
    if (!doc) return;
    runWholeFormatOnDocument(doc);
  }

  function ConsoleOpenTargetDocument() {
    if (openTargetDocument()) alertMessage("\u76ee\u6807\u6587\u6863\u5df2\u6253\u5f00\u5e76\u5207\u6362\u5230\u524d\u53f0\u3002");
  }

  function ConsoleUseActiveDocument() {
    ensureOfficialDocumentConsole();
    var doc = currentDocument();
    if (!doc) {
      alertMessage("\u672a\u627e\u5230\u5f53\u524d\u6d3b\u52a8\u6587\u6863\u3002");
      return;
    }
    var name = documentName(doc);
    if (sameDocument(doc, controllerDocument()) || isConsoleDocumentName(name) || documentHasConsoleMacroButton(doc)) {
      var chosen = chooseOpenTargetDocument(false, "\u8bbe\u4e3a\u76ee\u6807");
      if (!chosen) return;
      doc = chosen.doc;
      name = documentName(doc);
    }
    if (!askYesNo("\u786e\u8ba4\u5c06\u4ee5\u4e0b\u5f53\u524d\u6d3b\u52a8\u6587\u6863\u8bbe\u4e3a\u76ee\u6807\uff1f\n" + name, "\u516c\u6587\u683c\u5f0f\u63a7\u5236\u53f0")) return;
    OfficialDocumentConsole.targetDoc = doc;
    OfficialDocumentConsole.targetPath = documentPath(doc);
    clearSavedSelection();
    var captured = captureSelectionForDocument(doc, false);
    alertMessage(captured ?
      "\u5df2\u8bbe\u4e3a\u76ee\u6807\u6587\u6863\uff0c\u5e76\u8bc6\u522b\u5230\u5f53\u524d\u9009\u533a\u3002" :
      "\u5df2\u8bbe\u4e3a\u76ee\u6807\u6587\u6863\u3002");
  }

  function ConsoleChooseOpenDocument() {
    var chosen = chooseOpenTargetDocument(false, "\u8bbe\u4e3a\u76ee\u6807");
    if (!chosen) return;
    alertMessage("\u5df2\u9009\u62e9\u6253\u5f00\u7684\u76ee\u6807\u6587\u6863\uff1a\n" + documentName(chosen.doc));
  }

  function ConsoleCaptureTargetSelection() {
    var chosen = chooseOpenTargetDocument(true, "\u8bc6\u522b\u5f53\u524d\u9009\u533a");
    if (!chosen) return;
    captureSelectionForDocument(chosen.doc, true);
  }

  function ConsoleFormatTargetDocument() {
    var doc = targetDocument();
    if (!doc) return;
    runWholeFormatOnDocument(doc);
  }

  function ConsoleShowTargetDocument() {
    ensureOfficialDocumentConsole();
    var target = OfficialDocumentConsole.targetDoc;
    var path = OfficialDocumentConsole.targetPath;
    if (target) {
      alertMessage("\u5f53\u524d\u76ee\u6807\u6587\u6863\uff1a\n" + documentName(target) + (path ? "\n" + path : ""));
      return;
    }
    if (path) {
      alertMessage("\u5f53\u524d\u76ee\u6807\u6587\u6863\u8def\u5f84\uff1a\n" + path);
      return;
    }
    alertMessage("\u5c1a\u672a\u8bbe\u7f6e\u76ee\u6807\u6587\u6863\u3002");
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
    var target = chooseOpenTargetDocument(false, "\u5c45\u4e2d\u9875\u7801");
    if (!target) return;
    applyCenteredPageNumbers(target.doc);
    alertMessage("\u5df2\u5bf9\u4ee5\u4e0b\u6587\u6863\u5e94\u7528\u5c45\u4e2d\u9875\u7801\uff1a\n" + documentName(target.doc));
  }

  function ConsoleRunMenu() {
    ensureOfficialDocumentConsole();
    var menu =
      "\u8bf7\u8f93\u5165\u64cd\u4f5c\u5e8f\u53f7\uff1a\n" +
      "0  \u540c\u76ee\u5f55\u9009\u62e9\u5e76\u5168\u6587\u683c\u5f0f\u5316\n" +
      "1  \u540c\u76ee\u5f55\u9009\u62e9\u76ee\u6807\u6587\u6863\n" +
      "2  \u6253\u5f00\u76ee\u6807\u6587\u6863\n" +
      "3  \u5168\u6587\u683c\u5f0f\u5316\n" +
      "20 \u540c\u76ee\u5f55\u9009\u62e9\u5e76\u5168\u6587\u683c\u5f0f\u5316\n" +
      "22 \u4ece\u5df2\u6253\u5f00\u6587\u6863\u4e2d\u9009\u62e9\u76ee\u6807\n" +
      "4  \u5927\u6807\u9898\n" +
      "5  \u4e3b\u9001\u5355\u4f4d\n" +
      "6  \u6b63\u6587\n" +
      "7  \u4e00\u7ea7\u6807\u9898\n" +
      "8  \u4e8c\u7ea7\u6807\u9898\n" +
      "9  \u4e09\u7ea7\u6807\u9898\n" +
      "10 \u56db\u7ea7\u6807\u9898\n" +
      "11 \u4e00\u662f/\u4e8c\u662f\u52a0\u7c97\n" +
      "12 \u6210\u6587\u65e5\u671f\n" +
      "13 \u843d\u6b3e\u5355\u4f4d\n" +
      "14 \u9644\u4ef6\u8bf4\u660e\n" +
      "15 \u9644\u4ef6\u8bf4\u660e\u7eed\u884c\n" +
      "16 \u9644\u4ef6\u5e8f\u53f7\n" +
      "17 \u9644\u4ef6\u6807\u9898\n" +
      "18 \u5c45\u4e2d\u9875\u7801\n" +
      "19 \u8bc6\u522b\u5df2\u6253\u5f00\u6587\u6863\u7684\u5f53\u524d\u9009\u533a\n" +
      "21 \u67e5\u770b\u5f53\u524d\u76ee\u6807";
    var choice = trimText(consoleInput(menu, ""));
    if (!choice) return;
    switch (choice) {
      case "0": ConsoleSelectAndFormatDocument(); break;
      case "1": ConsoleSelectTargetDocument(); break;
      case "2": ConsoleOpenTargetDocument(); break;
      case "3": ConsoleFormatTargetDocument(); break;
      case "4": ConsoleApplyTitleFormat(); break;
      case "5": ConsoleApplyMainRecipientFormat(); break;
      case "6": ConsoleApplyBodyFormat(); break;
      case "7": ConsoleApplyLevel1HeadingFormat(); break;
      case "8": ConsoleApplyLevel2HeadingFormat(); break;
      case "9": ConsoleApplyLevel3HeadingFormat(); break;
      case "10": ConsoleApplyLevel4HeadingFormat(); break;
      case "11": ConsoleApplyYiShiFormat(); break;
      case "12": ConsoleApplyDateFormat(); break;
      case "13": ConsoleApplySignatureFormat(); break;
      case "14": ConsoleApplyAttachmentNoteFormat(); break;
      case "15": ConsoleApplyAttachmentNoteContinuationFormat(); break;
      case "16": ConsoleApplyAttachmentSequenceFormat(); break;
      case "17": ConsoleApplyAttachmentTitleFormat(); break;
      case "18": ConsoleApplyPageNumberFormat(); break;
      case "19": ConsoleCaptureTargetSelection(); break;
      case "20": ConsoleSelectAndFormatDocument(); break;
      case "21": ConsoleShowTargetDocument(); break;
      case "22": ConsoleChooseOpenDocument(); break;
      default: alertMessage("\u672a\u77e5\u64cd\u4f5c\u5e8f\u53f7\uff1a" + choice);
    }
  }

  function ConsoleSyntaxCheck() {
    ensureOfficialDocumentFormatter();
    ensureOfficialDocumentConsole();
    controllerDocument();
    alertMessage("\u516c\u6587\u683c\u5f0f\u63a7\u5236\u53f0\u811a\u672c\u5df2\u5b8c\u6574\u52a0\u8f7d\u3002");
  }
