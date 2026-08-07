/* \u516c\u6587\u683c\u5f0f\u5316 WPS JS \u5b8f
 * \u8fd0\u884c\u73af\u5883\uff1aWPS Office JS \u5b8f\u3002\u83dc\u5355\u6309\u94ae\u901a\u8fc7 customUI.xml \u7ed1\u5b9a\u672c\u6587\u4ef6\u4e2d\u7684\u5168\u5c40\u51fd\u6570\u3002
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
    OfficialDocumentFormatter.applyRecognizedSignatures = applyRecognizedSignatures;
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

  function findDateIndexInRange(items, startIndex, endIndex) {
    var idx = -1;
    var start = typeof startIndex === "number" && startIndex >= 0 ? startIndex : 0;
    var end = typeof endIndex === "number" && endIndex >= 0 ? endIndex : items.length;
    for (var i = start; i < end; i++) {
      if (isDateText(items[i].text)) idx = i;
    }
    return idx;
  }

  function findDateIndex(items, endIndex) {
    return findDateIndexInRange(items, 0, endIndex);
  }

  function findSignatureIndexInRange(items, dateIndex, startIndex) {
    var start = typeof startIndex === "number" && startIndex >= 0 ? startIndex : 0;
    if (dateIndex <= start) return -1;
    for (var i = dateIndex - 1; i >= start; i--) {
      if (trimText(items[i].text)) return i;
    }
    return -1;
  }

  function findSignatureIndex(items, dateIndex) {
    return findSignatureIndexInRange(items, dateIndex, 0);
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

  function findAttachmentSignatures(items, attachmentBody) {
    var pairs = [];
    for (var i = 0; i < attachmentBody.sequenceIndexes.length; i++) {
      var sequenceIndex = attachmentBody.sequenceIndexes[i];
      var endIndex = i + 1 < attachmentBody.sequenceIndexes.length ?
        attachmentBody.sequenceIndexes[i + 1] : items.length;
      var startIndex = sequenceIndex + 1;
      for (var t = 0; t < attachmentBody.titleIndexes.length; t++) {
        var titleIndex = attachmentBody.titleIndexes[t];
        if (titleIndex > sequenceIndex && titleIndex < endIndex) {
          startIndex = titleIndex + 1;
          break;
        }
      }
      var attachmentDateIndex = findDateIndexInRange(items, startIndex, endIndex);
      var attachmentSignatureIndex =
        findSignatureIndexInRange(items, attachmentDateIndex, startIndex);
      if (attachmentSignatureIndex >= 0 && attachmentDateIndex >= 0) {
        pairs.push({
          signatureIndex: attachmentSignatureIndex,
          dateIndex: attachmentDateIndex
        });
      }
    }
    return pairs;
  }

  function classifyDocument(items) {
    ensureOfficialDocumentFormatter();
    var titleIndexes = findTitleIndexes(items);
    var mainRecipientIndex = findMainRecipientIndex(items, titleIndexes);
    var attachmentBody = findAttachmentBody(items);
    var attachmentStartIndex = attachmentBody.sequenceIndexes.length ? attachmentBody.sequenceIndexes[0] : items.length;
    var dateIndex = findDateIndex(items, attachmentStartIndex);
    var signatureIndex = findSignatureIndex(items, dateIndex);
    var attachmentSignatures = findAttachmentSignatures(items, attachmentBody);
    var attachmentSignatureMap = {};
    var attachmentDateMap = {};
    for (var a = 0; a < attachmentSignatures.length; a++) {
      attachmentSignatureMap[attachmentSignatures[a].signatureIndex] = true;
      attachmentDateMap[attachmentSignatures[a].dateIndex] = true;
    }
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
      if (attachmentSignatureMap[i] || attachmentDateMap[i]) continue;
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
      attachmentSignatures: attachmentSignatures,
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

  function applyRecognizedSignatures(items, result, doc) {
    if (result.dateIndex >= 0) applyNamedFormat(items[result.dateIndex].paragraph, "date");
    if (result.signatureIndex >= 0 && result.dateIndex >= 0) {
      applySignatureFormat(items[result.signatureIndex].paragraph, items[result.dateIndex].text, doc);
    }
    for (var i = 0; i < result.attachmentSignatures.length; i++) {
      var attachmentSignature = result.attachmentSignatures[i];
      applyNamedFormat(items[attachmentSignature.dateIndex].paragraph, "date");
      applySignatureFormat(
        items[attachmentSignature.signatureIndex].paragraph,
        items[attachmentSignature.dateIndex].text,
        doc
      );
    }
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

  function signaturePairs(result) {
    var pairs = [];
    if (result.signatureIndex >= 0 && result.dateIndex >= 0) {
      pairs.push({ signatureIndex: result.signatureIndex, dateIndex: result.dateIndex });
    }
    for (var i = 0; i < result.attachmentSignatures.length; i++) {
      pairs.push(result.attachmentSignatures[i]);
    }
    pairs.sort(function (a, b) { return b.signatureIndex - a.signatureIndex; });
    return pairs;
  }

  function normalizeBlankLineBeforeSignaturePair(items, pair) {
    var previousIndex = previousNonEmptyIndex(items, pair.signatureIndex - 1);
    if (previousIndex < 0) return false;
    var changed = false;
    for (var i = pair.signatureIndex - 1; i > previousIndex; i--) {
      if (!trimText(items[i].text) && !items[i].hasBreak) {
        if (deleteParagraph(items[i].paragraph)) changed = true;
      }
    }
    if (insertParagraphAfter(items[previousIndex].paragraph)) changed = true;
    if (insertParagraphAfter(items[previousIndex].paragraph)) changed = true;
    return changed;
  }

  function normalizeBlankLineBeforeSignature(items, result) {
    var pairs = signaturePairs(result);
    var changed = false;
    for (var i = 0; i < pairs.length; i++) {
      if (normalizeBlankLineBeforeSignaturePair(items, pairs[i])) changed = true;
    }
    return changed;
  }

  function formatBlankLineBeforeSignature(items, result) {
    var pairs = signaturePairs(result);
    for (var p = 0; p < pairs.length; p++) {
      for (var i = pairs[p].signatureIndex - 1;
        i >= 0 && i >= pairs[p].signatureIndex - 2; i--) {
        if (trimText(items[i].text)) continue;
        if (!canFormatParagraph(items[i].paragraph)) continue;
        setFont(items[i].paragraph.Range, CONST.FONT_BODY, CONST.FONT_WEST, CONST.SIZE_BODY, false);
        setParagraph(items[i].paragraph, CONST.WD_ALIGN_LEFT, 0, 0, 0);
      }
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

  function rangeHasPageField(range) {
    var fields = [];
    try { fields = collectionToArray(range.Fields); } catch (e1) {}
    for (var i = 0; i < fields.length; i++) {
      if (isPageField(fields[i])) return true;
    }
    return false;
  }

  function isManualPageNumberText(text) {
    var s = trimText(text).replace(/\u00a0/g, " ");
    if (!s) return false;
    if (/^(?:[\u2014\u2013\u2212\uFF0D-]\s*)*(?:\d+|[\uFF10-\uFF19]+)(?:\s*(?:\/|of)\s*(?:\d+|[\uFF10-\uFF19]+))?(?:\s*[\u2014\u2013\u2212\uFF0D-])*$/i.test(s)) {
      return true;
    }
    return /^\u7b2c\s*(?:\d+|[\uFF10-\uFF19]+)\s*\u9875(?:\s*(?:\/|\u5171)\s*(?:\d+|[\uFF10-\uFF19]+)\s*\u9875?)?$/.test(s);
  }

  function clearParagraphContent(paragraph) {
    if (!paragraph) return;
    try {
      var content = paragraph.Range.Duplicate;
      var text = String(content.Text || "");
      while (text && /[\r\n\u0007\f]$/.test(text)) {
        text = text.substring(0, text.length - 1);
        content.End = content.End - 1;
      }
      content.Text = "";
    } catch (e1) {}
    try { paragraph.Range.Font.Reset(); } catch (e2) {}
    try { paragraph.Range.ParagraphFormat.Reset(); } catch (e3) {}
  }

  function clearExistingPageNumbers(footer) {
    var paragraphs = [];
    try { paragraphs = collectionToArray(footer.Range.Paragraphs); } catch (e1) {}
    for (var i = paragraphs.length - 1; i >= 0; i--) {
      var text = "";
      try { text = paragraphs[i].Range.Text; } catch (e2) {}
      if (rangeHasPageField(paragraphs[i].Range) || isManualPageNumberText(text)) {
        clearParagraphContent(paragraphs[i]);
      }
    }
    var fields = [];
    try { fields = collectionToArray(footer.Range.Fields); } catch (e3) {}
    for (var f = fields.length - 1; f >= 0; f--) {
      if (!isPageField(fields[f])) continue;
      try { fields[f].Delete(); } catch (e4) {}
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

  function paragraphIsEmpty(paragraph) {
    if (!paragraph) return true;
    try {
      if (trimText(paragraph.Range.Text)) return false;
    } catch (e1) {}
    try {
      if (Number(paragraph.Range.Fields.Count) > 0) return false;
    } catch (e2) {}
    return true;
  }

  function preparePageNumberParagraph(footer) {
    var paragraphs = [];
    try { paragraphs = collectionToArray(footer.Range.Paragraphs); } catch (e1) {}
    var emptyParagraph = null;
    var hasContent = false;
    for (var i = 0; i < paragraphs.length; i++) {
      if (paragraphIsEmpty(paragraphs[i])) emptyParagraph = paragraphs[i];
      else hasContent = true;
    }
    if (!hasContent) {
      try { footer.Range.Text = ""; } catch (e2) {}
      try { paragraphs = collectionToArray(footer.Range.Paragraphs); } catch (e3) {}
      return paragraphs.length ? paragraphs[0] : null;
    }
    if (emptyParagraph) return emptyParagraph;
    try {
      var appendRange = footer.Range.Duplicate;
      if (appendRange.End > appendRange.Start) appendRange.End = appendRange.End - 1;
      appendRange.Start = appendRange.End;
      appendRange.InsertAfter("\r");
    } catch (e4) {}
    try { paragraphs = collectionToArray(footer.Range.Paragraphs); } catch (e5) {}
    return paragraphs.length ? paragraphs[paragraphs.length - 1] : null;
  }

  function setPageNumberParagraphText(paragraph) {
    if (!paragraph) return;
    var content = paragraph.Range.Duplicate;
    var text = "";
    try { text = String(content.Text || ""); } catch (e1) {}
    while (text && /[\r\n\u0007\f]$/.test(text)) {
      text = text.substring(0, text.length - 1);
      content.End = content.End - 1;
    }
    content.Text = "\u2014\u2014";
  }

  function formatPageNumberParagraph(paragraph, alignment) {
    if (!paragraph) return;
    var range = paragraph.Range;
    try { range.Font.Reset(); } catch (e1) {}
    try { range.ParagraphFormat.Reset(); } catch (e2) {}
    setFont(range, CONST.FONT_PAGE_NUMBER, CONST.FONT_PAGE_NUMBER,
      CONST.SIZE_PAGE_NUMBER, false);
    var paragraphFormat = range.ParagraphFormat;
    try { paragraphFormat.Alignment = alignment; } catch (e3) {}
    try { paragraphFormat.CharacterUnitLeftIndent = 0; } catch (e4) {}
    try { paragraphFormat.CharacterUnitRightIndent = 0; } catch (e5) {}
    try { paragraphFormat.CharacterUnitFirstLineIndent = 0; } catch (e6) {}
    try {
      paragraphFormat.LeftIndent = alignment === CONST.WD_ALIGN_PAGE_NUMBER_LEFT ?
        CONST.PAGE_NUMBER_INDENT_PT : 0;
    } catch (e7) {}
    try {
      paragraphFormat.RightIndent = alignment === CONST.WD_ALIGN_PAGE_NUMBER_RIGHT ?
        CONST.PAGE_NUMBER_INDENT_PT : 0;
    } catch (e8) {}
    try { paragraphFormat.FirstLineIndent = 0; } catch (e9) {}
    try { paragraphFormat.SpaceBefore = 0; } catch (e10) {}
    try { paragraphFormat.SpaceAfter = 0; } catch (e11) {}
    try { paragraphFormat.LineSpacingRule = 0; } catch (e12) {}
  }

  function addOfficialPageNumber(footer, alignment) {
    var paragraph = preparePageNumberParagraph(footer);
    if (!paragraph) return false;
    clearParagraphContent(paragraph);
    setPageNumberParagraphText(paragraph);
    var paragraphRange = paragraph.Range.Duplicate;
    var insertion = paragraph.Range.Duplicate;
    insertion.Start = paragraphRange.Start + 1;
    insertion.End = insertion.Start;
    var field = null;
    try { field = paragraph.Range.Fields.Add(insertion, CONST.WD_FIELD_PAGE); } catch (e1) {}
    if (!field) {
      try { field = footer.Range.Fields.Add(insertion, CONST.WD_FIELD_PAGE); } catch (e2) {}
    }
    if (!field) return false;
    try { field.Update(); } catch (e3) {}
    formatPageNumberParagraph(paragraph, alignment);
    return true;
  }

  function updateFooterPageFields(footer) {
    var fields = [];
    try { fields = collectionToArray(footer.Range.Fields); } catch (e1) {}
    for (var i = 0; i < fields.length; i++) {
      if (!isPageField(fields[i])) continue;
      try { fields[i].Update(); } catch (e2) {}
    }
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
        var footer = footerByIndex(sections[i], footerIndex);
        if (!footer) continue;
        try { footer.LinkToPrevious = false; } catch (e1) {}
        clearExistingPageNumbers(footer);
        if (!footerIsEnabled(sections[i], footerIndex)) continue;
        var alignment = pageNumberAlignment(footerIndex);
        try {
          var pageNumbers = footer.PageNumbers;
          try { pageNumbers.RestartNumberingAtSection = false; } catch (e2) {}
        } catch (e3) {}
        if (!addOfficialPageNumber(footer, alignment)) {
          throw new Error("\u65e0\u6cd5\u5728\u9875\u811a\u4e2d\u521b\u5efa\u9875\u7801\u5b57\u6bb5\u3002");
        }
      }
    }
    try { doc.Repaginate(); } catch (e4) {}
    for (var s = 0; s < sections.length; s++) {
      for (var h = 0; h < footerIndexes.length; h++) {
        var updatedFooter = footerByIndex(sections[s], footerIndexes[h]);
        if (updatedFooter) updateFooterPageFields(updatedFooter);
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
    applyRecognizedSignatures(items, result, doc);
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
  function ApplyPageNumberFormatSilently() {
    ensureOfficialDocumentFormatter();
    applyOfficialPageNumbers(currentDocument());
  }

  function GetOfficialDocumentFormatter() {
    ensureOfficialDocumentFormatter();
    return OfficialDocumentFormatter;
  }
  /* SHARED_CORE_END */
