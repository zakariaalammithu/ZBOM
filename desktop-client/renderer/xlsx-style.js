// Professional styling + print setup for ZBOM Excel downloads.
//
// SheetJS (community edition) cannot write cell styles, so this module
// post-processes a SheetJS-generated .xlsx (a zip) with JSZip: it injects
// a styles.xml (fonts / fills / borders), sets each cell's style index
// in the sheet XML, and adds print setup so the file prints directly.
// The look matches the user's own formatted template: light-green header
// (C3D69B), dark bold text, thin borders on every cell, column widths
// fitted to content, header row frozen. Print setup (from the user's
// file, which prints at 78% portrait): portrait A4, fit all columns to
// one page wide, header row repeated on every printed page, 0.5" margins.
//
// Usage (browser): window.ZBOM_XLSX_STYLE.styleXlsxBlob(blob) returns a
// Promise<Blob> — the styled file. On any failure it resolves with the
// original blob unchanged, so a styling problem can never break a
// download.
(function (global) {
  "use strict";

  var HEADER_FILL_RGB = "FFC3D69B"; // light green (user's template: accent3 + 40% tint)
  var HEADER_FONT_RGB = "FF1F2A36"; // dark slate
  var BODY_FONT_RGB = "FF1F2A36";
  var BORDER_RGB = "FF808080"; // medium grey
  var HEADER_XF = 1;
  var BODY_XF = 2;

  var STYLES_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    "<numFmts count=\"0\"/>" +
    "<fonts count=\"2\">" +
    "<font><name val=\"Calibri\"/><sz val=\"11\"/><color rgb=\"" + BODY_FONT_RGB + "\"/></font>" +
    "<font><name val=\"Calibri\"/><sz val=\"11\"/><b/><color rgb=\"" + HEADER_FONT_RGB + "\"/></font>" +
    "</fonts>" +
    "<fills count=\"3\">" +
    "<fill><patternFill patternType=\"none\"/></fill>" +
    "<fill><patternFill patternType=\"gray125\"/></fill>" +
    "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"" + HEADER_FILL_RGB + "\"/><bgColor rgb=\"" + HEADER_FILL_RGB + "\"/></patternFill></fill>" +
    "</fills>" +
    "<borders count=\"2\">" +
    "<border><left/><right/><top/><bottom/><diagonal/></border>" +
    "<border>" +
    "<left style=\"thin\"><color rgb=\"" + BORDER_RGB + "\"/></left>" +
    "<right style=\"thin\"><color rgb=\"" + BORDER_RGB + "\"/></right>" +
    "<top style=\"thin\"><color rgb=\"" + BORDER_RGB + "\"/></top>" +
    "<bottom style=\"thin\"><color rgb=\"" + BORDER_RGB + "\"/></bottom>" +
    "<diagonal/>" +
    "</border>" +
    "</borders>" +
    "<cellStyleXfs count=\"1\">" +
    "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>" +
    "</cellStyleXfs>" +
    "<cellXfs count=\"3\">" +
    "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
    "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\"/></xf>" +
    "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyBorder=\"1\"/>" +
    "</cellXfs>" +
    "<cellStyles count=\"1\">" +
    "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>" +
    "</cellStyles>" +
    "<dxfs count=\"0\"/>" +
    "<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium9\" defaultPivotStyle=\"PivotStyleMedium4\"/>" +
    "</styleSheet>";

  // Print setup, matching the user's file (portrait, scaled to fit one
  // page wide). fitToPage + fitToWidth=1 adapts the scale to each file
  // (the user's 12-column file prints at ~78%, like their own file).
  var PAGE_SETUP_XML =
    "<pageMargins left=\"0.5\" right=\"0.5\" top=\"0.6\" bottom=\"0.6\" header=\"0.3\" footer=\"0.3\"/>" +
    "<pageSetup orientation=\"portrait\" paperSize=\"9\" fitToWidth=\"1\" fitToHeight=\"0\"/>";
  var SHEET_PR_XML = "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>";

  function colToNum(col) {
    var n = 0;
    for (var i = 0; i < col.length; i++) n = n * 26 + (col.charCodeAt(i) - 64);
    return n;
  }

  function decodeEntities(text) {
    return String(text || "")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
  }

  // Longest display length per column, from the <v> values in sheetData.
  function computeColumnWidths(sheetDataXml) {
    var maxLen = {};
    var cellRe = /<c r="([A-Z]+)(\d+)"[^>]*>(?:<v>([^<]*)<\/v>)?/g;
    var m;
    while ((m = cellRe.exec(sheetDataXml)) !== null) {
      var len = decodeEntities(m[3] || "").length;
      if (!maxLen[m[1]] || len > maxLen[m[1]]) maxLen[m[1]] = len;
    }
    var widths = {};
    Object.keys(maxLen).forEach(function (col) {
      widths[col] = Math.min(42, Math.max(10, maxLen[col] + 2));
    });
    return widths;
  }

  function buildColsXml(widths) {
    var cols = Object.keys(widths).sort(function (a, b) { return colToNum(a) - colToNum(b); });
    if (!cols.length) return "";
    var parts = ["<cols>"];
    cols.forEach(function (col) {
      var n = colToNum(col);
      parts.push("<col min=\"" + n + "\" max=\"" + n + "\" width=\"" + widths[col] + "\" customWidth=\"1\"/>");
    });
    parts.push("</cols>");
    return parts.join("");
  }

  // Style one worksheet XML string: header row (row 1) gets the green
  // header style, every other data cell gets the bordered body style,
  // column widths are fitted, the header row is frozen and given height
  // for its wrapped text, and print setup is added so the sheet prints
  // directly (portrait A4, all columns on one page wide).
  function styleSheetXml(xml) {
    var sheetDataMatch = /<sheetData>[\s\S]*?<\/sheetData>/.exec(xml);
    if (!sheetDataMatch) return xml;
    var sheetData = sheetDataMatch[0];
    var widths = computeColumnWidths(sheetData);

    // Set the style index on every cell: row 1 -> header, rest -> body.
    var styledData = sheetData.replace(/<c\b([^>]*)>/g, function (tag, attrs) {
      var rMatch = /r="([A-Z]+)(\d+)"/.exec(attrs);
      if (!rMatch) return tag;
      var row = parseInt(rMatch[2], 10);
      var cleaned = attrs.replace(/\ss="[^"]*"/, "");
      return "<c" + cleaned + " s=\"" + (row === 1 ? HEADER_XF : BODY_XF) + "\">";
    });

    var out = xml.replace(sheetData, styledData);

    // Header row height for wrapped text.
    out = out.replace(/<row r="1"(?![^>]*\bht=)/, "<row r=\"1\" ht=\"30\" customHeight=\"1\"");

    // Column widths, right before <sheetData>.
    var colsXml = buildColsXml(widths);
    if (colsXml) {
      out = out.replace(/<cols>[\s\S]*?<\/cols>/, "");
      out = out.replace("<sheetData>", colsXml + "<sheetData>");
    }

    // Freeze the header row.
    out = out.replace(
      /<sheetView\b([^>]*?)\/>/,
      "<sheetView$1><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\" activeCell=\"A2\" sqref=\"A2\"/></sheetView>",
    );

    // Print setup: enable fit-to-page (sheetPr) and add page margins +
    // page setup (portrait A4, fit all columns to one page wide).
    if (!/<pageSetUpPr\b/.test(out)) {
      out = out.replace(/<worksheet\b([^>]*)>/, "<worksheet$1>" + SHEET_PR_XML);
    }
    if (!/<pageSetup\b/.test(out)) {
      out = out.replace("</sheetData>", "</sheetData>" + PAGE_SETUP_XML);
    }
    return out;
  }

  // Repeat row 1 on every printed page, for every sheet. This is a
  // workbook-level defined name (_xlnm.Print_Titles) per sheet.
  function addPrintTitles(workbookXml) {
    var sheetNames = [];
    var sheetRe = /<sheet\b[^>]*\bname="([^"]*)"/g;
    var m;
    while ((m = sheetRe.exec(workbookXml)) !== null) sheetNames.push(m[1]);
    if (!sheetNames.length) return workbookXml;
    if (workbookXml.indexOf("_xlnm.Print_Titles") !== -1) return workbookXml;
    var entries = sheetNames.map(function (name, i) {
      var safe = String(name).replace(/'/g, "''");
      return "<definedName name=\"_xlnm.Print_Titles\" localSheetId=\"" + i + "\">'" + safe + "'!$1:$1</definedName>";
    }).join("");
    var definedNames = "<definedNames>" + entries + "</definedNames>";
    if (/<\/sheets>/.test(workbookXml)) {
      return workbookXml.replace(/<\/sheets>/, "</sheets>" + definedNames);
    }
    return workbookXml.replace(/<\/workbook>/, definedNames + "</workbook>");
  }

  async function styleXlsxBlob(blob) {
    try {
      var JSZip = global.JSZip;
      if (!JSZip || !blob) return blob;
      var buf = await blob.arrayBuffer();
      var zip = await JSZip.loadAsync(buf);
      var sheetFiles = Object.keys(zip.files).filter(function (name) {
        return /^xl\/worksheets\/sheet\d+\.xml$/.test(name);
      });
      if (!sheetFiles.length) return blob;
      zip.file("xl/styles.xml", STYLES_XML);
      for (var i = 0; i < sheetFiles.length; i++) {
        var name = sheetFiles[i];
        var xml = await zip.file(name).async("string");
        zip.file(name, styleSheetXml(xml));
      }
      var workbookFile = zip.file("xl/workbook.xml");
      if (workbookFile) {
        var workbookXml = await workbookFile.async("string");
        zip.file("xl/workbook.xml", addPrintTitles(workbookXml));
      }
      var out = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      return out;
    } catch (error) {
      // Never break a download because styling failed.
      if (typeof console !== "undefined" && console.warn) console.warn("ZBOM xlsx styling skipped:", error);
      return blob;
    }
  }

  var api = { styleXlsxBlob: styleXlsxBlob, styleSheetXml: styleSheetXml, addPrintTitles: addPrintTitles, STYLES_XML: STYLES_XML };
  global.ZBOM_XLSX_STYLE = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
