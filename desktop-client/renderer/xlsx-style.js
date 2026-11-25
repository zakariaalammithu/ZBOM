// Professional styling for ZBOM Excel downloads.
//
// SheetJS (community edition) cannot write cell styles, so this module
// post-processes a SheetJS-generated .xlsx (a zip) with JSZip: it injects
// a styles.xml (fonts / fills / borders) and sets each cell's style
// index in the sheet XML. The look matches the user's own formatted
// template: light-green header (C3D69B), dark bold text, thin borders
// on every cell, column widths fitted to content, header row frozen.
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
  // for its wrapped text.
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
    return out;
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
      var out = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      return out;
    } catch (error) {
      // Never break a download because styling failed.
      if (typeof console !== "undefined" && console.warn) console.warn("ZBOM xlsx styling skipped:", error);
      return blob;
    }
  }

  var api = { styleXlsxBlob: styleXlsxBlob, styleSheetXml: styleSheetXml, STYLES_XML: STYLES_XML };
  global.ZBOM_XLSX_STYLE = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
