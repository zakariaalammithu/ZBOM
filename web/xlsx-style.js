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
    // Merge into an existing <definedNames> block if one is already there
    // (SheetJS writes one for _FilterDatabase when a sheet has an
    // autofilter); a workbook may only have a single <definedNames>.
    if (/<definedNames>/.test(workbookXml)) {
      return workbookXml.replace(/(<definedNames>)/, "$1" + entries);
    }
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


  // ------------------------------------------------------------------
  // Production Calculation report styling (3-sheet professional report).
  //
  // The report has a fixed layout (matching ZBOM_Production_Result.xlsx):
  //   "Machine Details":     title row 1, subtitle row 2, header row 4,
  //                          data from row 5, one row per machine.
  //   "Material-wise Yarn %": title row 1, subtitle row 2, header row 4,
  //                          data from row 5, one row per material plus a
  //                          navy "M/C TOTAL" row per machine, then a
  //                          "Note:" row.
  //   "Production Summary":  title row 1, subtitle row 2, section header
  //                          row 4 ("Production Overview"), key-value rows
  //                          from row 5 (label in A, value in B).
  //
  // Design: navy title (#17365D, white bold 17), grey subtitle, blue
  // header (#0B5CAD, white bold 11, medium bottom border), data rows with
  // a thin bottom border and alternating light fill (#EAF2F8), navy TOTAL
  // rows (white bold), amber note row. Production kg as #,##0.00 /
  // #,##0.000 and percentages as 0.000"%". Landscape print, fit to one
  // page wide, header row (row 4) repeated on printed pages for the two
  // table sheets. The freeze pane is set below the header row (row 4).
  //
  // Strings are inline in the SheetJS output (t="str"), so the styler
  // can read cell values straight from the sheet XML to find TOTAL rows,
  // the note row, kg columns (header contains "Production (kg)") and the
  // summary value format (label ends with "(kg)" or "%").
  // ------------------------------------------------------------------

  var REPORT_NUMFMT_KG2 = 165; // #,##0.00
  var REPORT_NUMFMT_KG3 = 166; // #,##0.000
  var REPORT_NUMFMT_PCT = 167; // 0.000"%"

  var REPORT_STYLES_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="3">' +
    '<numFmt numFmtId="165" formatCode="#,##0.00"/>' +
    '<numFmt numFmtId="166" formatCode="#,##0.000"/>' +
    '<numFmt numFmtId="167" formatCode="0.000&quot;%&quot;"/>' +
    '</numFmts>' +
    '<fonts count="9">' +
    '<font><name val="Calibri"/><sz val="10"/><color rgb="FF263445"/></font>' +
    '<font><name val="Calibri"/><sz val="11"/><b/><color rgb="FFFFFFFF"/></font>' +
    '<font><name val="Calibri"/><sz val="17"/><b/><color rgb="FFFFFFFF"/></font>' +
    '<font><name val="Calibri"/><sz val="10"/><color rgb="FF64748B"/></font>' +
    '<font><name val="Calibri"/><sz val="11"/><b/><color rgb="FFFFFFFF"/></font>' +
    '<font><name val="Calibri"/><sz val="11"/><b/><color rgb="FF263445"/></font>' +
    '<font><name val="Calibri"/><sz val="11"/><color rgb="FF0B5CAD"/></font>' +
    '<font><name val="Calibri"/><sz val="12"/><b/><color rgb="FF17365D"/></font>' +
    '<font><name val="Calibri"/><sz val="9"/><color rgb="FF7C5A00"/></font>' +
    '</fonts>' +
    '<fills count="7">' +
    '<fill><patternFill patternType="none"/></fill>' +
    '<fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF17365D"/><bgColor rgb="FF17365D"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF0B5CAD"/><bgColor rgb="FF0B5CAD"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFEAF2F8"/><bgColor rgb="FFEAF2F8"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF4D6"/><bgColor rgb="FFFFF4D6"/></patternFill></fill>' +
    '</fills>' +
    '<borders count="3">' +
    '<border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left/><right/><top/><bottom style="thin"/></border>' +
    '<border><left/><right/><top/><bottom style="medium"/></border>' +
    '</borders>' +
    '<cellStyleXfs count="1">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>' +
    '</cellStyleXfs>' +
    '<cellXfs count="25">' +
    // 0 Normal
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    // 1 Title
    '<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
    // 2 Subtitle
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    // 3 Header
    '<xf numFmtId="0" fontId="1" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
    // 4 Body
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    // 5 Body Alt
    '<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    // 6 Body kg2
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    // 7 Body kg2 Alt
    '<xf numFmtId="165" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    // 8 Body kg3
    '<xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    // 9 Body kg3 Alt
    '<xf numFmtId="166" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    // 10 Body pct
    '<xf numFmtId="167" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    // 11 Body pct Alt
    '<xf numFmtId="167" fontId="0" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    // 12 Total
    '<xf numFmtId="0" fontId="4" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    // 13 Total kg3
    '<xf numFmtId="166" fontId="4" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    // 14 Total pct
    '<xf numFmtId="167" fontId="4" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>' +
    // 15 Note
    '<xf numFmtId="0" fontId="8" fillId="6" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>' +
    // 16 Section
    '<xf numFmtId="0" fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    // 17 Summary Label
    '<xf numFmtId="0" fontId="5" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    // 18 Summary Label Alt
    '<xf numFmtId="0" fontId="5" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    // 19 Summary Value
    '<xf numFmtId="0" fontId="6" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    // 20 Summary Value Alt
    '<xf numFmtId="0" fontId="6" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    // 21 Summary Value kg2
    '<xf numFmtId="165" fontId="6" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    // 22 Summary Value kg2 Alt
    '<xf numFmtId="165" fontId="6" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    // 23 Summary Value pct
    '<xf numFmtId="167" fontId="6" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    // 24 Summary Value pct Alt
    '<xf numFmtId="167" fontId="6" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>' +
    '</cellXfs>' +
    '<cellStyles count="1">' +
    '<cellStyle name="Normal" xfId="0" builtinId="0"/>' +
    '</cellStyles>' +
    '<dxfs count="0"/>' +
    '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>' +
    '</styleSheet>';

  var REPORT_PAGE_SETUP_XML =
    '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>' +
    '<pageSetup orientation="landscape" paperSize="9" fitToWidth="1" fitToHeight="0"/>';
  var REPORT_SHEET_PR_XML = "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>";

  // Read a cell's inline string value from sheet XML. SheetJS writes
  // strings inline (t="str"), so the literal text is in <v>.
  function reportCellValue(sheetDataXml, col, row) {
    var re = new RegExp("<c r=\"" + col + row + '\"[^>]*>(?:<v>([^<]*)<\\/v>)?', "");
    var m = re.exec(sheetDataXml);
    return m ? decodeEntities(m[1] || "") : "";
  }

  // Style one report worksheet. sheetName decides the layout; row 1 is
  // the navy title, row 2 the grey subtitle, row 4 the header (or the
  // section heading on the summary sheet), data from row 5.
  function styleReportSheetXml(xml, sheetName) {
    var sheetDataMatch = /<sheetData>[\s\S]*?<\/sheetData>/.exec(xml);
    if (!sheetDataMatch) return xml;
    var sheetData = sheetDataMatch[0];

    // Which columns hold a "Production (kg)" value on Machine Details:
    // read the header row (row 4) and mark every column whose header
    // contains "Production (kg)".
    var kgCols = {};
    if (sheetName === "Machine Details") {
      var headerRe = /<c r="([A-Z]+)4"[^>]*>(?:<v>([^<]*)<\/v>)?/g;
      var hm;
      while ((hm = headerRe.exec(sheetData)) !== null) {
        if (decodeEntities(hm[2] || "").indexOf("Production (kg)") !== -1) kgCols[hm[1]] = true;
      }
    }

    var alt = 0; // alternating data-row fill, reset after each TOTAL row
    var sumIdx = 0; // summary key-value row index (alternates fill 4 / 5)
    var styledData = sheetData.replace(/<row\b([^>]*)>([\s\S]*?)<\/row>/g, function (rowTag, rowAttrs, rowInner) {
      var rMatch = /r="(\d+)"/.exec(rowAttrs);
      if (!rMatch) return rowTag;
      var row = parseInt(rMatch[1], 10);
      var isTotal = false;
      var isNote = false;
      if (sheetName === "Material-wise Yarn %") {
        isTotal = reportCellValue(sheetData, "B", row) === "M/C TOTAL";
        isNote = reportCellValue(sheetData, "A", row).indexOf("Note:") === 0;
        if (isTotal) alt = 0;
      }
      var newInner = rowInner.replace(/<c\b([^>]*)>/g, function (tag, attrs) {
        var cMatch = /r="([A-Z]+)(\d+)"/.exec(attrs);
        if (!cMatch) return tag;
        var col = cMatch[1];
        var cleaned = attrs.replace(/\ss="[^"]*"/, "");
        var xf;
        if (row === 1) {
          xf = 1; // title
        } else if (row === 2) {
          xf = 2; // subtitle
        } else if (sheetName === "Production Summary" && row === 4) {
          xf = 16; // section heading
        } else if (row === 4) {
          xf = 3; // header
        } else if (isNote) {
          xf = 15; // note
        } else if (isTotal) {
          if (sheetName === "Material-wise Yarn %" && col === "D") xf = 13;
          else if (sheetName === "Material-wise Yarn %" && col === "E") xf = 14;
          else xf = 12;
        } else if (sheetName === "Production Summary" && row >= 5) {
          // Key-value rows: label in A, value in B, alternating fills.
          var label = reportCellValue(sheetData, "A", row);
          var isKg = label.slice(-4) === "(kg)";
          var isPct = label.slice(-7) === "Share %";
          if (col === "A") {
            xf = sumIdx % 2 === 0 ? 17 : 18;
          } else if (isKg) {
            xf = sumIdx % 2 === 0 ? 21 : 22;
          } else if (isPct) {
            xf = sumIdx % 2 === 0 ? 23 : 24;
          } else {
            xf = sumIdx % 2 === 0 ? 19 : 20;
          }
        } else if (row >= 5) {
          // Data rows.
          if (sheetName === "Material-wise Yarn %" && col === "D") xf = alt % 2 === 0 ? 8 : 9;
          else if (sheetName === "Material-wise Yarn %" && col === "E") xf = alt % 2 === 0 ? 10 : 11;
          else if (sheetName === "Machine Details" && kgCols[col]) xf = alt % 2 === 0 ? 6 : 7;
          else xf = alt % 2 === 0 ? 4 : 5;
        } else {
          xf = 0;
        }
        return "<c" + cleaned + ' s="' + xf + '">';
      });
      var out = "<row" + rowAttrs + ">" + newInner + "</row>";
      // Row heights: title and header taller for their text.
      if (row === 1 && !/\bht=/.test(rowAttrs)) out = out.replace("<row", '<row ht="28" customHeight="1"');
      if (row === 4 && !/\bht=/.test(rowAttrs)) out = out.replace("<row", '<row ht="30" customHeight="1"');
      if (isNote && !/\bht=/.test(rowAttrs)) out = out.replace("<row", '<row ht="42" customHeight="1"');
      if ((sheetName === "Material-wise Yarn %" || sheetName === "Machine Details") && row >= 5 && !isTotal && !isNote) alt += 1;
      if (sheetName === "Production Summary" && row >= 5) sumIdx += 1;
      return out;
    });

    var out = xml.replace(sheetData, styledData);

    // Freeze below the header row (row 4 -> topLeftCell A5). SheetJS does
    // not write !freeze, so the pane is added here.
    out = out.replace(/<sheetView\b([^>]*?)\/>/, '<sheetView$1><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A5" sqref="A5"/></sheetView>');
    out = out.replace(/<pane\b[^>]*\/>/, '<pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/>');

    // Print setup: landscape A4, fit all columns to one page wide.
    if (!/<pageSetUpPr\b/.test(out)) {
      out = out.replace(/<worksheet\b([^>]*)>/, "<worksheet$1>" + REPORT_SHEET_PR_XML);
    }
    if (!/<pageSetup\b/.test(out)) {
      out = out.replace("</sheetData>", "</sheetData>" + REPORT_PAGE_SETUP_XML);
    }
    return out;
  }

  // Repeat the header row (row 4) on every printed page, for the two
  // table sheets. The summary sheet is a short key-value list and gets
  // no repeated row.
  function addReportPrintTitles(workbookXml, sheetNames) {
    if (workbookXml.indexOf("_xlnm.Print_Titles") !== -1) return workbookXml;
    var entries = [];
    sheetNames.forEach(function (name, i) {
      if (name === "Machine Details" || name === "Material-wise Yarn %") {
        var safe = String(name).replace(/'/g, "''");
        entries.push('<definedName name="_xlnm.Print_Titles" localSheetId="' + i + '">' + "'" + safe + "'!$4:$4</definedName>");
      }
    });
    if (!entries.length) return workbookXml;
    var insert = entries.join("");
    // Merge into an existing <definedNames> block if one is already there
    // (SheetJS writes one for _FilterDatabase when a sheet has an
    // autofilter); a workbook may only have a single <definedNames>.
    if (/<definedNames>/.test(workbookXml)) {
      return workbookXml.replace(/(<definedNames>)/, "$1" + insert);
    }
    var definedNames = "<definedNames>" + insert + "</definedNames>";
    if (/<\/sheets>/.test(workbookXml)) {
      return workbookXml.replace(/<\/sheets>/, "</sheets>" + definedNames);
    }
    return workbookXml.replace(/<\/workbook>/, definedNames + "</workbook>");
  }

  // Map sheet names to their worksheet files, via workbook.xml and its
  // relationships, so styling follows the sheet by name (not by order).
  function reportSheetFiles(workbookXml, relsXml) {
    var map = {};
    var sheetRe = /<sheet\b[^>]*\bname="([^"]*)"[^>]*\br:id="([^"]*)"/g;
    var relRe = /<Relationship\b[^>]*\bId="([^"]*)"[^>]*\bTarget="([^"]*)"/g;
    var rels = {};
    var rm;
    while ((rm = relRe.exec(relsXml || "")) !== null) rels[rm[1]] = rm[2];
    var sm;
    while ((sm = sheetRe.exec(workbookXml)) !== null) {
      var target = rels[sm[2]] || "";
      var file = target.replace(/^\/?xl\//, "xl/");
      if (file) map[sm[1]] = file;
    }
    return map;
  }

  async function styleProductionReportBlob(blob) {
    try {
      var JSZip = global.JSZip;
      if (!JSZip || !blob) return blob;
      var buf = await blob.arrayBuffer();
      var zip = await JSZip.loadAsync(buf);
      var sheetFiles = Object.keys(zip.files).filter(function (name) {
        return /^xl\/worksheets\/sheet\d+\.xml$/.test(name);
      });
      if (!sheetFiles.length) return blob;
      zip.file("xl/styles.xml", REPORT_STYLES_XML);
      var workbookXml = "";
      var workbookFile = zip.file("xl/workbook.xml");
      if (workbookFile) workbookXml = await workbookFile.async("string");
      var relsXml = "";
      var relsFile = zip.file("xl/_rels/workbook.xml.rels");
      if (relsFile) relsXml = await relsFile.async("string");
      var nameToFile = reportSheetFiles(workbookXml, relsXml);
      var fileToName = {};
      Object.keys(nameToFile).forEach(function (name) { fileToName[nameToFile[name]] = name; });
      var sheetNames = [];
      var nameRe = /<sheet\b[^>]*\bname="([^"]*)"/g;
      var nm;
      while ((nm = nameRe.exec(workbookXml)) !== null) sheetNames.push(nm[1]);
      for (var i = 0; i < sheetFiles.length; i++) {
        var name = sheetFiles[i];
        var sheetName = fileToName[name] || sheetNames[i] || "";
        var xml = await zip.file(name).async("string");
        zip.file(name, styleReportSheetXml(xml, sheetName));
      }
      if (workbookFile) {
        zip.file("xl/workbook.xml", addReportPrintTitles(workbookXml, sheetNames));
      }
      var out = await zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      return out;
    } catch (error) {
      // Never break a download because styling failed.
      if (typeof console !== "undefined" && console.warn) console.warn("ZBOM production report styling skipped:", error);
      return blob;
    }
  }

  var api = { styleXlsxBlob: styleXlsxBlob, styleSheetXml: styleSheetXml, addPrintTitles: addPrintTitles, STYLES_XML: STYLES_XML, styleProductionReportBlob: styleProductionReportBlob, REPORT_STYLES_XML: REPORT_STYLES_XML };
  global.ZBOM_XLSX_STYLE = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof window !== "undefined" ? window : globalThis);
