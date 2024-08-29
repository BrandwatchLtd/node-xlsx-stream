(function() {
  var esc, utils, xml;

  utils = require('./utils');

  xml = utils.compress;

  esc = utils.escapeXML;

  module.exports = {
    worksheet: {
      header: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
  <sheetData>`),
      footer: xml(`</sheetData>`),
      hyperLinkStart: xml(`<hyperlinks>`),
      hyperLink: function(link, rId) {
        var escapedLink, parts, xmlString;
        parts = link.split('-');
        escapedLink = parts[1].replace(/&/g, '&amp;');
        xmlString = `<hyperlink display="${escapedLink}" r:id="rId${rId}" ref="${parts[0]}" />`;
        return xml(xmlString);
      },
      hyperLinkEnd: xml(`</hyperlinks>`),
      endSheet: xml(`</worksheet>`)
    },
    sheet_related: {
      "[Content_Types].xml": {
        header: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`),
        sheet: function(sheet) {
          return `<Override PartName="/${esc(sheet.path)}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
        },
        footer: xml(`</Types>`)
      },
      "xl/_rels/workbook.xml.rels": {
        header: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`),
        sheet: function(sheet) {
          return `<Relationship Id="rSheet${esc(sheet.index)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${esc(sheet.rel)}"/>`;
        },
        footer: xml(`  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`)
      },
      "xl/workbook.xml": {
        header: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9303"/>
  <workbookPr defaultThemeVersion="124226"/>
  <bookViews>
    <workbookView xWindow="480" yWindow="60" windowWidth="18195" windowHeight="8505"/>
  </bookViews>
  <sheets>`),
        sheet: function(sheet) {
          return xml(`<sheet name="${esc(sheet.name)}" sheetId="${esc(sheet.index)}" r:id="rSheet${esc(sheet.index)}"/>`);
        },
        footer: xml(`  </sheets>
  <calcPr calcId="145621"/>
</workbook>`)
      }
    },
    styles: function(styl) {
      var cellXfItems, cellXfs, i, item, j, len, len1, numFmtItems, numFmts, ref, ref1;
      numFmtItems = "";
      ref = styl.numFmts;
      for (i = 0, len = ref.length; i < len; i++) {
        item = ref[i];
        numFmtItems += `<numFmt numFmtId="${item.numFmtId}" formatCode="${esc(item.formatCode)}" />`;
      }
      numFmts = numFmtItems ? `<numFmts count="${styl.numFmts.length}">
  ${numFmtItems}
</numFmts>` : "";
      cellXfItems = "";
      ref1 = styl.cellStyleXfs;
      for (j = 0, len1 = ref1.length; j < len1; j++) {
        item = ref1[j];
        cellXfItems += `<xf xfId="0" fontId="0" fillId="0" borderId="0" numFmtId="${item.numFmtId}" applyNumberFormat="1"/>`;
      }
      cellXfs = cellXfItems ? `<cellXfs count="${Object.keys(styl.cellStyleXfs).length}">
  ${cellXfItems}
</cellXfs>` : "";
      return xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  ${numFmts}
  <fonts count="1" x14ac:knownFonts="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  ${cellXfs}
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
  <extLst>
    <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
      <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
    </ext>
  </extLst>
</styleSheet>`);
    },
    statics: {
      "_rels/.rels": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`),
      "xl/sharedStrings.xml": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>`),
      "docProps/app.xml": xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>node-xlsx-stream</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company>Microsoft Corporation</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>${require('../package.json').version}</AppVersion>
</Properties>`)
    },
    semiStatics: {
      "docProps/core.xml": function(opts) {
        var today;
        today = new Date().toISOString();
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>node-xlsx-stream</dc:creator>
  <cp:lastModifiedBy>node-xlsx-stream</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${today}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${today}</dcterms:modified>
</cp:coreProperties>`;
      }
    },
    rels: {
      "xl/worksheets/_rels/sheet1.xml.rels": function(links) {
        var escapedLink, i, len, link, linksCounter, parts, xmlString;
        xmlString = xml(`<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`);
        linksCounter = 0;
        for (i = 0, len = links.length; i < len; i++) {
          link = links[i];
          linksCounter++;
          parts = link.split('-');
          escapedLink = parts[1].replace(/&/g, '&amp;');
          xmlString += xml(`<Relationship Id="rId${linksCounter}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapedLink}" TargetMode="External"/>`);
        }
        xmlString += xml(`</Relationships>`);
        return xmlString;
      }
    }
  };

}).call(this);
