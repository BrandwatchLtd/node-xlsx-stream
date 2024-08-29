(function() {
  var _, sheetStream, template, through, utils, worksheetTemplates;

  _ = require("lodash");

  through = require('through');

  utils = require('./utils');

  template = require('./templates');

  worksheetTemplates = template.worksheet;

  module.exports = sheetStream = function(zip, sheet, opts = {}) {
    var colChar, converter, links, nRow, onData, onEnd;
    // 列番号の26進表記(A, B, .., Z, AA, AB, ..)
    // 一度計算したらキャッシュしておく。
    colChar = _.memoize(utils.colChar);
    links = [];
    // 行ごとに変換してxl/worksheets/sheet1.xml に追加
    nRow = 0;
    onData = function(row) {
      var buf, col, i, j, k, len, len1, ref, val;
      nRow++;
      buf = `<row r='${nRow}'>`;
      if (opts.columns != null) {
        ref = opts.columns;
        for (i = j = 0, len = ref.length; j < len; i = ++j) {
          col = ref[i];
          buf += utils.buildCell(`${colChar(i)}${nRow}`, row[col], sheet.styles);
        }
      } else {
        for (i = k = 0, len1 = row.length; k < len1; i = ++k) {
          val = row[i];
          buf += utils.buildCell(`${colChar(i)}${nRow}`, val, sheet.styles);
        }
      }
      buf += '</row>';
      return this.queue(buf);
    };
    onEnd = function() {
      var converter, func, j, len, link, linkCounter, name, rel;
      this.queue(worksheetTemplates.footer);
      if (links.length > 0) {
        rel = template.rels;
        for (name in rel) {
          func = rel[name];
          zip.append(func(links), {
            name: name
          });
        }
        this.queue(worksheetTemplates.hyperLinkStart);
        linkCounter = 0;
        for (j = 0, len = links.length; j < len; j++) {
          link = links[j];
          linkCounter++;
          this.queue(worksheetTemplates.hyperLink(link, linkCounter));
        }
        this.queue(worksheetTemplates.hyperLinkEnd);
      }
      this.queue(worksheetTemplates.endSheet);
      this.queue(null);
      return converter = colChar = zip = null;
    };
    converter = through(onData, onEnd);
    zip.append(converter, {
      name: sheet.path,
      store: opts.store
    });
    // ヘッダ部分を追加
    converter.queue(worksheetTemplates.header);
    return converter;
  };

}).call(this);
