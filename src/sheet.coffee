_ = require "lodash"
through = require('through')

utils = require('./utils')
template = require('./templates')
worksheetTemplates = template.worksheet

module.exports = sheetStream = (zip, sheet, opts={})->
  # 列番号の26進表記(A, B, .., Z, AA, AB, ..)
  # 一度計算したらキャッシュしておく。
  colChar = _.memoize utils.colChar
  links = []

  # 行ごとに変換してxl/worksheets/sheet1.xml に追加
  nRow = 0
  onData = (row)->
    nRow++
    buf = "<row r='#{nRow}'>"
    if opts.columns?
      buf += utils.buildCell("#{colChar(i)}#{nRow}", row[col], sheet.styles) for col, i in opts.columns
    else
      buf += utils.buildCell("#{colChar(i)}#{nRow}", val, sheet.styles) for val, i in row
    buf += '</row>'
    @queue buf

  onEnd = ->
    @queue worksheetTemplates.footer

    if links.length > 0
      rel = template.rels
      for name, func of rel
        zip.append func(links), name: name

      @queue worksheetTemplates.hyperLinkStart
      linkCounter = 0
      for link in links
        linkCounter++
        @queue worksheetTemplates.hyperLink(link, linkCounter)
      @queue worksheetTemplates.hyperLinkEnd

    @queue worksheetTemplates.endSheet
    @queue null
    converter = colChar = zip = null

  converter = through(onData, onEnd)
  zip.append converter, name: sheet.path, store: opts.store

  # ヘッダ部分を追加
  converter.queue worksheetTemplates.header

  return converter
