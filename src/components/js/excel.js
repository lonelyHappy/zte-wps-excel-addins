

/**
 * 选中文本
 * @returns Array: {text:string, isTranslate:boolean} []
 */
function GetSeletionWord() {
  const excel = wps.EtApplication().ActiveWorkbook
  if (!excel) {
      throw new Error("当前没有打开任何工作簿")
  }
  const sheet = excel.ActiveSheet;
  if (!sheet) {
    throw new Error("当前没有打开任何工作表")
  }
  const selection = wps.EtApplication().Selection
  if (!selection) {
    throw new Error("没有选中任何单元格");
  }
  const selectedCells = selection.Cells;
  const results = []
  for (let row = selection.Row; row <= selection.Row + selectedCells.Rows.Count - 1; row++) {
    for (let col = selection.Column; col <= selection.Column + selectedCells.Columns.Count - 1; col++) {
      const cell = sheet.Cells.Item(row, col);
      const resultsItem = {
        text: '',
        isTranslate: false,
        row: row,
        col: col
      }
      // console.log(cell)
      try {
        if (cell.Value2 && cell.Value2.trim().length > 0) {
          resultsItem.text = cell.Value2
          resultsItem.isTranslate = true
        }
        results.push(resultsItem)
      } catch (error) {
        console.log(`无法访问单元格(${row}, ${col})的值: ${error}`);
      }
    }
  }
  // Sentences.split(rgSel.Text, rgSel.Start).forEach(item => results.push(item))
  return results
}
/**
 * 替换文本
 * @param {*} item 当前文本 {text:string, isTranslate:boolean, translate:string}
 * @param {*} prev 上一个文本
 * @param {*} next 下一个文本
 * @returns Promise
 */
function ReplaceWord (item) {
  // console.log(item)
  const excel = wps.EtApplication().ActiveWorkbook
  if (!excel) {
      throw new Error("当前没有打开任何工作簿")
  }
  const sheet = excel.ActiveSheet;
  if (!sheet) {
    throw new Error("当前没有打开任何工作表")
  }
  const selection = wps.EtApplication().Selection
  if (!selection) {
    throw new Error("没有选中任何单元格");
  }
  if (!item.isTranslate) {
    return Promise.resolve()
  }
  // console.log('ssss')
  // const cell = sheet.Cells.Item(1, 1);
  // console.log(cell.Value2)
  sheet.Cells.Item(item.row, item.col).Value2 = item.translate
  if (selection)
    selection.Range.Select()
  return Promise.resolve()
}
/**
 * 原文后插入译文
 * @param {*} item 当前文本 {text:string, isTranslate:boolean, translate:string}
 * @param {*} prev 上一个文本
 * @param {*} next 下一个文本
 * @returns Promise
 */
function InsertAfter(item) {
  const excel = wps.EtApplication().ActiveWorkbook
  if (!excel) {
      throw new Error("当前没有打开任何工作簿")
  }
  const sheet = excel.ActiveSheet;
  if (!sheet) {
    throw new Error("当前没有打开任何工作表")
  }
  const selection = wps.EtApplication().Selection
  if (!selection) {
    throw new Error("没有选中任何单元格");
  }
  // if (prev) {
  //   // 将item的start设置为prev的end，同时item的end根据item的length位移
  //   item.end = prev.end + item.length
  //   item.start = prev.end
  // }
  if (!item.isTranslate) {
    return Promise.resolve()
  }
  // const selectedCells = wps.EtApplication().Selection.Cells
  // excel.Range(item.end, item.end).Text = item.translate
  item.translate = item.text + item.translate;
  sheet.Cells.Item(item.row, item.col).Value2 = item.translate
  // Sentences.updateData(item, item.translate);
  if (selection)
    selection.Range.Select()
  return Promise.resolve()
}

function InsertRight (item) {
  const excel = wps.EtApplication().ActiveWorkbook
  if (!excel) {
      throw new Error("当前没有打开任何工作簿")
  }
  const sheet = excel.ActiveSheet;
  if (!sheet) {
    throw new Error("当前没有打开任何工作表")
  }
  const selection = wps.EtApplication().Selection
  if (!selection) {
    throw new Error("没有选中任何单元格");
  }
  if (!item.isTranslate) {
    return Promise.resolve()
  }
  sheet.Cells.Item(item.row, item.col + 1).Value2 = item.translate
  if (selection)
    selection.Range.Select()
  return Promise.resolve()
}
 
function InsertUnder (item) {
  const excel = wps.EtApplication().ActiveWorkbook
  if (!excel) {
      throw new Error("当前没有打开任何工作簿")
  }
  const sheet = excel.ActiveSheet;
  if (!sheet) {
    throw new Error("当前没有打开任何工作表")
  }
  const selection = wps.EtApplication().Selection
  if (!selection) {
    throw new Error("没有选中任何单元格");
  }
  if (!item.isTranslate) {
    return Promise.resolve()
  }
  sheet.Cells.Item(item.row + 1, item.col).Value2 = item.translate
  if (selection)
    selection.Range.Select()
  return Promise.resolve()
 }

export default {
  GetSeletionWord,
  ReplaceWord,
  InsertAfter,
  InsertRight,
  InsertUnder,
}