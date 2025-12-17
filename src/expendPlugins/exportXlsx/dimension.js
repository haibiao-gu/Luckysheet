// 设置工作表的行高和列宽
function setDimensions(config, worksheet) {
  // 处理列宽
  if (config && config.columnlen) {
    const columnLengths = config.columnlen;
    // console.log("开始设置列宽", columnLengths);
    for (const colIndex in columnLengths) {
      // ExcelJS中的列索引从1开始，而Luckysheet从0开始
      const excelColIndex = parseInt(colIndex) + 1;
      // Luckysheet的列宽单位是像素，ExcelJS使用字符宽度
      worksheet.getColumn(excelColIndex).width = columnLengths[colIndex] / 7;
    }
  }

  // 处理行高
  if (config && config.rowlen) {
    const rowLengths = config.rowlen;
    // console.log("开始设置行高", rowLengths);
    for (const rowIndex in rowLengths) {
      // ExcelJS中的行索引从1开始，而Luckysheet从0开始
      const excelRowIndex = parseInt(rowIndex) + 1;
      // Luckysheet的行高单位是像素，ExcelJS也使用像素
      worksheet.getRow(excelRowIndex).height = rowLengths[rowIndex] * 0.75;
    }
  }
}

export { setDimensions };
