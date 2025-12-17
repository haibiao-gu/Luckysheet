function setHidden(config, worksheet) {
  if (!config) return;

  // 处理隐藏的行
  if (config.rowhidden) {
    Object.keys(config.rowhidden).forEach(rowIndex => {
      const row = worksheet.getRow(parseInt(rowIndex) + 1);
      row.hidden = true;
    });
  }

  // 处理隐藏的列
  if (config.colhidden) {
    Object.keys(config.colhidden).forEach(colIndex => {
      const col = worksheet.getColumn(parseInt(colIndex) + 1);
      col.hidden = true;
    });
  }
}

export { setHidden };
