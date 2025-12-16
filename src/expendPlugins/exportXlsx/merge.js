// 设置合并数据
function setMerge(luckyMerge = {}, worksheet) {
  const mergearr = Object.values(luckyMerge);
  mergearr.forEach(function (elem) {
    // elem格式：{r: 0, c: 0, rs: 1, cs: 2}
    // 按开始行，开始列，结束行，结束列合并（相当于 K10:M12）
    worksheet.mergeCells(
      elem.r + 1,
      elem.c + 1,
      elem.r + elem.rs,
      elem.c + elem.cs
    );
  });
}

export { setMerge };
