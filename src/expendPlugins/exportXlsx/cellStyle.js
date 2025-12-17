import { fillConvert, fontConvert, alignmentConvert } from "./styleUtils";

// 设置单元格样式和值
function setStyleAndValue(cellArr, worksheet) {
  if (!Array.isArray(cellArr)) return;
  // console.log("开始设置样式和值", cellArr);

  cellArr.forEach((row, rowid) => {
    row.every((cell, columnid) => {
      if (!cell) return true;
      let fill = fillConvert(cell.bg);
      let font = fontConvert(
        cell.ff || "Times New Roman",
        cell.fc,
        cell.bl,
        cell.it,
        cell.fs,
        cell.cl,
        cell.ul
      );
      let alignment = alignmentConvert(cell.vt, cell.ht, cell.tb, cell.tr);
      let value;

      let v = "";
      if (cell.ct && cell.ct.t === "inlineStr") {
        const s = cell.ct.s;
        s.forEach(val => {
          v += val.v;
        });
      } else {
        //导出后取显示值
        v = cell.m;
      }

      // 将数字字符串转换为数字类型
      if (typeof v === "string" && !isNaN(v) && !isNaN(parseFloat(v))) {
        v = parseFloat(v);
      }

      if (cell.f) {
        value = { formula: cell.f, result: v };
      } else {
        value = v;
      }
      let target = worksheet.getCell(rowid + 1, columnid + 1);
      //添加批注
      if (cell.ps) {
        let ps = cell.ps;
        target.note = ps.value;
      }
      //单元格填充
      target.fill = fill;
      //单元格字体
      target.font = font;
      target.alignment = alignment;
      target.value = value;
      return true;
    });
  });
}

export { setStyleAndValue };
