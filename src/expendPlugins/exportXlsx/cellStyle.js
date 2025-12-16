import { luckysheet } from "../../core";
import { fillConvert, fontConvert, alignmentConvert } from "./styleUtils";

// 设置单元格样式和值
function setStyleAndValue(cellArr, worksheet) {
  if (!Array.isArray(cellArr)) return;

  cellArr.forEach(function (row, rowid) {
    const dbrow = worksheet.getRow(rowid + 1);
    //设置单元格行高,默认乘以0.8倍
    dbrow.height = luckysheet.getRowHeight([rowid])[rowid] * 0.8;
    row.every(function (cell, columnid) {
      if (!cell) return true;
      if (rowid === 0) {
        const dobCol = worksheet.getColumn(columnid + 1);
        //设置单元格列宽除以8
        dobCol.width = luckysheet.getColumnWidth([columnid])[columnid] / 8;
      }
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
