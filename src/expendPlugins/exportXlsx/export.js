import { luckysheet } from "../../core";
import Excel from "exceljs";
import FileSaver from "file-saver";
import { setStyleAndValue } from "./cellStyle";
import { setMerge } from "./merge";
import { setBorder } from "./border";
import { setImages } from "./image";
import { setHyperlink } from "./hyperlink";
import { setFrozen } from "./frozen";

function localExport(order, success) {
  const sheetInfo = luckysheet.toJson();
  console.log("开始导出", order, sheetInfo);
  // 获取需要导出的工作表
  const exportSheet =
    order === "all" ? sheetInfo.data : [sheetInfo.data[order]];

  const workbook = new Excel.Workbook();
  // 写入工作薄
  exportSheet.forEach(sheet => {
    const worksheet = workbook.addWorksheet(sheet.name);
    // 设置工作表样式
    setStyleAndValue(sheet.data, worksheet);
    setMerge((sheet.config && sheet.config.merge) || {}, worksheet);
    setBorder((sheet.config && sheet.config.borderInfo) || {}, worksheet);
    setImages(sheet, worksheet, workbook);
    setHyperlink(sheet.hyperlink, worksheet);
    setFrozen(sheet.frozen, worksheet);
  });
  // 写入 buffer
  workbook.xlsx.writeBuffer().then(data => {
    const blob = new Blob([data], {
      type: "application/vnd.ms-excel;charset=utf-8",
    });
    FileSaver.saveAs(blob, `${sheetInfo.title}.xlsx`);
  });
  success && success();
}

export { localExport };
