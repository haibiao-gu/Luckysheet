import { luckysheet } from "../../core";
import Excel from "exceljs";
import FileSaver from "file-saver";
import { setStyleAndValue } from "./cellStyle";
import { setMerge } from "./merge";
import { setBorder } from "./border";
import { setImages } from "./image";
import { setHyperlink } from "./hyperlink";
import { setFrozen } from "./frozen";
import { setDimensions } from "./dimension";
import { setHidden } from "./hidden";
import { setEcharts } from "./echarts";

async function localExport(order, success) {
  const sheetInfo = luckysheet.toJson();
  // console.log("开始导出", order, sheetInfo);
  // 获取需要导出的工作表
  const exportSheet =
    order === "all" ? sheetInfo.data : [sheetInfo.data[order]];

  const workbook = new Excel.Workbook();
  // 写入工作薄
  let i = 0;
  for (const sheet of exportSheet) {
    // 跳转到工作表 i
    luckysheet.setSheetActive(i);

    // 等待500ms确保工作表切换完成
    await new Promise(resolve => setTimeout(resolve, 500));

    const worksheet = await workbook.addWorksheet(sheet.name, {
      properties: {
        // 工作表标签颜色
        tabColor: {
          argb: sheet.color?.replace("#", "") || "",
        },
      },
      // 网格线
      views: [{ showGridLines: sheet.showGridLines === "1" }],
    });

    // 设置工作表样式
    setDimensions(sheet.config, worksheet);
    setStyleAndValue(sheet.data, worksheet);
    setMerge((sheet.config && sheet.config.merge) || {}, worksheet);
    setBorder(
      (sheet.config && sheet.config.borderInfo) || {},
      worksheet,
      (sheet.config && sheet.config.merge) || {}
    );
    // 处理隐藏的行和列
    setHidden(sheet.config, worksheet);
    await setEcharts(sheet, worksheet, workbook);
    await setImages(sheet, worksheet, workbook);
    setHyperlink(sheet.hyperlink, worksheet);
    setFrozen(sheet.frozen, worksheet);
    i++;
  }

  // 写入 buffer
  const data = await workbook.xlsx.writeBuffer();
  const blob = new Blob([data], {
    type: "application/vnd.ms-excel;charset=utf-8",
  });
  FileSaver.saveAs(blob, `${sheetInfo.title}.xlsx`);
  success && success();
}

export { localExport };
