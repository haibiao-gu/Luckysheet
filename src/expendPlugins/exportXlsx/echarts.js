import html2canvas from "html2canvas";
import { getImagePosition, calculatePositionsFromLengths } from "./image";

export async function setEcharts(sheet, worksheet, workbook) {
  // 从 sheet 数据中获取 echarts 图表信息
  const echarts = sheet.chart || [];
  const config = sheet.config;
  if (typeof echarts != "object" || !echarts) return;
  // console.log("开始导出图表", echarts);

  // 获取列宽和行高配置，如果没有则使用空对象
  const columnLen = (config && config.columnlen) || {};
  const rowLen = (config && config.rowlen) || {};

  // 将对象格式转换为数组格式
  const columnPositions = calculatePositionsFromLengths(columnLen, true);
  const rowPositions = calculatePositionsFromLengths(rowLen, false);

  for (const chart of echarts) {
    // 将 echarts 图表转换为图片并添加到工作表
    // 具体实现取决于 echarts 图表在 sheet 中的存储格式
    if (chart.chart_id) {
      const chartElement = document.getElementById(chart.chart_id);
      // console.log("chartElement", chartElement);
      if (chartElement) {
        try {
          // 使用 html2canvas 将 DOM 元素转换为 canvas，然后获取图片数据
          const canvas = await html2canvas(chartElement, {
            backgroundColor: "#fff",
            scale: 2, // 提高图片质量
          });

          const imgData = canvas.toDataURL("image/png");

          //开始行 开始列 结束行 结束列
          const imageId = workbook.addImage({
            base64: imgData,
            extension: "png",
          });

          const col_st = getImagePosition(chart.left, columnPositions);
          const row_st = getImagePosition(chart.top, rowPositions);
          // console.log("tl", col_st, row_st);
          // console.log("ext", chart.width, chart.height);
          //模式1，图片左侧与luckysheet位置一样，像素比例保持不变，但是，右侧位置可能与原图所在单元格不一致
          worksheet.addImage(imageId, {
            tl: { col: col_st, row: row_st },
            ext: { width: chart.width, height: chart.height },
          });
        } catch (error) {
          console.error("图表导出失败:", error);
        }
      }
    }
  }
}
