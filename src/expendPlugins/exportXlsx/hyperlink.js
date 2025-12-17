import { fontConvert } from "./styleUtils";

// 设置超链接
function setHyperlink(hyperlink, worksheet) {
  if (!hyperlink) return;
  // console.log("开始设置超链接", hyperlink);
  for (const key in hyperlink) {
    const row_col = key.split("_");
    let cell = worksheet.getCell(
      Number(row_col[0]) + 1,
      Number(row_col[1]) + 1
    );
    let font = cell.style.font;
    //设置导出后超链接的样式
    cell.font = fontConvert(font.name, "#0000ff", 0, 0, font.size, 0, true);
    if (hyperlink[key].linkType === "external") {
      //外部链接
      cell.value = {
        text: cell.value,
        hyperlink: hyperlink[key].linkAddress,
        tooltip: hyperlink[key].linkTooltip,
      };
    } else {
      // 内部链接
      const linkArr = hyperlink[key].linkAddress.split("!");
      let hyper = "#\\" + linkArr[0] + "\\" + "!" + linkArr[1];
      cell.value = {
        text: cell.value,
        hyperlink: hyper,
        tooltip: hyperlink[key].linkTooltip,
      };
    }
  }
}

export { setHyperlink };
