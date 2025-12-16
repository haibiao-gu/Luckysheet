// 冻结视图
function setFrozen(frozen, worksheet) {
  //不存在冻结或取消冻结，则不执行后续代码
  if (!frozen || frozen.type === "cancel") return;
  //执行冻结操作代码
  let views = [];
  switch (frozen.type) {
    //冻结首行
    case "row":
      views = [{ state: "frozen", xSplit: 0, ySplit: 1 }];
      break;
    //冻结首列
    case "column":
      views = [{ state: "frozen", xSplit: 1, ySplit: 0 }];
      break;
    //冻结首行首列
    case "both":
      views = [{ state: "frozen", xSplit: 1, ySplit: 1 }];
      break;
    //冻结行至选区
    case "rangeRow":
      views = [
        { state: "frozen", xSplit: 0, ySplit: frozen.range.row_focus + 1 },
      ];
      break;
    //冻结列至选区
    case "rangeColumn":
      views = [
        { state: "frozen", xSplit: frozen.range.column_focus + 1, ySplit: 0 },
      ];
      break;
    //冻结至选区
    case "rangeBoth":
      views = [
        {
          state: "frozen",
          xSplit: frozen.range.column_focus + 1,
          ySplit: frozen.range.row_focus + 1,
        },
      ];
      break;
  }
  worksheet.views = views;
}

export { setFrozen };
