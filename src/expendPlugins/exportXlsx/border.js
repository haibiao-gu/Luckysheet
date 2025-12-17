// 重新设置边框
function setBorder(luckyBorderInfo, worksheet, mergeConfig = {}) {
  if (!Array.isArray(luckyBorderInfo)) return;
  // console.log("开始设置边框", luckyBorderInfo);

  // 构建合并单元格映射，用于快速查找某个单元格是否属于合并区域
  const mergedCellsMap = {};
  if (mergeConfig) {
    for (const key in mergeConfig) {
      const merge = mergeConfig[key];
      const startRow = merge.r;
      const endRow = merge.r + merge.rs - 1;
      const startCol = merge.c;
      const endCol = merge.c + merge.cs - 1;

      // 标记合并区域的所有单元格
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
          mergedCellsMap[`${r}-${c}`] = {
            isMainCell: r === startRow && c === startCol, // 主单元格
            mainCellPos: { r: startRow, c: startCol }, // 主单元格位置
            isMergedCell: true,
            range: { startRow, endRow, startCol, endCol },
          };
        }
      }
    }
  }

  luckyBorderInfo.forEach(elem => {
    // 现在只兼容到borderType 为range的情况
    if (elem.rangeType === "range") {
      let rang = elem.range[0];
      let row = rang.row;
      let column = rang.column;

      // 特殊处理区域边框类型
      if (
        elem.borderType === "border-inside" ||
        elem.borderType === "border-outside" ||
        elem.borderType === "border-horizontal" ||
        elem.borderType === "border-vertical"
      ) {
        let border = borderConvert(elem.borderType, elem.style, elem.color);

        // 根据不同的边框类型进行处理
        if (elem.borderType === "border-inside") {
          // 内边框：只绘制单元格之间的边框
          for (let i = row[0] + 1; i <= row[1] + 1; i++) {
            for (let y = column[0] + 1; y <= column[1] + 1; y++) {
              let borderSettings = {};
              const cell = worksheet.getCell(i, y);

              // 右边框（如果不是最右列）
              if (y < column[1] + 1) {
                borderSettings.right = border.inside;
              }
              // 下边框（如果不是最下行）
              if (i < row[1] + 1) {
                borderSettings.bottom = border.inside;
              }
              // 左边框（如果不是最左列，用于已有左边单元格的右边框）
              if (y > column[0] + 1) {
                borderSettings.left = border.inside;
              }
              // 上边框（如果不是最上行，用于已有上边单元格的下边框）
              if (i > row[0] + 1) {
                borderSettings.top = border.inside;
              }

              if (Object.keys(borderSettings).length > 0) {
                cell.border = Object.assign(
                  {},
                  cell.border || {},
                  borderSettings
                );
              }
            }
          }
        } else if (elem.borderType === "border-outside") {
          // 外边框：只绘制区域边缘的边框
          for (let i = row[0] + 1; i <= row[1] + 1; i++) {
            for (let y = column[0] + 1; y <= column[1] + 1; y++) {
              let borderSettings = {};
              const cell = worksheet.getCell(i, y);

              // 上边框（如果是第一行）
              if (i === row[0] + 1) {
                borderSettings.top = border.outside;
              }
              // 下边框（如果是最后一行）
              if (i === row[1] + 1) {
                borderSettings.bottom = border.outside;
              }
              // 左边框（如果是第一列）
              if (y === column[0] + 1) {
                borderSettings.left = border.outside;
              }
              // 右边框（如果是最后一列）
              if (y === column[1] + 1) {
                borderSettings.right = border.outside;
              }

              if (Object.keys(borderSettings).length > 0) {
                cell.border = Object.assign(
                  {},
                  cell.border || {},
                  borderSettings
                );
              }
            }
          }
        } else if (elem.borderType === "border-horizontal") {
          // 水平边框：只绘制水平线
          for (let i = row[0] + 1; i <= row[1] + 1; i++) {
            for (let y = column[0] + 1; y <= column[1] + 1; y++) {
              const cell = worksheet.getCell(i, y);
              let borderSettings = {};

              // 下边框（如果不是最下行）
              if (i < row[1] + 1) {
                borderSettings.bottom = border.horizontal;
              }

              if (Object.keys(borderSettings).length > 0) {
                cell.border = Object.assign(
                  {},
                  cell.border || {},
                  borderSettings
                );
              }
            }
          }
        } else if (elem.borderType === "border-vertical") {
          // 垂直边框：只绘制垂直线
          for (let i = row[0] + 1; i <= row[1] + 1; i++) {
            for (let y = column[0] + 1; y <= column[1] + 1; y++) {
              const cell = worksheet.getCell(i, y);
              let borderSettings = {};

              // 右边框（如果不是最右列）
              if (y < column[1] + 1) {
                borderSettings.right = border.vertical;
              }

              if (Object.keys(borderSettings).length > 0) {
                cell.border = Object.assign(
                  {},
                  cell.border || {},
                  borderSettings
                );
              }
            }
          }
        }
      } else {
        let border = borderConvert(elem.borderType, elem.style, elem.color);
        for (let i = row[0] + 1; i < row[1] + 2; i++) {
          for (let y = column[0] + 1; y < column[1] + 2; y++) {
            worksheet.getCell(i, y).border = border;
          }
        }
      }
    }
    if (elem.rangeType === "cell") {
      const { col_index, row_index } = elem.value;
      const borderData = Object.assign({}, elem.value);
      delete borderData.col_index;
      delete borderData.row_index;

      // 检查当前单元格是否是合并单元格的一部分
      const cellKey = `${row_index}-${col_index}`;
      const mergedInfo = mergedCellsMap[cellKey];

      const cell = worksheet.getCell(row_index + 1, col_index + 1);
      let borderSettings = {};
      const border = addborderToCell(borderData);

      // 对于合并单元格，使用最边缘的边框格式
      if (mergedInfo) {
        const range = mergedInfo.range;

        // 上边框（如果在合并区域的最上方）
        if (row_index === range.startRow) {
          if (border.top) borderSettings.top = border.top;
        }
        // 下边框（如果在合并区域的最下方）
        if (row_index === range.endRow) {
          if (border.bottom) borderSettings.bottom = border.bottom;
        }
        // 左边框（如果在合并区域的最左方）
        if (col_index === range.startCol) {
          if (border.left) borderSettings.left = border.left;
        }
        // 右边框（如果在合并区域的最右方）
        if (col_index === range.endCol) {
          if (border.right) borderSettings.right = border.right;
        }

        if (Object.keys(borderSettings).length > 0) {
          cell.border = Object.assign({}, cell.border || {}, borderSettings);
        }
      } else {
        // 非合并单元格直接应用边框
        cell.border = border;
      }
    }
  });
}

// 边框转换
function borderConvert(borderType, style = 1, color = "#000") {
  // 对应luckysheet的config中borderinfo的的参数
  if (!borderType) {
    return {};
  }
  const luckyToExcel = {
    type: {
      "border-all": "all",
      "border-top": "top",
      "border-right": "right",
      "border-bottom": "bottom",
      "border-left": "left",
      "border-none": "none", // 添加无边框情况
      "border-inside": "inside", // 内边框
      "border-outside": "outside", // 外边框
      "border-horizontal": "horizontal", // 水平边框
      "border-vertical": "vertical", // 垂直边框
    },
    style: {
      0: "none",
      1: "thin",
      2: "hair",
      3: "dotted",
      4: "dashDot", // 'Dashed',
      5: "dashDot",
      6: "dashDotDot",
      7: "double",
      8: "medium",
      9: "mediumDashed",
      10: "mediumDashDot",
      11: "mediumDashDotDot",
      12: "slantDashDot",
      13: "thick",
    },
  };
  let template = {
    style: luckyToExcel.style[style],
    color: { argb: color.replace("#", "") },
  };
  let border = {};
  if (luckyToExcel.type[borderType] === "all") {
    border["top"] = template;
    border["right"] = template;
    border["bottom"] = template;
    border["left"] = template;
  } else if (luckyToExcel.type[borderType] === "none") {
    // 处理无边框情况，返回空边框或重置边框
    border = {
      top: { style: "none" },
      right: { style: "none" },
      bottom: { style: "none" },
      left: { style: "none" },
    };
  } else if (luckyToExcel.type[borderType] === "inside") {
    // 内边框 - 需要在调用处特殊处理，这里返回模板
    border = { inside: template };
  } else if (luckyToExcel.type[borderType] === "outside") {
    // 外边框 - 需要在调用处特殊处理，这里返回模板
    border = { outside: template };
  } else if (luckyToExcel.type[borderType] === "horizontal") {
    // 水平边框 - 需要在调用处特殊处理，这里返回模板
    border = { horizontal: template };
  } else if (luckyToExcel.type[borderType] === "vertical") {
    // 垂直边框 - 需要在调用处特殊处理，这里返回模板
    border = { vertical: template };
  } else {
    border[luckyToExcel.type[borderType]] = template;
  }
  return border;
}

// 向单元格添加边框
function addborderToCell(borders) {
  let border = {};
  const luckyExcel = {
    type: {
      l: "left",
      r: "right",
      b: "bottom",
      t: "top",
    },
    style: {
      0: "none",
      1: "thin",
      2: "hair",
      3: "dotted",
      4: "dashDot", // 'Dashed',
      5: "dashDot",
      6: "dashDotDot",
      7: "double",
      8: "medium",
      9: "mediumDashed",
      10: "mediumDashDot",
      11: "mediumDashDotDot",
      12: "slantDashDot",
      13: "thick",
    },
  };
  for (const bor in borders) {
    if (!borders[bor]) continue;
    if (borders[bor].color.indexOf("rgb") === -1) {
      border[luckyExcel.type[bor]] = {
        style: luckyExcel.style[borders[bor].style],
        color: { argb: borders[bor].color.replace("#", "") },
      };
    } else {
      border[luckyExcel.type[bor]] = {
        style: luckyExcel.style[borders[bor].style],
        color: { argb: borders[bor].color },
      };
    }
  }

  return border;
}

export { setBorder, borderConvert, addborderToCell };
