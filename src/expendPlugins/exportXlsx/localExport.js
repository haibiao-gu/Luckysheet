// import { luckysheet } from "../../core";
// import Excel from "exceljs";
// import FileSaver from "file-saver";
//
// function localExport(order, success) {
//   const sheetInfo = luckysheet.toJson();
//   console.log("开始导出", order, sheetInfo);
//   // 获取需要导出的工作表
//   const exportSheet =
//     order === "all" ? sheetInfo.data : [sheetInfo.data[order]];
//
//   const workbook = new Excel.Workbook();
//   // 写入工作薄
//   exportSheet.forEach(sheet => {
//     const worksheet = workbook.addWorksheet(sheet.name);
//     // 设置工作表样式
//     setStyleAndValue(sheet.data, worksheet);
//     setMerge((sheet.config && sheet.config.merge) || {}, worksheet);
//     setBorder((sheet.config && sheet.config.borderInfo) || {}, worksheet);
//     setImages(sheet, worksheet, workbook);
//     setHyperlink(sheet.hyperlink, worksheet);
//     setFrozen(sheet.frozen, worksheet);
//     setConditions(sheet.luckysheet_conditionformat_save, worksheet);
//   });
//   // 写入 buffer
//   workbook.xlsx.writeBuffer().then(data => {
//     const blob = new Blob([data], {
//       type: "application/vnd.ms-excel;charset=utf-8",
//     });
//     FileSaver.saveAs(blob, `${sheetInfo.title}.xlsx`);
//   });
//   success && success();
// }
//
// // 设置单元格样式和值
// function setStyleAndValue(cellArr, worksheet) {
//   if (!Array.isArray(cellArr)) return;
//
//   cellArr.forEach(function (row, rowid) {
//     const dbrow = worksheet.getRow(rowid + 1);
//     //设置单元格行高,默认乘以0.8倍
//     dbrow.height = luckysheet.getRowHeight([rowid])[rowid] * 0.8;
//     row.every(function (cell, columnid) {
//       if (!cell) return true;
//       if (rowid === 0) {
//         const dobCol = worksheet.getColumn(columnid + 1);
//         //设置单元格列宽除以8
//         dobCol.width = luckysheet.getColumnWidth([columnid])[columnid] / 8;
//       }
//       let fill = fillConvert(cell.bg);
//       let font = fontConvert(
//         cell.ff || "Times New Roman",
//         cell.fc,
//         cell.bl,
//         cell.it,
//         cell.fs,
//         cell.cl,
//         cell.ul
//       );
//       let alignment = alignmentConvert(cell.vt, cell.ht, cell.tb, cell.tr);
//       let value;
//
//       let v = "";
//       if (cell.ct && cell.ct.t === "inlineStr") {
//         const s = cell.ct.s;
//         s.forEach(val => {
//           v += val.v;
//         });
//       } else {
//         //导出后取显示值
//         v = cell.m;
//       }
//       if (cell.f) {
//         value = { formula: cell.f, result: v };
//       } else {
//         value = v;
//       }
//       let target = worksheet.getCell(rowid + 1, columnid + 1);
//       //添加批注
//       if (cell.ps) {
//         let ps = cell.ps;
//         target.note = ps.value;
//       }
//       //单元格填充
//       target.fill = fill;
//       //单元格字体
//       target.font = font;
//       target.alignment = alignment;
//       target.value = value;
//       return true;
//     });
//   });
// }
//
// // 单元格背景填充色处理
// function fillConvert(bg) {
//   if (!bg) {
//     return null;
//   }
//   bg = bg.indexOf("rgb") > -1 ? rgb2hex(bg) : bg;
//   return {
//     type: "pattern",
//     pattern: "solid",
//     fgColor: { argb: bg.replace("#", "") },
//   };
// }
//
// // 转换颜色
// function rgb2hex(rgb) {
//   if (rgb.charAt(0) === "#") {
//     return rgb;
//   }
//
//   let ds = rgb.split(/\D+/);
//   let decimal = Number(ds[1]) * 65536 + Number(ds[2]) * 256 + Number(ds[3]);
//   return "#" + zero_fill_hex(decimal, 6);
//
//   function zero_fill_hex(num, digits) {
//     let s = num.toString(16);
//     while (s.length < digits) s = "0" + s;
//     return s;
//   }
// }
// // 字体转换处理
// function fontConvert(
//   ff = 0,
//   fc = "#000000",
//   bl = 0,
//   it = 0,
//   fs = 10,
//   cl = 0,
//   ul = 0
// ) {
//   // luckysheet：ff(样式), fc(颜色), bl(粗体), it(斜体), fs(大小), cl(删除线), ul(下划线)
//   const luckyToExcel = {
//     0: "微软雅黑",
//     1: "宋体（Song）",
//     2: "黑体（ST Heiti）",
//     3: "楷体（ST Kaiti）",
//     4: "仿宋（ST FangSong）",
//     5: "新宋体（ST Song）",
//     6: "华文新魏",
//     7: "华文行楷",
//     8: "华文隶书",
//     9: "Arial",
//     10: "Times New Roman",
//     11: "Tahoma ",
//     12: "Verdana",
//     num2bl: function (num) {
//       return !(num === 0 || false);
//     },
//   };
//
//   // 改进的字体颜色处理
//   let fontColor = "#000000"; // 默认黑色
//   if (fc) {
//     // 处理RGB格式的颜色
//     if (fc.indexOf("rgb") > -1) {
//       fontColor = rgb2hex(fc);
//     } else {
//       fontColor = fc;
//     }
//   }
//
//   return {
//     name: ff,
//     family: 1,
//     size: fs,
//     color: { argb: fontColor.replace("#", "") },
//     bold: luckyToExcel.num2bl(bl),
//     italic: luckyToExcel.num2bl(it),
//     underline: luckyToExcel.num2bl(ul),
//     strike: luckyToExcel.num2bl(cl),
//   };
// }
// // 对齐转换
// function alignmentConvert(
//   vt = "default",
//   ht = "default",
//   tb = "default",
//   tr = "default"
// ) {
//   // luckysheet:vt(垂直), ht(水平), tb(换行), tr(旋转)
//   const luckyToExcel = {
//     vertical: {
//       0: "middle",
//       1: "top",
//       2: "bottom",
//       default: "top",
//     },
//     horizontal: {
//       0: "center",
//       1: "left",
//       2: "right",
//       default: "left",
//     },
//     wrapText: {
//       0: false,
//       1: false,
//       2: true,
//       default: false,
//     },
//     textRotation: {
//       0: 0,
//       1: 45,
//       2: -45,
//       3: "vertical",
//       4: 90,
//       5: -90,
//       default: 0,
//     },
//   };
//
//   return {
//     vertical: luckyToExcel.vertical[vt],
//     horizontal: luckyToExcel.horizontal[ht],
//     wrapText: luckyToExcel.wrapText[tb],
//     textRotation: luckyToExcel.textRotation[tr],
//   };
// }
//
// // 设置合并数据
// function setMerge(luckyMerge = {}, worksheet) {
//   const mergearr = Object.values(luckyMerge);
//   mergearr.forEach(function (elem) {
//     // elem格式：{r: 0, c: 0, rs: 1, cs: 2}
//     // 按开始行，开始列，结束行，结束列合并（相当于 K10:M12）
//     worksheet.mergeCells(
//       elem.r + 1,
//       elem.c + 1,
//       elem.r + elem.rs,
//       elem.c + elem.cs
//     );
//   });
// }
//
// // 重新设置边框
// function setBorder(luckyBorderInfo, worksheet) {
//   if (!Array.isArray(luckyBorderInfo)) return;
//   luckyBorderInfo.forEach(function (elem) {
//     // 现在只兼容到borderType 为range的情况
//     if (elem.rangeType === "range") {
//       let rang = elem.range[0];
//       let row = rang.row;
//       let column = rang.column;
//
//       // 特殊处理区域边框类型
//       if (
//         elem.borderType === "border-inside" ||
//         elem.borderType === "border-outside" ||
//         elem.borderType === "border-horizontal" ||
//         elem.borderType === "border-vertical"
//       ) {
//         let border = borderConvert(elem.borderType, elem.style, elem.color);
//
//         // 根据不同的边框类型进行处理
//         if (elem.borderType === "border-inside") {
//           // 内边框：只绘制单元格之间的边框
//           for (let i = row[0] + 1; i <= row[1] + 1; i++) {
//             for (let y = column[0] + 1; y <= column[1] + 1; y++) {
//               let borderSettings = {};
//               const cell = worksheet.getCell(i, y);
//
//               // 右边框（如果不是最右列）
//               if (y < column[1] + 1) {
//                 borderSettings.right = border.inside;
//               }
//               // 下边框（如果不是最下行）
//               if (i < row[1] + 1) {
//                 borderSettings.bottom = border.inside;
//               }
//               // 左边框（如果不是最左列，用于已有左边单元格的右边框）
//               if (y > column[0] + 1) {
//                 borderSettings.left = border.inside;
//               }
//               // 上边框（如果不是最上行，用于已有上边单元格的下边框）
//               if (i > row[0] + 1) {
//                 borderSettings.top = border.inside;
//               }
//
//               if (Object.keys(borderSettings).length > 0) {
//                 cell.border = Object.assign(
//                   {},
//                   cell.border || {},
//                   borderSettings
//                 );
//               }
//             }
//           }
//         } else if (elem.borderType === "border-outside") {
//           // 外边框：只绘制区域边缘的边框
//           for (let i = row[0] + 1; i <= row[1] + 1; i++) {
//             for (let y = column[0] + 1; y <= column[1] + 1; y++) {
//               let borderSettings = {};
//               const cell = worksheet.getCell(i, y);
//
//               // 上边框（如果是第一行）
//               if (i === row[0] + 1) {
//                 borderSettings.top = border.outside;
//               }
//               // 下边框（如果是最后一行）
//               if (i === row[1] + 1) {
//                 borderSettings.bottom = border.outside;
//               }
//               // 左边框（如果是第一列）
//               if (y === column[0] + 1) {
//                 borderSettings.left = border.outside;
//               }
//               // 右边框（如果是最后一列）
//               if (y === column[1] + 1) {
//                 borderSettings.right = border.outside;
//               }
//
//               if (Object.keys(borderSettings).length > 0) {
//                 cell.border = Object.assign(
//                   {},
//                   cell.border || {},
//                   borderSettings
//                 );
//               }
//             }
//           }
//         } else if (elem.borderType === "border-horizontal") {
//           // 水平边框：只绘制水平线
//           for (let i = row[0] + 1; i <= row[1] + 1; i++) {
//             for (let y = column[0] + 1; y <= column[1] + 1; y++) {
//               const cell = worksheet.getCell(i, y);
//               let borderSettings = {};
//
//               // 下边框（如果不是最下行）
//               if (i < row[1] + 1) {
//                 borderSettings.bottom = border.horizontal;
//               }
//
//               if (Object.keys(borderSettings).length > 0) {
//                 cell.border = Object.assign(
//                   {},
//                   cell.border || {},
//                   borderSettings
//                 );
//               }
//             }
//           }
//         } else if (elem.borderType === "border-vertical") {
//           // 垂直边框：只绘制垂直线
//           for (let i = row[0] + 1; i <= row[1] + 1; i++) {
//             for (let y = column[0] + 1; y <= column[1] + 1; y++) {
//               const cell = worksheet.getCell(i, y);
//               let borderSettings = {};
//
//               // 右边框（如果不是最右列）
//               if (y < column[1] + 1) {
//                 borderSettings.right = border.vertical;
//               }
//
//               if (Object.keys(borderSettings).length > 0) {
//                 cell.border = Object.assign(
//                   {},
//                   cell.border || {},
//                   borderSettings
//                 );
//               }
//             }
//           }
//         }
//       } else {
//         let border = borderConvert(elem.borderType, elem.style, elem.color);
//         for (let i = row[0] + 1; i < row[1] + 2; i++) {
//           for (let y = column[0] + 1; y < column[1] + 2; y++) {
//             worksheet.getCell(i, y).border = border;
//           }
//         }
//       }
//     }
//     if (elem.rangeType === "cell") {
//       // col_index: 2
//       // row_index: 1
//       // b: {
//       //   color: '#d0d4e3'
//       //   style: 1
//       // }
//       const { col_index, row_index } = elem.value;
//       const borderData = Object.assign({}, elem.value);
//       delete borderData.col_index;
//       delete borderData.row_index;
//
//       worksheet.getCell(row_index + 1, col_index + 1).border =
//         addborderToCell(borderData);
//     }
//   });
// }
//
// // 边框转换
// function borderConvert(borderType, style = 1, color = "#000") {
//   // 对应luckysheet的config中borderinfo的的参数
//   if (!borderType) {
//     return {};
//   }
//   const luckyToExcel = {
//     type: {
//       "border-all": "all",
//       "border-top": "top",
//       "border-right": "right",
//       "border-bottom": "bottom",
//       "border-left": "left",
//       "border-none": "none", // 添加无边框情况
//       "border-inside": "inside", // 内边框
//       "border-outside": "outside", // 外边框
//       "border-horizontal": "horizontal", // 水平边框
//       "border-vertical": "vertical", // 垂直边框
//     },
//     style: {
//       0: "none",
//       1: "thin",
//       2: "hair",
//       3: "dotted",
//       4: "dashDot", // 'Dashed',
//       5: "dashDot",
//       6: "dashDotDot",
//       7: "double",
//       8: "medium",
//       9: "mediumDashed",
//       10: "mediumDashDot",
//       11: "mediumDashDotDot",
//       12: "slantDashDot",
//       13: "thick",
//     },
//   };
//   let template = {
//     style: luckyToExcel.style[style],
//     color: { argb: color.replace("#", "") },
//   };
//   let border = {};
//   if (luckyToExcel.type[borderType] === "all") {
//     border["top"] = template;
//     border["right"] = template;
//     border["bottom"] = template;
//     border["left"] = template;
//   } else if (luckyToExcel.type[borderType] === "none") {
//     // 处理无边框情况，返回空边框或重置边框
//     border = {
//       top: { style: "none" },
//       right: { style: "none" },
//       bottom: { style: "none" },
//       left: { style: "none" },
//     };
//   } else if (luckyToExcel.type[borderType] === "inside") {
//     // 内边框 - 需要在调用处特殊处理，这里返回模板
//     border = { inside: template };
//   } else if (luckyToExcel.type[borderType] === "outside") {
//     // 外边框 - 需要在调用处特殊处理，这里返回模板
//     border = { outside: template };
//   } else if (luckyToExcel.type[borderType] === "horizontal") {
//     // 水平边框 - 需要在调用处特殊处理，这里返回模板
//     border = { horizontal: template };
//   } else if (luckyToExcel.type[borderType] === "vertical") {
//     // 垂直边框 - 需要在调用处特殊处理，这里返回模板
//     border = { vertical: template };
//   } else {
//     border[luckyToExcel.type[borderType]] = template;
//   }
//   return border;
// }
//
// // 向单元格添加边框
// function addborderToCell(borders) {
//   let border = {};
//   const luckyExcel = {
//     type: {
//       l: "left",
//       r: "right",
//       b: "bottom",
//       t: "top",
//     },
//     style: {
//       0: "none",
//       1: "thin",
//       2: "hair",
//       3: "dotted",
//       4: "dashDot", // 'Dashed',
//       5: "dashDot",
//       6: "dashDotDot",
//       7: "double",
//       8: "medium",
//       9: "mediumDashed",
//       10: "mediumDashDot",
//       11: "mediumDashDotDot",
//       12: "slantDashDot",
//       13: "thick",
//     },
//   };
//   for (const bor in borders) {
//     if (!borders[bor]) continue;
//     if (borders[bor].color.indexOf("rgb") === -1) {
//       border[luckyExcel.type[bor]] = {
//         style: luckyExcel.style[borders[bor].style],
//         color: { argb: borders[bor].color.replace("#", "") },
//       };
//     } else {
//       border[luckyExcel.type[bor]] = {
//         style: luckyExcel.style[borders[bor].style],
//         color: { argb: borders[bor].color },
//       };
//     }
//   }
//
//   return border;
// }
//
// // 设置图片
// function setImages(thesheet, worksheet, workbook) {
//   // 安全地访问属性
//   const images = thesheet.images;
//   const config = thesheet.config;
//   if (typeof images != "object" || !images) return;
//   // 获取列宽和行高配置，如果没有则使用空对象
//   const columnLen = (config && config.columnlen) || {};
//   const rowLen = (config && config.rowlen) || {};
//
//   // 将对象格式转换为数组格式
//   const columnPositions = calculatePositionsFromLengths(columnLen, true);
//   const rowPositions = calculatePositionsFromLengths(rowLen, false);
//
//   for (let key in images) {
//     // 检查图片数据完整性
//     if (!images[key] || !images[key].src || !images[key].default) continue;
//
//     // 通过 base64  将图像添加到工作簿
//     const myBase64Image = images[key].src;
//     //开始行 开始列 结束行 结束列
//     const item = images[key];
//     const imageId = workbook.addImage({
//       base64: myBase64Image,
//       extension: "png",
//     });
//
//     const col_st = getImagePosition(item.default.left, columnPositions);
//     const row_st = getImagePosition(item.default.top, rowPositions);
//
//     //模式1，图片左侧与luckysheet位置一样，像素比例保持不变，但是，右侧位置可能与原图所在单元格不一致
//     worksheet.addImage(imageId, {
//       tl: { col: col_st, row: row_st },
//       ext: { width: item.default.width, height: item.default.height },
//     });
//   }
// }
// // 根据长度配置计算位置数组
// function calculatePositionsFromLengths(lengthConfig, isColumn = true) {
//   const positions = [0]; // 必须从0开始
//   let currentPosition = 0;
//
//   // 获取配置中的最大索引
//   const indices = Object.keys(lengthConfig)
//     .map(Number)
//     .filter(n => !isNaN(n))
//     .sort((a, b) => a - b);
//
//   // 确定要计算的最大索引（至少计算到100，或根据配置中的最大值）
//   const maxIndex = Math.max(indices.length > 0 ? Math.max(...indices) : 0, 100);
//
//   // 逐个计算每个位置的累计值
//   for (let i = 0; i <= maxIndex; i++) {
//     const length = lengthConfig[i] || (isColumn ? 72 : 19);
//     currentPosition += length;
//     positions.push(currentPosition);
//   }
//
//   return positions;
// }
//
// // 获取图片在单元格的位置
// function getImagePosition(num, arr) {
//   // 添加数组验证
//   if (!Array.isArray(arr) || arr.length === 0) {
//     return 0;
//   }
//
//   // 数字验证
//   if (typeof num !== "number" || isNaN(num) || num < 0) {
//     return 0;
//   }
//
//   // 如果坐标超过最大位置，则放在最后
//   if (num >= arr[arr.length - 1]) {
//     return arr.length - 1;
//   }
//
//   // 找到对应区间
//   for (let i = 0; i < arr.length - 1; i++) {
//     if (num >= arr[i] && num < arr[i + 1]) {
//       // 计算在该区间内的相对位置
//       const segmentStart = arr[i];
//       const segmentEnd = arr[i + 1];
//       const segmentLength = segmentEnd - segmentStart;
//
//       if (segmentLength === 0) {
//         return i;
//       }
//
//       // 返回浮点数位置，精确到小数点后几位
//       return i + (num - segmentStart) / segmentLength;
//     }
//   }
//
//   return 0;
// }
//
// // 设置超链接
// function setHyperlink(hyperlink, worksheet) {
//   if (!hyperlink) return;
//   for (const key in hyperlink) {
//     const row_col = key.split("_");
//     let cell = worksheet.getCell(
//       Number(row_col[0]) + 1,
//       Number(row_col[1]) + 1
//     );
//     let font = cell.style.font;
//     //设置导出后超链接的样式
//     cell.font = fontConvert(font.name, "#0000ff", 0, 0, font.size, 0, true);
//     if (hyperlink[key].linkType === "external") {
//       //外部链接
//       cell.value = {
//         text: cell.value,
//         hyperlink: hyperlink[key].linkAddress,
//         tooltip: hyperlink[key].linkTooltip,
//       };
//     } else {
//       // 内部链接
//       const linkArr = hyperlink[key].linkAddress.split("!");
//       let hyper = "#\\" + linkArr[0] + "\\" + "!" + linkArr[1];
//       cell.value = {
//         text: cell.value,
//         hyperlink: hyper,
//         tooltip: hyperlink[key].linkTooltip,
//       };
//     }
//   }
// }
//
// // 冻结视图
// function setFrozen(frozen, worksheet) {
//   //不存在冻结或取消冻结，则不执行后续代码
//   if (!frozen || frozen.type === "cancel") return;
//   //执行冻结操作代码
//   let views = [];
//   switch (frozen.type) {
//     //冻结首行
//     case "row":
//       views = [{ state: "frozen", xSplit: 0, ySplit: 1 }];
//       break;
//     //冻结首列
//     case "column":
//       views = [{ state: "frozen", xSplit: 1, ySplit: 0 }];
//       break;
//     //冻结首行首列
//     case "both":
//       views = [{ state: "frozen", xSplit: 1, ySplit: 1 }];
//       break;
//     //冻结行至选区
//     case "rangeRow":
//       views = [
//         { state: "frozen", xSplit: 0, ySplit: frozen.range.row_focus + 1 },
//       ];
//       break;
//     //冻结列至选区
//     case "rangeColumn":
//       views = [
//         { state: "frozen", xSplit: frozen.range.column_focus + 1, ySplit: 0 },
//       ];
//       break;
//     //冻结至选区
//     case "rangeBoth":
//       views = [
//         {
//           state: "frozen",
//           xSplit: frozen.range.column_focus + 1,
//           ySplit: frozen.range.row_focus + 1,
//         },
//       ];
//       break;
//   }
//   worksheet.views = views;
// }
//
// // 条件格式设置
// function setConditions(conditions, worksheet) {
//   //条件格式不存在，则不执行后续代码
//   if (conditions === undefined) return;
//
//   //循环遍历规则列表
//   conditions.forEach(item => {
//     let ruleObj = {
//       ref: createCellRange(item.cellrange[0].row, item.cellrange[0].column),
//       rules: [],
//     };
//     //lucksheet对应的为----突出显示单元格规则和项目选区规则
//     if (item.type === "default") {
//       //excel中type为cellIs的条件下
//       if (
//         item.conditionName === "equal" ||
//         "greaterThan" ||
//         "lessThan" ||
//         "betweenness"
//       ) {
//         ruleObj.rules = setDefaultRules({
//           type: "cellIs",
//           operator:
//             item.conditionName === "betweenness"
//               ? "between"
//               : item.conditionName,
//           condvalue: item.conditionValue,
//           colorArr: [item.format.cellColor, item.format.textColor],
//         });
//         worksheet.addConditionalFormatting(ruleObj);
//       }
//       //excel中type为containsText的条件下
//       if (item.conditionName === "textContains") {
//         ruleObj.rules = [
//           {
//             type: "containsText",
//             operator: "containsText", //表示如果单元格值包含在text 字段中指定的值，则应用格式
//             text: item.conditionValue[0],
//             style: setStyle([item.format.cellColor, item.format.textColor]),
//           },
//         ];
//         worksheet.addConditionalFormatting(ruleObj);
//       }
//       //发生日期--时间段
//       if (item.conditionName === "occurrenceDate") {
//         ruleObj.rules = [
//           {
//             type: "timePeriod",
//             timePeriod: "today", //表示如果单元格值包含在text 字段中指定的值，则应用格式
//             style: setStyle([item.format.cellColor, item.format.textColor]),
//           },
//         ];
//         worksheet.addConditionalFormatting(ruleObj);
//       }
//       //项目选区规则--top10前多少项的操作
//       if (item.conditionName === "top10" || "top10%" || "last10" || "last10%") {
//         ruleObj.rules = [
//           {
//             type: "top10",
//             rank: item.conditionValue[0], //指定格式中包含多少个顶部（或底部）值
//             percent: !(item.conditionName === "top10" || "last10"),
//             bottom: !(item.conditionName === "top10" || "top10%"),
//             style: setStyle([item.format.cellColor, item.format.textColor]),
//           },
//         ];
//         worksheet.addConditionalFormatting(ruleObj);
//       }
//       //项目选区规则--高于/低于平均值的操作
//       if (item.conditionName === "AboveAverage" || "SubAverage") {
//         ruleObj.rules = [
//           {
//             type: "aboveAverage",
//             aboveAverage: item.conditionName === "AboveAverage",
//             style: setStyle([item.format.cellColor, item.format.textColor]),
//           },
//         ];
//         worksheet.addConditionalFormatting(ruleObj);
//       }
//       return;
//     }
//
//     //数据条
//     if (item.type === "dataBar") {
//       ruleObj.rules = [
//         {
//           type: "dataBar",
//           style: {},
//         },
//       ];
//       worksheet.addConditionalFormatting(ruleObj);
//       return;
//     }
//     //色阶
//     if (item.type === "colorGradation") {
//       ruleObj.rules = [
//         {
//           type: "colorScale",
//           color: item.format,
//           style: {},
//         },
//       ];
//       worksheet.addConditionalFormatting(ruleObj);
//       return;
//     }
//     //图标集
//     if (item.type === "icons") {
//       ruleObj.rules = [
//         {
//           type: "iconSet",
//           iconSet: item.format.len,
//         },
//       ];
//       worksheet.addConditionalFormatting(ruleObj);
//     }
//   });
// }
//
// // 创建单元格范围
// function createCellRange(rowArr, colArr) {
//   const startCell = createCellPos(colArr[0]) + (rowArr[0] + 1);
//   const endCell = createCellPos(colArr[1]) + (rowArr[1] + 1);
//
//   return startCell + ":" + endCell;
// }
//
// // 创建单元格所在列的列的字母
// function createCellPos(n) {
//   let ordA = "A".charCodeAt(0);
//
//   let ordZ = "Z".charCodeAt(0);
//   let len = ordZ - ordA + 1;
//   let s = "";
//   while (n >= 0) {
//     s = String.fromCharCode((n % len) + ordA) + s;
//
//     n = Math.floor(n / len) - 1;
//   }
//   return s;
// }
//
// function setDefaultRules(obj) {
//   return [
//     {
//       type: obj.type,
//       operator: obj.operator,
//       formulae: obj.condvalue,
//       style: setStyle(obj.colorArr),
//     },
//   ];
// }
//
// function setStyle(colorArr) {
//   return {
//     fill: {
//       type: "pattern",
//       pattern: "solid",
//       bgColor: { argb: colorArr[0].replace("#", "") },
//     },
//     font: { color: { argb: colorArr[1].replace("#", "") } },
//   };
// }
//
// export { localExport };
