// 设置图片
function setImages(thesheet, worksheet, workbook) {
  // 安全地访问属性
  const images = thesheet.images;
  const config = thesheet.config;
  if (typeof images != "object" || !images) return;
  // console.log("开始设置图片", images);
  // 获取列宽和行高配置，如果没有则使用空对象
  const columnLen = (config && config.columnlen) || {};
  const rowLen = (config && config.rowlen) || {};

  // 将对象格式转换为数组格式
  const columnPositions = calculatePositionsFromLengths(columnLen, true);
  const rowPositions = calculatePositionsFromLengths(rowLen, false);

  for (let key in images) {
    // 检查图片数据完整性
    if (!images[key] || !images[key].src || !images[key].default) continue;

    // 通过 base64  将图像添加到工作簿
    const myBase64Image = images[key].src;
    //开始行 开始列 结束行 结束列
    const item = images[key];
    const imageId = workbook.addImage({
      base64: myBase64Image,
      extension: "png",
    });

    const col_st = getImagePosition(item.default.left, columnPositions);
    const row_st = getImagePosition(item.default.top, rowPositions);

    //模式1，图片左侧与luckysheet位置一样，像素比例保持不变，但是，右侧位置可能与原图所在单元格不一致
    worksheet.addImage(imageId, {
      tl: { col: col_st, row: row_st },
      ext: { width: item.default.width, height: item.default.height },
    });
  }
}

// 根据长度配置计算位置数组
function calculatePositionsFromLengths(lengthConfig, isColumn = true) {
  const positions = [0]; // 必须从0开始
  let currentPosition = 0;

  // 获取配置中的最大索引
  const indices = Object.keys(lengthConfig)
    .map(Number)
    .filter(n => !isNaN(n))
    .sort((a, b) => a - b);

  // 确定要计算的最大索引（至少计算到100，或根据配置中的最大值）
  const maxIndex = Math.max(indices.length > 0 ? Math.max(...indices) : 0, 100);

  // 逐个计算每个位置的累计值
  for (let i = 0; i <= maxIndex; i++) {
    const length = lengthConfig[i] || (isColumn ? 72 : 19);
    currentPosition += length;
    positions.push(currentPosition);
  }

  return positions;
}

// 获取图片在单元格的位置
function getImagePosition(num, arr) {
  // 添加数组验证
  if (!Array.isArray(arr) || arr.length === 0) {
    return 0;
  }

  // 数字验证
  if (typeof num !== "number" || isNaN(num) || num < 0) {
    return 0;
  }

  // 如果坐标超过最大位置，则放在最后
  if (num >= arr[arr.length - 1]) {
    return arr.length - 1;
  }

  // 找到对应区间
  for (let i = 0; i < arr.length - 1; i++) {
    if (num >= arr[i] && num < arr[i + 1]) {
      // 计算在该区间内的相对位置
      const segmentStart = arr[i];
      const segmentEnd = arr[i + 1];
      const segmentLength = segmentEnd - segmentStart;

      if (segmentLength === 0) {
        return i;
      }

      // 返回浮点数位置，精确到小数点后几位
      return i + (num - segmentStart) / segmentLength;
    }
  }

  return 0;
}

export { setImages, calculatePositionsFromLengths, getImagePosition };
