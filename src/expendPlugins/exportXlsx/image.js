// 设置图片
async function setImages(sheet, worksheet, workbook) {
  // 安全地访问属性
  const images = sheet.images;
  const config = sheet.config;
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
    const item = images[key];
    if (!item || !item.src || !item.default) continue;
    // console.log("item", item);

    // 网络图片转base64
    if (item.src.startsWith("http")) {
      item.src = await convertImageUrlToBase64(item.src);
    }

    //开始行 开始列 结束行 结束列
    const imageId = workbook.addImage({
      base64: item.src,
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
  if (num == null) {
    return 0;
  }

  // 尝试将 num 转换为数字
  num = Number(num);

  if (isNaN(num) || num < 0) {
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

async function convertImageUrlToBase64(imageUrl) {
  return new Promise((resolve, reject) => {
    // 1. 发起请求获取图片Blob数据
    fetch(imageUrl)
      .then(async response => {
        // 检查请求是否成功
        if (!response.ok) {
          reject(
            new Error(`请求图片失败：${response.status} ${response.statusText}`)
          );
          return;
        }
        // 将响应转换为Blob格式
        const blob = await response.blob();

        // 2. 使用FileReader转换Blob为Base64
        const reader = new FileReader();
        // 读取成功的回调
        reader.onload = () => {
          // result即为Base64编码字符串
          const base64Str = reader.result;
          resolve(base64Str);
        };
        // 读取失败的回调
        reader.onerror = error => {
          reject(new Error(`转换Base64失败：${error}`));
        };
        // 开始读取Blob数据为DataURL（Base64格式）
        reader.readAsDataURL(blob);
      })
      .catch(error => {
        // 捕获网络请求错误（如跨域、URL无效）
        reject(new Error(`网络请求错误：${error.message}`));
      });
  });
}

export { setImages, calculatePositionsFromLengths, getImagePosition };
