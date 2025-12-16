// 单元格背景填充色处理
function fillConvert(bg) {
  if (!bg) {
    return null;
  }
  bg = bg.indexOf("rgb") > -1 ? rgb2hex(bg) : bg;
  return {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: bg.replace("#", "") },
  };
}

// 转换颜色
function rgb2hex(rgb) {
  if (rgb.charAt(0) === "#") {
    return rgb;
  }

  let ds = rgb.split(/\D+/);
  let decimal = Number(ds[1]) * 65536 + Number(ds[2]) * 256 + Number(ds[3]);
  return "#" + zero_fill_hex(decimal, 6);

  function zero_fill_hex(num, digits) {
    let s = num.toString(16);
    while (s.length < digits) s = "0" + s;
    return s;
  }
}

// 字体转换处理
function fontConvert(
  ff = 0,
  fc = "#000000",
  bl = 0,
  it = 0,
  fs = 10,
  cl = 0,
  ul = 0
) {
  // luckysheet：ff(样式), fc(颜色), bl(粗体), it(斜体), fs(大小), cl(删除线), ul(下划线)
  const luckyToExcel = {
    0: "微软雅黑",
    1: "宋体（Song）",
    2: "黑体（ST Heiti）",
    3: "楷体（ST Kaiti）",
    4: "仿宋（ST FangSong）",
    5: "新宋体（ST Song）",
    6: "华文新魏",
    7: "华文行楷",
    8: "华文隶书",
    9: "Arial",
    10: "Times New Roman",
    11: "Tahoma ",
    12: "Verdana",
    num2bl: function (num) {
      return !(num === 0 || false);
    },
  };

  // 改进的字体颜色处理
  let fontColor = "#000000"; // 默认黑色
  if (fc) {
    // 处理RGB格式的颜色
    if (fc.indexOf("rgb") > -1) {
      fontColor = rgb2hex(fc);
    } else {
      fontColor = fc;
    }
  }

  return {
    name: ff,
    family: 1,
    size: fs,
    color: { argb: fontColor.replace("#", "") },
    bold: luckyToExcel.num2bl(bl),
    italic: luckyToExcel.num2bl(it),
    underline: luckyToExcel.num2bl(ul),
    strike: luckyToExcel.num2bl(cl),
  };
}

// 对齐转换
function alignmentConvert(
  vt = "default",
  ht = "default",
  tb = "default",
  tr = "default"
) {
  // luckysheet:vt(垂直), ht(水平), tb(换行), tr(旋转)
  const luckyToExcel = {
    vertical: {
      0: "middle",
      1: "top",
      2: "bottom",
      default: "top",
    },
    horizontal: {
      0: "center",
      1: "left",
      2: "right",
      default: "left",
    },
    wrapText: {
      0: false,
      1: false,
      2: true,
      default: false,
    },
    textRotation: {
      0: 0,
      1: 45,
      2: -45,
      3: "vertical",
      4: 90,
      5: -90,
      default: 0,
    },
  };

  return {
    vertical: luckyToExcel.vertical[vt],
    horizontal: luckyToExcel.horizontal[ht],
    wrapText: luckyToExcel.wrapText[tb],
    textRotation: luckyToExcel.textRotation[tr],
  };
}

export { fillConvert, fontConvert, alignmentConvert, rgb2hex };
