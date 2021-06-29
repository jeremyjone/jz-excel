const columeHeader_word = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
  "Q",
  "R",
  "S",
  "T",
  "U",
  "V",
  "W",
  "X",
  "Y",
  "Z"
];

const columeHeader_word_index = {
  A: 0,
  B: 1,
  C: 2,
  D: 3,
  E: 4,
  F: 5,
  G: 6,
  H: 7,
  I: 8,
  J: 9,
  K: 10,
  L: 11,
  M: 12,
  N: 13,
  O: 14,
  P: 15,
  Q: 16,
  R: 17,
  S: 18,
  T: 19,
  U: 20,
  V: 21,
  W: 22,
  X: 23,
  Y: 24,
  Z: 25
};

const indexedColors = {
  "0": "00000000",
  "1": "00FFFFFF",
  "2": "00FF0000",
  "3": "0000FF00",
  "4": "000000FF",
  "5": "00FFFF00",
  "6": "00FF00FF",
  "7": "0000FFFF",
  "8": "00000000",
  "9": "00FFFFFF",
  "10": "00FF0000",
  "11": "0000FF00",
  "12": "000000FF",
  "13": "00FFFF00",
  "14": "00FF00FF",
  "15": "0000FFFF",
  "16": "00800000",
  "17": "00008000",
  "18": "00000080",
  "19": "00808000",
  "20": "00800080",
  "21": "00008080",
  "22": "00C0C0C0",
  "23": "00808080",
  "24": "009999FF",
  "25": "00993366",
  "26": "00FFFFCC",
  "27": "00CCFFFF",
  "28": "00660066",
  "29": "00FF8080",
  "30": "000066CC",
  "31": "00CCCCFF",
  "32": "00000080",
  "33": "00FF00FF",
  "34": "00FFFF00",
  "35": "0000FFFF",
  "36": "00800080",
  "37": "00800000",
  "38": "00008080",
  "39": "000000FF",
  "40": "0000CCFF",
  "41": "00CCFFFF",
  "42": "00CCFFCC",
  "43": "00FFFF99",
  "44": "0099CCFF",
  "45": "00FF99CC",
  "46": "00CC99FF",
  "47": "00FFCC99",
  "48": "003366FF",
  "49": "0033CCCC",
  "50": "0099CC00",
  "51": "00FFCC00",
  "52": "00FF9900",
  "53": "00FF6600",
  "54": "00666699",
  "55": "00969696",
  "56": "00003366",
  "57": "00339966",
  "58": "00003300",
  "59": "00333300",
  "60": "00993300",
  "61": "00993366",
  "62": "00333399",
  "63": "00333333",
  "64": null,
  "65": null
};

export function getptToPxRatioByDPI() {
  return 72 / 96;
}

export function getXmlAttibute(dom, attr, d) {
  let value = dom[attr];
  value = value == null ? d : value;
  return value;
}

export function getPxByEMUs(emus) {
  if (emus == null) {
    return 0;
  }

  const inch = emus / 914400;
  const pt = inch * 72;
  return pt / getptToPxRatioByDPI();
}

export function getcellrange(txt, sheets, sheetId) {
  if (sheets === void 0) {
    sheets = {};
  }

  if (sheetId === void 0) {
    sheetId = "1";
  }

  const val = txt.split("!");
  let sheettxt = "",
    rangetxt = "",
    sheetIndex = -1;

  if (val.length > 1) {
    sheettxt = val[0];
    rangetxt = val[1];
    const si = sheets[sheettxt];

    if (si == null) {
      sheetIndex = parseInt(sheetId);
    } else {
      sheetIndex = parseInt(si);
    }
  } else {
    sheetIndex = parseInt(sheetId);
    rangetxt = val[0];
  }

  if (rangetxt.indexOf(":") == -1) {
    const row = parseInt(rangetxt.replace(/[^0-9]/g, "")) - 1;
    const col = ABCatNum(rangetxt.replace(/[^A-Za-z]/g, ""));

    if (!isNaN(row) && !isNaN(col)) {
      return {
        row: [row, row],
        column: [col, col],
        sheetIndex: sheetIndex
      };
    } else {
      return null;
    }
  } else {
    const rangetxtArray = rangetxt.split(":");
    const row = [],
      col = [];
    row[0] = parseInt(rangetxtArray[0].replace(/[^0-9]/g, "")) - 1;
    row[1] = parseInt(rangetxtArray[1].replace(/[^0-9]/g, "")) - 1; // if (isNaN(row[0])) {
    //     row[0] = 0;
    // }
    // if (isNaN(row[1])) {
    //     row[1] = sheetdata.length - 1;
    // }

    if (row[0] > row[1]) {
      return null;
    }

    col[0] = ABCatNum(rangetxtArray[0].replace(/[^A-Za-z]/g, ""));
    col[1] = ABCatNum(rangetxtArray[1].replace(/[^A-Za-z]/g, "")); // if (isNaN(col[0])) {
    //     col[0] = 0;
    // }
    // if (isNaN(col[1])) {
    //     col[1] = sheetdata[0].length - 1;
    // }

    if (col[0] > col[1]) {
      return null;
    }

    return {
      row: row,
      column: col,
      sheetIndex: sheetIndex
    };
  }
}

function ABCatNum(abc) {
  abc = abc.toUpperCase();
  const abc_len = abc.length;

  if (abc_len == 0) {
    return NaN;
  }

  const abc_array = abc.split("");
  const wordlen = columeHeader_word.length;
  let ret = 0;

  for (let i = abc_len - 1; i >= 0; i--) {
    if (i == abc_len - 1) {
      ret += columeHeader_word_index[abc_array[i]];
    } else {
      ret +=
        Math.pow(wordlen, abc_len - i - 1) *
        (columeHeader_word_index[abc_array[i]] + 1);
    }
  }

  return ret;
}

export function generateRandomIndex(prefix) {
  if (prefix == null) {
    prefix = "Sheet";
  }

  const userAgent = window.navigator.userAgent
    .replace(/[^a-zA-Z0-9]/g, "")
    .split("");
  let mid = "";

  for (let i = 0; i < 5; i++) {
    mid += userAgent[Math.round(Math.random() * (userAgent.length - 1))];
  }

  const time = new Date().getTime();
  return prefix + "_" + mid + "_" + time;
}

export function escapeCharacter(str) {
  if (str == null || str.length == 0) {
    return str;
  }

  return str
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&nbsp;/g, " ")
    .replace(/&apos;/g, "'")
    .replace(/&iexcl;/g, "¡")
    .replace(/&cent;/g, "¢")
    .replace(/&pound;/g, "£")
    .replace(/&curren;/g, "¤")
    .replace(/&yen;/g, "¥")
    .replace(/&brvbar;/g, "¦")
    .replace(/&sect;/g, "§")
    .replace(/&uml;/g, "¨")
    .replace(/&copy;/g, "©")
    .replace(/&ordf;/g, "ª")
    .replace(/&laquo;/g, "«")
    .replace(/&not;/g, "¬")
    .replace(/&shy;/g, "­")
    .replace(/&reg;/g, "®")
    .replace(/&macr;/g, "¯")
    .replace(/&deg;/g, "°")
    .replace(/&plusmn;/g, "±")
    .replace(/&sup2;/g, "²")
    .replace(/&sup3;/g, "³")
    .replace(/&acute;/g, "´")
    .replace(/&micro;/g, "µ")
    .replace(/&para;/g, "¶")
    .replace(/&middot;/g, "·")
    .replace(/&cedil;/g, "¸")
    .replace(/&sup1;/g, "¹")
    .replace(/&ordm;/g, "º")
    .replace(/&raquo;/g, "»")
    .replace(/&frac14;/g, "¼")
    .replace(/&frac12;/g, "½")
    .replace(/&frac34;/g, "¾")
    .replace(/&iquest;/g, "¿")
    .replace(/&times;/g, "×")
    .replace(/&divide;/g, "÷")
    .replace(/&Agrave;/g, "À")
    .replace(/&Aacute;/g, "Á")
    .replace(/&Acirc;/g, "Â")
    .replace(/&Atilde;/g, "Ã")
    .replace(/&Auml;/g, "Ä")
    .replace(/&Aring;/g, "Å")
    .replace(/&AElig;/g, "Æ")
    .replace(/&Ccedil;/g, "Ç")
    .replace(/&Egrave;/g, "È")
    .replace(/&Eacute;/g, "É")
    .replace(/&Ecirc;/g, "Ê")
    .replace(/&Euml;/g, "Ë")
    .replace(/&Igrave;/g, "Ì")
    .replace(/&Iacute;/g, "Í")
    .replace(/&Icirc;/g, "Î")
    .replace(/&Iuml;/g, "Ï")
    .replace(/&ETH;/g, "Ð")
    .replace(/&Ntilde;/g, "Ñ")
    .replace(/&Ograve;/g, "Ò")
    .replace(/&Oacute;/g, "Ó")
    .replace(/&Ocirc;/g, "Ô")
    .replace(/&Otilde;/g, "Õ")
    .replace(/&Ouml;/g, "Ö")
    .replace(/&Oslash;/g, "Ø")
    .replace(/&Ugrave;/g, "Ù")
    .replace(/&Uacute;/g, "Ú")
    .replace(/&Ucirc;/g, "Û")
    .replace(/&Uuml;/g, "Ü")
    .replace(/&Yacute;/g, "Ý")
    .replace(/&THORN;/g, "Þ")
    .replace(/&szlig;/g, "ß")
    .replace(/&agrave;/g, "à")
    .replace(/&aacute;/g, "á")
    .replace(/&acirc;/g, "â")
    .replace(/&atilde;/g, "ã")
    .replace(/&auml;/g, "ä")
    .replace(/&aring;/g, "å")
    .replace(/&aelig;/g, "æ")
    .replace(/&ccedil;/g, "ç")
    .replace(/&egrave;/g, "è")
    .replace(/&eacute;/g, "é")
    .replace(/&ecirc;/g, "ê")
    .replace(/&euml;/g, "ë")
    .replace(/&igrave;/g, "ì")
    .replace(/&iacute;/g, "í")
    .replace(/&icirc;/g, "î")
    .replace(/&iuml;/g, "ï")
    .replace(/&eth;/g, "ð")
    .replace(/&ntilde;/g, "ñ")
    .replace(/&ograve;/g, "ò")
    .replace(/&oacute;/g, "ó")
    .replace(/&ocirc;/g, "ô")
    .replace(/&otilde;/g, "õ")
    .replace(/&ouml;/g, "ö")
    .replace(/&oslash;/g, "ø")
    .replace(/&ugrave;/g, "ù")
    .replace(/&uacute;/g, "ú")
    .replace(/&ucirc;/g, "û")
    .replace(/&uuml;/g, "ü")
    .replace(/&yacute;/g, "ý")
    .replace(/&thorn;/g, "þ")
    .replace(/&yuml;/g, "ÿ");
}

export function getColor(color, styles, type) {
  if (type === void 0) {
    type = "g";
  }

  const attrList = color.attributeList;
  const clrScheme = styles["clrScheme"];
  const indexedColorsInner = styles["indexedColors"];
  // const mruColorsInner = styles["mruColors"];

  let indexedColorsList = {};
  if (indexedColorsInner == null || indexedColorsInner.length == 0) {
    return indexedColors;
  }

  for (const key in indexedColors) {
    const value = indexedColors[key],
      kn = parseInt(key);
    const inner = indexedColorsInner[kn];

    if (inner == null) {
      indexedColorsList[key] = value;
    } else {
      const rgb = inner.attributeList.rgb;
      indexedColorsList[key] = rgb;
    }
  }

  let indexed = attrList.indexed,
    rgb = attrList.rgb,
    theme = attrList.theme,
    tint = attrList.tint;
  let bg;

  if (indexed != null) {
    const indexedNum = parseInt(indexed);
    bg = indexedColorsList[indexedNum];

    if (bg != null) {
      bg = bg.substring(bg.length - 6, bg.length);
      bg = "#" + bg;
    }
  } else if (rgb != null) {
    rgb = rgb.substring(rgb.length - 6, rgb.length);
    bg = "#" + rgb;
  } else if (theme != null) {
    let themeNum = parseInt(theme);

    if (themeNum == 0) {
      themeNum = 1;
    } else if (themeNum == 1) {
      themeNum = 0;
    } else if (themeNum == 2) {
      themeNum = 3;
    } else if (themeNum == 3) {
      themeNum = 2;
    }

    const clrSchemeElement = clrScheme[themeNum];

    if (clrSchemeElement != null) {
      const clrs = clrSchemeElement.getInnerElements("a:sysClr|a:srgbClr");

      if (clrs != null) {
        const clr = clrs[0];
        const clrAttrList = clr.attributeList; // console.log(clr.container, );

        if (clr.container.indexOf("sysClr") > -1) {
          if (clrAttrList.lastClr != null) {
            bg = "#" + clrAttrList.lastClr;
          } else if (clrAttrList.val != null) {
            bg = "#" + clrAttrList.val;
          }
        } else if (clr.container.indexOf("srgbClr") > -1) {
          // console.log(clrAttrList.val);
          bg = "#" + clrAttrList.val;
        }
      }
    }
  }

  if (tint != null) {
    const tintNum = parseFloat(tint);

    if (bg != null) {
      bg = LightenDarkenColor(bg, tintNum);
    }
  }

  return bg;
}

function LightenDarkenColor(sixColor, tint) {
  const hex = sixColor.substring(sixColor.length - 6, sixColor.length);
  const rgbArray = hexToRgbArray("#" + hex);
  const hslArray = rgbToHsl(rgbArray[0], rgbArray[1], rgbArray[2]);

  if (tint > 0) {
    hslArray[2] = hslArray[2] * (1.0 - tint) + tint;
  } else if (tint < 0) {
    hslArray[2] = hslArray[2] * (1.0 + tint);
  } else {
    return "#" + hex;
  }

  const newRgbArray = hslToRgb(hslArray[0], hslArray[1], hslArray[2]);
  return rgbToHex("RGB(" + newRgbArray.join(",") + ")");
}

function hexToRgbArray(hex) {
  let sColor = hex.toLowerCase(); //十六进制颜色值的正则表达式

  const reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{6})$/; // 如果是16进制颜色

  if (sColor && reg.test(sColor)) {
    if (sColor.length === 4) {
      let sColorNew = "#";

      for (let i = 1; i < 4; i += 1) {
        sColorNew += sColor.slice(i, i + 1).concat(sColor.slice(i, i + 1));
      }

      sColor = sColorNew;
    } //处理六位的颜色值

    const sColorChange = [];

    for (let i = 1; i < 7; i += 2) {
      sColorChange.push(parseInt("0x" + sColor.slice(i, i + 2)));
    }

    return sColorChange;
  }

  return null;
}

function rgbToHex(rgb) {
  //十六进制颜色值的正则表达式
  const reg = /^#([0-9a-fA-f]{3}|[0-9a-fA-f]{6})$/; // 如果是rgb颜色表示

  if (/^(rgb|RGB)/.test(rgb)) {
    const aColor = rgb.replace(/(?:\(|\)|rgb|RGB)*/g, "").split(",");
    let strHex = "#";

    for (let i = 0; i < aColor.length; i++) {
      let hex = Number(aColor[i]).toString(16);

      if (hex.length < 2) {
        hex = "0" + hex;
      }

      strHex += hex;
    }

    if (strHex.length !== 7) {
      strHex = rgb;
    }

    return strHex;
  } else if (reg.test(rgb)) {
    const aNum = rgb.replace(/#/, "").split("");

    if (aNum.length === 6) {
      return rgb;
    } else if (aNum.length === 3) {
      let numHex = "#";

      for (let i = 0; i < aNum.length; i += 1) {
        numHex += aNum[i] + aNum[i];
      }

      return numHex;
    }
  }

  return rgb;
}

function hslToRgb(h, s, l) {
  let r, g, b;

  if (s == 0) {
    r = g = b = l; // achromatic
  } else {
    const hue2rgb = function hue2rgb(p, q, t) {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };

    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hue2rgb(p, q, h + 1 / 3);
    g = hue2rgb(p, q, h);
    b = hue2rgb(p, q, h - 1 / 3);
  }

  return [Math.round(r * 255), Math.round(g * 255), Math.round(b * 255)];
}

function rgbToHsl(r, g, b) {
  (r /= 255), (g /= 255), (b /= 255);
  const max = Math.max(r, g, b),
    min = Math.min(r, g, b);
  let h,
    s,
    l = (max + min) / 2;

  if (max == min) {
    h = s = 0; // achromatic
  } else {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);

    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;

      case g:
        h = (b - r) / d + 2;
        break;

      case b:
        h = (r - g) / d + 4;
        break;
    }

    h /= 6;
  }

  return [h, s, l];
}

export function isChinese(temp) {
  const re = /[^\u4e00-\u9fa5]/;
  const reg = /[\u3002|\uff1f|\uff01|\uff0c|\u3001|\uff1b|\uff1a|\u201c|\u201d|\u2018|\u2019|\uff08|\uff09|\u300a|\u300b|\u3008|\u3009|\u3010|\u3011|\u300e|\u300f|\u300c|\u300d|\ufe43|\ufe44|\u3014|\u3015|\u2026|\u2014|\uff5e|\ufe4f|\uffe5]/;
  if (reg.test(temp)) return true;
  if (re.test(temp)) return false;
  return true;
}

export function isJapanese(temp) {
  const re = /[^\u0800-\u4e00]/;
  if (re.test(temp)) return false;
  return true;
}

export function isKoera(chr) {
  if ((chr > 0x3130 && chr < 0x318f) || (chr >= 0xac00 && chr <= 0xd7a3)) {
    return true;
  }

  return false;
}

export function getlineStringAttr(frpr, attr) {
  const attrEle = frpr.getInnerElements(attr);
  let value;

  if (attrEle != null && attrEle.length > 0) {
    if (attr == "b" || attr == "i" || attr == "strike") {
      value = "1";
    } else if (attr == "u") {
      const v = attrEle[0].attributeList.val;

      if (v == "double") {
        value = "2";
      } else if (v == "singleAccounting") {
        value = "3";
      } else if (v == "doubleAccounting") {
        value = "4";
      } else {
        value = "1";
      }
    } else if (attr == "vertAlign") {
      const v = attrEle[0].attributeList.val;

      if (v == "subscript") {
        value = "1";
      } else if (v == "superscript") {
        value = "2";
      }
    } else {
      value = attrEle[0].attributeList.val;
    }
  }

  return value;
}

export function getAspectRatio() {
  return 16 / 9;
}

export function getBase64Image(url) {
  return new Promise((resolve, reject) => {
    const image = new Image();

    image.onload = function() {
      const canvas = document.createElement("canvas");
      canvas.width = image.width;
      canvas.height = image.height;

      canvas.getContext("2d")?.drawImage(image, 0, 0);
      const dataURI = canvas.toDataURL("image/png");
      resolve(dataURI);
    };

    image.onerror = function() {
      // on failure
      reject("Error Loading Image");
    };

    image.setAttribute("crossOrigin", "anonymous");
    image.src = url;
  });
}

export const boringWait = (cond, timeout = 10000) => {
  return new Promise((resolve, reject) => {
    let time = 0;
    const t = setInterval(() => {
      time++;
      if (cond()) {
        clearInterval(t);
        resolve();
      }
      if (time > timeout / 100) {
        // 大于10秒，超时
        clearInterval(t);
        reject();
      }
    }, 100);
  });
};
