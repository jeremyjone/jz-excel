import JSZip from "jszip";
import { ParseExcelFile } from "./file";

/**
 * 惊喜！嗯，第一次看到 LuckySheet 的时候，就是惊喜。丫居然是 MIT 的，我需要的就是能解析图片的玩意。
 * 之前一直是把 excel 给后端，解析后再返回来，好麻烦，有了这玩意，简直不要太好用。
 *
 * 说明：这个解析方法是魔改的 LuckySheet，功能很强大，效果非常好。
 * 我只是对代码修修剪剪，补补改改，然后就成了现在这个样子。它相当于是一个 excel2json 的增强版本。
 * 我把暂时用不到的数据都删了，只保留了需要用到的最基本数据的字段，但是保留的数据字段都没有修改键值。
 * 不修改键值，这给之后如果替换成 LuckySheet 提供了便利，相当于之后可以直接扩展。
 * 主要是这个玩意可以直接读取数据和图片，这样就省事儿多了
 *
 * 吐槽一下：丫里面各种 console.log，实在受不了。
 *
 * 使用方式：把 excel 文件作为参数传进来，然后跟着一个回调函数。
 * 回调函数的参数就是抛出去的最终对象:
 * parseExcel(ExcelFile, data => {console.log(data)})
 *
 * 整个 data 是这个鸟样子的：
 * {
 *   info: {...各种文件信息，自己打印看吧},
 *   sheets: [...所有表格信息，主要用到的是：
 *     {
 *       celldata: [...每一格的数据,
 *         {
 *           c: number, // 列
 *           r: number, // 行
 *           v: { v: any, [other_key]: any }, // v 是数据内容
 *         }
 *       ],
 *       images: {
 *         image_key: {
 *           pathName: relative path,
 *           src: base64,
 *           fromRow: number, // 起始自哪一行
 *           toRow: number, // 结束自哪一行
 *           fromCol: number, // 起始自哪一列
 *           toRow: number, // 结束自哪一列
 *           [other_key]: 自己打印看吧
 *         }
 *       }
 *     }
 *   ]
 * }
 *
 * @param {File} excel
 * @param {(data: null | object, err: undefined | any) => any} cb
 */
export function parseExcel(excel, cb) {
  if (excel.name.split(".").pop().toLowerCase() !== "xlsx") {
    cb(null, "It not be a excel file. Only can parse 'xlsx' file.");
    return;
  }

  const jszip = new JSZip();

  jszip.loadAsync(excel).then(
    zip => {
      const fileList = {},
        lastIndex = Object.keys(zip.files).length;
      let index = 0;

      Object.keys(zip.files).forEach(filename => {
        const fileNameArr = filename.split(".");
        const suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();
        let fileType = "string";

        if (
          suffix in
          {
            png: 1,
            jpeg: 1,
            jpg: 1,
            gif: 1,
            bmp: 1,
            tif: 1,
            webp: 1
          }
        ) {
          fileType = "base64";
        } else if (suffix == "emf") {
          fileType = "arraybuffer";
        }

        // FIXME: 一个奇怪的问题，async 返回的 Promise 在浏览器中一直处于 Pending。
        // 但是之前 v1.0.2 打包的文件在浏览器中可以执行，函数没有动过，看过源码，没啥问题。
        // 已经提问，无人问津，苦涩。。。🤯 https://stackoverflow.com/questions/68172627/zipobject-async-function-not-work-in-browser
        // 如果亲爱的你看到了这个问题，并且有方案，欢迎 PR 一下。🤓 https://github.com/jeremyjone/jz-excel
        zip.files[filename].async(fileType).then(data => {
          if (fileType == "base64") {
            data = "data:image/" + suffix + ";base64," + data;
          }

          fileList[filename] = data;

          if (lastIndex == index + 1) {
            const f = new ParseExcelFile(fileList, excel.name);
            cb(JSON.parse(f.Parse()));
          }

          index++;
        });
      });
    },
    err => cb(null, err)
  );
}
