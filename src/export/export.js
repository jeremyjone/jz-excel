import ExcelJS from "exceljs";
import {
  getAspectRatio,
  getptToPxRatioByDPI,
  getBase64Image,
  boringWait
} from "../tools/func";

export class ExportExcel {
  constructor(companyName = "") {
    this.companyName = companyName;

    this.workbook = new ExcelJS.Workbook();
    this.workbook.creator = companyName;
    this.workbook.lastModifiedBy = companyName;
    this.workbook.created = new Date();
    this.workbook.modified = new Date();

    this.imageWidth = 300;
    this.imageHeight = this.imageWidth / getAspectRatio();
  }

  /**
   * 设置工作簿的公司名称
   * @param {string} name 公司名称
   */
  setCompanyName(name) {
    this.companyName = companyName;
    this.workbook.creator = companyName;
    this.workbook.lastModifiedBy = companyName;
  }

  /**
   * 添加工作表
   * @param {string} name 工作表名。默认 sheet1
   * @returns 返回添加的工作表名。因为实际创建可能存在重名，如果重名就添加一个后缀
   */
  addSheet(name = "sheet1") {
    // sheet 名字长度最大为30。超过的话，取中间部分内容转为 ~ 符号
    if (name.length > 25) {
      name = name.replace(/^(.{20})(.*)(.{5})$/g, "$1~$3");
    }

    let i = 1;
    this.workbook.eachSheet((sheet, id) => {
      if (sheet.name === name) {
        name = `${name}_${i}`;
        i++;
      }
    });

    this.workbook.addWorksheet(name);

    return name;
  }

  /**
   * 添加内容
   * @param {string} sheetName 工作表名
   * @param {Array<Array<{value: ExcelJS.CellValue, options?: {width?: number, height?: number, font?: Partial<ExcelJS.Font>, alignment?: Partial<ExcelJS.Alignment>}}>>} contents 要添加的表内容
   */
  addContents(sheetName, contents) {
    const sheet = this.workbook.getWorksheet(sheetName);
    if (!sheet) throw Error(`sheet named "${sheetName}" is not exist.`);

    // 写入内容
    contents.forEach(content => {
      // 添加行
      const row = sheet.addRow([]);
      let height;

      // 行内循环，设置属性
      content.forEach((item, i) => {
        const index = ++i;
        item?.options?.height &&
          item.options.height > height &&
          (height = item.options.height) &&
          (row.height = height);

        item?.options?.font && (row.font = item.options.font);
        item?.options?.alignment && (row.alignment = item.options.alignment);

        const cell = row.getCell(index);
        item?.options?.width && (cell.width = item.options.width);
        cell.value = item.value;
      });
    });
  }

  /**
   * 添加图片。该方法是异步方法
   * @param {string} sheetName 工作表名
   * @param {Array<{value: string, isBase64: boolean, r: number, c: number, offset?: number}>} images 要添加的图片列表
   */
  addImagesAsync(sheetName, images) {
    const sheet = this.workbook.getWorksheet(sheetName);
    if (!sheet) throw Error(`sheet named "${sheetName}" is not exist.`);

    const threads = [];

    return new Promise((resolve, reject) => {
      // 写入内容
      images.map(image => {
        const totalRows = sheet.rowCount;

        // 如果图片位置超出了当前表的范围，添加行
        if (image.r > totalRows)
          sheet.addRows(new Array(image.r - totalRows).fill(""));

        // 设置图片高度
        sheet.getRow(image.r).height =
          this.imageHeight * getptToPxRatioByDPI() + 5;

        sheet.getColumn(image.c).width = this.imageWidth * 0.2;

        let pic = "";
        if (image.isBase64) pic = image.value;
        else {
          threads.push(Symbol());
          getBase64Image(image.value)
            .then(res => {
              pic = res;

              const imgId = this.workbook.addImage({
                base64: pic,
                extension: "png"
              });

              const nativeRow = image.r - 1;
              const nativeCol = image.c - 1 + (image?.offset ?? 0);

              sheet.addImage(imgId, {
                tl: {
                  col: nativeCol,
                  row: nativeRow
                },

                // br: {
                //   col: nativeCol + 1,
                //   row: nativeRow + 1
                // },
                ext: {
                  width: this.imageWidth * 0.6,
                  height: this.imageHeight * 0.8
                },
                editAs: "undefined"
              });

              threads.pop();
            })
            .catch(err => {});
        }
      });

      // 图片获取时异步的，需要等待
      boringWait(() => threads.length === 0)
        .then(result => {
          resolve(result);
        })
        .catch(err => {
          reject(err);
        });
    });
  }

  /**
   * 导出 Excel
   * @param {string} filename 下载后保存的 Excel 文件名
   */
  async export(filename) {
    if (!filename) throw Error("Excel filename is required.");

    // 保证图片完全获取完成，将下载方法放进队列
    return new Promise((resolve, reject) => {
      // 保存数据并下载
      this.workbook.xlsx
        .writeBuffer()
        .then(buffer => {
          const blob = new Blob([buffer], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          });
          const url = window.URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url;
          a.download = `${filename.replace(/.xls[x]?$/, "")}.xlsx`;
          a.style.display = "none";
          a.click();
          window.URL.revokeObjectURL(url);

          resolve("downloaded");
        })
        .catch(err => {
          reject(err);
        });
    });
  }
}
