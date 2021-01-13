import { getXmlAttibute } from "./func";
import { ParseImage } from "./image";
import { Sheet } from "./sheet";
import { ReadXml } from "./xml";

const coreFile = "docProps/core.xml";
const appFile = "docProps/app.xml";
const workBookFile = "xl/workbook.xml";
const stylesFile = "xl/styles.xml";
const sharedStringsFile = "xl/sharedStrings.xml";
const theme1File = "xl/theme/theme1.xml";
const workbookRels = "xl/_rels/workbook.xml.rels";
const numFmtDefault = {
  "0": "General",
  "1": "0",
  "2": "0.00",
  "3": "#,##0",
  "4": "#,##0.00",
  "9": "0%",
  "10": "0.00%",
  "11": "0.00E+00",
  "12": "# ?/?",
  "13": "# ??/??",
  "14": "m/d/yy",
  "15": "d-mmm-yy",
  "16": "d-mmm",
  "17": "mmm-yy",
  "18": "h:mm AM/PM",
  "19": "h:mm:ss AM/PM",
  "20": "h:mm",
  "21": "h:mm:ss",
  "22": "m/d/yy h:mm",
  "37": "#,##0 ;(#,##0)",
  "38": "#,##0 ;[Red](#,##0)",
  "39": "#,##0.00;(#,##0.00)",
  "40": "#,##0.00;[Red](#,##0.00)",
  "45": "mm:ss",
  "46": "[h]:mm:ss",
  "47": "mmss.0",
  "48": "##0.0E+0",
  "49": "@"
};

export class ParseExcelFile {
  constructor(files, fileName) {
    this.files = files;
    this.fileName = fileName;
    this.xml = new ReadXml(this.files);

    this.sheetNameList = this.getSheetNameList();

    this.sharedStrings = this.xml.getElementsByTagName(
      "sst/si",
      sharedStringsFile
    );
    this.styles = {};
    this.styles["cellXfs"] = this.xml.getElementsByTagName(
      "cellXfs/xf",
      stylesFile
    );
    this.styles["cellStyleXfs"] = this.xml.getElementsByTagName(
      "cellStyleXfs/xf",
      stylesFile
    );
    this.styles["cellStyles"] = this.xml.getElementsByTagName(
      "cellStyles/cellStyle",
      stylesFile
    );
    this.styles["fonts"] = this.xml.getElementsByTagName(
      "fonts/font",
      stylesFile
    );
    this.styles["fills"] = this.xml.getElementsByTagName(
      "fills/fill",
      stylesFile
    );
    this.styles["borders"] = this.xml.getElementsByTagName(
      "borders/border",
      stylesFile
    );
    this.styles["clrScheme"] = this.xml.getElementsByTagName(
      "a:clrScheme/a:dk1|a:lt1|a:dk2|a:lt2|a:accent1|a:accent2|a:accent3|a:accent4|a:accent5|a:accent6|a:hlink|a:folHlink",
      theme1File
    );
    this.styles["indexedColors"] = this.xml.getElementsByTagName(
      "colors/indexedColors/rgbColor",
      stylesFile
    );
    this.styles["mruColors"] = this.xml.getElementsByTagName(
      "colors/mruColors/color",
      stylesFile
    );
    this.imageList = new ParseImage(this.files);

    const numfmts = this.xml.getElementsByTagName("numFmt/numFmt", stylesFile);

    const numFmtDefaultC = numFmtDefault;

    for (let i = 0; i < numfmts.length; i++) {
      const attrList = numfmts[i].attributeList;
      const numfmtid = getXmlAttibute(attrList, "numFmtId", "49");
      const formatcode = getXmlAttibute(attrList, "formatCode", "@"); // console.log(numfmtid, formatcode);

      if (!(numfmtid in numFmtDefault)) {
        numFmtDefaultC[numfmtid] = formatcode;
      }
    }

    this.styles["numfmts"] = numFmtDefaultC;
  }

  getSheetNameList() {
    const workbookRelList = this.xml.getElementsByTagName(
      "Relationships/Relationship",
      workbookRels
    );

    if (workbookRelList == null) {
      return;
    }

    const regex = new RegExp("worksheets/[^/]*?.xml");
    const sheetNames = {};

    for (let i = 0; i < workbookRelList.length; i++) {
      const rel = workbookRelList[i],
        attrList = rel.attributeList;
      const id = attrList["Id"],
        target = attrList["Target"];

      if (regex.test(target)) {
        sheetNames[id] = "xl/" + target;
      }
    }

    return sheetNames;
  }

  getSheetFileBysheetId(sheetId) {
    return this.sheetNameList[sheetId];
  }

  getWorkBookInfo() {
    var Company = this.xml.getElementsByTagName("Company", appFile);
    var AppVersion = this.xml.getElementsByTagName("AppVersion", appFile);
    var creator = this.xml.getElementsByTagName("dc:creator", coreFile);
    var lastModifiedBy = this.xml.getElementsByTagName(
      "cp:lastModifiedBy",
      coreFile
    );
    var created = this.xml.getElementsByTagName("dcterms:created", coreFile);
    var modified = this.xml.getElementsByTagName("dcterms:modified", coreFile);
    this.info = {};
    this.info.name = this.fileName;
    this.info.creator = creator.length > 0 ? creator[0].value : "";
    this.info.lastmodifiedby =
      lastModifiedBy.length > 0 ? lastModifiedBy[0].value : "";
    this.info.createdTime = created.length > 0 ? created[0].value : "";
    this.info.modifiedTime = modified.length > 0 ? modified[0].value : "";
    this.info.company = Company.length > 0 ? Company[0].value : "";
    this.info.appversion = AppVersion.length > 0 ? AppVersion[0].value : "";
  }

  getSheetsFull() {
    const sheets = this.xml.getElementsByTagName("sheets/sheet", workBookFile);
    const sheetList = {};

    for (const key in sheets) {
      const sheet = sheets[key];
      sheetList[sheet.attributeList.name] = sheet.attributeList["sheetId"];
    }

    this.sheets = [];
    let order = 0;

    for (const key in sheets) {
      const sheet = sheets[key];
      const sheetName = sheet.attributeList.name;
      const sheetId = sheet.attributeList["sheetId"];
      const rid = sheet.attributeList["r:id"];
      const sheetFile = this.getSheetFileBysheetId(rid);
      const drawing = this.xml.getElementsByTagName(
        "worksheet/drawing",
        sheetFile
      );
      let drawingFile = void 0,
        drawingRelsFile = void 0;

      if (drawing != null && drawing.length > 0) {
        const attrList = drawing[0].attributeList;
        const rid_1 = getXmlAttibute(attrList, "r:id", null);

        if (rid_1 != null) {
          drawingFile = this.getDrawingFile(rid_1, sheetFile);
          drawingRelsFile = this.getDrawingRelsFile(drawingFile);
        }
      }

      if (sheetFile != null) {
        const sheet = new Sheet(sheetName, sheetId, order, {
          sheetFile: sheetFile,
          xml: this.xml,
          sheetList: sheetList,
          styles: this.styles,
          sharedStrings: this.sharedStrings,
          imageList: this.imageList,
          drawingFile: drawingFile,
          drawingRelsFile: drawingRelsFile
        });
        this.columnWidthSet = [];
        this.rowHeightSet = [];
        this.imagePositionCaculation(sheet);
        this.sheets.push(sheet);
        order++;
      }
    }
  }

  imagePositionCaculation(sheet) {
    const images = sheet.images,
      defaultColWidth = sheet.defaultColWidth,
      defaultRowHeight = sheet.defaultRowHeight;
    let colhidden = {};
    let columnlen = {};
    let rowhidden = {};
    let rowlen = {};

    for (const key in images) {
      const imageObject = images[key];

      const fromCol = imageObject.fromCol;
      const fromColOff = imageObject.fromColOff;
      const fromRow = imageObject.fromRow;
      const fromRowOff = imageObject.fromRowOff;
      const toCol = imageObject.toCol;
      const toColOff = imageObject.toColOff;
      const toRow = imageObject.toRow;
      const toRowOff = imageObject.toRowOff;
      let x_n = 0,
        y_n = 0,
        cx_n = 0,
        cy_n = 0;

      if (fromCol >= this.columnWidthSet.length) {
        this.extendArray(
          fromCol,
          this.columnWidthSet,
          defaultColWidth,
          colhidden,
          columnlen
        );
      }

      if (fromCol == 0) {
        x_n = 0;
      } else {
        x_n = this.columnWidthSet[fromCol - 1];
      }

      x_n = x_n + fromColOff;

      if (fromRow >= this.rowHeightSet.length) {
        this.extendArray(
          fromRow,
          this.rowHeightSet,
          defaultRowHeight,
          rowhidden,
          rowlen
        );
      }

      if (fromRow == 0) {
        y_n = 0;
      } else {
        y_n = this.rowHeightSet[fromRow - 1];
      }

      y_n = y_n + fromRowOff;

      if (toCol >= this.columnWidthSet.length) {
        this.extendArray(
          toCol,
          this.columnWidthSet,
          defaultColWidth,
          colhidden,
          columnlen
        );
      }

      if (toCol == 0) {
        cx_n = 0;
      } else {
        cx_n = this.columnWidthSet[toCol - 1];
      }

      cx_n = cx_n + toColOff - x_n;

      if (toRow >= this.rowHeightSet.length) {
        this.extendArray(
          toRow,
          this.rowHeightSet,
          defaultRowHeight,
          rowhidden,
          rowlen
        );
      }

      if (toRow == 0) {
        cy_n = 0;
      } else {
        cy_n = this.rowHeightSet[toRow - 1];
      }

      cy_n = cy_n + toRowOff - y_n;
      imageObject.originWidth = cx_n;
      imageObject.originHeight = cy_n;
      imageObject.crop.height = cy_n;
      imageObject.crop.width = cx_n;
      imageObject["default"].height = cy_n;
      imageObject["default"].left = x_n;
      imageObject["default"].top = y_n;
      imageObject["default"].width = cx_n;
    }
  }

  extendArray(index, sets, def, hidden, lens) {
    if (index < sets.length) {
      return;
    }

    const startIndex = sets.length,
      endIndex = index;
    let allGap = 0;

    if (startIndex > 0) {
      allGap = sets[startIndex - 1];
    } // else{
    //     sets.push(0);
    // }

    for (let i = startIndex; i <= endIndex; i++) {
      let gap = def,
        istring = i.toString();

      if (istring in hidden) {
        gap = 0;
      } else if (istring in lens) {
        gap = lens[istring];
      }

      allGap += Math.round(gap + 1);
      sets.push(allGap);
    }
  }

  getDrawingFile(rid, sheetFile) {
    const sheetRelsPath = "xl/worksheets/_rels/";
    const sheetFileArr = sheetFile.split("/");
    const sheetRelsName = sheetFileArr[sheetFileArr.length - 1];
    const sheetRelsFile = sheetRelsPath + sheetRelsName + ".rels";
    const drawing = this.xml.getElementsByTagName(
      "Relationships/Relationship",
      sheetRelsFile
    );

    if (drawing.length > 0) {
      for (let i = 0; i < drawing.length; i++) {
        const attrList = drawing[i].attributeList;
        const relationshipId = getXmlAttibute(attrList, "Id", null);

        if (relationshipId == rid) {
          const target = getXmlAttibute(attrList, "Target", null);

          if (target != null) {
            return target.replace(/\.\.\//g, "");
          }
        }
      }
    }

    return null;
  }

  getDrawingRelsFile(drawingFile) {
    const drawingRelsPath = "xl/drawings/_rels/";
    const drawingFileArr = drawingFile.split("/");
    const drawingRelsName = drawingFileArr[drawingFileArr.length - 1];
    const drawingRelsFile = drawingRelsPath + drawingRelsName + ".rels";
    return drawingRelsFile;
  }

  Parse() {
    this.getWorkBookInfo();
    this.getSheetsFull();

    return this.toJsonString(this);
  }

  toJsonString(file) {
    const output = {};
    output.info = file.info;
    output.sheets = [];
    file.sheets.forEach(sheet => {
      const sheetout = {};

      if (sheet.name != null) {
        sheetout.name = sheet.name;
      }

      if (sheet.index != null) {
        sheetout.index = sheet.index;
      }

      if (sheet.order != null) {
        sheetout.order = sheet.order;
      }

      if (sheet.celldata != null) {
        // sheetout.celldata = sheet.celldata;
        sheetout.celldata = [];
        sheet.celldata.forEach(cell => {
          const cellout = {};
          cellout.r = cell.r;
          cellout.c = cell.c;
          cellout.v = cell.v;
          sheetout.celldata.push(cellout);
        });
      }

      if (sheet.images != null) {
        sheetout.images = sheet.images;
      }

      output.sheets.push(sheetout);
    });
    return JSON.stringify(output);
  }
}
