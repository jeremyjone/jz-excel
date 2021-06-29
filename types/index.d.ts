import ExcelJS from "exceljs";

export type JExcelInfo = {
  appversion: string;
  company: string;
  createdTime: string;
  cretor: string;
  lastmodifiedby: string;
  modifiedTime: string;
  name: string;
  [key: string]: any;
};

export interface JExcelData {
  c: number;
  r: number;
  /**
   * 这是为了更好的扩展，对数据结构不做改动。
   * - 数据如果没有特殊格式，直接 v.v
   * - 相反的，如果内容有格式，按照 LuckySheet 的处理，它将把内容进行拆解，所以返回 ct.s 的数组，其中 v 是内容，拼接即可
   */
  v: {
    v?: string;
    ct?: { s: Array<{ v: string; [key: string]: any }>; [key: string]: any };
  };
}

export type JExcelImage = {
  pathName: string;
  /**
   * base64
   */
  src: string;
  fromCol: number;
  fromColOff: number;
  fromRow: number;
  fromRowOff: number;
  toCol: number;
  toColOff: number;
  toRow: number;
  toRowOff: number;
  [key: string]: any;
};

export type JExcelSheets = {
  celldata: Array<JExcelData>;
  images: {
    [key: string]: JExcelImage;
  };
  index: string;
  name: string;
  order: string;
  [key: string]: any;
};

declare function parseExcel(
  excel: File,
  cb: (
    data: null | {
      info: JExcelInfo;
      sheets: Array<JExcelSheets>;
    },
    err: undefined | any
  ) => any
): void;

declare class ExportExcel {
  constructor(companyName?: string);

  setCompanyName(name: string): void;
  addSheet(name: string): string;
  addContents(
    sheetName: string,
    contents: Array<
      Array<{
        value: ExcelJS.CellValue;
        options?: {
          width?: number;
          height?: number;
          font?: Partial<ExcelJS.Font>;
          alignment?: Partial<ExcelJS.Alignment>;
        };
      }>
    >
  ): void;
  addImagesAsync(
    sheetName: string,
    images: Array<{
      value: string;
      isBase64: boolean;
      r: number;
      c: number;
      offset?: number;
    }>
  ): Promise<any>;
  export(filename: string): Promise<any>
}
