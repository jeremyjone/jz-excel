# Jz-Excel

[[English](./README.md)] [[中文](./README_cn.md)]

[[GitHub](https://github.com/jeremyjone/jz-excel)]

## A Excel ← Transform → Json tool.

Using this tool, you can convert Excel to Json, and also can convert Json to Excel. This works well for exporting tabular data from the page.

- Excel to Json： is the overall simplified parsing tool for the `LuckySheet` open source framework. According to the content of its framework, all the contents are sorted out and modified one by one to create the present parsing tool.

- Json to Excel：Use the `exceljs` framework for parsing assembly.

### Features

- Convert data by row and by column
- You can convert images
- Matches the `Luckysheet` format

## Demo

[Demo](https://desktop.jeremyjone.com/example/jz-excel.html)

## Install

```bash
npm i jz-excel
```

or

```bash
yarn add jz-excel
```

## Link

```html
<script src="https://cdn.jsdelivr.net/npm/jz-excel@1.0.2/dist/index.min.js"></script>
```

## Use

It very easy to use.

### Parse Excel

```js
import { parseExcel } from "jz-excel"
```

and use the function:

```js
/**
 * @param {File} excel
 * @param {(data: null | object, err: undefined | any) => any} cb
 */
parseExcel(excel, function (data, err) {
  if (data) {
    // success

    // NOTICE: In order not to change Luckysheet's data structure, the text content may come in two formats that require some processing
    data.sheets[0].celldata.forEach(item => {
      // extract all the content
      let content = item?.v?.v ?? "";
      if (item?.v?.ct) {
        content = item?.v.ct.s.reduce((per, cur) => per + cur.v, "");
      }

      // other code...
    });
  } else {
    // failure
  }
});
```

`data` is `object | null`, includes `info` and `sheets`, same as Luckysheet.

### Export Excel

```js
import { ExportExcel } from "jz-excel";
```

It's a `class` that adds `worksheet`, `contents`, `images`(if any) in that order, and then it's ready to export.

```js
var excel = new JzExcel.ExportExcel("my company");

var sheetName = "mySheet";
excel.addSheet(sheetName);
excel.addContents(sheetName, data);
excel.addImagesAsync(sheetName, images).then(() => {
  excel.export("my-file");
});
```

For data formats, refer to the corresponding `TypeScript`.

### Typescript support

I wrote a simple type support.

Parse method declaration:

```ts
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
```

Export class declaration:

```ts
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
  export(filename: string): Promise<any>;
}
```

## Special

[Luckysheet](https://github.com/mengshukeji/Luckysheet)

Note: if there is any improper, please contact me to delete.

## License

MIT
