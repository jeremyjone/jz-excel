# Jz-Excel

[[English](./README.md)] [[中文](./README_cn.md)]

[[GitHub](https://github.com/jeremyjone/jz-excel)]

## 一个 Excel ← 转换 → Json 的工具。

使用该工具，可以任意将 Excel 转换为 Json，同时可以将 Json 转换为 Excel。这对于导出页面中的表格数据有很好的效果。

- Excel 转换 Json：是对 `Luckysheet` 开源框架的整体简化解析工具。根据其框架的内容，将所有内容逐一整理并加以修改，创建了现在的解析工具。

- Json 转换 Excel：使用 `exceljs` 框架进行解析拼装。

### 特色

- 按行按列转换数据
- 可以转换图片
- 匹配 `Luckysheet` 格式

## 样例

[在线样例](https://desktop.jeremyjone.com/example/jz-excel.html)

## 安装

```bash
npm i jz-excel
```

或者：v

```bash
yarn add jz-excel
```

## 直接引入

```html
<script src="https://cdn.jsdelivr.net/npm/jz-excel@1.0.4/dist/index.min.js"></script>
```

## 如何使用

它的使用非常简单。

### 解析 Excel

```js
import { parseExcel } from "jz-excel";
```

然后直接使用这个函数即可：

```js
/**
 * @param {File} excel
 * @param {(data: null | object, err: undefined | any) => any} cb
 */
parseExcel(excel, function (data, err) {
  if (data) {
    // 成功

    // 注意: 为了不改变 Luckysheet 的数据结构，文本内容可能有两种格式，需要一些处理
    data.sheets[0].celldata.forEach(item => {
      // 提取所有内容
      let content = item?.v?.v ?? "";
      if (item?.v?.ct) {
        content = item?.v.ct.s.reduce((per, cur) => per + cur.v, "");
      }

      // 其他代码...
    });
  } else {
    // 失败
  }
});
```

`data` 是 `object | null` 类型的，它包括了 `info` 和 `sheets`, 和 Luckysheet 是一样的，这给扩展留了空间。

### 导出 Excel

```js
import { ExportExcel } from "jz-excel";
```

它是一个类，按照顺序依次添加 `工作表`、`内容`、`图片`(如果有)，然后就可以导出了。

```js
var excel = new JzExcel.ExportExcel("my company");

var sheetName = "mySheet";
excel.addSheet(sheetName);
excel.addContents(sheetName, data);
excel.addImagesAsync(sheetName, images).then(() => {
  excel.export("my-file");
});
```

至于数据格式，参照对应的 `TypeScript` 即可。

### TypeScript 支持

我写了一个简单的 TS 声明，把用到的字段都声明了。

解析方法声明：

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

导出类声明：

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

## 特别感谢

[Luckysheet](https://github.com/mengshukeji/Luckysheet)

注：如有不当，请联系我删除。

## License

MIT
