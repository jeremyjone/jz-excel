# Jz-Excel

[[English](./README.md)] [[中文](./README_cn.md)]

[[GitHub](https://github.com/jeremyjone/jz-excel)]

Parse excel file(only .xlsx) to JSON based on LuckySheet code. It parses everything, including images, by row and column.

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
<script src="https://desktop.jeremyjone.com/resource/js/jz-excel.min.js"></script>
```

## Use

It very easy to use.

```bash
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

### Typescript support

I wrote a simple type support.

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

## Special

[Luckysheet](https://github.com/mengshukeji/Luckysheet)

Note: if there is any improper, please contact me to delete.

## License

MIT
