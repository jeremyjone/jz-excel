# Jz-Excel

[[English](./README.md)] [[中文](./README_cn.md)]

[[GitHub](https://github.com/jeremyjone/jz-excel)]

这是基于 `Luckysheet` 开源框架的源码，解析 Excel 文件内容为 JSON 的一个插件，它可以同时按行和列分别解析文字内容和图片内容。

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
<script src="https://desktop.jeremyjone.com/resource/js/jz-excel.min.js"></script>
```

## 如何使用

它的使用非常简单。

```bash
import { parseExcel } from "jz-excel"
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

    // 注意: 为了不改变Lucky的数据结构，文本内容可能有两种格式，需要一些处理
    data.sheets[0].celldata.forEach(item => {
      // 提取所有内容
      let content = item?.v?.v ?? "";
      if (item?.v?.ct) {
        content = item?.v.ct.s.reduce((per, cur) => per + cur.v, "");
      }

      // 其他代码...
  } else {
    // 失败
  }
});
```

`data` 是 `object | null` 类型的，它包括了 `info` 和 `sheets`, 和 Luckysheet 是一样的，这给扩展留了空间。

### Typescript 支持

我写了一个简单的 TS 声明，把用到的字段都声明了。

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

## 特别感谢

[Luckysheet](https://github.com/mengshukeji/Luckysheet)

注：如有不当，请联系我删除。

## License

MIT
