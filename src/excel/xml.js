class XmlBase {
  getElementsByOneTag(tag, file) {
    let readTagReg;

    if (tag.indexOf("|") > -1) {
      const tags = tag.split("|");
      let tagsRegTxt = "";

      for (let i = 0; i < tags.length; i++) {
        const t = tags[i];
        tagsRegTxt +=
          "|<" +
          t +
          " [^>]+?[^/]>[\\s\\S]*?</" +
          t +
          ">|<" +
          t +
          " [^>]+?/>|<" +
          t +
          ">[\\s\\S]*?</" +
          t +
          ">|<" +
          t +
          "/>";
      }

      tagsRegTxt = tagsRegTxt.substr(1, tagsRegTxt.length);
      readTagReg = new RegExp(tagsRegTxt, "g");
    } else {
      readTagReg = new RegExp(
        "<" +
          tag +
          " [^>]+?[^/]>[\\s\\S]*?</" +
          tag +
          ">|<" +
          tag +
          " [^>]+?/>|<" +
          tag +
          ">[\\s\\S]*?</" +
          tag +
          ">|<" +
          tag +
          "/>",
        "g"
      );
    }

    const ret = file.match(readTagReg);

    if (ret == null) {
      return [];
    } else {
      return ret;
    }
  }
}

class Element extends XmlBase {
  constructor(str) {
    super();
    this.eleStr = str;
    this.setValue();

    const readAttrReg = new RegExp('[a-zA-Z0-9_:]*?=".*?"', "g");

    const attrList = this.container.match(readAttrReg);

    this.attributeList = {};

    if (attrList != null) {
      for (const key in attrList) {
        const attrFull = attrList[key]; // let al= attrFull.split("=");

        if (attrFull.length == 0) {
          continue;
        }

        const attrKey = attrFull.substr(0, attrFull.indexOf("="));
        const attrValue = attrFull.substr(attrFull.indexOf("=") + 1);

        if (
          attrKey == null ||
          attrValue == null ||
          attrKey.length == 0 ||
          attrValue.length == 0
        ) {
          continue;
        }

        this.attributeList[attrKey] = attrValue.substr(1, attrValue.length - 2);
      }
    }
  }

  get(name) {
    return this.attributeList[name];
  }

  getInnerElements(tag) {
    const ret = this.getElementsByOneTag(tag, this.eleStr);
    const elements = [];

    for (let i = 0; i < ret.length; i++) {
      const ele = new Element(ret[i]);
      elements.push(ele);
    }

    if (elements.length == 0) {
      return null;
    }

    return elements;
  }

  setValue() {
    const str = this.eleStr;

    if (str.substr(str.length - 2, 2) == "/>") {
      this.value = "";
      this.container = str;
    } else {
      const firstTag = this.getFirstTag();
      const firstTagReg = new RegExp(
        "(<" +
          firstTag +
          " [^>]+?[^/]>)([\\s\\S]*?)</" +
          firstTag +
          ">|(<" +
          firstTag +
          ">)([\\s\\S]*?)</" +
          firstTag +
          ">",
        "g"
      );
      const result = firstTagReg.exec(str);

      if (result != null) {
        if (result[1] != null) {
          this.container = result[1];
          this.value = result[2];
        } else {
          this.container = result[3];
          this.value = result[4];
        }
      }
    }
  }

  getFirstTag() {
    const str = this.eleStr;
    let firstTag = str.substr(0, str.indexOf(" "));

    if (firstTag == "" || firstTag.indexOf(">") > -1) {
      firstTag = str.substr(0, str.indexOf(">"));
    }

    firstTag = firstTag.substr(1, firstTag.length);
    return firstTag;
  }
}

export class ReadXml extends XmlBase {
  constructor(files) {
    super();
    this.originFile = files;
  }

  getElementsByTagName(path, fileName) {
    const file = this.getFileByName(fileName);
    const pathArr = path.split("/");
    let ret;

    for (const key in pathArr) {
      const path_1 = pathArr[key];

      if (ret == undefined) {
        ret = this.getElementsByOneTag(path_1, file);
      } else {
        if (ret instanceof Array) {
          let items = [];

          for (const key_1 in ret) {
            const item = ret[key_1];
            items = items.concat(this.getElementsByOneTag(path_1, item));
          }

          ret = items;
        } else {
          ret = this.getElementsByOneTag(path_1, ret);
        }
      }
    }

    const elements = [];

    for (let i = 0; i < ret.length; i++) {
      const ele = new Element(ret[i]);
      elements.push(ele);
    }

    return elements;
  }

  getFileByName(name) {
    for (const fileKey in this.originFile) {
      if (fileKey.indexOf(name) > -1) {
        return this.originFile[fileKey];
      }
    }

    return "";
  }
}
