import {
  getXmlAttibute,
  getcellrange,
  getPxByEMUs,
  generateRandomIndex,
  escapeCharacter,
  getColor,
  isChinese,
  isJapanese,
  isKoera,
  getlineStringAttr
} from "../tools/func";

const ST_CellType = {
  Boolean: "b",
  Date: "d",
  Error: "e",
  InlineString: "inlineStr",
  Number: "n",
  SharedString: "s",
  String: "str"
};

export class Sheet {
  constructor(name, id, order, options) {
    this.xml = options.xml;
    this.sheetFile = options.sheetFile;
    this.styles = options.styles;
    this.sharedStrings = options.sharedStrings;
    this.sheetList = options.sheetList;
    this.imageList = options.imageList;

    this.name = name;
    this.index = id;
    this.order = order.toString();
    this.config = {};
    this.celldata = [];
    this.mergeCells = this.xml.getElementsByTagName(
      "mergeCells/mergeCell",
      this.sheetFile
    );

    this.generateConfigRowLenAndHiddenAddCell();

    const drawingFile = options.drawingFile,
      drawingRelsFile = options.drawingRelsFile;

    if (drawingFile != null && drawingRelsFile != null) {
      const oneCellAnchors = this.xml.getElementsByTagName(
        "xdr:oneCellAnchor",
        drawingFile
      );

      const twoCellAnchors = this.xml.getElementsByTagName(
        "xdr:twoCellAnchor",
        drawingFile
      );

      let cellAnchors,
        finishPosTag,
        isOneCellAnchor = false;

      if (oneCellAnchors != null && oneCellAnchors.length > 0) {
        cellAnchors = oneCellAnchors;
        finishPosTag = "xdr:ext";
        isOneCellAnchor = true;
      } else if (twoCellAnchors != null && twoCellAnchors.length > 0) {
        cellAnchors = twoCellAnchors;
        finishPosTag = "xdr:to";
      }

      if (cellAnchors.length > 0) {
        for (let i = 0; i < cellAnchors.length; i++) {
          const cellAnchor = cellAnchors[i];
          const editAs = getXmlAttibute(
            cellAnchor.attributeList,
            "editAs",
            "oneCell"
          );

          const xdrFroms = cellAnchor.getInnerElements("xdr:from"),
            xdrTos = cellAnchor.getInnerElements(finishPosTag),
            xdr_blipfills = cellAnchor.getInnerElements("a:blip");

          if (
            xdrFroms != null &&
            xdr_blipfills != null &&
            xdrFroms.length > 0 &&
            xdr_blipfills.length > 0
          ) {
            const xdrFrom = xdrFroms[0],
              xdrTo = xdrTos[0],
              xdr_blipfill = xdr_blipfills[0];
            const rembed = getXmlAttibute(
              xdr_blipfill.attributeList,
              "r:embed",
              null
            );

            const imageObject = this.getBase64ByRid(rembed, drawingRelsFile); // let aoff = xdr_xfrm.getInnerElements("a:off"), aext = xdr_xfrm.getInnerElements("a:ext");

            let x_n = 0,
              y_n = 0,
              cx_n = 0,
              cy_n = 0;
            imageObject.fromCol = this.getXdrValue(
              xdrFrom.getInnerElements("xdr:col")
            );
            imageObject.fromColOff = getPxByEMUs(
              this.getXdrValue(xdrFrom.getInnerElements("xdr:colOff"))
            );
            imageObject.fromRow = this.getXdrValue(
              xdrFrom.getInnerElements("xdr:row")
            );
            imageObject.fromRowOff = getPxByEMUs(
              this.getXdrValue(xdrFrom.getInnerElements("xdr:rowOff"))
            );
            imageObject.toCol = isOneCellAnchor
              ? imageObject.fromCol
              : this.getXdrValue(xdrTo.getInnerElements("xdr:col"));
            imageObject.toColOff = getPxByEMUs(
              this.getXdrValue(xdrTo.getInnerElements("xdr:colOff"))
            );
            imageObject.toRow = isOneCellAnchor
              ? imageObject.fromRow
              : this.getXdrValue(xdrTo.getInnerElements("xdr:row"));
            imageObject.toRowOff = getPxByEMUs(
              this.getXdrValue(xdrTo.getInnerElements("xdr:rowOff"))
            );
            imageObject.originWidth = cx_n;
            imageObject.originHeight = cy_n;

            if (editAs == "absolute") {
              imageObject.type = "3";
            } else if (editAs == "oneCell") {
              imageObject.type = "2";
            } else {
              imageObject.type = "1";
            }

            imageObject.isFixedPos = false;
            imageObject.fixedLeft = 0;
            imageObject.fixedTop = 0;
            const imageBorder = {
              color: "#000",
              radius: 0,
              style: "solid",
              width: 0
            };
            imageObject.border = imageBorder;
            const imageCrop = {
              height: cy_n,
              offsetLeft: 0,
              offsetTop: 0,
              width: cx_n
            };
            imageObject.crop = imageCrop;
            const imageDefault = {
              height: cy_n,
              left: x_n,
              top: y_n,
              width: cx_n
            };
            imageObject["default"] = imageDefault;

            if (this.images == null) {
              this.images = {};
            }

            this.images[generateRandomIndex("image")] = imageObject;
          }
        }
      }
    }
  }

  getBase64ByRid(rid, drawingRelsFile) {
    const Relationships = this.xml.getElementsByTagName(
      "Relationships/Relationship",
      drawingRelsFile
    );

    if (Relationships != null && Relationships.length > 0) {
      for (let i = 0; i < Relationships.length; i++) {
        const Relationship = Relationships[i];
        const attrList = Relationship.attributeList;
        const Id = getXmlAttibute(attrList, "Id", null);
        let src = getXmlAttibute(attrList, "Target", null);

        if (Id == rid) {
          src = src.replace(/\.\.\//g, "");
          src = "xl/" + src;
          return this.imageList.getImageByName(src);
        }
      }
    }

    return null;
  }

  getXdrValue(ele) {
    if (ele == null || ele.length == 0) {
      return null;
    }

    return parseInt(ele[0].value);
  }

  generateConfigRowLenAndHiddenAddCell() {
    const rows = this.xml.getElementsByTagName("sheetData/row", this.sheetFile);

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i],
        attrList = row.attributeList;
      const rowNo = getXmlAttibute(attrList, "r", null);

      if (rowNo == null) {
        continue;
      }

      const cells = row.getInnerElements("c");

      for (const key in cells) {
        const cell = cells[key];
        const cellValue = new SheetCelldata(
          cell,
          this.styles,
          this.sharedStrings,
          this.mergeCells,
          this.sheetFile,
          this.xml
        );

        if (cellValue._borderObject != null) {
          delete cellValue._borderObject;
        }

        this.celldata.push(cellValue);
      }
    }
  }
}

class SheetCelldata {
  constructor(cell, styles, sharedStrings, mergeCells, sheetFile, xml) {
    this.cell = cell;
    this.sheetFile = sheetFile;
    this.styles = styles;
    this.sharedStrings = sharedStrings;
    this.xml = xml;
    this.mergeCells = mergeCells;
    const attrList = cell.attributeList;
    const r = attrList.r,
      s = attrList.s,
      t = attrList.t;
    const range = getcellrange(r);
    this.r = range.row[0];
    this.c = range.column[0];
    this.v = this.generateValue(s, t);
  }

  generateValue(s, t) {
    let v = this.cell.getInnerElements("v");
    const f = this.cell.getInnerElements("f");

    if (v == null) {
      v = this.cell.getInnerElements("t");
    }

    const sharedStrings = this.sharedStrings;
    const cellValue = {};

    if (f != null) {
      const formula = f[0],
        attrList = formula.attributeList;
      const t_1 = attrList.t,
        ref = attrList.ref,
        si = attrList.si;
      let formulaValue = f[0].value;

      if (t_1 == "shared") {
        this._fomulaRef = ref;
        this._formulaType = t_1;
        this._formulaSi = si;
      }

      if (ref != null || (formulaValue != null && formulaValue.length > 0)) {
        formulaValue = escapeCharacter(formulaValue);
        cellValue.f = "=" + formulaValue;
      }
    }

    let familyFont = null;
    let quotePrefix;

    if (v != null) {
      let value = v[0].value;

      if (/&#\d+;/.test(value)) {
        value = this.htmlDecode(value);
      }

      if (t == ST_CellType["SharedString"]) {
        const siIndex = parseInt(v[0].value);
        const sharedSI = sharedStrings[siIndex];
        const rFlag = sharedSI.getInnerElements("r");

        if (rFlag == null) {
          const tFlag = sharedSI.getInnerElements("t");

          if (tFlag != null) {
            let text_1 = "";
            tFlag.forEach(t => {
              text_1 += t.value;
            });
            text_1 = escapeCharacter(text_1); //isContainMultiType(text) &&

            if (familyFont == "Roman" && text_1.length > 0) {
              const textArray = text_1.split("");
              let preWordType = null,
                wordText = "",
                preWholef = null;
              let wholef = "Times New Roman";

              if (cellValue.ff != null) {
                wholef = cellValue.ff;
              }

              let cellFormat = cellValue.ct;

              if (cellFormat == null) {
                cellFormat = {};
              }

              if (cellFormat.s == null) {
                cellFormat.s = [];
              }

              for (let i = 0; i < textArray.length; i++) {
                const w = textArray[i];
                let type = null,
                  ff = wholef;

                if (isChinese(w)) {
                  type = "c";
                  ff = "宋体";
                } else if (isJapanese(w)) {
                  type = "j";
                  ff = "Yu Gothic";
                } else if (isKoera(w)) {
                  type = "k";
                  ff = "Malgun Gothic";
                } else {
                  type = "e";
                }

                if (
                  (type != preWordType && preWordType != null) ||
                  i == textArray.length - 1
                ) {
                  const InlineString = {};
                  InlineString.ff = preWholef;

                  if (cellValue.fc != null) {
                    InlineString.fc = cellValue.fc;
                  }

                  if (cellValue.fs != null) {
                    InlineString.fs = cellValue.fs;
                  }

                  if (cellValue.cl != null) {
                    InlineString.cl = cellValue.cl;
                  }

                  if (cellValue.un != null) {
                    InlineString.un = cellValue.un;
                  }

                  if (cellValue.bl != null) {
                    InlineString.bl = cellValue.bl;
                  }

                  if (cellValue.it != null) {
                    InlineString.it = cellValue.it;
                  }

                  if (i == textArray.length - 1) {
                    if (type == preWordType) {
                      InlineString.ff = ff;
                      InlineString.v = wordText + w;
                    } else {
                      InlineString.ff = preWholef;
                      InlineString.v = wordText;
                      cellFormat.s.push(InlineString);
                      const InlineStringLast = {};
                      InlineStringLast.ff = ff;
                      InlineStringLast.v = w;

                      if (cellValue.fc != null) {
                        InlineStringLast.fc = cellValue.fc;
                      }

                      if (cellValue.fs != null) {
                        InlineStringLast.fs = cellValue.fs;
                      }

                      if (cellValue.cl != null) {
                        InlineStringLast.cl = cellValue.cl;
                      }

                      if (cellValue.un != null) {
                        InlineStringLast.un = cellValue.un;
                      }

                      if (cellValue.bl != null) {
                        InlineStringLast.bl = cellValue.bl;
                      }

                      if (cellValue.it != null) {
                        InlineStringLast.it = cellValue.it;
                      }

                      cellFormat.s.push(InlineStringLast);
                      break;
                    }
                  } else {
                    InlineString.v = wordText;
                  }

                  cellFormat.s.push(InlineString);
                  wordText = w;
                } else {
                  wordText += w;
                }

                preWordType = type;
                preWholef = ff;
              }

              cellFormat.t = "inlineStr"; // cellFormat.s = [InlineString];

              cellValue.ct = cellFormat; // console.log(cellValue);
            } else {
              text_1 = this.replaceSpecialWrap(text_1);

              if (text_1.indexOf("\r\n") > -1 || text_1.indexOf("\n") > -1) {
                const InlineString = {};
                InlineString.v = text_1;
                let cellFormat = cellValue.ct;

                if (cellFormat == null) {
                  cellFormat = {};
                }

                if (cellValue.ff != null) {
                  InlineString.ff = cellValue.ff;
                }

                if (cellValue.fc != null) {
                  InlineString.fc = cellValue.fc;
                }

                if (cellValue.fs != null) {
                  InlineString.fs = cellValue.fs;
                }

                if (cellValue.cl != null) {
                  InlineString.cl = cellValue.cl;
                }

                if (cellValue.un != null) {
                  InlineString.un = cellValue.un;
                }

                if (cellValue.bl != null) {
                  InlineString.bl = cellValue.bl;
                }

                if (cellValue.it != null) {
                  InlineString.it = cellValue.it;
                }

                cellFormat.t = "inlineStr";
                cellFormat.s = [InlineString];
                cellValue.ct = cellFormat;
              } else {
                cellValue.v = text_1;
                quotePrefix = "1";
              }
            }
          }
        } else {
          const styles_1 = [];
          rFlag.forEach(r => {
            const tFlag = r.getInnerElements("t");
            const rPr = r.getInnerElements("rPr");
            const InlineString = {};

            if (tFlag != null && tFlag.length > 0) {
              let text = tFlag[0].value;
              text = this.replaceSpecialWrap(text);
              text = escapeCharacter(text);
              InlineString.v = text;
            }

            if (rPr != null && rPr.length > 0) {
              const frpr = rPr[0];
              const sz = getlineStringAttr(frpr, "sz"),
                rFont = getlineStringAttr(frpr, "rFont"),
                // family = getlineStringAttr(frpr, "family"),
                // charset = getlineStringAttr(frpr, "charset"),
                // scheme = getlineStringAttr(frpr, "scheme"),
                b = getlineStringAttr(frpr, "b"),
                i = getlineStringAttr(frpr, "i"),
                u = getlineStringAttr(frpr, "u"),
                strike = getlineStringAttr(frpr, "strike"),
                vertAlign = getlineStringAttr(frpr, "vertAlign");
              const cEle = frpr.getInnerElements("color");
              let color = void 0;

              if (cEle != null && cEle.length > 0) {
                color = getColor(cEle[0], this.styles, "t");
              }

              let ff = void 0;

              if (rFont != null) {
                ff = rFont;
              }

              if (ff != null) {
                InlineString.ff = ff;
              } else if (cellValue.ff != null) {
                InlineString.ff = cellValue.ff;
              }

              if (color != null) {
                InlineString.fc = color;
              } else if (cellValue.fc != null) {
                InlineString.fc = cellValue.fc;
              }

              if (sz != null) {
                InlineString.fs = parseInt(sz);
              } else if (cellValue.fs != null) {
                InlineString.fs = cellValue.fs;
              }

              if (strike != null) {
                InlineString.cl = parseInt(strike);
              } else if (cellValue.cl != null) {
                InlineString.cl = cellValue.cl;
              }

              if (u != null) {
                InlineString.un = parseInt(u);
              } else if (cellValue.un != null) {
                InlineString.un = cellValue.un;
              }

              if (b != null) {
                InlineString.bl = parseInt(b);
              } else if (cellValue.bl != null) {
                InlineString.bl = cellValue.bl;
              }

              if (i != null) {
                InlineString.it = parseInt(i);
              } else if (cellValue.it != null) {
                InlineString.it = cellValue.it;
              }

              if (vertAlign != null) {
                InlineString.va = parseInt(vertAlign);
              } // ff:string | undefined //font family
              // fc:string | undefined//font color
              // fs:number | undefined//font size
              // cl:number | undefined//strike
              // un:number | undefined//underline
              // bl:number | undefined//blod
              // it:number | undefined//italic
              // v:string | undefined
            } else {
              if (InlineString.ff == null && cellValue.ff != null) {
                InlineString.ff = cellValue.ff;
              }

              if (InlineString.fc == null && cellValue.fc != null) {
                InlineString.fc = cellValue.fc;
              }

              if (InlineString.fs == null && cellValue.fs != null) {
                InlineString.fs = cellValue.fs;
              }

              if (InlineString.cl == null && cellValue.cl != null) {
                InlineString.cl = cellValue.cl;
              }

              if (InlineString.un == null && cellValue.un != null) {
                InlineString.un = cellValue.un;
              }

              if (InlineString.bl == null && cellValue.bl != null) {
                InlineString.bl = cellValue.bl;
              }

              if (InlineString.it == null && cellValue.it != null) {
                InlineString.it = cellValue.it;
              }
            }

            styles_1.push(InlineString);
          });
          let cellFormat = cellValue.ct;

          if (cellFormat == null) {
            cellFormat = {};
          }

          cellFormat.t = "inlineStr";
          cellFormat.s = styles_1;
          cellValue.ct = cellFormat;
        }
      } // else if(t==ST_CellType["InlineString"] && v!=null){
      // }
      else {
        value = escapeCharacter(value);
        cellValue.v = value;
      }
    }

    if (quotePrefix != null) {
      cellValue.qp = parseInt(quotePrefix);
    }

    return cellValue;
  }

  htmlDecode(str) {
    return str.replace(/&#(x)?([^&]{1,5});?/g, function ($, $1, $2) {
      return String.fromCharCode(parseInt($2, $1 ? 16 : 10));
    });
  }

  replaceSpecialWrap(text) {
    return text
      .replace(/_x000D_/g, "")
      .replace(/&#13;&#10;/g, "\r\n")
      .replace(/&#13;/g, "\r")
      .replace(/&#10;/g, "\n");
  }
}
