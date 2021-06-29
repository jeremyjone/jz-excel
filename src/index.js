import { ExportExcel as Export } from "./export/export";
import { parseExcel as parse } from "./excel/index"

export const parseExcel = parse;
export const ExportExcel = Export;

export default {
  parseExcel: parse,
  ExportExcel: Export
};
