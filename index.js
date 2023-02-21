import * as fs from "fs";
import XLSX, { readFile, set_fs } from "xlsx";
set_fs(fs);

const workbook = readFile("teaor08_tartalom_2021_06_01.xlsx");
const wsnames = workbook.SheetNames;

if (wsnames.length === 0) {
  throw new Error("No sheets found in workbook");
}
const worksheet = workbook.Sheets[wsnames[0]];

const teorData = XLSX.utils.sheet_to_json(worksheet, { header: "A" });
//generate JSON from the sheet loopin through the rows

let sector = ""; //szakágazat
let category = ""; //ágazat
let subcategory = "";

let sectorObject = {
  code: "",
  title: "",
  codeAndTitle: "",
  description: "",
  inThisSector: "",
  inThisSectorMore: "",
  notInThisSector: "",
  notInThisSectorMore: "",
};
let categoryObject = {
  code: "",
  title: "",
  codeAndTitle: "",
  description: "",
  inThisCategory: "",
  inThisCategoryMore: "",
  notInThisCategory: "",
  notInThisCategoryMore: "",
};
let subcategoryObject = {
  code: "",
  title: "",
  codeAndTitle: "",
  description: "",
  inThisSubcategory: "",
  inThisSubcategoryMore: "",
  notInThisSubcategory: "",
  notInThisSubcategoryMore: "",
};

const activies = [];
const activies_slim = [];
const categories = {};

for (const item of teorData) {
  const code = item.A;
  switch (code.length) {
    case 1:
      sector = code;
      sectorObject = {
        code: code,
        title: item.B,
        codeAndTitle: code + " " + item.B,
        description: item.C,
        inThisSector: item.D,
        inThisSectorMore: item.E,
        notInThisSector: item.F,
      };
      categories[code] = sectorObject;
      break;
    case 2:
      category = code;
      categoryObject = {
        code: code,
        title: item.B,
        codeAndTitle: code + " " + item.B,
        description: item.C,
        inThisCategory: item.D,
        inThisCategoryMore: item.E,
        notInThisCategory: item.F,
      };
      categories[code] = categoryObject;

      break;
    case 4:
      subcategory = code;
      subcategoryObject = {
        code: code,
        title: item.B,
        codeAndTitle: code + " " + item.B,
        description: item.C,
        inThisSubcategory: item.D,
        inThisSubcategoryMore: item.E,
        notInThisSubcategory: item.F,
      };
      categories[code] = subcategoryObject;
      break;
    case 5:
      activies.push({
        code: code,
        title: item.B,
        codeAndTitle: code + " " + item.B,
        description: item.C,
        inThisActivity: item.D,
        inThisActivityMore: item.E,
        notInThisActivity: item.F,
        sectorID: sector,
        categoryID: category,
        subcategoryID: subcategory,
        sectorObject: sectorObject,
        categoryObject: categoryObject,
        subcategoryObject: subcategoryObject,
      });
      activies_slim.push({
        code: code,
        title: item.B,
        codeAndTitle: code + " " + item.B,
        description: item.C,
        inThisActivity: item.D,
        inThisActivityMore: item.E,
        notInThisActivity: item.F,
        sectorID: sector,
        categoryID: category,
        subcategoryID: subcategory,
      });
      break;
    default:
      break;
  }
}
fs.writeFileSync("teaor_slim.json", JSON.stringify(activies_slim, null, 2));
fs.writeFileSync("teaor.json", JSON.stringify(activies, null, 2));
fs.writeFileSync("teaor_categories.json", JSON.stringify(categories, null, 2));
