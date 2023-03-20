const cheerio = require("cheerio");
const axios = require("axios");
const j2cp = require("json2csv").Parser;
const fs = require("fs");
const XLSX = require("xlsx");

const exelDataFilePath = `./Zipcodes_Germany.xlsx`;

function friteCSVdoc(data) {
  const parser = new j2cp();
  const csv = parser.parse(data);
  fs.writeFileSync("./books.csv", csv);
}

//03119

async function getGenre(zipCode) {
  const url = `https://www.justizadressen.nrw.de/de/justiz/gericht?ang=grundbuch&plzort=${zipCode}`;

  try {
    const response = await axios.get(url);
    const $ = cheerio.load(response.data);

    const infoBody = $("main .border.rounded.bg-white.p-3.mb-3");

    const title = infoBody.find("h5.mt-3").text();
    const title2 = infoBody.find("h6").text();
    const transactions_list = infoBody
      .find("ul")
      .text()
      .split("\n")
      .filter((el) => el.trim() !== "")
      .map((el) => el.trim())
      .join("; ");

    const addressBody = infoBody
      .find(".row.clearfix")
      .text()
      .split("\n")
      .filter((el) => el.trim() !== "")
      .map((el) => el.trim());

    const dataForCvs = {
      zip_code: `${zipCode}`,
      title: `${title}`,
      sub_title: `${addressBody[0]}`,
      delivery_address: `${addressBody[0]},${addressBody[1]}, ${addressBody[2]}`,
      postal_address: `${addressBody[4]},${addressBody[5]}`,
      contact_title: `${addressBody[6]}`,
      phone: `${addressBody[7]}`,
      fax: `${addressBody[8]}`,
      internet: `${addressBody[9]}`,
      email: `${addressBody[10]}`,
      x_justiz_id: `${addressBody[11]}`,
      transactions_title: `${addressBody[12]}`,
      transactions_list: `${transactions_list}`,
    };

    console.log(`${zipCode} - completed !`);

    return dataForCvs;
    // return data;
    // return zipCode;
  } catch (err) {
    console.log("error===", zipCode);
    // throw new Error(err);
  }
}

// TODO first variant
function readExelFile(path) {
  const workbook = XLSX.readFile(path);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];

  const zipDataArr = [];

  for (let i = 2; i <= 100; i++) {
    const zipCode = worksheet[`A${i}`].v;
    zipDataArr.push(zipCode);
  }

  // console.log(zipDataArr);
  // return ["97769"];
  return zipDataArr;
}

//second variant
// function readExelFile(path) {
//   const workbook = XLSX.readFile(path);
//   const worksheet = {};

//   for (const sheetName of workbook.SheetNames) {
//     worksheet[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
//   }

//   const zipCodes = worksheet[Object.keys(worksheet)[0]].map(
//     (el) => el[Object.keys(el)[0]]
//   );

//   return zipCodes;
// }
// console.log(readExelFile(exelDataFilePath));
// console.log(readExelFile(exelDataFilePath));

async function getDataFromSrappin(exelDataFilePath) {
  const zipCodeArrData = readExelFile(exelDataFilePath);
  // console.log("zipCodeArrData===", zipCodeArrData);

  console.log("start get scrapping data ...");

  return await Promise.all(
    zipCodeArrData.map(async (zipCode) => {
      const res = await getGenre(String(zipCode));

      // console.log(res);

      return res;
    })
  );
}
///**** */
const dataFromSrappin = getDataFromSrappin(exelDataFilePath);

dataFromSrappin.then((data) => {
  createNewXLSXFile(data);
  console.log("completed get scrapping data !");
  // console.log(data);
});
//** */

// function addNewSheet =

// console.log(getDataFromSrappin(exelDataFilePath));

// const res = getGenre("09518").then((result) => {
//   console.log(result);

//   return result;
// });

// console.log(res);

//////////////////////////////////////////////////////////////////////////////////

// const newBook = XLSX.utils.book_new("ssss");

// const DATA = [
//   {
//     zip_code: "01594",
//     title: "Gesucht wurde: Grundbuchsachen, 01594 ",
//     sub_title: "Amtsgericht Riesa - Grundbuchamt -",
//     delivery_address: "Lieferanschrift,Lauchhammerstraße 10, 01591 Riesa",
//     postal_address: "Lauchhammerstraße 10,01591 Riesa",
//     contact_title: "Kontakt",
//     phone: "Telefon: 03525 745-10",
//     fax: "Fax: 03525 745-111",
//     internet: "Internet: http://www.justiz.sachsen.de/agrie",
//     x_justiz_id: "XJustiz-ID: U1115G",
//     field: "Der elektronische Rechtsverkehr ist zugelassen.",
//   },
//   {
//     zip_code: "02899",
//     title: "Gesucht wurde: Grundbuchsachen, 02899 ",
//     sub_title: "Amtsgericht Zittau - Grundbuchamt -",
//     delivery_address: "Lieferanschrift,Lessingstraße 1, 02763 Zittau",
//     postal_address: "Postfach 15 55,02755 Zittau",
//     contact_title: "Kontakt",
//     phone: "Telefon: 03583 759-0",
//     fax: "Fax: 03583 759-1030",
//     internet: "Internet: http://www.justiz.sachsen.de/agzi",
//     x_justiz_id: "XJustiz-ID: U1118G",
//     field: "Der elektronische Rechtsverkehr ist zugelassen.",
//   },
// ];

// const newbook = XLSX.utils.book_new();
// const newSheet = XLSX.utils.json_to_sheet(DATA);
// XLSX.utils.book_append_sheet(newbook, newSheet, "zipDataInfo");
// XLSX.writeFile(newbook, "data-info.xlsx");

function createNewXLSXFile(
  data,
  sheetName = "zipDataInfo",
  fileName = "data-info.xlsx"
) {
  const newbook = XLSX.utils.book_new();
  const newSheet = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(newbook, newSheet, sheetName);
  XLSX.writeFile(newbook, fileName);
}

// createNewXLSXFile(DATA);
