const cheerio = require("cheerio");
const axios = require("axios");
const XLSX = require("xlsx");

const exelDataFilePath = `./Zipcodes_Germany.xlsx`;

const xlsxFile = "data-info.xlsx";
const xlsxSheetName = "zipDataInfo";
const baseUrl =
  "https://www.justizadressen.nrw.de/de/justiz/gericht?ang=grundbuch&plzort=$";

const startTime = Date.now();

function getTime() {
  const currentTime = Date.now();

  return (currentTime - startTime) / 1000;
}

let start = 0;

async function getGenre(zipCode) {
  const url = `${baseUrl}${zipCode}`;

  try {
    const response = await axios.get(url);
    const $ = cheerio.load(response.data);

    const infoBody = $("main .border.rounded.bg-white.p-3.mb-3");

    const variantList = infoBody.find("ul.list-unstyled:first-of-type")
      ? infoBody
          .find("ul.list-unstyled:first-of-type")
          .text()
          .split("\n")
          .filter((el) => el.trim() !== "")
          .map((el) => el.trim())
      : "";

    const buildData = (infoBody) => {
      const main_title = infoBody.find("h5.mt-3").text() || "undefined";
      const title = infoBody.find("h6").text() || "undefined";

      const transactions_list = infoBody.find("ul")
        ? infoBody
            .find("ul")
            .text()
            .split("\n")
            .filter((el) => el.trim() !== "")
            .map((el) => el.trim())
            .join("; ")
        : "undefined";

      const addressBody = infoBody.find(".row.clearfix")
        ? infoBody
            .find(".row.clearfix")
            .text()
            .split("\n")
            .filter((el) => el.trim() !== "")
            .map((el) => el.trim())
        : "undefined";

      const dataForCvs = {
        zip_code: `${zipCode}`,
        details: `${main_title ? main_title.split(",")[1] : ""}`,
        main_title: `${main_title}`,
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

      return dataForCvs;
    };

    if (variantList.length) {
      const variants = variantList.map((el) => el.split(" ")[1]);

      console.log("variants===", variants);

      const nextData = [];

      for (let k = 0; k < variants.length; k++) {
        // "https://www.justizadressen.nrw.de/de/justiz/gericht?ang=grundbuch&plz=08393&ort=Sch%C3%B6nberg";
        // https://www.justizadressen.nrw.de/de/justiz/gericht?ang=grundbuch&plzort=$08393&ort=Schönberg

        const nextUrl = `https://www.justizadressen.nrw.de/de/justiz/gericht?ang=grundbuch&plz=${zipCode}&ort=${variants[k]}`;
        // console.log(" \n\n nextUrl", nextUrl, "\n\n");

        const nextResponse = await axios.get(nextUrl);
        const $ = cheerio.load(nextResponse.data);

        const infoBody = $("main .border.rounded.bg-white.p-3.mb-3");

        const nextResData = buildData(infoBody);

        // console.log("nextResData===", nextResData);

        nextData.push(nextResData);
      }

      // const nextUrl = `${baseUrl}${zipCode}&ort=${}`;

      // const nextResponse = await axios.get(nextUrl);
      // const $ = cheerio.load(nextResponse.data);

      // const infoBody = $("main .border.rounded.bg-white.p-3.mb-3");

      return nextData;
    } else {
      console.log(`${zipCode} - completed !`, `${getTime()}sec`, `#${start++}`);
      // console.log(`dataForCvs===`, dataForCvs);

      return [buildData(infoBody)];
    }

    // console.log(`${zipCode} - completed !`);
    // // console.log(`dataForCvs===`, dataForCvs);

    // return dataForCvs;
  } catch (err) {
    console.log("error===", zipCode);
  }
}

// function readExelFile(path) {
//   const workbook = XLSX.readFile(path);
//   const worksheet = workbook.Sheets[workbook.SheetNames[0]];

//   console.log("worksheet.length===", worksheet.length);

//   const zipDataArr = [];

//   for (let i = 2; i <= 8175; i++) {
//     const zipCode = worksheet[`A${i}`].v;
//     zipDataArr.push(zipCode);
//   }

//   return zipDataArr;
// }

function readExelFile(path) {
  const workbook = XLSX.readFile(path);
  const worksheet = {};

  for (const sheetName of workbook.SheetNames) {
    worksheet[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  }

  const zipCodes = worksheet[Object.keys(worksheet)[0]].map(
    (el) => el[Object.keys(el)[0]]
  );

  return zipCodes;
}

async function getDataFromSrappin(exelDataFilePath) {
  const zipCodeArrData = readExelFile(exelDataFilePath);
  console.log("zipCodeArrData===", zipCodeArrData.length);

  // const zipCodeArrData = ["08393"];
  // const zipCodeArrData = ["08058"];

  createNewXLSXFile({ fileName: xlsxFile, sheetName: xlsxSheetName });

  console.log("start get scrapping data ...");

  // await Promise.all(
  //   zipCodeArrData.map(async (zipCode) => {
  //     const res = await getGenre(String(zipCode));

  //     updateXLSXFile({
  //       data: res,
  //       xlsxFile: xlsxFile,
  //       xlsxSheetName: xlsxSheetName,
  //     });

  //     // return res;
  //   })
  // );

  for (let i = 0; i < zipCodeArrData.length; i++) {
    const res = await getGenre(String(zipCodeArrData[i]));

    // console.log("res===", res);

    res.forEach((el) =>
      updateXLSXFile({
        data: el,
        xlsxFile: xlsxFile,
        xlsxSheetName: xlsxSheetName,
      })
    );

    // updateXLSXFile({
    //   data: res,
    //   xlsxFile: xlsxFile,
    //   xlsxSheetName: xlsxSheetName,
    // });
  }

  console.log("completed scrapping data !");
}

getDataFromSrappin(exelDataFilePath);

function createNewXLSXFile({
  data = [{}],
  sheetName = xlsxSheetName,
  fileName = xlsxFile,
}) {
  const newbook = XLSX.utils.book_new();
  const newSheet = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(newbook, newSheet, sheetName);
  XLSX.writeFile(newbook, fileName);
}

function updateXLSXFile({ data, xlsxFile, xlsxSheetName }) {
  const workbook = XLSX.readFile(xlsxFile);

  const worksheet = workbook.Sheets[xlsxSheetName];

  let worksheets = {};
  for (const sheetName of workbook.SheetNames) {
    worksheets[sheetName] = XLSX.utils.sheet_to_json(
      workbook.Sheets[sheetName]
    );
  }

  worksheets[xlsxSheetName].push(data);

  // const newData = workbook.Sheets[xlsxSheetName].push(data);

  XLSX.utils.sheet_add_json(worksheet, worksheets[xlsxSheetName]);
  // XLSX.utils.sheet_add_json(worksheet, newData);
  XLSX.writeFile(workbook, xlsxFile);
}
