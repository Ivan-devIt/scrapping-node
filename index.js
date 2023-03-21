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

  return ((currentTime - startTime) / 1000).toFixed(1);
}

let start = 0;

async function getGenre(zipCode) {
  const url = `${baseUrl}${zipCode}`;

  try {
    const response = await axios.get(url);
    const $ = cheerio.load(response.data);

    const infoBody = $("main .border.rounded.bg-white.p-3.mb-3");

    function getNextLinks(infoBody) {
      const links = [];

      infoBody
        .find("ul.list-unstyled:first-of-type li a")
        .each((index, value) => {
          const link = $(value).attr("href");
          links.push(link);
        });

      return links;
    }

    const linksLinks = getNextLinks(infoBody);
    console.log("linksLinks===", linksLinks);

    const buildData = (infoBody) => {
      const details = infoBody.find("h5.mt-3").text() || "undefined";
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

      const addressData = [];

      const newAdress = infoBody
        .find(".row.clearfix div address")
        .each((index, value) => {
          addressData.push(
            $(value)
              .text()
              .split("\n")
              .filter((el) => el.trim() !== "")
              .map((el) => el.trim())
          );
        });

      const dataForCvs = {
        zip_code: `${zipCode}`,
        details: `${details}`,
        title: `${title}`,
        delivery_address: `${addressData[0][1]}, ${addressData[0][2]}`,
        postal_address: `${addressData[1][0]},${addressData[1][1]}`,
        phone: `${addressData[2][1]}`,
        fax: `${addressData[2][2]}`,
        internet: `${addressData[2][3]}`,
        email: `${addressData[2][4]}`,
        x_justiz_id: `${addressData[2][5]}`,
        transactions_title: `${addressBody[12]}`,
        transactions_list: `${transactions_list}`,
      };

      return dataForCvs;
    };

    if (linksLinks.length) {
      const nextData = [];

      for (let k = 0; k < linksLinks.length; k++) {
        const nextUrl = `${linksLinks[k]}`;
        const nextResponse = await axios.get(nextUrl);
        const $ = cheerio.load(nextResponse.data);
        const infoBody = $("main .border.rounded.bg-white.p-3.mb-3");
        const nextResData = buildData(infoBody);

        nextData.push(nextResData);
      }

      return nextData;
    } else {
      console.log(`${zipCode} - completed !`, `${getTime()}sec`, `#${start++}`);
      // console.log(`dataForCvs===`, dataForCvs);

      return [buildData(infoBody)];
    }
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

async function getDataFromSrappin(exelDataFilePath) {
  // const zipCodeArrData = readExelFile(exelDataFilePath);
  const zipCodeArrData = ["16259"];
  // const zipCodeArrData = ["99958"];
  // const zipCodeArrData = ["22297"];

  // console.log("zipCodeArrData===", zipCodeArrData.length);

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

    console.log("res===", res);

    // TODO add check
    res.forEach((el) =>
      updateXLSXFile({
        data: el,
        xlsxFile: xlsxFile,
        xlsxSheetName: xlsxSheetName,
      })
    );
  }

  console.log("completed scrapping data !");
}

getDataFromSrappin(exelDataFilePath);

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
