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

    const linksData = getNextLinks(infoBody);
    // console.log("linksData===", linksData);

    const buildData = (infoBody) => {
      const details = infoBody.find("h5.mt-3").text() || " ";
      const title = infoBody.find("h6").text() || " ";

      const transactions_list = infoBody.find("ul")
        ? infoBody
            .find("ul")
            .text()
            .split("\n")
            .filter((el) => el.trim() !== "")
            .map((el) => el.trim())
            .join("; ")
        : " ";

      const addressBody = infoBody.find(".row.clearfix")
        ? infoBody
            .find(".row.clearfix")
            .text()
            .split("\n")
            .filter((el) => el.trim() !== "")
            .map((el) => el.trim())
        : " ";

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

      if (!addressData.length) {
        return [];
      }

      const dataForCvs = {
        zip_code: `${zipCode}`,
        details: `${details}`,
        title: `${title}`,
        delivery_address: `${addressData[0][1]}, ${addressData[0][2]}`,
        // postal_address: `${addressData[1][1]},${addressData[1][2]}`,
        postal_address: `${addressData[1].slice(1).join(",")}`,
        phone: `${addressData[2][1].split(" ").slice(1).join(" ")}`,
        fax: `${addressData[2][2].split(" ").slice(1).join(" ")}`,
        internet: `${addressData[2][3].split(" ").slice(1).join(" ")}`,
        email: `${addressData[2][4].split(" ").slice(1).join(" ")}`,
        x_justiz_id: `${addressData[2][5].split(" ").slice(1).join(" ")}`,
        transactions_title: `${addressBody[12]}`,
        transactions_list: `${transactions_list}`,
      };

      // console.log("dataForCvs", dataForCvs);

      return dataForCvs;
    };

    if (linksData.length) {
      const nextData = [];

      for (let k = 0; k < linksData.length; k++) {
        const nextResponse = await axios.get(linksData[k]);
        const $ = cheerio.load(nextResponse.data);
        const infoBody = $("main .border.rounded.bg-white.p-3.mb-3");
        const nextStepLinksData = getNextLinks(infoBody);

        if (nextStepLinksData.length) {
          const nextStepData = [];

          for (let j = 0; j < nextStepLinksData.length; j++) {
            const nextStepResponse = await axios.get(nextStepLinksData[j]);
            const $ = cheerio.load(nextStepResponse.data);
            const nextInfoBody = $("main .border.rounded.bg-white.p-3.mb-3");

            const nextStepResData = buildData(nextInfoBody);

            nextStepData.push(nextStepResData);
          }

          nextData.push(...nextStepData);
        } else {
          const nextResData = buildData(infoBody);

          nextData.push(nextResData);
        }
      }

      return nextData;
    } else {
      console.log(`${zipCode} - completed !`, `${getTime()}sec`, `#${start++}`);

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
  const zipCodeArrData = readExelFile(exelDataFilePath);
  // const zipCodeArrData = readExelFile(exelDataFilePath).slice(4190);
  // const zipCodeArrData = ["01099"];
  // const zipCodeArrData = ["99958"];
  // const zipCodeArrData = ["22297"];

  // console.log("zipCodeArrData===", zipCodeArrData.length);

  createNewXLSXFile({ fileName: xlsxFile, sheetName: xlsxSheetName });

  console.log("start get scrapping data ...");

  for (let i = 0; i < zipCodeArrData.length; i++) {
    const res = await getGenre(String(zipCodeArrData[i]));

    // console.log("res===", res);

    // TODO add check
    if (res) {
      res.forEach((el) =>
        updateXLSXFile({
          data: el,
          xlsxFile: xlsxFile,
          xlsxSheetName: xlsxSheetName,
        })
      );
    }
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
