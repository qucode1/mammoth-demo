(async () => {
  const mammoth = require("mammoth");
  const x1 = require("excel4node");
  const fs = require("fs");

  const re = /<tr><td><p><strong>\w*\W*<\/strong>/gm;

  const preRe = /<\w*>/gm;
  const postRe = /<\/\w*>/gm;

  const dataHeader = "Titles";

  // Create Workbook, Worksheet, set header
  const wb = new x1.Workbook();
  const ws = wb.addWorksheet("Sheet 1");
  ws.cell(1, 1)
    .string(dataHeader)
    .style({ font: { bold: true } });

  ws.row(1).freeze();

  const getSourceFileName = () => {
    return new Promise((resolve, reject) => {
      fs.readdir("./source", (err, files) => {
        if (err) {
          reject(err);
        }
        resolve(files.length ? files[0].replace(".docx", "") : null);
      });
    });
  };

  const convertDocxToHtml = async fileName => {
    try {
      const result = await mammoth.convertToHtml({
        path: `./source/${fileName}.docx`
      });
      return result.value;
    } catch (e) {
      console.error(e);
    }
  };

  const getDataFromHtml = htmlResult => {
    const dataArray = htmlResult.match(re).map(data => {
      // clean up data
      const filteredData = data.replace(preRe, "").replace(postRe, "");
      return filteredData;
    });
    return dataArray;
  };

  const writeDataToSheet = (sheet, data) => {
    data.forEach((x, index) => {
      sheet.cell(1 + index + 1, 1).string(x);
    });
  };

  const run = async () => {
    try {
      const sourceFileName = await getSourceFileName();
      if (!sourceFileName) {
        throw new Error(
          "Invalid Source File: Please make sure to put a '.docx' file into the source directory!"
        );
      }
      const htmlResult = await convertDocxToHtml(sourceFileName);
      const data = getDataFromHtml(htmlResult);
      writeDataToSheet(ws, data);
      // generate random String as file name to avoid file name collisions
      const randomString = `${sourceFileName}__${Math.ceil(
        Math.random() * 100000000
      )}`;
      // create results.xlsx
      wb.write(`./results/result__${randomString}.xlsx`);
      console.log(
        `New file: 'result__${randomString}.xlsx' has successfully been created in './results/'`
      );
    } catch (err) {
      console.error(err);
    }
  };

  run();
})();
