const reader = require("xlsx");
const fs = require("fs");

let data = [];
let id = [];

function ExcelDateToJSDate(serial) {
  var utc_days = Math.floor(serial - 25569);
  var utc_value = utc_days * 86400;
  var date_info = new Date(utc_value * 1000);

  var fractional_day = serial - Math.floor(serial) + 0.0000001;

  var total_seconds = Math.floor(86400 * fractional_day);

  var seconds = total_seconds % 60;

  total_seconds -= seconds;

  var hours = Math.floor(total_seconds / (60 * 60));
  var minutes = Math.floor(total_seconds / 60) % 60;

  return new Date(
    date_info.getFullYear(),
    date_info.getMonth(),
    date_info.getDate()
  );
}

const pathFolder = "./PM2.5/";
fs.readdir(pathFolder, async (err, files) => {
  await files.forEach((fileName) => {
    const file = reader.readFile("./PM2.5/" + fileName);

    const raw = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[0]]);
    const header = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[1]]);
    header.forEach((res, index) => {
      if (index > 0) {
        const keys = Object.keys(res);
        const length = keys.length;
        if (length == 3) {
          id.push({
            id: res[keys[0]],
            name: res[keys[1]].trim() + " " + res[keys[2]].trim(),
          });
        } else if (length == 4) {
          id.push({
            id: res[keys[1]],
            name: res[keys[2]].trim() + " " + res[keys[3]].trim(),
          });
        }
      }
    });

    raw.forEach((res, index) => {
      if (index > 0) {
        const date = res["Date"];
        for (let [header, detail] of Object.entries(res)) {
          if (typeof date != "number") break;
          if (header != "Date") {
            const location = id.find((item) => item.id == header);
            if (location) {
              const dateFormat = ExcelDateToJSDate(date)
                .toISOString()
                .split("T")[0];
              let pm25 = "";
              if (typeof(detail) == 'number') pm25 = detail
              data.push({
                date: dateFormat,
                location: location.id,
                pm25: pm25,
              });
            }
          }
        }
      }
    });
  });
  const workSheet = reader.utils.json_to_sheet(data);
  const workSheetHeader = reader.utils.json_to_sheet(id);
  const workBook = reader.utils.book_new();
  reader.utils.book_append_sheet(workBook, workSheet, "dataset");
  reader.utils.book_append_sheet(workBook, workSheetHeader, "header");
  reader.writeFile(workBook, "./dataset.xlsx");
});
