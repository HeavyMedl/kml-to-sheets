const xml2json = require('xml2json');
const SheetsAPI = require('sheets-api');
const spreadsheet = "1_Hp3-5cQJjFmZZL9zim7YH3ZNUQ8bCeItKJ_n-bns5E";
const url = `https://docs.google.com/spreadsheets/d/${spreadsheet}/edit#gid=0`;
const client = new SheetsAPI(spreadsheet);
const fs = require('fs');
const logger = require('./logger.js')({
  id: 'index.js'
}, 'verbose');
client.authorize()
  .then(auth => {
    try {
      logger.info('Sheets API authorized', auth);
      logger.info('Parsing XML to JSON string and converting to object');
      const kml = fs.readFileSync('./United\ States.kml', 'utf8');
      const kml_obj = JSON.parse(xml2json.toJson(kml));
      logger.info('Building requests array to be sent to Google Sheets');
      let requests = [{ // Clear operation.
        updateCells: {
          range: {
            sheetId: 0,
            startRowIndex: 1
          },
          fields: "userEnteredValue"
        }
      }, { // Append cells operation.
        appendCells: {
          sheetId: 0,
          rows: [],
          fields: '*'
        }
      }];
      kml_obj.kml.Document.Folder.forEach(folder => {
        logger.verbose('Folder name:', folder.name);
        folder.Placemark.forEach(placemark => {
          let row = {
            values: []
          };
          let arr = placemark.ExtendedData.Data;
          let num = get_property("Location Number", arr);
          let address = get_property("Address", arr);
          let region = get_property("Region", arr);
          let phone = get_property("Phone Number", arr);
          logger.silly('Location name:', placemark.name);
          logger.silly('Location number:', num);
          logger.silly('Location address:', address);
          logger.silly('Location region:', region);
          logger.silly('Location Phone Number:', phone);
          row.values.push(make_string_cell(placemark.name)); // Name
          row.values.push(make_num_cell(num)); // Number
          row.values.push(make_string_cell(address)); // Address
          row.values.push(make_string_cell(region)); // Region
          row.values.push(make_string_cell(phone)); // Phone
          logger.silly("Row: ", row);
          requests[1].appendCells.rows.push(row);
        });
      });
      logger.info('Pushing to Google Sheets. Spreadsheet ID: ', spreadsheet);
      logger.verbose('Requests going to batchUpdate: ', requests);
      return client.spreadsheets('batchUpdate', auth, {
        spreadsheetId: spreadsheet,
        resource: {
          requests: requests
        }
      })
    } catch (e) {
      console.error(e);
    }
  })
  .then(auth => logger.info('Process finished successfully. Visit: ', url))
  .catch(reason => {
    console.error(reason);
  });
/**
 * Returns an object representing cell data. Must be a string.
 * @param  {string} str The string that becomes the cell data
 * @return {object}     The object (cell) to be pushed into a row.
 */
function make_string_cell(str) {
  return {
    userEnteredValue: {
      stringValue: typeof str === "string" ? str : "Missing"
    }
  }
}

function make_num_cell(num_str) {
  let num = parseInt(num_str);
  return {
    userEnteredValue: {
      numberValue: !isNaN(num) ? num : 0
    }
  }
}

function get_property(property, arr) {
  let value = "Missing";
  arr.forEach(data_obj => {
    if (data_obj.name == property) {
      value = data_obj.value;
      return;
    }
  });
  return value;
}

function writeToDisk() {
  const fileName = 'test.json'
  const kml = fs.readFileSync('./United\ States.kml', 'utf8');
  const kml_obj = JSON.parse(xml2json.toJson(kml));
  fs.writeFile(fileName, JSON.stringify(kml_obj, null, 2), err => {
    if (err) throw err;
    logger.info(`${fileName} done writing to ${__dirname}`);
  })
}
