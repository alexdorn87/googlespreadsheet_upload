var fs = require('fs');
var readline = require('readline');
var google = require('googleapis');
var googleAuth = require('google-auth-library');
var random = require("random-js")(); // uses the nativeMath engine
var request = require('request');
var obj = require('./document.json');
var spreadsheetid = "1_bEBzuXxEtR8voJNpd8ICampqe_2cEuy_YH84bKw9vA";
var sheets = google.sheets('v4');

// If modifying these scopes, delete your previously saved credentials
// at ~/.credentials/sheets.googleapis.com-nodejs-quickstart.json
var SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
var TOKEN_DIR = (process.env.HOME || process.env.HOMEPATH ||
    process.env.USERPROFILE) + '/.credentials/';
var TOKEN_PATH = TOKEN_DIR + 'rabih.json';

// Load client secrets from a local file.
fs.readFile('client_secret.json', function processClientSecrets(err, content) {
  if (err) {
    console.log('Error loading client secret file: ' + err);
    return;
  }
  // Authorize a client with the loaded credentials, then call the
  // Google Sheets API.
  //console.log (Object.keys(obj));
  authorize(JSON.parse(content), addNewSheet);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 *
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  var clientSecret = credentials.installed.client_secret;
  var clientId = credentials.installed.client_id;
  var redirectUrl = credentials.installed.redirect_uris[0];
  var auth = new googleAuth();
  var oauth2Client = new auth.OAuth2(clientId, clientSecret, redirectUrl);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, function(err, token) {
    if (err) {
      getNewToken(oauth2Client, callback);
    } else {
      oauth2Client.credentials = JSON.parse(token);
      callback(oauth2Client);
    }
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 *
 * @param {google.auth.OAuth2} oauth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback to call with the authorized
 *     client.
 */
function getNewToken(oauth2Client, callback) {
  var authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES
  });
  console.log('Authorize this app by visiting this url: ', authUrl);
  var rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  rl.question('Enter the code from that page here: ', function(code) {
    rl.close();
    oauth2Client.getToken(code, function(err, token) {
      if (err) {
        console.log('Error while trying to retrieve access token', err);
        return;
      }
      oauth2Client.credentials = token;
      storeToken(token);
      callback(oauth2Client);
    });
  });
}

/**
 * Store token to disk be used in later program executions.
 *
 * @param {Object} token The token to store to disk.
 */
function storeToken(token) {
  try {
    fs.mkdirSync(TOKEN_DIR);
  } catch (err) {
    if (err.code != 'EEXIST') {
      throw err;
    }
  }
  fs.writeFile(TOKEN_PATH, JSON.stringify(token));
  console.log('Token stored to ' + TOKEN_PATH);
}

function addNewSheet(auth, callback) {
  sheets.spreadsheets.batchUpdate({
    auth: auth,
    spreadsheetId: spreadsheetid,//"1_bEBzuXxEtR8voJNpd8ICampqe_2cEuy_YH84bKw9vA",
    resource: {
      "requests": [
        {
          "addSheet": {
            "properties": {
              "title": new Date().toISOString(),
              /*"gridProperties": {
               "rowCount": 86,
               "columnCount": 20
               },*/
              "tabColor": {
                "red": random.integer(1, 100)/100,
                "green": random.integer(1, 100)/100,
                "blue": random.integer(1, 100)/100
              }
            }
          }
        }
      ]
    }
  }, (err, res1) => {
    if (err) {
      console.log('The API returned an error: ' + err);
      return;
    } else {
      console.log("New Sheet Appended");
      console.log (res1);
      //getsheetid
      getsheetid(auth, batchupdate)
    }
  });
}

function getsheetid(auth, callback) {
  sheets.spreadsheets.get({
    auth: auth,
    spreadsheetId: spreadsheetid,//"1_bEBzuXxEtR8voJNpd8ICampqe_2cEuy_YH84bKw9vA",
    includeGridData: false,

  }, (err, res2) => {
    if (err) {
      console.log('The API returned an error: ' + err);
      return;
    } else {
      //console.log("sheet info");
      //console.log (res2.sheets[res2.sheets.length-1].properties.sheetId);
      request('http://api.openratio.com/v4/app/getV4FullProperty?slugline=barnc', function (error, response, body) {
        //console.log('error:', error); // Print the error if one occurred
        //console.log('statusCode:', response && response.statusCode); // Print the response status code if a response was received
        //console.log(JSON.parse(body)); // Print the HTML for the Google homepage.
        var obj = JSON.parse(body).modules;
        var row = 0;
        var tmp = {};
        //console.log (obj);
        Object.keys(obj).forEach(function(item){
          tmp = {};
          tmp [item] = obj[item];
          //console.log (obj[item]);
          //batchupdate
          row = callback (auth, res2.sheets[res2.sheets.length-1].properties.sheetId, tmp, row, null);
        });

      });

    }
  });
}

function makestr(obj_tmp) {
  var str = '';
  if (obj_tmp != undefined)
    str = obj_tmp.toString();
  return str;
}

function batchupdate(auth, sheetId, module_obj, startrow, callback) {
  var requests = [];
  var row = startrow;
  var col = 0;
  var ent = module_obj;

  requests.push({
    repeatCell: {
      range: {sheetId: sheetId, startRowIndex: 0, endRowIndex:1000},
      cell: {
        userEnteredFormat: {
          /*"backgroundColor": {
           "red": 0.0,
           "green": 0.0,
           "blue": 0.0
           },*/
          "horizontalAlignment" : "CENTER",
          "textFormat": {
            /*"foregroundColor": {
             "red": 1.0,
             "green": 1.0,
             "blue": 1.0
             },*/
            "fontSize": 12,
            "bold": true
          }
        }
      },
      fields: 'userEnteredValue,userEnteredFormat.horizontalAlignment'
    }
  });

  requests.push({
    "mergeCells": {
      "range": {
        "sheetId": sheetId,
        "startRowIndex": row,
        "endRowIndex": row+1,
        "startColumnIndex": col,
        "endColumnIndex": col+19,
      },
      "mergeType": 'MERGE_ALL'
    }
  });

  requests.push({
    "mergeCells": {
      "range": {
        "sheetId": sheetId,
        "startRowIndex": row+1,
        "endRowIndex": row+2,
        "startColumnIndex": col+4,
        "endColumnIndex": col+12,
      },
      "mergeType": 'MERGE_ALL'
    }
  });

  requests.push({
    "mergeCells": {
      "range": {
        "sheetId": sheetId,
        "startRowIndex": row+1,
        "endRowIndex": row+2,
        "startColumnIndex": col+12,
        "endColumnIndex": col+18,
      },
      "mergeType": 'MERGE_ALL'
    }
  });

  requests.push({
    "mergeCells": {
      "range": {
        "sheetId": sheetId,
        "startRowIndex": row+2,
        "endRowIndex": row+3,
        "startColumnIndex": col+4,
        "endColumnIndex": col+12,
      },
      "mergeType": 'MERGE_ALL'
    }
  });

  requests.push({
    "mergeCells": {
      "range": {
        "sheetId": sheetId,
        "startRowIndex": row+2,
        "endRowIndex": row+3,
        "startColumnIndex": col+12,
        "endColumnIndex": col+18,
      },
      "mergeType": 'MERGE_ALL'
    }
  });

  requests.push({
    "mergeCells": {
      "range": {
        "sheetId": sheetId,
        "startRowIndex": row+3,
        "endRowIndex": row+4,
        "startColumnIndex": col+6,
        "endColumnIndex": col+12,
      },
      "mergeType": 'MERGE_ALL'
    }
  });

  requests.push({
    "mergeCells": {
      "range": {
        "sheetId": sheetId,
        "startRowIndex": row+3,
        "endRowIndex": row+4,
        "startColumnIndex": col+15,
        "endColumnIndex": col+18,
      },
      "mergeType": 'MERGE_ALL'
    }
  });
  requests.push({
    updateCells: {
      start: {sheetId: sheetId, rowIndex: row, columnIndex: col},
      rows: [
          {values: [{
            userEnteredValue: {stringValue: Object.keys(ent)[0]},
            userEnteredFormat: {backgroundColor: {blue: 1}}
          }]
        }],
      fields: 'userEnteredValue,userEnteredFormat.backgroundColor'
    }
  });
  ent = ent[Object.keys(ent)[0]];

  requests.push({
    updateCells: {
      start: {sheetId: sheetId, rowIndex: row+1, columnIndex: col},
      rows: [
        {values: [{
            userEnteredValue: {stringValue: 'id'},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8}}
          },
          {
            userEnteredValue: {stringValue: 'type'},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8}}
          },
          {
            userEnteredValue: {stringValue: 'title'},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8}}
          },
          {
            userEnteredValue: {stringValue: 'style'},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8}}
          },
          {
            userEnteredValue: {stringValue: 'options'},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8}}
          }
          ]
        },
        {values: [{
            userEnteredValue: {stringValue: makestr(ent.id)},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          },
          {
            userEnteredValue: {stringValue: makestr(ent.type)},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          },
          {
            userEnteredValue: {stringValue: makestr(ent.title)},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          },
          {
            userEnteredValue: {stringValue: makestr(ent.style)},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          },
          {
            userEnteredValue: {stringValue: 'options []'},
            userEnteredFormat: {backgroundColor: {red:1, green: 0.8, blue:0.5}}
          }
        ]
        }
      ],
      fields: 'userEnteredValue,userEnteredFormat.backgroundColor'
    }
  });

  requests.push({
    updateCells: {
      start: {sheetId: sheetId, rowIndex: row+2, columnIndex: col+18},
      rows: [
        {values: [{
          userEnteredValue: {stringValue: ent[Object.keys(ent)[5]].toString()},
          userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
        }
        ]
        }
      ],
      fields: 'userEnteredValue,userEnteredFormat.backgroundColor'
    }
  });

  requests.push({
    updateCells: {
      start: {sheetId: sheetId, rowIndex: row+1, columnIndex: col+12},
      rows: [
        {values: [{
          userEnteredValue: {stringValue: Object.keys(ent)[4]+' []'},
          userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8}}
        }]
        }],
      fields: 'userEnteredValue,userEnteredFormat.backgroundColor'
    }
  });
  requests.push({
    updateCells: {
      start: {sheetId: sheetId, rowIndex: row+1, columnIndex: col+18},
      rows: [
        {values: [{
          userEnteredValue: {stringValue: Object.keys(ent)[5]},
          userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8}}
        }]
        }],
      fields: 'userEnteredValue,userEnteredFormat.backgroundColor'
    }
  });
  //console.log (ent[Object.keys(ent)[4]]);
  requests.push({
    updateCells: {
      start: {sheetId: sheetId, rowIndex: row+2, columnIndex: col+12},
      rows: [
        {values: [{
          userEnteredValue: {stringValue: 'data []'},
          userEnteredFormat: {backgroundColor: {red: random.integer(1, 100)/100, green: random.integer(1, 100)/100, blue:random.integer(1, 100)/100}}
        }]
        }],
      fields: 'userEnteredValue,userEnteredFormat.backgroundColor'
    }
  });

  //var ent1 = ent[Object.keys(ent)[3]] [Object.keys(ent[Object.keys(ent)[3]])[0]][0];
  console.log (ent);
  var ent1 = ent.options[Object.keys(ent.options)[0]];
  if (Array.isArray(ent1))
    ent1 = ent1[0];
  var col1 = {red: random.integer(1, 100)/100, green: random.integer(1, 100)/100, blue:random.integer(1, 100)/100};
  var col2 = {red: random.integer(1, 100)/100, green: random.integer(1, 100)/100, blue:random.integer(1, 100)/100};

  requests.push({
    updateCells: {
      start: {sheetId: sheetId, rowIndex: row+3, columnIndex: col+4},
      rows: [
        {values: [{
          userEnteredValue: {stringValue: Object.keys(ent1)[0]},
          userEnteredFormat: {backgroundColor: col1}
        },
          {
            userEnteredValue: {stringValue: Object.keys(ent1)[1]},
            userEnteredFormat: {backgroundColor: col1}
          },
          {
            userEnteredValue: {stringValue: Object.keys(ent1)[2]},
            userEnteredFormat: {backgroundColor: col1}
          }
        ]
        },
        {values: [
          {
            userEnteredValue: {stringValue: ""},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          },
          {
            userEnteredValue: {stringValue: ""},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          },
          {
            userEnteredValue: {stringValue: Object.keys(ent1[Object.keys(ent1)[2]])[0].toString()},
            userEnteredFormat: {backgroundColor: col2}
          },
          {
            userEnteredValue: {stringValue: Object.keys(ent1[Object.keys(ent1)[2]])[1].toString()},
            userEnteredFormat: {backgroundColor: col2}
          },
          {
            userEnteredValue: {stringValue: Object.keys(ent1[Object.keys(ent1)[2]])[2].toString()},
            userEnteredFormat: {backgroundColor: col2}
          },
          {
            userEnteredValue: {stringValue: Object.keys(ent1[Object.keys(ent1)[2]])[3].toString()},
            userEnteredFormat: {backgroundColor: col2}
          },
          {
            userEnteredValue: {stringValue: Object.keys(ent1[Object.keys(ent1)[2]])[4].toString()},
            userEnteredFormat: {backgroundColor: col2}
          },
          {
            userEnteredValue: {stringValue: Object.keys(ent1[Object.keys(ent1)[2]])[5].toString()},
            userEnteredFormat: {backgroundColor: col2}
          }
        ]
        },
        {values: [
          {
            userEnteredValue: {stringValue: ent1[Object.keys(ent1)[0]].toString()},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          },
          {
            userEnteredValue: {stringValue: ent1[Object.keys(ent1)[1]].toString()},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          },
          {
            userEnteredValue: {stringValue: ent1[Object.keys(ent1)[2]][Object.keys(ent1[Object.keys(ent1)[2]])[0]].toString()},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          },
          {
            userEnteredValue: {stringValue: ent1[Object.keys(ent1)[2]][Object.keys(ent1[Object.keys(ent1)[2]])[1]].toString()},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          },
          {
            userEnteredValue: {stringValue: ent1[Object.keys(ent1)[2]][Object.keys(ent1[Object.keys(ent1)[2]])[2]].toString()},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          },
          {
            userEnteredValue: {stringValue: ent1[Object.keys(ent1)[2]][Object.keys(ent1[Object.keys(ent1)[2]])[3]].toString()},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          },
          {
            userEnteredValue: {stringValue: ent1[Object.keys(ent1)[2]][Object.keys(ent1[Object.keys(ent1)[2]])[4]].toString()},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          },
          {
            userEnteredValue: {stringValue: ent1[Object.keys(ent1)[2]][Object.keys(ent1[Object.keys(ent1)[2]])[5]].toString()},
            userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8, red: 0.8}}
          }
        ]
        }
      ],
      fields: 'userEnteredValue,userEnteredFormat.backgroundColor'
    }
  });

  var ent2 = ent.data[0].data;
  var col3 = {red: random.integer(1, 100)/100, green: random.integer(1, 100)/100, blue:random.integer(1, 100)/100};
  var col4 = {red: random.integer(1, 100)/100, green: random.integer(1, 100)/100, blue:random.integer(1, 100)/100};
  //console.log (ent2);

  requests.push({
    updateCells: {
      start: {sheetId: sheetId, rowIndex: row+3, columnIndex: col+12},
      rows: [
        {values: [{
          userEnteredValue: {stringValue: 'type'},
          userEnteredFormat: {backgroundColor: {blue:0.8}}
        },
          {
            userEnteredValue: {stringValue: 'title'},
            userEnteredFormat: {backgroundColor: {blue:0.8}}
          },
          {
            userEnteredValue: {stringValue: 'staffId'},
            userEnteredFormat: {backgroundColor: {blue:0.8}}
          },
          {
            userEnteredValue: {stringValue: 'data'},
            userEnteredFormat: {backgroundColor: {blue:0.8}}
          }
          ]
        }],
      fields: 'userEnteredValue,userEnteredFormat.backgroundColor'
    }
  });
  var drow = row+4;
  var dcol = col+12;

  ent2.forEach(function(item){
    requests.push({
      updateCells: {
        start: {sheetId: sheetId, rowIndex: drow, columnIndex: dcol},
        rows: [
          {values: [{
            userEnteredValue: {stringValue: ''},
            userEnteredFormat: {backgroundColor: {blue:0.8, red:0.8, green:0.8}}
          },
            {
              userEnteredValue: {stringValue: ''},
              userEnteredFormat: {backgroundColor: {blue:0.8, red:0.8, green:0.8}}
            },
            {
              userEnteredValue: {stringValue: ''},
              userEnteredFormat: {backgroundColor: {blue:0.8, red:0.8, green:0.8}}
            },
            {
              userEnteredValue: {stringValue: 'description'},
              userEnteredFormat: {backgroundColor: {blue:0.2, red:0.2, green:0.2},
                "textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 1.0},"fontSize": 12,/*"bold": true*/}}
            },
            {
              userEnteredValue: {stringValue: 'phone'},
              userEnteredFormat: {backgroundColor: {blue:0.2, red:0.2, green:0.2},
                "textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 1.0},"fontSize": 12,/*"bold": true*/}}
            },
            {
              userEnteredValue: {stringValue: 'email'},
              userEnteredFormat: {backgroundColor: {blue:0.2, red:0.2, green:0.2},
                "textFormat": {"foregroundColor": {"red": 1.0,"green": 1.0,"blue": 1.0},"fontSize": 12,/*"bold": true*/}
              }
            }
          ]
          }],
        fields: 'userEnteredValue,userEnteredFormat(backgroundColor,textFormat)'
      }
    });
    requests.push({
      updateCells: {
        start: {sheetId: sheetId, rowIndex: drow+1, columnIndex: dcol},
        rows: [
          {values: [{
            userEnteredValue: {stringValue: makestr(item.type)},
            userEnteredFormat: {backgroundColor: {blue:0.8, red:0.8, green:0.8},
              wrapStrategy: 'WRAP'}
          },
            {
              userEnteredValue: {stringValue: makestr(item.title)},
              userEnteredFormat: {backgroundColor: {blue:0.8, red:0.8, green:0.8},
                wrapStrategy: 'WRAP'}
            },
            {
              userEnteredValue: {stringValue: makestr(item.staffId)},
              userEnteredFormat: {backgroundColor: {blue:0.8, red:0.8, green:0.8},
                wrapStrategy: 'WRAP'}
            },
            {
              userEnteredValue: {stringValue: makestr(item.data.description)},
              userEnteredFormat: {backgroundColor: {blue:0.8, red:0.8, green:0.8},
                wrapStrategy: 'WRAP'
              }
            },
            {
              userEnteredValue: {stringValue: makestr(item.data.phone)},
              userEnteredFormat: {backgroundColor: {blue:0.8, red:0.8, green:0.8},
                wrapStrategy: 'WRAP'}
            },
            {
              userEnteredValue: {stringValue: makestr(item.data.email)},
              userEnteredFormat: {backgroundColor: {blue:0.8, red:0.8, green:0.8},
                wrapStrategy: 'WRAP'}
            }
          ]
          }],
        fields: 'userEnteredValue,userEnteredFormat(backgroundColor,wrapStrategy)'
      }
    });
    drow = drow+2;
  });

  /*requests.push({
    updateCells: {
      start: {sheetId: sheetId, rowIndex: row+1, columnIndex: col+1},
      rows: [
        {values: [{
          userEnteredValue: {stringValue: Object.keys(ent)[1]},
          userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8}}
        }]
        }],
      fields: 'userEnteredValue,userEnteredFormat.backgroundColor'
    }
  });
  requests.push({
    updateCells: {
      start: {sheetId: sheetId, rowIndex: row+1, columnIndex: col+2},
      rows: [
        {values: [{
          userEnteredValue: {stringValue: Object.keys(ent)[2]},
          userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8}}
        }]
        }],
      fields: 'userEnteredValue,userEnteredFormat.backgroundColor'
    }
  });
  requests.push({
    updateCells: {
      start: {sheetId: sheetId, rowIndex: row+1, columnIndex: col+3},
      rows: [
        {values: [{
          userEnteredValue: {stringValue: Object.keys(ent)[3]},
          userEnteredFormat: {backgroundColor: {green: 0.8, blue:0.8}}
        }]
        }],
      fields: 'userEnteredValue,userEnteredFormat.backgroundColor'
    }
  });
  */


  sheets.spreadsheets.batchUpdate({
    auth: auth,
    spreadsheetId: "1_bEBzuXxEtR8voJNpd8ICampqe_2cEuy_YH84bKw9vA",
    resource: {
      "requests": requests
    }
  }, (err, res3) => {
    if (err) {
      console.log('The API returned an error: ' + err);
      return -1;
    } else {
      //console.log("finished!");
      return drow;
    }
  });
}

/**
 * Print the names and majors of students in a sample spreadsheet:
 * https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
 */
function mergecell(auth, sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex, callback) {
  sheets.spreadsheets.batchUpdate({
    auth: auth,
    spreadsheetId: "1_bEBzuXxEtR8voJNpd8ICampqe_2cEuy_YH84bKw9vA",
    resource: {
      "requests": [
        {
          "mergeCells": {
            "range": {
              "sheetId": sheetId,
              "startRowIndex": startRowIndex,
              "endRowIndex": endRowIndex,
              "startColumnIndex": startColumnIndex,
              "endColumnIndex": endColumnIndex,
            },
            "mergeType": 'MERGE_ALL'
          }
        }
      ]
    }
  }, (err, res3) => {
    if (err) {
      console.log('The API returned an error: ' + err);
      return;
    } else {
      console.log("merged");
    }
  });
}


function updatecell(auth, start, rows, fields, callback) {
  sheets.spreadsheets.batchUpdate({
    auth: auth,
    spreadsheetId: "1_bEBzuXxEtR8voJNpd8ICampqe_2cEuy_YH84bKw9vA",
    resource: {
      "requests": [
        {
          "mergeCells": {
            "range": {
              "sheetId": res2.sheets[res2.sheets.length-1].properties.sheetId,
              "startRowIndex": 0,
              "endRowIndex": 1,
              "startColumnIndex": 0,
              "endColumnIndex": 18,
            },
            "mergeType": 'MERGE_ALL'
          }
        }
      ]
    }
  }, (err, res3) => {
    if (err) {
      console.log('The API returned an error: ' + err);
      return;
    } else {
      console.log("merged");

    }
  });
}
