const fs = require('fs')
const readline = require('readline')
const { google } = require('googleapis')
const items = require('./items.js')

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json'
const SHEET_ID = '1fayZ9ySip-ZWkYc0j5V6nxxjJXMMzYMzTEclYApEYuA'

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
  if (err) return console.log('Error loading client secret file:', err);
  // Authorize a client with credentials, then call the Google Sheets API.
  authorize(JSON.parse(content), writeData);
});

let sheets

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  const {client_secret, client_id, redirect_uris} = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
      client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error while trying to retrieve access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

function writeData (auth) {
  sheets = google.sheets({version: 'v4', auth});

  for (let category in items) {
    createSheetIfNeeded(auth, category, function () {
      buildArrayOfData(category).then(data => {
        const request = {
          spreadsheetId: SHEET_ID,
          resource: {
            valueInputOption: 'USER_ENTERED',
            data: data,
          },
          auth: auth,
        }
        sheets.spreadsheets.values.batchUpdate(request, function (err, response) {
          if (err) {
            console.error(err.errors)
          }
        })
      })
    })
  }
}

async function buildArrayOfData(category) {
  let data = []
  let column = null
  let row = null

  return await Promise.all(items[category].map(item => {
    if (column == null) {
      column = 'B'
      row = 1
    } else {
      column = nextLetter(column)
      if (column == 'I') {
        column = 'B'
        row += 5
      }
    }

    let image = getImage(category, item['name'])
    if ( ! image) {
      console.error('No image found for', item['name'], category)
    }

    let name = item['name']
    if (item.hasOwnProperty('prefix')) {
      name = item['prefix'] + ': ' + name
    }

    const range = category.toUpperCase() + '!' + column + row + ':' + column + (row + 3)

    return {
      range: range,
      majorDimension: 'COLUMNS',
      values: [
        [
          '=IMAGE("' + image.url + '", 2)',
          name,
          item['painted'] || '-',
          item['certified'] || '-',
        ]
      ]
    }
  }))
}

function nextLetter (currentLetter) {
  let charCode = currentLetter.charCodeAt(currentLetter.length - 1)
  return String.fromCharCode(charCode + 1).toUpperCase()
}

function createSheetIfNeeded (authClient, category, callback) {
  sheets.spreadsheets.get({
    spreadsheetId: SHEET_ID,
  }, function (err, response) {

    let exists = response.data.sheets.map(sheet => {
      let name = sheet.properties.title
      return name === category.toUpperCase()
    })

    if (exists.some(boolean => boolean)) {
      callback()
      return
    }

    let requests = [{
      addSheet: {
        properties: {
          title: category.toUpperCase(),
        },
      }
    }]

    const batchUpdateRequest = { requests };

    sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      resource: batchUpdateRequest,
    }, (err, response) => {
      if (err) {
        console.error(err.errors)
        return
      }

      sheetId = response.data.replies[0].addSheet.properties.sheetId
      formatSheet(sheetId, category, callback)
    })
  })
}

function formatSheet(sheetId, category, callback) {
  let requests = [{
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: 0,
        endRowIndex: 99,
      },
      cell: {
        userEnteredFormat: {
          horizontalAlignment: 'CENTER',
          backgroundColor: {
            red: 0.0,
            green: 0.0,
            blue: 0.0,
          },
          textFormat: {
            foregroundColor: {
              red: 1.0,
              green: 1.0,
              blue: 1.0,
            }
          }
        }
      },
      fields: "userEnteredFormat(horizontalAlignment,textFormat,backgroundColor)",
    }
  }, {
    updateDimensionProperties: {
      range: {
        sheetId,
        dimension: 'COLUMNS',
        startIndex: 1,
        endIndex: 8,
      },
      properties: {
        pixelSize: 200,
      },
      fields: 'pixelSize'
    }
  }]

  for (var i = 0; i < 50; i+=5) {
    for (var j = 1; j <= 8; j++) {
      requests.push({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: i,
            endRowIndex: i+1,
            startColumnIndex: j,
            endColumnIndex: 8,
          },
          cell: {
            userEnteredFormat: {
              borders: {
                top: {
                  width: 1,
                  style: 'SOLID',
                  color: {
                    red: 1.0,
                    green: 1.0,
                    blue: 1.0,
                  }
                },
                bottom: {
                  width: 1,
                  style: 'SOLID',
                  color: {
                    red: 1.0,
                    green: 1.0,
                    blue: 1.0,
                  }
                },
                left: {
                  width: 1,
                  style: 'SOLID',
                  color: {
                    red: 1.0,
                    green: 1.0,
                    blue: 1.0,
                  }
                },
                right: {
                  width: 1,
                  style: 'SOLID',
                  color: {
                    red: 1.0,
                    green: 1.0,
                    blue: 1.0,
                  }
                }
              },
            }
          },
          fields: "userEnteredFormat(borders)",
        }
      })
    }
  }

  for (var i = 1; i < 50; i+=5) {
    for (var j = 1; j <= 8; j++) {
      requests.push({
        repeatCell: {
          range: {
            sheetId,
            startRowIndex: i,
            endRowIndex: i+3,
            startColumnIndex: j,
            endColumnIndex: 8,
          },
          cell: {
            userEnteredFormat: {
              borders: {
                top: {
                  width: 1,
                  style: 'SOLID',
                  color: {
                    red: 1.0,
                    green: 1.0,
                    blue: 1.0,
                  }
                },
                bottom: {
                  width: 1,
                  style: 'SOLID',
                  color: {
                    red: 1.0,
                    green: 1.0,
                    blue: 1.0,
                  }
                },
                left: {
                  width: 1,
                  style: 'SOLID',
                  color: {
                    red: 1.0,
                    green: 1.0,
                    blue: 1.0,
                  }
                },
                right: {
                  width: 1,
                  style: 'SOLID',
                  color: {
                    red: 1.0,
                    green: 1.0,
                    blue: 1.0,
                  }
                }
              },
            }
          },
          fields: "userEnteredFormat(borders)",
        }
      })
    }
  }

  // Update the row heights
  for (var i = 0; i <= 50; i += 5) {
    requests.push({
      updateDimensionProperties: {
        range: {
          sheetId,
          dimension: 'ROWS',
          startIndex: i,
          endIndex: i + 1,
        },
        properties: {
          pixelSize: 200,
        },
        fields: 'pixelSize'
      }
    })
  }

  const batchUpdateRequest = { requests };

  sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    resource: batchUpdateRequest,
  }, (err, response) => {
    if (err) {
      console.error(err.errors)
      return
    }

    setDefaultData(category, callback)
  })
}

function setDefaultData(category, callback) {
  let rows = Math.ceil(items[category].length / 7)
  console.log('Rows', rows, items[category].length)


  for (let i = 2; i < 5 * rows; i += 5) {
    let range = category + '!A' + i + ':A' + (i + 2)
    let values = [
      [
        'Name',
        'Painted?',
        'Certified?',
      ]
    ]

    const resource = { values, majorDimension: 'COLUMNS' }
    
    sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range,
      valueInputOption: 'USER_ENTERED',
      resource
    }, (err, result) => {
      if (err) { 
        console.error(err.errors) 
      }
    })
  }

  callback()
}

function getImage (category, itemName) {
  let json = fs.readFileSync('images/' + category + '.json', 'utf8')
  let images = JSON.parse(json)
  return images.find(image => image.name === itemName)
}

