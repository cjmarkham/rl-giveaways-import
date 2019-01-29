const fs = require('fs')
const readline = require('readline')
const { google } = require('googleapis')
const items = require('./items.js')
const COLORS = require('./colors.js')

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
    createSheetIfNeeded(auth, category, function (sheetId) {
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

          let requests = []

          data.forEach(d => {
            const colorRange = d.range.replace(/\w+![A-Z][0-9]+:([A-Z+][0-9]+)/, '$1')
            const colorRow = colorRange[0]
            let colorColumn = colorRange[1]
            if (colorRange.length === 3) {
              colorColumn = colorRange[1] + colorRange[2]
            }
            let color;
            Object.keys(COLORS).map(function (c) {
              if (COLORS[c].name == d.values[0][2]) {
                color = COLORS[c]
              }
            })
            let backgroundColor;
            let foregroundColor = {r: 0, g: 0, b: 0}

            if ( ! color) {
              backgroundColor = {r: 0, g: 0, b: 0}
            } else {
              backgroundColor = hexToRgb('#' + color.color)
              if (color.text) {
                foregroundColor = hexToRgb(color.text)
              }
            }

            if (d.values[0][2] === 'Not painted') {
              foregroundColor = {r: 255, g: 255, b: 255}
            }

            requests.push({
              repeatCell: {
                range: {
                  sheetId,
                  startRowIndex: colorColumn - 2,
                  endRowIndex: colorColumn - 1,
                  startColumnIndex: (colorRow.charCodeAt(0) - 64) - 1,
                  endColumnIndex: (colorRow.charCodeAt(0) - 64),
                },
                cell: {
                  userEnteredFormat: {
                    backgroundColor: {
                      red: backgroundColor.r / 255,
                      green: backgroundColor.g / 255,
                      blue: backgroundColor.b / 255,
                    },
                    textFormat: {
                      foregroundColor: {
                        red: foregroundColor.r / 255,
                        green: foregroundColor.g / 255,
                        blue: foregroundColor.b / 255,
                      }
                    }
                  }
                },
                fields: "userEnteredFormat(textFormat,backgroundColor)",
              }
            })
          })

          setTimeout(() => {
            const batchUpdateRequest = { requests };

            sheets.spreadsheets.batchUpdate({
              spreadsheetId: SHEET_ID,
              resource: batchUpdateRequest,
            }, (err, response) => {
              if (err) {
                console.error(err.errors)
                return
              }
            }, 2000)
          })
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
      if (column == 'P') {
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
          name,
          '=IMAGE("' + image.url + '", 2)',
          item['painted'] ? item['painted']['name'] : 'Not painted',
          item['certified'] ? item['certified'].toUpperCase() : 'Not certified',
        ]
      ]
    }
  }))
}

function nextLetter (currentLetter) {
  let charCode = currentLetter.charCodeAt(currentLetter.length - 1)
  return String.fromCharCode(charCode + 2).toUpperCase()
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
            red: 0,
            green: 0,
            blue: 0,
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
  }]

  for (var o = 1; o < 14; o+=2) {
    requests.push({
      updateDimensionProperties: {
        range: {
          sheetId,
          dimension: 'COLUMNS',
          startIndex: o,
          endIndex: o+1,
        },
        properties: {
          pixelSize: 200,
        },
        fields: 'pixelSize'
      }
    })
  }

  for (var o = 0; o < 15; o+=2) {
    requests.push({
      updateDimensionProperties: {
        range: {
          sheetId,
          dimension: 'COLUMNS',
          startIndex: o,
          endIndex: o+1,
        },
        properties: {
          pixelSize: 20,
        },
        fields: 'pixelSize'
      }
    })
  }

  // Update the row heights
  for (var i = 0; i <= 50; i += 5) {
    requests.push({
      updateDimensionProperties: {
        range: {
          sheetId,
          dimension: 'ROWS',
          startIndex: i + 1,
          endIndex: i + 2,
        },
        properties: {
          pixelSize: 200,
        },
        fields: 'pixelSize'
      }
    })
  }

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
          pixelSize: 40,
        },
        fields: 'pixelSize'
      }
    })
  }

  for (var i = 0; i <= 50; i += 5) {
    requests.push({
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: i,
          endRowIndex: i + 1,
        },
        cell: {
          userEnteredFormat: {
            textFormat: {
              fontSize: 12,
              bold: true,
              foregroundColor: {
                red: 1.0,
                green: 1.0,
                blue: 1.0,
              }
            },
            verticalAlignment: 'MIDDLE',
          }
        }, fields: 'userEnteredFormat(textFormat,verticalAlignment)',
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

    callback(sheetId)
  })
}

function getImage (category, itemName) {
  let json = fs.readFileSync('images/' + category + '.json', 'utf8')
  let images = JSON.parse(json)
  return images.find(image => image.name === itemName)
}

function hexToRgb(hex) {
    var shorthandRegex = /^#?([a-f\d])([a-f\d])([a-f\d])$/i;
    hex = hex.replace(shorthandRegex, function(m, r, g, b) {
        return r + r + g + g + b + b;
    });

    var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
    } : null;
}