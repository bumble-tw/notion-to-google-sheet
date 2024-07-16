const fs = require("fs")
const readline = require("readline")
const { google } = require("googleapis")

// 加載憑證
const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
const TOKEN_PATH = "token.json"

// 讀取並加載憑證
fs.readFile("credentials.json", async (err, content) => {
  if (err) return console.log("Error loading client secret file:", err)
  // await authorize(JSON.parse(content), listMajors)
  await authorize(JSON.parse(content), writeData)
})

// 創建 OAuth2 客戶端並授權
async function authorize(credentials, callback) {
  const { client_secret, client_id, redirect_uris } = credentials.installed
  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
  )

  // 讀取 token
  try {
    const token = fs.readFileSync(TOKEN_PATH)
    oAuth2Client.setCredentials(JSON.parse(token))
    await callback(oAuth2Client)
  } catch (err) {
    await getNewToken(oAuth2Client, callback)
  }
}

// 獲取並存儲新 Token，然後執行回調
async function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    scope: SCOPES,
  })
  console.log("Authorize this app by visiting this url:", authUrl)
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  })

  rl.question("Enter the code from that page here: ", async (code) => {
    rl.close()
    try {
      const { tokens } = await oAuth2Client.getToken(code)
      oAuth2Client.setCredentials(tokens)
      fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens))
      console.log("Token stored to", TOKEN_PATH)
      await callback(oAuth2Client)
    } catch (err) {
      console.error("Error retrieving access token", err)
    }
  })
}

// 列出 Google Sheets 的主要數據
async function listMajors(auth) {
  const sheets = google.sheets({ version: "v4", auth })
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: "1tpg6_zKzZTTLIjj-eeNw1wDIhqU0Dd0785NzyKwYQW4", // 替換為您的試算表 ID
      range: "A1:A1", // 替換為您要讀取的範圍
    })
    const rows = res.data.values
    if (rows.length) {
      console.log("Name, Major:")
      rows.map((row) => {
        console.log(`${row[0]}, ${row[1]}`)
      })
    } else {
      console.log("No data found.")
    }
  } catch (err) {
    console.log("The API returned an error: " + err)
  }
}

async function writeData(
  auth,
  spreadsheetId = "1tpg6_zKzZTTLIjj-eeNw1wDIhqU0Dd0785NzyKwYQW4",
  range = "A1:B2",
  values = [
    ["Value1", "Value2"],
    ["Value3", "Value4"],
  ]
) {
  const sheets = google.sheets({ version: "v4", auth })
  const resource = {
    values,
  }
  try {
    const result = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption: "RAW",
      resource,
    })
    console.log(`${result.data.updatedCells} cells updated.`)
  } catch (err) {
    console.error("The API returned an error: " + err)
  }
}
