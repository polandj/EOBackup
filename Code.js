// Public Wave API URL - Shouldn't need to change
const EO_HOST = 'https://emailoctopus.com'
const EO_API_URL = EO_HOST + '/api/1.5/'
const EO_FILE_RE = /^EO-export-\d{8}$/g
const EO_KEEP_DAYS = 365

// Need to get an EO token from https://emailoctopus.com/api-documentation and then set it as a
// property for this script.  In the classic editor, go to File -> Project Properties -> Script Properties
// Add row with name EO_API_TOKEN and value of your token.  This setting is only available in the classic editor.
// You can switch back to new editor once it's set, if you want.
const EO_API_TOKEN = PropertiesService.getScriptProperties().getProperty('EO_API_TOKEN')

// Get the list of all the lists from EO
function getAllLists() {
  return JSON.parse(UrlFetchApp.fetch(EO_API_URL + 'lists?api_key=' + EO_API_TOKEN).getContentText())
}

// Get the contacts for the given list object.  Typically, this just gets the first page of 100,
// because the API only provides 100 at a time.
function getContacts(listObj) {
  var url = EO_API_URL + 'lists/' + listObj.id + '/contacts'
  return JSON.parse(UrlFetchApp.fetch(url + '?api_key=' + EO_API_TOKEN).getContentText())
}

// Sum all the values in an object
function sumValues(obj) {
  return Object.values(obj).reduce((a, b) => a + b)
}

// Prepend a zero to the given number if it's less than 10
function zeroPad(num) {
  if (num < 10) {
    num = '0' + num
  }
  return num
}

// Create a new spreadsheet with a summary sheet and one sheet per given list
function makeNewBackup(lists) {
  var today = new Date()
  var month = zeroPad(today.getMonth() + 1)
  var day = zeroPad(today.getDate())
  var ss = SpreadsheetApp.create('EO-export-' + today.getFullYear() + month + day)
  var firstSheet = ss.getActiveSheet()
  firstSheet.setName('Summary')
  firstSheet.appendRow(['This is an export of all the lists in Email Octopus'])
  firstSheet.appendRow(['Generated', today])
  for (var i=0; i < lists.length; i++) {
    firstSheet.appendRow([lists[i].name, sumValues(lists[i].counts)])
    backupToSheet(ss, lists[i])
  }
  return ss
}

// Make a new sheet in the given spreadsheet and populate it with the contacts for that list from EO
function backupToSheet(ss, list) {
  var newSheet = ss.insertSheet(list.name)
  // Make first row with field names
  var row = ['Status']
  for (var i = 0; i < list.fields.length; i++) {
    row.push(list.fields[i].label)
  }
  newSheet.appendRow(row)
  // Fetch and add contact data
  var rowsAdded = 1
  var contacts
  do {
    var rows = []
    if (!contacts) {
      contacts = getContacts(list)
    } else {
      contacts = JSON.parse(UrlFetchApp.fetch(EO_HOST + contacts.paging.next).getContentText())
    }

    for (var i = 0; i < contacts.data.length; i++) {
      row = []
      row.push(contacts.data[i].status)
      row.push(contacts.data[i].email_address)
      // Start at 1 to skip over email_address (it's in fields but not in data)
      for (var j = 1; j < list.fields.length; j++) {
        row.push(contacts.data[i].fields[list.fields[j].tag])
      }
      rows.push(row)
    }
    newSheet.getRange(rowsAdded + 1, 1, rows.length, rows[0].length).setValues(rows)
    rowsAdded += contacts.data.length
  } while (contacts.paging.next)
}

// The primary entrypoint to backup all the lists in EO
function backupAllLists() {
  if (!EO_API_TOKEN) {
    console.log('Error: Your EO_API_TOKEN script property is NOT set!')
  } else {
    const lists = getAllLists().data
    console.log('Exporting ' + lists.length + ' lists')
    if (lists.length > 0) {
      ss = makeNewBackup(lists)
    }
  }
}

// Get a list of file IDs for EO backups older than EO_KEEP_DAYS
function getOldFiles() {
  var arrFileIDs = []

  var Threshold = new Date().getTime()-3600*1000*24*EO_KEEP_DAYS
  var CullDate = new Date(Threshold)
  var strCullDate = Utilities.formatDate(CullDate, "GMT", "yyyy-MM-dd")

  //Create an array of file ID's by date criteria
  var files = DriveApp.searchFiles(
     'modifiedDate < "' + strCullDate + '"')

  while (files.hasNext()) {
    var file = files.next()
    var FileID = file.getId()
    var FileName = file.getName()

    if (FileName.match(EO_FILE_RE)) {
      arrFileIDs.push(FileID)
    }
  }

  return arrFileIDs
}

// External entrypoint to call periodically to delete old backup files
function removeOldBackups() {
  var arrayIDs = getOldFiles()

  for (var i=0; i < arrayIDs.length; i++) {
    console.info('arrayIDs[i]: ' + arrayIDs[i])
    DriveApp.getFileById(arrayIDs[i]).setTrashed(true)
  }

  console.log(`Cleaned up ${arrayIDs.length} old EO backups`)
}

