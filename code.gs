const DATA_SHEET = 'Data'
const LOGIN_SHEET = 'Login'
const OPTIONS_SHEET = 'Options'

function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate().setTitle('CRUD App')
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).getRawContent()
}

function userLogin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const loginSheet = ss.getSheetByName(LOGIN_SHEET)
  const data = loginSheet.getDataRange().getValues()
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === password) {
      return { success: true, userType: data[i][2] }
    }
  }
  return { success: false, message: 'Invalid username or password' }
}

function registerUser(newUsername, newPassword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const loginSheet = ss.getSheetByName(LOGIN_SHEET)
  const data = loginSheet.getDataRange().getValues()
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === newUsername) {
      return { success: false, message: 'Username already exists' }
    }
  }
  loginSheet.appendRow([ newUsername, newPassword, 'User', '' ])
  return { success: true, message: 'Registered successfully. You may now log in.' }
}

function forgotPasswordRequest(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const loginSheet = ss.getSheetByName(LOGIN_SHEET)
  const data = loginSheet.getDataRange().getValues()
  let userRow = -1
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      userRow = i
      break
    }
  }
  if (userRow === -1) {
    return { success: false, message: 'Username not found.' }
  }
  const otp = Math.floor(100000 + Math.random() * 900000).toString()
  loginSheet.getRange(userRow + 1, 4).setValue(otp)
  try {
    MailApp.sendEmail({
      to: username,
      subject: 'Your OTP for Password Reset',
      body: 'Your OTP is: ' + otp
    })
  } catch (err) {
    return { success: false, message: 'Error sending OTP email: ' + err }
  }
  return { success: true, message: 'OTP sent to your email.' }
}

function forgotPasswordVerify(username, userOTP, newPassword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const loginSheet = ss.getSheetByName(LOGIN_SHEET)
  const data = loginSheet.getDataRange().getValues()
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      if (data[i][3] && data[i][3].toString() === userOTP) {
        loginSheet.getRange(i + 1, 2).setValue(newPassword)
        loginSheet.getRange(i + 1, 4).setValue('')
        return { success: true, message: 'Password updated successfully' }
      } else {
        return { success: false, message: 'Invalid OTP' }
      }
    }
  }
  return { success: false, message: 'Username not found.' }
}

function getDropdownOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(OPTIONS_SHEET)
  const range = sheet.getDataRange().getValues()
  const classList = []
  const categoryList = []
  const branchList = []
  const staffBranchList = []
  for (let i = 1; i < range.length; i++) {
    const row = range[i]
    if (row[0] && classList.indexOf(row[0]) === -1) classList.push(row[0])
    if (row[1] && categoryList.indexOf(row[1]) === -1) categoryList.push(row[1])
    if (row[2] && branchList.indexOf(row[2]) === -1) branchList.push(row[2])
    if (row[3] && staffBranchList.indexOf(row[3]) === -1) staffBranchList.push(row[3])
  }
  return {
    classList: classList,
    categoryList: categoryList,
    branchList: branchList,
    staffBranchList: staffBranchList
  }
}

function getData(username, userType, filterClass, filterCategory, filterBranch, filterStaffBranch, dateFrom, dateTo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(DATA_SHEET)
  const values = sheet.getDataRange().getValues()
  const dataRows = values.slice(1)
  const results = []
  let fromDate = dateFrom ? new Date(dateFrom + 'T00:00:00') : null
  let toDate = dateTo ? new Date(dateTo + 'T23:59:59') : null
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i]
    const admissionNo = row[0]
    const branchName = row[1]
    const studentName = row[2]
    const fatherName = row[3]
    const classValue = row[4]
    const categoryValue = row[5]
    const admissionDate = row[6]
    const staffName = row[7]
    const designation = row[8]
    const staffBranch = row[9]
    const recordOwner = row[10]
    const staffCode = row[11]
    if (userType.toLowerCase() === 'user' && recordOwner !== username) continue
    if (filterClass && filterClass !== '' && filterClass !== classValue) continue
    if (filterCategory && filterCategory !== '' && filterCategory !== categoryValue) continue
    if (filterBranch && filterBranch !== '' && filterBranch !== branchName) continue
    if (filterStaffBranch && filterStaffBranch !== '' && filterStaffBranch !== staffBranch) continue
    if ((fromDate || toDate) && admissionDate instanceof Date) {
      if (fromDate && admissionDate < fromDate) continue
      if (toDate && admissionDate > toDate) continue
    } else if ((fromDate || toDate) && !(admissionDate instanceof Date)) {
      continue
    }
    let formattedDate = ''
    if (admissionDate instanceof Date) {
      formattedDate = Utilities.formatDate(admissionDate, 'Asia/Kolkata', 'yyyy-MM-dd')
    }
    results.push({
      admissionNo: admissionNo,
      branchName: branchName,
      studentName: studentName,
      fatherName: fatherName,
      classValue: classValue,
      categoryValue: categoryValue,
      admissionDate: formattedDate,
      staffName: staffName,
      designation: designation,
      staffBranch: staffBranch2,
      staffCode: staffCode2,
      rowIndex: i + 2
    })
  }
  return results
}

function createRecord(dataObj, loggedInUser) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(DATA_SHEET)
  let newDate = dataObj.admissionDate ? new Date(dataObj.admissionDate + 'T00:00:00') : ''
  sheet.appendRow([
    dataObj.admissionNo,
    dataObj.branchName,
    dataObj.studentName,
    dataObj.fatherName,
    dataObj.classValue,
    dataObj.categoryValue,
    newDate,
    dataObj.staffName,
    dataObj.designation,
    dataObj.staffBranch,
    loggedInUser,
    dataObj.staffCode
  ])
  return { success: true, message: 'Record created successfully' }
}

function updateRecord(dataObj, loggedInUser) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(DATA_SHEET)
  const rowIndex = dataObj.rowIndex
  if (!rowIndex) {
    return { success: false, message: 'No rowIndex' }
  }
  let newDate = dataObj.admissionDate ? new Date(dataObj.admissionDate + 'T00:00:00') : ''
  sheet.getRange(rowIndex, 1).setValue(dataObj.admissionNo)
  sheet.getRange(rowIndex, 2).setValue(dataObj.branchName)
  sheet.getRange(rowIndex, 3).setValue(dataObj.studentName)
  sheet.getRange(rowIndex, 4).setValue(dataObj.fatherName)
  sheet.getRange(rowIndex, 5).setValue(dataObj.classValue)
  sheet.getRange(rowIndex, 6).setValue(dataObj.categoryValue)
  sheet.getRange(rowIndex, 7).setValue(newDate)
  sheet.getRange(rowIndex, 8).setValue(dataObj.staffName)
  sheet.getRange(rowIndex, 9).setValue(dataObj.designation)
  sheet.getRange(rowIndex, 10).setValue(dataObj.staffBranch)
  sheet.getRange(rowIndex, 11).setValue(loggedInUser)
  sheet.getRange(rowIndex, 12).setValue(dataObj.staffCode)
  return { success: true, message: 'Record updated successfully' }
}

function deleteRecord(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName(DATA_SHEET)
  if (!rowIndex) {
    return { success: false, message: 'No rowIndex' }
  }
  sheet.deleteRow(rowIndex)
  return { success: true, message: 'Record deleted successfully' }
}
