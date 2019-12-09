namespace EventPoint{

  /*
  * Prepares a Regional Master Session Grid (Approved) report dump for processing
  */
  export function processMasterSessionGridReport(reportSheetName: string){
    var sheet = Settings.SS.getSheetByName(reportSheetName)

    // Freeze the Top Row
    sheet.setFrozenRows(1);

    // Format the Time Columns
    var timeRangeString = reportSheetName+'!'+Util.lookupColumnInfo('Start',0)+':'+Util.lookupColumnInfo('Finish',0)
    var timeRange = sheet.getRange(timeRangeString)
    timeRange.setNumberFormat('h:mm am/pm')

    // Delete Rows Outside of Event Range
    let startDate = new Date(Settings.MONDAY_DATE)
    let endDate = new Date(Settings.FRIDAY_DATE)
    let lastCol = Util.columnToLetter(sheet.getLastColumn())
    let lastRow = sheet.getLastRow()
    let dataRange = sheet.getRange(reportSheetName+'!A2:'+lastCol+lastRow)
    let values = dataRange.getValues()
    let dateIndex = Util.lookupColumnInfo('Date',2)

    for(r = 0; r < values.length; r++){
      let dateVal = values[r][dateIndex]
      let date = new Date(dateVal)
      if(date < startDate || endDate < date || date == "Invalid Date"){
        values.splice(r,1)
        r--
      }
    }

    // Clear Existing Data and Add the Cleaned Data
    dataRange.clear()
    dataRange = sheet.getRange(reportSheetName+'!A2:'+lastCol+(values.length+1))
    dataRange.setValues(values)

    // Remove the Empty Rows from the Bottom of the Sheet
    SheetUtil.removeEmptyRows(reportSheetName)
  }

  /*
  * Downloads the Regional Master Session Grid (Approved) CSV Report from Event Point
  */
  export function downloadMasterSessionGridReport(){

    // Login
    var payload =
     {
       "username" : "api-saunders",
       "password" : "Event99point",
     }
    var options =
     {
       "method" : "post",
       "payload" : payload,
       "followRedirects" : false
     }
    var login = UrlFetchApp.fetch("https://rhte2019.eventpoint.com/SignIn" , options)
    var sessionDetails = login.getAllHeaders()['Set-Cookie']
    Logger.log(sessionDetails)

    // Get CSV
    let csvUrl = Settings.REGIONAL_MSG_APPROVED_REPORT_URL;
    let csvContent = UrlFetchApp.fetch(csvUrl, {"headers" : {"Cookie" : sessionDetails} }).getContentText();
    Logger.log(csvContent)// <-- Currently its not Taking the Cookie or something and just sending back to the Login Page

    // Parse CSV & Set Data to Sheet
    //let csvData = Utilities.parseCsv(csvContent,',');
    //let sheet = Settings.SS.getSheetByName('EPDump')
    //sheet.clear()
    //sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);

    // Logoff
    UrlFetchApp.fetch(Settings.LOGOFF_URL)
  }

}
