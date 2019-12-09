namespace Agenda {

  /*
  * Clears an Agenda Sheet of all Data
  */
  export function clearAgendaSheets(){
    // Clear Agenda
    var dayList = ['Monday','Tuesday','Wednesday','Thursday','Friday']
    for (i = 0; i < dayList.length; i++){
      var sheetName = Settings.SS.getSheetByName(dayList[i])
      if (sheetName) {
        Settings.SS.getRange(dayList[i]+'!D3:BJ'+(Settings.ROOM_COUNT+2)).clear()
      }
    }
  }

  /*
  * Creates the Monday-Friday Agenda sheets based on the BackupSheet
  */
  export function createAgendaSheets(){
    var dayList = ['Friday','Thursday','Wednesday','Tuesday','Monday']

    // Delete Existing Agenda sheets
    for (i = 0; i < dayList.length; i++){
      var sheet = Settings.SS.getSheetByName(dayList[i])
      if (sheet) {
        Settings.SS.deleteSheet(sheet)
      }
    }

    // Create New sheets from Template
    var templateSheet = Settings.SS.getSheetByName('AgendaDayTemplate')
    for (i = 0; i < dayList.length; i++){
      var newSheet = templateSheet.copyTo(Settings.SS)
      newSheet.setName(dayList[i])
      newSheet.setTabColor('#05c9ff')
      Settings.SS.setActiveSheet(newSheet)
      Settings.SS.moveActiveSheet(3)
    }

  }

}
