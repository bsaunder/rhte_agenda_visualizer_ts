namespace SheetUtil {

  /*
  * Removes Empty Rows from the bottom of the Named Sheet
  */
  export function removeEmptyRows(sheetName: string){
    var sheet = Settings.SS.getSheetByName(sheetName)

    var maxRows = sheet.getMaxRows();
    var lastRow = sheet.getLastRow();

    if (maxRows-lastRow != 0){
      sheet.deleteRows(lastRow+1, maxRows-lastRow);
    }
  }

  /*
  * Gets a list of all of the Times from the Backup Sheet
  */
  export function getTimeList(): [string]{
    var dataRange = 'AgendaDayTemplate!D2:BJ2'
    var times = Settings.SS.getRange(dataRange).getDisplayValues()

    return times
  }

  /*
  * Gets a list of all of the Rooms from the Backup Sheet
  */
  export function getRoomList(): [string]{
    var dataRange = 'AgendaDayTemplate!A3:A'+(Settings.ROOM_COUNT+2)
    var rooms = Settings.SS.getRange(dataRange).getValues()

    return rooms
  }

  /*
  * Applies Standard Format to a Cell Range
  */
  export function formatCell(range: Range){
    range.setBorder(true, true, true, true, true, true, 'Black', null)
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    range.setFontSize(8)
    range.setHorizontalAlignment("center")
    range.setVerticalAlignment("middle")
  }

  /*
  * Gets the End Column for a Row
  */
  export function getMergeEndCol(startCol: string, record: [string]): string{
    var minutes = SheetUtil.getSessionLengthInMinutes(record)
    return SheetUtil.getMergeEndColMinutes(startCol, minutes)
}

  /*
  * Gets the End Column for a Cell Merge based on Duration in Minutes
  */
  export function getMergeEndColMinutes(startCol: string, minutes: number): string{
    let mergeLength = 0
    mergeLength = Math.ceil(minutes / 15)

    if(mergeLength <= 1){
      return startCol
    }else{
      let startColInt = Util.letterToColumn(startCol)
      return Util.columnToLetter(startColInt + (mergeLength - 1))
    }
  }

  /*
  * Computes the Duration in Minutes of a Session based on its Start and End Times
  */
  export function getSessionLengthInMinutes(record): number{
    let dayString = record[Util.lookupColumnInfo('Date',2)]
    let date = new Date(dayString)

    let startString = record[Util.lookupColumnInfo('Start',2)]
    let start = new Date(dayString + " " + startString)

    let endString = record[Util.lookupColumnInfo('Finish',2)]
    let end = new Date(dayString + " " + endString)

    let timeDiff = end - start

    let minutes = parseInt((timeDiff/(1000*60)))

    return minutes
  }

  /*
  * Determines the correct Background Color for a Row
  */
  export function getBackgroundColor(record): string{

    let type = record[Util.lookupColumnInfo('Type',2)]
    let color = '#d7ffff'

    // Is this longer than 2 Hours? If so it probably falls into Internal or Extra Things like Exams
    let minutes = SheetUtil.getSessionLengthInMinutes(record)
    if(minutes > 120){
      color = '#b6d7a8'
    }

    switch(type){
      case 'Presentation':
        color = '#67deff'
        break
      case 'Panel':
        color = '#52b2cd'
        break
      case 'Workshop':
        color = '#518bcd'
        break
      case 'Hands-On Lab':
        color = '#60a5f4'
        break
    }

    var room = record[Util.lookupColumnInfo('Room',2)]

    switch(room){
      // Make Meals Purple
      case 'Kingston Hall':
      case 'Cayman Court Pavilion':
        color = '#b4a7d6'
        break
      // Make Keynotes Bright Aqua
      case 'Grand Caribbean 1 - 12':
        color = '#05c9ff'
        break
    }

    // Session is TBD
    var tbd = record[Util.lookupColumnInfo('Title',2)].indexOf('TBD')
    if(tbd >= 0){
      color = '#ffff00'
    }

    return color
  }

}
