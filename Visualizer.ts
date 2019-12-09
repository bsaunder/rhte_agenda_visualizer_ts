namespace Visualizer {

  /*
  * Update the Agenda using the Data from the EP Dump
  */
  export function updateAgenda(){
    // Set Status Tab to In Progress
    let lastRunRange = Settings.SS.getRange('Status!B9')
    let runStatusRange = Settings.SS.getRange('Status!B10')
    let logRange = Settings.SS.getRange('Status!J2')
    logRange.setValue('')
    lastRunRange.setValue('')
    runStatusRange.setValue('In Progress')

    // Process Dump and Update Agenda
    let sheet = Settings.SS.getSheetByName('EPDump')
    let lastRow = sheet.getLastRow()
    let lastCol = sheet.getLastColumn()
    let dataRange = 'EPDump!A2:'+lastCol+lastRow
    let values = Settings.SS.getRange(dataRange).getDisplayValues()

    if (!values) {
    } else {
      let numRows = values ? values.length : 0
      for (r = 0 ; r < numRows ; r++) {
        let row = values[r]
        let titleCol = Util.lookupColumnInfo('Title',2)
        if(row[titleCol] != '' && row[titleCol] != null){
          processRow(row)
        }
      }
    }

    // Set Status Tab to Complete
    lastRunRange.setValue(new Date())
    runStatusRange.setValue('Success')
    logRange.setValue(Logger.getLog())
  }

  /*
  * Process a Single Row from an EP Dump and mark it on the Agenda
  */
  function processRow(record: [string]){
    let col = ''
    let row = ''
    let sheet = ''

    // Find Target Sheet (Day)
    let date = new Date(record[Util.lookupColumnInfo('Date',2)])
    switch(date.getDay()){
      case 1:
        sheet = 'Monday'
        break
      case 2:
        sheet = 'Tuesday'
        break
      case 3:
        sheet = 'Wednesday'
        break
      case 4:
        sheet = 'Thursday'
        break
      case 5:
        sheet = 'Friday'
        break
      default:
        return
    }

    // Find Target Column (Start Time)
    let times = SheetUtil.getTimeList()
    let targetTime = record[Util.lookupColumnInfo('Start',2)]
    let colValue = times[0].indexOf(targetTime)
    if(colValue >= 0){
      // Convert Col Value to Number
      col = Util.columnToLetter(colValue+4)
    }else{
      return
    }

    // Find Target Row (Room)
    let targetRoom = record[Util.lookupColumnInfo('Room',2)]
    let slashIndex = targetRoom.indexOf("/")
    if(slashIndex >= 0){
      targetRoom = targetRoom.substring(0, slashIndex - 1)
    }

    let roomList = SheetUtil.getRoomList()
    for (let i = 0 ; i < roomList.length ; i++) {
      let roomRow = roomList[i]
      if(roomRow[0] == targetRoom){
        row = i+3
        break
      }
    }

    // Update Cell
    let targetRange = sheet+'!'+col+row

    // Merge Cells to the Right
    let mergeEndCol = SheetUtil.getMergeEndCol(col, record)
    let mergedRange = targetRange

    // Check if we need to Merge Down a row as well
    let mergeCells = ''
    if(slashIndex >= 0){
      mergeCells = sheet+'!'+col+row+':'+mergeEndCol+(row+1)
    }else{
      mergeCells = sheet+'!'+col+row+':'+mergeEndCol+row
    }

    let idCol = Util.lookupColumnInfo('ID',2)
    let mergedRange = Settings.SS.getSheetByName(sheet).getRange(mergeCells)
    if(mergedRange.isPartOfMerge()){
      Logger.log('Failed to Process Row (ID: %s) using Target Range: %s - Overlaps Existing Session.', record[idCol], mergedRange.getA1Notation())
      return
    }else{
      mergedRange.merge()
    }

    // Set New Cell Value
    var newValue = record[Util.lookupColumnInfo('Title',2)]
    if(record[idCol] != '') {
      newValue += ' ('+record[idCol]+')'
    }

    // Update Cell
    Settings.SS.getRange(targetRange).setValue(newValue)

    // Format Cell
    SheetUtil.formatCell(mergedRange)

    // Set the Background Color based on the Session Type
    var color = SheetUtil.getBackgroundColor(record)
    mergedRange.setBackground(color)

  }

  /*
  * Finds and Marks Free Space on the Agenda
  */
  export function findFreeSpace(){
    // Get the Time Slot & Room list
    let timeSlots = SheetUtil.getTimeList()
    let roomList = SheetUtil.getRoomList()

    // Loop over Free Space Slots and Mark Free Times
    let totalCount = 0
    for(let i = 0; i < Settings.FREE_SPACE_SLOTS.length; i++){
      let day = Settings.FREE_SPACE_SLOTS[i][0]
      let start = Settings.FREE_SPACE_SLOTS[i][1]
      let duration = Settings.FREE_SPACE_SLOTS[i][2]

      // Find the Column Letter for the Time (Coming from Array Index)
      let timeIndex = timeSlots[0].indexOf(start)
      let timeColumn = '';
      if(timeIndex >= 0){
        timeColumn = Util.columnToLetter(timeIndex+4)
      }

      // Get the Sheet for the Day
      let sheet = Settings.SS.getSheetByName(day)

      // Get the Cells to Check for Sessions Starting
      let endRow = sheet.getLastRow()
      var rangeValues = sheet.getRange(timeColumn+'3:'+timeColumn+endRow).getValues()

      // Loop over each Cell and Check it
      for(r = 0; r < rangeValues.length; r++){
        let cell = rangeValues[r]
        if(cell == ''){
          // Cell is Empty, so lets make sure its ok to mark as Free

          // Get the End Column for this Duration
          let endColumn = SheetUtil.getMergeEndColMinutes(timeColumn, duration)
          let row = r + 3
          let mergeCells = timeColumn+row+':'+endColumn+row
          let mergedRange = sheet.getRange(mergeCells)

          // Skip if Part of a Merge
          let skip = false
          if(mergedRange.isPartOfMerge()){
            skip = true
          }

          // Skip if the Room is in FORBIDDEN_ROOM_LIST
          if(!skip){
            let room = roomList[r].toString()
            let index = Settings.FORBIDDEN_ROOM_LIST.indexOf(room)
            if(index >= 0){
              skip = true
            }
          }

          if(!skip){
            mergedRange.merge()
            mergedRange.setBackground('pink')
            SheetUtil.formatCell(mergedRange)
            totalCount++
          }
        }
      }
    }

    // Set the Overflow Count on the Status Tab
    Settings.SS.getRange('Status!B12').setValue(totalCount)

  }

}
