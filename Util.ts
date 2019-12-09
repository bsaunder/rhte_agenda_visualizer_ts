namespace Util {

  // Current Field List, In Order, Starting at A
  // TID	Session Source	Approved for regions	Submitted for Region	Title	Description	Americas speaker	Americas speaker email	Americas Lab Assistant	Americas Lab Assistant email	APAC speaker	APAC speaker email	APAC Lab Assistant	APAC Lab Assistant email	EMEA speaker	EMEA speaker email	EMEA Lab Assistant	EMEA Lab Assistant email	Duration	Session type	PnT Pillar	Red Hat Conversation Alignment	Audience	Mobile App Tag	Inc. Tech Demo	Demo Type	Takeaway 1	Takeaway 2	Takeaway 3	Technical Difficulty	Physical/Tech Assets	Preparation or Experience	Speakers	Who owns content?	Content Owner name	Content Owner email	Previously Presented?	Link to past/existing presentation	Approval Status	Publishing Status	Day	Start	Finish	Location	Room	EMEA Slides URL	APAC Slides URL	Americas Slides URL	Notification Phase	Submitter FirstName	Submitter LastName	Submitter Email	Date Created	Date Last Modified
  const EP_ID_LETTER = 'A'
  const EP_REGIONS_LETTER = 'C'
  const EP_TITLE_LETTER = 'E'
  const EP_TYPE_LETTER = 'T'
  const EP_DATE_LETTER = 'AO'
  const EP_START_LETTER = 'AP'
  const EP_END_LETTER = 'AQ'
  const EP_ROOM_LETTER = 'AS'

  const LETTER = 0
  const COLUMN = 1
  const ARRAY = 2

  /*
  * Looks up Column Information for a specific Column in a Data Dump and Returns its Value as the Sheets Letter, Sheets Index, or Array Index.
  */
  export function lookupColumnInfo(columnName: string, columnType: number): any{
    let returnVal = '';
    switch(columnName){
      case 'ID':
        returnVal = EP_ID_LETTER
        break;
      case 'Regions':
        returnVal = EP_REGIONS_LETTER
        break;
      case 'Title':
        returnVal = EP_TITLE_LETTER
        break;
      case 'Type':
        returnVal = EP_TYPE_LETTER
        break;
      case 'Date':
        returnVal = EP_DATE_LETTER
        break;
      case 'Start':
        returnVal = EP_START_LETTER
        break;
      case 'Finish':
        returnVal = EP_END_LETTER
        break;
      case 'Room':
        returnVal = EP_ROOM_LETTER
        break;
      }

    if(columnType == COLUMN){
      // Need Column Number
      returnVal = Util.letterToColumn(returnVal)
    } else if(columnType == ARRAY){
      // Need Array Index
      returnVal = Util.letterToColumn(returnVal) - 1
    }

    return returnVal
  }

  /*
  * Converts Column Numbers to Letters
  */
  export function columnToLetter(column: number): string{
    let temp = ''
    let letter = ''

    while (column > 0)  {
      temp = (column - 1) % 26
      letter = String.fromCharCode(temp + 65) + letter
      column = (column - temp - 1) / 26
    }
    return letter
  }

  /*
  * Converts Column Letters to Numbers for Arrays
  */
  export function letterToColumn(letter: string): number{
    let column = 0
    let length = letter.length

    for (var i = 0 ; i < length ; i++)  {
      column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1)
    }
    return column
  }
}
