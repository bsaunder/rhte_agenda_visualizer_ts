export class Settings {

  // App Script Globals
  public static UI = SpreadsheetApp.getUi()
  public static SS = SpreadsheetApp.getActiveSpreadsheet()

  // General Constants
  public static ROOM_COUNT = 60

  // Date Constants
  public static MONDAY_DATE = '10/14/2019'
  public static TUESDAY_DATE = '10/15/2019'
  public static WEDNESDAY_DATE = '10/16/2019'
  public static THURSDAY_DATE = '10/17/2019'
  public static FRIDAY_DATE = '10/18/2019'

  // EventPoint Constants
  public static REGIONAL_MSG_APPROVED_REPORT_URL = 'https://rhte2019.eventpoint.com/admin/reports/report.aspx?Format=csv&ReportID=60b4a08b-6aba-4e46-b17d-b74f49994051'
  public static QUICK_LOGIN_URL = 'https://rhte2019.eventpoint.com/ws/login.aspx?qk=NB24c6a6f9edbff41038263de8d66d29894'
  public static LOGOFF_URL = 'https://rhte2019.eventpoint.com/WS/Logoff.aspx'

  // Free Space Days, Start Times, and Durations
 public static FREE_SPACE_SLOTS = [
   ['Tuesday', '12:30 PM', 45],
   ['Tuesday', '1:30 PM', 45],
   ['Tuesday', '3:00 PM', 45],
   ['Tuesday', '4:00 PM', 45],
   ['Wednesday', '12:30 PM', 45],
   ['Wednesday', '1:30 PM', 45],
   ['Wednesday', '3:00 PM', 45],
   ['Wednesday', '4:00 PM', 45],
   ['Thursday', '9:30 AM', 45],
   ['Thursday', '12:30 PM', 45],
   ['Thursday', '1:30 PM', 45],
   ['Thursday', '3:00 PM', 45],
   ['Thursday', '10:30 AM', 45]
 ]

 public static FORBIDDEN_ROOM_LIST = [
   'Grand Caribbean 1 - 12',
   'Kingston Hall',
   'Cayman Court Pavilion',
   'Admiralty Room',
   'Barbados Boardroom',
   'SF Planner Office 1 - 3',
   'St. Croix 1',
   'St. Croix 2',
   'St. Croix 3',
   'RP Planner Office 1 - 3',
   'RP Kitchen',
   'Mobile Portfolio Center'
 ]

}
