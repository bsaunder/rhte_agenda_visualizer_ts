export class Menu {

  function onOpen() {
  // Build Application Menu
  Settings.UI.createMenu('RHTE Tools')
    .addItem('Rebuild Agenda', 'fullRebuild')
    .addSubMenu(Settings.UI.createMenu('Event Point')
      .addItem('Pull Data from Event Point', 'pullEventPointData')
      .addItem('Prepare Data Dump', 'prepareEventPointData'))
    .addSubMenu(Settings.UI.createMenu('Agenda')
      .addItem('Update Agenda from Dump', 'updateAgenda')
      .addItem('Find Free Space', 'findFreeSpace'))
    .addSubMenu(Settings.UI.createMenu('Utilities')
       .addItem('Clear Agendas', 'clearAgendaSheets')
       .addItem('Create Days from Backup Sheet', 'copyFromBackupSheet'))
    .addToUi();
  }

  function fullRebuild(){
    let result = Settings.UI.alert(
       'Warning',
       'This will remove ALL Agenda data rebuild it from the EPDump sheet. This may take several minutes to complete. Proceed?',
        Settings.UI.ButtonSet.YES_NO)

    // Process the user's response.
    if (result == Settings.UI.Button.NO) {
      return
    }

    EventPoint.processMasterSessionGridReport('EPDump')
    Agenda.createAgendaSheets()
    Visualizer.updateAgenda()
    Visualizer.findFreeSpace()

    Settings.UI.alert('Agenda Rebuild Complete.')
  }

  function prepareEventPointData(){
    EventPoint.processMasterSessionGridReport('EPDump')
    Settings.UI.alert('EventPoint Data Ready.')
  }

  function pullEventPointData(){
    //EventPoint.downloadMasterSessionGridReport()
    //Settings.UI.alert('Report Downloaded to EPDump Sheet.')
    Settings.UI.alert('Not Yet Implemented')
  }

  function updateAgenda(){
    let result = Settings.UI.alert(
       'Warning',
       'This may take several minutes to complete. Proceed?',
        Settings.UI.ButtonSet.YES_NO)

    // Process the user's response.
    if (result == Settings.UI.Button.NO) {
      return
    }

    Visualizer.updateAgenda()

    Settings.UI.alert('Agenda Update Complete.')
  }

  function findFreeSpace(){
    Visualizer.findFreeSpace()
    Settings.UI.alert('Free Space marked in Pink.')
  }

  function clearAgendaSheets(){
    Agenda.clearAgendaSheets()
    Settings.UI.alert('Agenda Sheets Cleared.')
  }

  function copyFromBackupSheet(){
    Agenda.createAgendaSheets()
    Settings.UI.alert('Agenda Sheets Created.')
  }

}
