// Edit only this file for SharePoint site/list and static column settings.
window.ATTENDANCE_STATIC_CONFIG = Object.freeze({
  appTitle: 'Attendance',
  demo: {
    enabled: false,
    useDummyData: false,
    dataUrl: 'demo/attendance-demo-data.json'
  },
  sharepoint: {
    siteUrl: 'http://base0028.sites.airnet/minhala2025/coahadam',
    settingsListTitle: 'AttendanceAppSettings',
    settingsItemTitle: 'main',
    sourceListTitle: 'AttendanceMainDb',
    attendanceListTitle: 'AttendanceDailyDb',
    chatListTitle: 'AttendanceChatMessages',
    operationsLogListTitle: 'AttendanceOperationsLog',
    casualtyListTitle: 'AttendanceCasualties',
    autoProvisionLists: true
  },
  workbook: {
    mainSheet: '\u05de\u05e6\u05d1\u05d4',
    dataSheet: 'data',
    auditSheet: 'AttendanceAudit',
    columnCount: 16,
    dropdownSources: {
      O: 'A',
      P: 'B',
      F: 'C'
    }
  },
  defaults: {
    openDates: [],
    adminPasswordHash: '',
    groupColumns: ['A'],
    filterColumns: ['F'],
    editableColumns: ['M', 'N', 'O', 'P'],
    addRowColumns: ['A', 'B', 'C', 'D', 'F', 'M', 'N', 'O', 'P'],
    phonebookColumns: [],
    dropdownConfigs: [],
    crossTabs: [
      { id: 'unit-status', label: '\u05d9\u05d7\u05d9\u05d3\u05d4 \u05de\u05d5\u05dc O', rowColumns: ['A'], colColumns: ['O'] },
      { id: 'service-result', label: '\u05e1\u05d5\u05d2 \u05e9\u05d9\u05e8\u05d5\u05ea \u05de\u05d5\u05dc P', rowColumns: ['F'], colColumns: ['P'] }
    ]
  }
});
