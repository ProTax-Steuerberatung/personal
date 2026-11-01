/* Zeiterfassung ProTax – Konfiguration */
window.ZE_CONFIG = {
  // false = Echtbetrieb mit Microsoft-Anmeldung und SharePoint-Listen
  demo: false,

  tenantId: 'c630873a-2003-4e72-975c-9ba6c08254f5',
  clientId: '832dcf0f-0f6b-4a68-9ffc-45f7c4c213c2',
  redirectUri: '',

  sharepoint: {
    hostname: 'protaxsteuerberatung.sharepoint.com',
    sitePath: '/sites/Personal'
  },

  // Leitung wird in der App unter Stammdaten markiert – hier keine Namen eintragen (Datei ist öffentlich)
  leitung: [],

  mariaeHimmelfahrt: true,
  halbeTage: [],
  korrekturFristTage: 7,
  orte: ['Büro', 'Homeoffice', 'Extern / Mandant'],
  firma: 'ProTax Steuerberatung'
};
