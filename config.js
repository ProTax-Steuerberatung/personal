/* =====================================================================
   Zeiterfassung ProTax – Konfiguration
   Nur diese Datei muss nach der Einrichtung angepasst werden.
   ===================================================================== */
window.ZE_CONFIG = {
  // true  = Demo-Modus: Testdaten im Browser, keine Anmeldung, nichts geht an Microsoft 365
  // false = Echtbetrieb mit Microsoft-Anmeldung und SharePoint-Listen
  demo: true,

  // Aus der App-Registrierung in Microsoft Entra ID (Anleitung Schritt 2)
  tenantId: 'HIER-VERZEICHNIS-ID-EINTRAGEN',
  clientId: 'HIER-ANWENDUNGS-ID-EINTRAGEN',
  // Genau die Adresse, unter der die App erreichbar ist (mit abschließendem /)
  // Leer lassen = aktuelle Adresse wird verwendet
  redirectUri: '',

  // SharePoint-Website, in der die Listen liegen (Anleitung Schritt 1)
  sharepoint: {
    hostname: 'protaxsteuerberatung.sharepoint.com',
    sitePath: '/sites/Zeiterfassung'
  },

  // Leitung wird in der App unter Stammdaten markiert (Haken „Leitung“) – hier nichts eintragen,
  // da diese Datei öffentlich lesbar sein kann (GitHub Pages). Nur als Notfall-Zugang nutzen.
  leitung: [],

  // Feiertage Bayern; Mariä Himmelfahrt gilt nur in Gemeinden mit überwiegend katholischer Bevölkerung
  mariaeHimmelfahrt: true,
  // Halbe Arbeitstage als 'MM-TT', z. B. ['12-24', '12-31'] – nur eintragen, wenn das in der Kanzlei so geregelt ist
  halbeTage: [],

  // Wie viele Tage rückwirkend dürfen Mitarbeitende eigene Zeiten nachtragen/korrigieren? Danach nur die Leitung.
  korrekturFristTage: 7,

  orte: ['Büro', 'Homeoffice', 'Extern / Mandant'],
  firma: 'ProTax Steuerberatung'
};
