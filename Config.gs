/**
 * @file Config.gs
 * @description Az alkalmazás globális konfigurációs beállításai.
 */

const APP = Object.freeze({
  NAME: "PayKal - Wallet of Calasanz",
  VERSION: "0.2.0",
  TIMEZONE: "Europe/Budapest",
  SHEETS: Object.freeze({
    USERS: "Felhasználók",
    GROUPS: "Csoportok",
    MEMBERS: "KM Tagok", // Frissítve a pontos munkalap névre!
    CITIES: "Városok",
    SETTINGS: "Beállítások",
    BALANCE: "Egyenleg",
    EXPENSES: "Költségek",
    INCOME: "Bevételek",
    TRANSFERS: "Pénzmozgások",
    SUMMARY: "Összesítő",
    LOG: "Log"
  })
});

const CONFIG = {
  SHEET_ID: "1vjjkOriMEJVKBPzY_k3NNJItmLQySWZTQCCZHiovDYE"
};