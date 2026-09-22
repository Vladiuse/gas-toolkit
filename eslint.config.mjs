import globals from "globals";
import pluginJs from "@eslint/js";
import html from "eslint-plugin-html"


// Глобалы рантайма Apps Script: они существуют всегда, перечислять их
// в директиве /* global */ каждого файла не нужно.
const appsScriptGlobals = {
  SpreadsheetApp: "readonly",
  UrlFetchApp: "readonly",
  PropertiesService: "readonly",
  CacheService: "readonly",
  LockService: "readonly",
  DriveApp: "readonly",
  GmailApp: "readonly",
  MailApp: "readonly",
  CalendarApp: "readonly",
  DocumentApp: "readonly",
  FormApp: "readonly",
  SlidesApp: "readonly",
  HtmlService: "readonly",
  ScriptApp: "readonly",
  Session: "readonly",
  Utilities: "readonly",
  Logger: "readonly",
  Browser: "readonly",
};

/** @type {import('eslint').Linter.Config[]} */
export default [
  {
    files: ["**/*.js"],
    languageOptions: {
      sourceType: "module",
      globals: { ...globals.browser, ...appsScriptGlobals }
    },
  },
  {languageOptions: { globals: { ...globals.browser, ...appsScriptGlobals } }},
  pluginJs.configs.recommended,
  {
    files: ["**/*.html"],
    plugins: { html },
    rules: {
      "no-unused-vars": "warn"
    }
  },
  {
    files: ["**/*.js"],
    rules: {
      "no-unused-vars": "warn",
    }
  }
];