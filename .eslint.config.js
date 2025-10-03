// eslint.config.js
const tsParser = require("@typescript-eslint/parser");
const tsPlugin = require("@typescript-eslint/eslint-plugin");

module.exports = [
  {
    files: ["VITAFilterInjector/**/*.ts"],
    languageOptions: {
      ecmaVersion: 2020,
      sourceType: "module",
      parser: tsParser,
      parserOptions: { project: "./tsconfig.json" }
    },
    plugins: {
      "@typescript-eslint": tsPlugin
    },
    rules: {
      // keep empty for now
    }
  }
];
