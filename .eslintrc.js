module.exports = {
  extends: ["eslint:recommended", "google"],
  env: {
    browser: true,
    es2021: true,
    es6: true,
    mocha: true,
  },
  overrides: [],
  parserOptions: {
    ecmaVersion: "latest",
  },
  rules: {
    "no-undef": "off",
  },
};
