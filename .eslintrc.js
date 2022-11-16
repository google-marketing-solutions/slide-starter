{
  "extends": ["eslint:recommended", "google"],
  "parserOptions": {
    "ecmaVersion": 8,
    "sourceType": "module",
    "ecmaFeatures": {
      "jsx": true,
    }
  },
  "env": {
    "node": true,
    "es6": true,
    "mocha": true,
  },
  "rules": {
    "prefer-promise-reject-errors": "off",
  }
}
