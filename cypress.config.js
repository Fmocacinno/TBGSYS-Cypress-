const { defineConfig } = require("cypress");

module.exports = defineConfig({
  viewportWidth: 1280, // Default width
  viewportHeight: 800, // Default height
  e2e: {
    setupNodeEvents(on, config) {
      // implement node event listeners here
    },
    chromeWebSecurity: false,
  },

  component: {
    devServer: {
      framework: "react",
      bundler: "vite",
    },
  },
});
