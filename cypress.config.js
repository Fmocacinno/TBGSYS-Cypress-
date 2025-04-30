const { defineConfig } = require("cypress");

module.exports = defineConfig({
  viewportWidth: 1280,
  viewportHeight: 800,

  e2e: {
    setupNodeEvents(on, config) {
      // implement node event listeners here
    },

    chromeWebSecurity: false,

    // ✅ Memory optimization
    experimentalMemoryManagement: true,   // Enable new memory handling
    numTestsKeptInMemory: 0,              // Don't keep previous test snapshots
  },

  component: {
    devServer: {
      framework: "react",
      bundler: "vite",
    },
  },
});
