const { defineConfig } = require("cypress");

// Flag to toggle debugging mode
const isDebugMode = process.env.CYPRESS_DEBUG === "true";

module.exports = defineConfig({
  viewportWidth: 1280,
  viewportHeight: 800,

  e2e: {
    setupNodeEvents(on, config) {
      // implement node event listeners here if needed
    },
    chromeWebSecurity: false,

    // ✅ Conditional config based on debug mode
    experimentalMemoryManagement: !isDebugMode,
    numTestsKeptInMemory: isDebugMode ? 50 : 1,
  },

  component: {
    devServer: {
      framework: "react",
      bundler: "vite",
    },
  },
});
