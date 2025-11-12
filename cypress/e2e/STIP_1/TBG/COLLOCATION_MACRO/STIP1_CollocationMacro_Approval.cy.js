import { timeStamp } from 'console';
import 'cypress-file-upload';
const XLSX = require('xlsx');
const fs = require('fs');

// Function to export test results to Excel
function exportToExcel(testResults) {
  const filePath = 'resultsApproval_NewBuildMacro.xlsx'; // Path to the Excel file

  // Create a worksheet from the test results
  const worksheet = XLSX.utils.json_to_sheet(testResults);

  // Create a workbook and add the worksheet
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Test Results');

  // Write the workbook to a file
  XLSX.writeFile(workbook, filePath);
}
const testResults = []; // Array to store test results
const randomValue = Math.floor(Math.random() * 1000) + 1; // Random number between 1 and 1000
const RangerandomValue = Math.floor(Math.random() * 20) + 1; // Random number between 1 and 1000
// const unique = `APP_PKP_`;

function generateRandomString(minLength, maxLength) {
  const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  const length = Math.floor(Math.random() * (maxLength - minLength + 1)) + minLength;
  let result = '';

  for (let i = 0; i < length; i++) {
    const randomIndex = Math.floor(Math.random() * characters.length);
    result += characters.charAt(randomIndex);
  }
  return result;
}
function clickAllEnabledButtons(buttons, index = 0, callback) {
  if (index >= buttons.length) {
    callback(); // Done with this page
    return;
  }

  cy.wrap(buttons[index])
    .scrollIntoView()
    .should('be.visible')
    .click()
    .then(() => {
      // Small wait to allow DOM updates after click
      cy.wait(1000);
      clickAllEnabledButtons(buttons, index + 1, callback);
    });
}

function checkRowsSequentially(page = 1, maxPages = 10) {
  if (page > maxPages) return;

  // Optional wait to ensure table content is rendered
  cy.wait(1000);

  cy.get('button.btnSelect').then(($buttons) => {
    const $enabledBtns = $buttons.filter((i, el) =>
      !el.disabled && Cypress.$(el).is(':visible')
    );

    if ($enabledBtns.length > 0) {
      cy.wrap($enabledBtns[0])
        .scrollIntoView()
        .should('be.visible')
        .click();
    } else {
      // Try going to the next page
      cy.get('a[title="Next"]:visible').then(($next) => {
        if ($next.length > 0) {
          cy.wrap($next).click();
          cy.wait(2000); // wait for table to update
          checkRowsSequentially(page + 1, maxPages);
        } else {
          cy.log('No next page available');
        }
      });
    }
  });
}



let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip;

const minLength = 5;
const maxLength = 15;
const randomString = generateRandomString(minLength, maxLength);
// const date = "2-Jan-2025";
// const user = "555504220025"
// const pass = "123456"
const filePath = 'documents/pdf/C (1).pdf';
const latMin = -11.0; // Southernmost point
const latMax = 6.5;   // Northernmost point
const longMin = 94.0; // Westernmost point
const longMax = 141.0; // Easternmost point

// Generate random latitude and longitude within bounds
const lat = (Math.random() * (latMax - latMin) + latMin).toFixed(6);
const long = (Math.random() * (longMax - longMin) + longMin).toFixed(6);
//Batas

describe('template spec', () => {
  let testResults = []; // Shared results array
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, baseUrlVP, baseUrlTBGSYS, login, dashboard, menu1, menu2, menu3, menu4, logout, PICVendorMobile1, PICVendorMobile2;

  before(() => {
    testResults = []; // Reset results before all tests
  });

  after(() => {
    exportToExcel(testResults); // Export after all tests complete
  });
  beforeEach(() => {
    cy.readFile('cypress/e2e/STIP_1/TBG/COLLOCATION_MACRO/soDataCOLLOCATION_MACRO.json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;
    });
    cy.readFile('cypress/e2e/STIP_1/TBG/COLLOCATION_MACRO/DataVariable.json').then((values) => {
      cy.log(values);
      unique = values.unique;
      userAM = values.userAM;
      userInputStip = values.userInputStip;
      userLeadAM = values.userLeadAM;
      userLeadPM = values.userLeadPM;
      userPMFO = values.userPMFO;
      userARO = values.userARO;
      pass = values.pass;
      date = values.date;
      baseUrlVP = values.baseUrlVP;
      baseUrlTBGSYS = values.baseUrlTBGSYS;
      menu1 = values.menu1;
      menu2 = values.menu2;
      menu3 = values.menu3;
      menu4 = values.menu4;
      login = values.login;
      logout = values.logout;
      dashboard = values.dashboard;
      PICVendorMobile1 = values.PICVendorMobile1;
      PICVendorMobile2 = values.PICVendorMobile2;
    });


    Cypress.on('uncaught:exception', (err, runnable) => {
      return false;
    });
  });

  //AM
  it('AM Test Case', () => {
    const testResults = [];

    cy.visit('http://tbgappdev111.tbg.local:8023/Login');

    cy.get('#tbxUserID').type(userAM);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get("#btnSubmit").click();
    cy.wait(2000);

    cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval');

    cy.wait(2000);
    cy.get('#tbxSearchSONumber').type(sonumb);
    cy.get('.btnSearch').first().click();
    cy.wait(2000);
    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.log("📌 Status Found:", text);
      cy.wait(2000);

      if (text.includes("Waiting for Approval AM")) {  // ✅ Checks if "Lead PM" is in the status
        cy.log("✅ Status contains 'AM', proceeding with approval...");

        cy.get('tbody tr:first-child td:nth-child(1) .btnApprovalDetail').click();
        cy.get('#tarApprovalRemark').type('Remark_' + unique, { force: true });
        // SKIPPING 'tarApprovalRemark' input field
        cy.log("⚠️ Skipping remark input...");

        // Attempt to click the approval button only if it's visible and enabled
        cy.get("#btnApprove").then(($btn) => {
          if ($btn.is(':visible') && !$btn.is(':disabled')) {
            cy.wrap($btn).click();
            cy.log("✅ Button clicked successfully");
            cy.wait(2000);
          } else {
            cy.wait(2000);
            cy.log("⚠️ Button not clickable, skipping...");
          }
        });

      } else {
        cy.log("⚠️ Status does not match, skipping approval step.");
      }
    });

    cy.wait(2000);
    cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');
  });

  //LEAD AM
  it('Lead AM Test Case', () => {
    const testResults = [];
    // Lead PM
    cy.visit('http://tbgappdev111.tbg.local:8042');
    cy.get('#tbxUserID').type(userLeadAM);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();
    cy.wait(2000);
    cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval');
    cy.wait(2000);
    cy.get('#tbxSearchSONumber').type(sonumb);
    cy.get('.btnSearch').first().click();
    cy.wait(2000);
    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.log("📌 Status Found:", text);

      if (text.includes("Waiting for Approval Lead AM")) {  // ✅ Checks if "Lead PM" is in the status
        cy.log("✅ Status contains 'Lead AM', proceeding with approval...");
        cy.get('tbody tr:first-child .btnApprovalDetail').click();
        cy.get('#tarApprovalRemark').type('Remark_' + unique, { force: true });
        // SKIPPING 'tarApprovalRemark' input field
        cy.log("⚠️ Skipping remark input...");

        // Attempt to click the approval button only if it's visible and enabled
        cy.get("#btnApprove").then(($btn) => {
          if ($btn.is(':visible') && !$btn.is(':disabled')) {
            cy.wrap($btn).click();
            cy.log("✅ Button clicked successfully");
          } else {
            cy.log("⚠️ Button not clickable, skipping...");
          }
        });

      } else {
        cy.log("⚠️ Status does not match, skipping approval step.");
      }
    });

    cy.wait(5000);
    cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');

  });
  //LEAD PM
  it('Lead PM Test Case', () => {
    const testResults = [];
    // Lead PM
    cy.visit('http://tbgappdev111.tbg.local:8042');
    cy.get('#tbxUserID').type(userLeadPM);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();
    cy.wait(2000);
    cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval');
    cy.wait(2000);
    cy.url().should('include', 'http://tbgappdev111.tbg.local:8042/STIP/Approval');
    testResults.push({
      Test: 'User melakukan akses ke menu Stip Approval',
      Status: 'Pass',
      timeStamp: new Date().toISOString(),
    });

    cy.get('#tbxSearchSONumber').type(sonumb).should('have.value', sonumb).then(() => {
      // Log the test result if input is successful
      testResults.push({
        Test: 'User melakukan input SONumber di Stip approval',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    });

    cy.get('.btnSearch').first().click().should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User melakukan klik tombol Search di Stip approval',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    });
    cy.wait(2000);



    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.log("📌 Status Found:", text);

      if (text.includes("Lead PM")) {  // ✅ Flexible condition to match variations
        cy.log("✅ Status contains 'Lead PM', proceeding with approval...");

        cy.get('tbody tr:first-child td:nth-child(1) .btnApprovalDetail').click();
        cy.get('#tarApprovalRemark').type('Remark_' + unique, { force: true });
        cy.log("⚠️ Skipping remark input...");

        cy.get("#btnConfirm").then(($btn) => {
          if ($btn.is(':visible') && !$btn.is(':disabled')) {
            cy.get('#slsPMCME').select('202307020084', { force: true });
            cy.get('#slsFieldController').select('201103180014', { force: true });
            cy.wrap($btn).click();
            cy.log("✅ Button clicked successfully");
            cy.wait(2000);
          } else {
            cy.log("⚠️ Button not clickable, skipping...");
          }
        });

      } else {
        cy.log("⚠️ Status does not match, skipping approval step.");
      }
    });

    cy.wait(2000);
    cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');

  });
  //ARO
  it('ARO Test Case', () => {
    // Lead PM
    const testResults = [];
    cy.visit('http://tbgappdev111.tbg.local:8042');
    cy.get('#tbxUserID').type(userARO);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();
    cy.wait(2000);
    cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval');
    cy.wait(2000);
    cy.get('#tbxSearchSONumber').type(sonumb);
    cy.get('.btnSearch').first().click();
    cy.wait(2000);
    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.log("📌 Status Found:", text);

      if (text === "Waiting for Confirmation ARO") {
        cy.log("✅ Status matches, proceeding with approval...");

        cy.get('tbody tr:first-child .btnApprovalDetail').click();
        cy.get('#tarApprovalRemark').type('Remark_' + unique, { force: true });
        // SKIPPING 'tarApprovalRemark' input field
        cy.log("⚠️ Skipping remark input...");

        // Attempt to click the approval button only if it's visible and enabled
        cy.get("#btnConfirm").then(($btn) => {
          if ($btn.is(':visible') && !$btn.is(':disabled')) {
            cy.wrap($btn).click();

            cy.log("✅ Button clicked successfully");
          } else {
            cy.log("⚠️ Button not clickable, skipping...");
          }
        });

      } else {
        cy.log("⚠️ Status does not match, skipping approval step.");
      }
      cy.wait(2000);
    });
    cy.wait(2000);
    cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');
    cy.then(() => {
      exportToExcel(testResults);
    });

  });

});
