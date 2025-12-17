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
    cy.readFile('cypress/e2e/STIP_1/TBG/COLLOCATION_MACRO_FWA/SURGE/CollocationMacroFWASURGE/soDataCOLLOCATION_MACRO_FWA(SURGE).json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;
    });
    cy.readFile('cypress/e2e/STIP_1/TBG/COLLOCATION_MACRO_FWA/SURGE/CollocationMacroFWASURGE/DataVariable.json').then((values) => {
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

  //ARO
  it('ARO Test Case', () => {
    // ARO
    cy.visit(`${baseUrlTBGSYS}${login}`);
    cy.get('#tbxUserID').type(userARO);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();
    cy.wait(3000);
    cy.visit(`${baseUrlTBGSYS}/STIP/Approval`);
    cy.url().should('include', `${baseUrlTBGSYS}/STIP/Approval`);
    cy.wait(3000);
    cy.contains('label', /^\s *By SO Number\s*$/)
      .click(); // search By Radio Button SONumber
    cy.get('#tbxApprovalSONumber').type(sonumb).should('have.value', sonumb).then(() => {
      // Log the test result if input is successful
      testResults.push({
        Test: 'User ARO melakukan input SONumber di Stip approval',
        Status: 'Pass',
        TimeStamp: new Date().toISOString(),
      });
    });

    cy.get('.btnSearch').first().click().should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User ARO melakukan klik tombol Search di Stip approval',
        Status: 'Pass',
        TimeStamp: new Date().toISOString(),
      });
    });
    cy.wait(3000);
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
      cy.wait(3000);
    });
    cy.wait(3000);
    cy.contains('a', 'Log Out').click({ force: true });
    cy.wait(4000);

    cy.then(() => {
      exportToExcel(testResults);
    });

  });

});
