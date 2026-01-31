import { timeStamp } from 'console';

import 'cypress-file-upload';
const XLSX = require('xlsx');
const fs = require('fs');

// Function to export test results to Excel
function exportToExcel(testResults) {
  const filePath = 'resultsApproval_Assignmaintenace.xlsx'; // Path to the Excel file

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
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, baseUrlVP, baseUrlTBGSYS, login, dashboard, menu1, menu2, menu3, menu4, logout, PICVendorMobile1, PICVendorMobile2, baseUrlCompass, Company;

  before(() => {
    testResults = []; // Reset results before all tests
  });

  after(() => {
    exportToExcel(testResults); // Export after all tests complete
  });
  beforeEach(() => {
    cy.readFile('cypress/e2e/ASSET/TBG/MODUL_PV/TOWER/TSEL/soDataMaintenance.json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;
    });
    cy.readFile('cypress/e2e/ASSET/TBG/MODUL_PV/TOWER/TSEL/DataVariable.json').then((values) => {
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
      baseUrlCompass = values.baseUrlCompass;
      Company = values.Company;
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
  const loopCount = 1; // Jumlah iterasi loop

  for (let i = 0; i < loopCount; i++) {
    it(`THiS FOR CHECK ALL FILTER ${i + 1}`, () => {

      cy.visit(`${baseUrlTBGSYS}${login}`);

      cy.get('#tbxUserID').type(userInputStip).should('have.value', userInputStip).then(() => {
        // Log the test result if input is successful
        testResults.push({
          Test: 'User ID Input',
          Status: 'Pass',
          Timestamp: new Date().toISOString(),
        });
      });

      cy.get('#tbxPassword').type(pass).should('have.value', pass).then(() => {
        // Log the test result if input is successful
        testResults.push({
          Test: 'Password has been inputed',
          Status: 'Pass',
          Timestamp: new Date().toISOString(),
        });
      });

      cy.get('#RefreshButton').click();

      cy.window().then((window) => {
        const rightCode = window.rightCode;
        cy.log('Right Code:', rightCode);

        cy.get('#captchaInsert').type(rightCode);
      });

      cy.get('#btnSubmit').click();
      cy.wait(5000)
      cy.visit(`${baseUrlTBGSYS}${dashboard}`); // Ensure the page changes or some result occurs
      testResults.push({
        Test: 'Button Clicked',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });

      // Export results to Excel after the test
      cy.wait(2000)
      cy.visit(`${baseUrlTBGSYS}/pv/master/maintenance`);
      cy.url().should('include', `${baseUrlTBGSYS}/pv/master/maintenance`);
      // Ensure the page changes or some result occurs
      testResults.push({
        Test: 'User masuk ke Page PV Master List Maintenance',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
      cy.wait(2000)

      // Check label "Company" is shown (visible)
      cy.contains('label.control-label', 'Company')
        .should('be.visible');

      // (optional) log to your testResults (soft)
      cy.contains('label.control-label', 'Company').then(($el) => {
        const ok = Cypress.$($el).is(':visible');
        testResults.push({
          Test: 'Label Company is shown',
          Status: ok ? 'Pass' : 'Fail',
          Timestamp: new Date().toISOString(),
          ...(ok ? {} : { Error: 'Company label not visible' }),
        });
      });
      // Collect all Company options
      cy.get('#ddlCompany').then(($select) => {
        const options = [...$select.find('option')];

        options.forEach((opt) => {
          const value = opt.value;
          const text = opt.textContent.trim();

          // Optional: skip empty option if you want
          // if (value === "") return;

          const isValid =
            value !== undefined &&
            text !== "";

          testResults.push({
            Test: `Company filter option exists: ${text}`,
            Status: isValid ? 'Pass' : 'Fail',
            Timestamp: new Date().toISOString(),
            ...(isValid ? {} : { Error: 'Invalid company option' }),
          });
        });
      });



      // Add this section to extract values from popup
      // cy.get('p.lead.text-muted').should('be.visible').then(($el) => {
      //   const text = $el.text();

      //   // Extract SO Number using regex
      //   const soNumber = text.match(/\bSO Number = (\d+)\b/)[1];
      //   // Extract Site ID using regex
      //   const siteId = text.match(/\bSite ID = (\d+)\b/)[1];

      //   // Save to aliases for later use
      //   cy.wrap(soNumber).as('soNumber');
      //   cy.wrap(siteId).as('siteId');

      //   // Optional: log the values to Cypress console
      //   cy.log(`Captured SO Number: ${soNumber}`);
      //   cy.log(`Captured Site ID: ${siteId}`);
      // });

      // To use these values later in the test or other tests:
      // cy.get('@soNumber').then((soNumber) => {
      //   cy.log(`Using SO Number: ${soNumber}`);
      //   // Add your logic here using the SO Number
      // });

      // cy.get('@siteId').then((siteId) => {
      //   cy.log(`Using Site ID: ${siteId}`);

      //   cy.get('@soNumber').then((soNumber) => {
      //     cy.get('@siteId').then((siteId) => {
      //       const filePath = Cypress.config('fileServerFolder') + '/cypress/e2e/STIP_1/TBG/INTERSITE_FO/soDataIntersiteFo.json';
      //       cy.writeFile(filePath, { soNumber, siteId });

      //     });
      //   });
      // Add your logic here using the Site ID
      // });

    })
    it(`THIS FLOW FOR ASSIGN MAINTENANCE ${i + 1}`, () => {

      cy.visit(`${baseUrlTBGSYS}${login}`);

      cy.get('#tbxUserID').type(userInputStip).should('have.value', userInputStip).then(() => {
        // Log the test result if input is successful
        testResults.push({
          Test: 'User ID Input',
          Status: 'Pass',
          Timestamp: new Date().toISOString(),
        });
      });

      cy.get('#tbxPassword').type(pass).should('have.value', pass).then(() => {
        // Log the test result if input is successful
        testResults.push({
          Test: 'Password has been inputed',
          Status: 'Pass',
          Timestamp: new Date().toISOString(),
        });
      });

      cy.get('#RefreshButton').click();

      cy.window().then((window) => {
        const rightCode = window.rightCode;
        cy.log('Right Code:', rightCode);

        cy.get('#captchaInsert').type(rightCode);
      });

      cy.get('#btnSubmit').click();
      cy.wait(5000)
      cy.visit(`${baseUrlTBGSYS}${dashboard}`); // Ensure the page changes or some result occurs
      testResults.push({
        Test: 'Button Clicked',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });

      // Export results to Excel after the test
      cy.wait(2000)
      cy.visit(`${baseUrlTBGSYS}/pv/master/maintenance`);
      cy.url().should('include', `${baseUrlTBGSYS}/pv/master/maintenance`);
      // Ensure the page changes or some result occurs
      testResults.push({
        Test: 'User masuk ke Page PV Master List Maintenance',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
      cy.wait(2000)

      // Check label "Company" is shown (visible)
      // cy.contains('label.control-label', 'Company')
      //   .should('be.visible');

      // // (optional) log to your testResults (soft)
      // cy.contains('label.control-label', 'Company').then(($el) => {
      //   const ok = Cypress.$($el).is(':visible');
      //   testResults.push({
      //     Test: 'Label Company is shown',
      //     Status: ok ? 'Pass' : 'Fail',
      //     Timestamp: new Date().toISOString(),
      //     ...(ok ? {} : { Error: 'Company label not visible' }),
      //   });
      // });
      cy.get('#ddlCompany').then(($select) => {
        cy.wrap($select).select('TB', { force: true })
      })


      cy.contains('label.mt-radio', 'No').then(($label) => {
        const visible = $label.is(':visible');

        if (visible) {
          cy.wrap($label).click();
        }

        testResults.push({
          Test: 'Radio "No" can be clicked via label',
          Status: visible ? 'Pass' : 'Fail',
          Timestamp: new Date().toISOString(),
          ...(visible ? {} : { Error: 'Radio label not visible' }),
        });
      });
      cy.get('#btMtSearch').then(($btn) => {
        const ok = $btn.is(':visible') && !$btn.is(':disabled');

        if (ok) {
          cy.wrap($btn).click();
        }

        testResults.push({
          Test: 'Search button can be clicked',
          Status: ok ? 'Pass' : 'Fail',
          Timestamp: new Date().toISOString(),
          ...(ok ? {} : { Error: 'Search button not clickable' }),
        });
      });
      cy.get('#tblMasterlistMaintenance_processing', { timeout: 30000 })
        .should('not.be.visible')
        .then(() => {
          testResults.push({
            Test: 'Maintenance table loading spinner disappears',
            Status: 'Pass',
            Timestamp: new Date().toISOString(),
          });
        });

      cy.wait(6000)


      // Add this section to extract values from popup
      // cy.get('p.lead.text-muted').should('be.visible').then(($el) => {
      //   const text = $el.text();

      //   // Extract SO Number using regex
      //   const soNumber = text.match(/\bSO Number = (\d+)\b/)[1];
      //   // Extract Site ID using regex
      //   const siteId = text.match(/\bSite ID = (\d+)\b/)[1];

      //   // Save to aliases for later use
      //   cy.wrap(soNumber).as('soNumber');
      //   cy.wrap(siteId).as('siteId');

      //   // Optional: log the values to Cypress console
      //   cy.log(`Captured SO Number: ${soNumber}`);
      //   cy.log(`Captured Site ID: ${siteId}`);
      // });

      // To use these values later in the test or other tests:
      // cy.get('@soNumber').then((soNumber) => {
      //   cy.log(`Using SO Number: ${soNumber}`);
      //   // Add your logic here using the SO Number
      // });

      // cy.get('@siteId').then((siteId) => {
      //   cy.log(`Using Site ID: ${siteId}`);

      //   cy.get('@soNumber').then((soNumber) => {
      //     cy.get('@siteId').then((siteId) => {
      //       const filePath = Cypress.config('fileServerFolder') + '/cypress/e2e/STIP_1/TBG/INTERSITE_FO/soDataIntersiteFo.json';
      //       cy.writeFile(filePath, { soNumber, siteId });

      //     });
      //   });
      // Add your logic here using the Site ID
      // });
      cy.contains('a', 'Log Out').click({ force: true });
      cy.then(() => {
        exportToExcel(testResults);
      });
    })


  }
})

