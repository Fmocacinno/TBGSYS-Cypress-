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
    cy.readFile('cypress/e2e/STIP_8/TBG/FTTH_SURVEY/soDataSurvey.json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;
    });

    cy.readFile('cypress/e2e/STIP_8/TBG/FTTH_SURVEY/DataVariable.json').then((values) => {
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
  const loopCount = 1; // Jumlah iterasi loop

  for (let i = 0; i < loopCount; i++) {
    it.only(`passes iteration ${i + 1}`, () => {
      // cy.visit('http://tbgappdev111.tbg.local:8127/Login')
      cy.visit(`${baseUrlTBGSYS}${login}`);
      // cy.get('#tbxUserID').type(userInputStip);
      // cy.get('#tbxPassword').type(pass);


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
      cy.visit(`${baseUrlTBGSYS}${dashboard}`); // Ensure the page changes or some result occurs
      testResults.push({
        Test: 'Button Clicked',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });

      // Export results to Excel after the test




      cy.wait(2000)

      cy.visit(`${baseUrlTBGSYS}/STIP/SiteSurvey`);
      cy.url().should('include', `${baseUrlTBGSYS}/STIP/SiteSurvey`);
      // Ensure the page changes or some result occurs
      testResults.push({
        Test: 'User masuk ke Page Stip Input',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });

      cy.wait(4000)
      cy.get('#slsSiteSurveyCompany').select('TB', { force: true });
      cy.wait(2000);
      cy.get('#slsSiteSurveyCustomer').select('XL', { force: true });
      cy.wait(2000);
      cy.get('#slsSiteSurveyRegion').select('1', { force: true });
      cy.wait(2000);
      cy.get('#slsSiteSurveyProvince').select('11', { force: true });
      cy.wait(2000);
      cy.get('#slsSiteSurveyResidence').select('178', { force: true });
      cy.wait(2000);
      cy.get('#tbxSiteSurveyNomLatitude').type(lat);
      cy.wait(2000);
      cy.get('#tbxSiteSurveyNomLongitude').type(long);
      cy.wait(2000);
      cy.get('input[name="rdoSiteSurveyProjectType"][value="2"]').parent().click()
      cy.wait(2000);
      cy.get('#tbxSiteSurveySiteName').type('Site_' + unique + randomValue);
      cy.get('#tbxSiteSurveyCustomerSiteID').type('Site_' + unique + randomValue);
      cy.get('#tbxSiteSurveyTargetHomepass').type(RangerandomValue);
      cy.wait(2000);
      cy.get('#slsSiteSurveyBatch2').select('UMU - Q3 AOP 2024 Batch 2', { force: true });
      //UMU - Q3 AOP 2024 Batch 2
      cy.get('#slsSiteSurveyPIC').select('Project', { force: true });
      cy.wait(2000);
      cy.get('#slsSiteSurveyLeadProjectManager').select(userLeadPM, { force: true });
      cy.wait(2000);
      cy.get('#slsSiteSurveyAccountManager').select(userLeadAM, { force: true });
      cy.get('#tarSiteSurveyRemark').type('REMARK NIH DARY ' + randomString);
      cy.wait(2000);
      cy.get('#btnSubmitSiteSurvey').click();

      cy.wait(2000);
      cy.get('.sa-confirm-button-container button.confirm').click();
      cy.wait(10000);

      // Add this section to extract values from popup
      cy.get('p.lead.text-muted').should('be.visible').then(($el) => {
        const text = $el.text();

        // Extract SO Number using regex
        const soNumber = text.match(/\bSO Number Survey = (\d+)\b/)[1];
        // Extract Site ID using regex
        const siteId = text.match(/\bSite ID Survey = (\d+)\b/)[1];

        // Save to aliases for later use
        cy.wrap(soNumber).as('soNumber');
        cy.wrap(siteId).as('siteId');

        // Optional: log the values to Cypress console
        cy.log(`Captured SO Number: ${soNumber}`);
        cy.log(`Captured Site ID: ${siteId}`);
      });

      // To use these values later in the test or other tests:
      cy.get('@soNumber').then((soNumber) => {
        cy.log(`Using SO Number: ${soNumber}`);
        // Add your logic here using the SO Number
      });

      cy.get('@siteId').then((siteId) => {
        cy.log(`Using Site ID: ${siteId}`);

        cy.get('@soNumber').then((soNumber) => {
          cy.get('@siteId').then((siteId) => {

            // const backhaulFilePath = Cypress.config('fileServerFolder') + '/cypress/e2e/STIP_4/PKP/FTTH_BACKHAUL/DataVariable.json';
            // const newBuildFilePath = Cypress.config('fileServerFolder') + '/cypress/e2e/STIP_4/PKP/FTTH_NEWBUILDOLT/DataVariable.json';
            const filePath = Cypress.config('fileServerFolder') + '/cypress/e2e/STIP_8/TBG/FTTH_SURVEY/soDataSurvey.json';
            // cy.readFile(backhaulFilePath).then((backhaulData) => {
            //   backhaulData.SONumberBackhaul = soNumber; // ✅ Update the field
            //   cy.writeFile(backhaulFilePath, backhaulData);
            //   cy.log('✅ Updated SONumberBackhaul in FTTH_BACKHAUL DataVariable.json');
            // });

            // cy.readFile(newBuildFilePath).then((newBuildData) => {
            //   newBuildData.SONumberBackhaul = soNumber; // ✅ Update the field
            //   cy.writeFile(newBuildFilePath, newBuildData);
            //   cy.log('✅ Updated SONumberBackhaul in FTTH_NEWBUILDOLT DataVariable.json');
            // });




            cy.writeFile(filePath, { soNumber, siteId });

          });
        });

        // Add your logic here using the Site ID
      });

      cy.get('@soNumber').then((soNumber) => {
        cy.get('@siteId').then((siteId) => {
          const filePath = Cypress.config('fileServerFolder') + '/cypress/e2e/STIP_4/PKP/FTTH_BACKHAUL/soDataBackhaul.json';
          const backhaulFilePath = Cypress.config('fileServerFolder') + '/cypress/e2e/STIP_4/PKP/FTTH_BACKHAUL/DataVariable.json';
          const newBuildFilePath = Cypress.config('fileServerFolder') + '/cypress/e2e/STIP_4/PKP/FTTH_NEWBUILDOLT/DataVariable.json';
          // Read, update, and write for FTTH_BACKHAUL
          cy.readFile(backhaulFilePath).then((backhaulData) => {
            backhaulData.SONumberBackhaul = soNumber; // ✅ Update the field
            cy.writeFile(backhaulFilePath, backhaulData);
            cy.log('✅ Updated SONumberBackhaul in FTTH_BACKHAUL DataVariable.json');
          });

          cy.readFile(newBuildFilePath).then((newBuildData) => {
            newBuildData.SONumberBackhaul = soNumber; // ✅ Update the field
            cy.writeFile(newBuildFilePath, newBuildData);
            cy.log('✅ Updated SONumberBackhaul in FTTH_NEWBUILDOLT DataVariable.json');
          });
          cy.writeFile(filePath, { soNumber, siteId });

        });
      });
      cy.contains('a', 'Log Out').click({ force: true });
      cy.then(() => {
        exportToExcel(testResults);
      });
    })
  }
})