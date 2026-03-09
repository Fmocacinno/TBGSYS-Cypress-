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
const filePath = 'documents/pdf/receipt.pdf';
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
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip;

  before(() => {
    testResults = []; // Reset results before all tests
  });

  after(() => {
    exportToExcel(testResults); // Export after all tests complete
  });
  beforeEach(() => {
    cy.readFile('cypress/e2e/STIP_1/NEW_BUILD_MACRO/soDataNewBuild.json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;
    });

    cy.readFile('cypress/e2e/STIP_1/NEW_BUILD_MACRO/DataVariable.json').then((values) => {
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

    });


    Cypress.on('uncaught:exception', (err, runnable) => {
      return false;
    });
  });
  const loopCount = 1; // Jumlah iterasi loop

  for (let i = 0; i < loopCount; i++) {
    it.only(`passes iteration ${i + 1}`, () => {

      cy.visit('http://tbgappdev111.tbg.local:8127/Login')

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
          Test: 'Passworhasben ',
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
      cy.url().should('include', 'http://tbgappdev111.tbg.local:8127/Dashboard'); // Ensure the page changes or some result occurs
      testResults.push({
        Test: 'Button Clicked',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
      cy.wait(2000)

      cy.visit('http://tbgappdev111.tbg.local:8127/STIP/Input')
      cy.url().should('include', 'http://tbgappdev111.tbg.local:8127/STIP/Input'); // Ensure the page changes or some result occurs
      testResults.push({
        Test: 'User masuk ke Page Stip Input',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
      cy.wait(2000)

      cy.get('#slsSTIPCategory').then(($select) => {
        cy.wrap($select).select('1', { force: true })
      })

      cy.get('#slsProduct').then(($select) => {
        cy.wrap($select).select('18', { force: true })
      })

      cy.get('#slsNewCompany').then(($select) => {
        cy.wrap($select).select('PKP', { force: true })
      })

      cy.get('#slsNewCustomer').then(($select) => {
        cy.wrap($select).select('XL', { force: true })
      })

      cy.get('#slsNewRegion').then(($select) => {
        cy.wrap($select).select('1', { force: true })
      })

      cy.get('#btnNewPriceAmountPopUp').click();

      cy.get('tbody tr:first-child .btnSelect').click();

      cy.wait(2000)

      // handle HTML

      cy.get('.slsBatchSLD').eq(0)
        .select('180', { force: true });

      cy.get('.slsBatchSLD').eq(1)
        .select('180', { force: true });


      cy.get('.slsBatchSLD').eq(2)
        .select('180', { force: true });

      cy.get('.slsBatchSLD').eq(3)
        .select('180', { force: true });

      cy.get('.slsBatchSLD').eq(4)
        .select('180', { force: true });


      cy.get('.slsBatchSLD').eq(5)
        .select('180', { force: true });

      cy.get('.slsBatchSLD').eq(6)
        .select('180', { force: true });

      cy.get('.slsBatchSLD').eq(7)
        .select('180', { force: true });

      cy.get('.slsBatchSLD').eq(8)
        .select('180', { force: true });
      // handle HTML

      cy.get('#tbxNewSiteName').type('Site_' + unique + RangerandomValue);
      cy.get('#tbxNewCustomerSiteID').type('Cust_' + unique + RangerandomValue);

      cy.get('#slsNewDocumentOrder').then(($select) => {
        cy.wrap($select).select('7', { force: true })
      })

      cy.get('#tbxNewDocumentName').type('DoctName_' + unique + RangerandomValue);

      cy.get('#fleNewDocument').attachFile(filePath);

      cy.get('#slsNewProvince').then(($select) => {
        cy.wrap($select).select('11', { force: true })
      })
      cy.wait(2000)
      cy.get('#slsNewResidence').then(($select) => {
        cy.wrap($select).select('178', { force: true })
      })


      cy.get('#tbxNewNomLatitude').type(lat);
      cy.get('#tbxNewNomLongitude').type(long);

      cy.get('#slsNewLeadProjectManager').then(($select) => {
        cy.wrap($select).select(userLeadPM, { force: true })
      })
      cy.get('#slsNewAccountManager').then(($select) => {
        cy.wrap($select).select(userLeadAM, { force: true })
      })

      cy.get('#slsNewTowerHeight').then(($select) => {
        cy.wrap($select).select('32', { force: true })
      })

      cy.get('#slsNewShelterType').then(($select) => {
        cy.wrap($select).select('3', { force: true })
      })
      cy.get('#tbxNewPLNPowerKVA').type(5);

      cy.get('#tbxNewNumberOfAntenna').type(RangerandomValue);
      cy.get('#tbxNewNumberOfSectoral').type(RangerandomValue);



      cy.get('#dpkNewRFITarget')
        .invoke('val', date)
        .trigger('change');

      cy.get('#tbxNewLeasePeriod').type(RangerandomValue);



      cy.get('#slsNewMLANumber').then(($select) => {
        cy.wrap($select).select('0031-14-F07-121782', { force: true })
      })

      cy.get('#tbxNewLeasePeriod').type(5);

      cy.get('#dpkNewMLADate')
        .invoke('val', date)
        .trigger('change');

      cy.get('#tarNewRemark').type('Remark' + randomString);


      cy.get("#btnSubmitNew").click();

      cy.wait(2000)

      cy.get('.sa-confirm-button-container button.confirm').click();

      cy.wait(15000)

      // Add this section to extract values from popup
      cy.get('p.lead.text-muted').should('be.visible').then(($el) => {
        const text = $el.text();

        // Extract SO Number using regex
        const soNumber = text.match(/\bSO Number = (\d+)\b/)[1];
        // Extract Site ID using regex
        const siteId = text.match(/\bSite ID = (\d+)\b/)[1];

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
            const filePath = Cypress.config('fileServerFolder') + '/cypress/e2e/STIP_1/MCP_NON_FO/soDataMCPNONFO.json';
            cy.writeFile(filePath, { soNumber, siteId });

          });
        });
        // Add your logic here using the Site ID
      });


      cy.visit('http://tbgappdev111.tbg.local:8127/Login/Logout')
      cy.then(() => {
        exportToExcel(testResults);
      });

    })
  }
})