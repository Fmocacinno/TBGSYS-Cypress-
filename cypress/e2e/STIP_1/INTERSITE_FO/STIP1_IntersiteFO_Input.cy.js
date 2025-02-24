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
    cy.readFile('cypress/e2e/STIP_1/INTERSITE_FO/soDataIntersiteFO.json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;
    });

    cy.readFile('cypress/e2e/STIP_1/INTERSITE_FO/DataVariable.json').then((values) => {
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

      cy.get('#tbxUserID').type(userInputStip);
      cy.get('#tbxPassword').type(pass);


      // cy.get('#tbxUserID').type(userLeadAM).should('have.value', userLeadAM).then(() => {
      //   // Log the test result if input is successful
      //   testResults.push({
      //     Test: 'User ID Input',
      //     Status: 'Pass',
      //     Timestamp: new Date().toISOString(),
      //   });
      // });

      // cy.get('#tbxPassword').type(pass).should('have.value', pass).then(() => {
      //   // Log the test result if input is successful
      //   testResults.push({
      //     Test: 'Password has been inputed',
      //     Status: 'Pass',
      //     Timestamp: new Date().toISOString(),
      //   });
      // });

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

      // Export results to Excel after the test




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
        cy.wrap($select).select('83', { force: true })
      })

      cy.get('#slsIntersiteFOCompany').then(($select) => {
        cy.wrap($select).select('TB', { force: true })
      })

      cy.get('#slsIntersiteFOCustomer').then(($select) => {
        cy.wrap($select).select('XL', { force: true })
      })

      cy.get('#slsIntersiteFORegion').then(($select) => {
        cy.wrap($select).select('1', { force: true })
      })
      cy.wait(1000); // Ensure dropdown selection is applied
      /// radio button with regex
      cy.contains('label', /^\s *Segment\s*$/)
        .click(); // Click the label

      // Assert that the "Segment" radio button is selected
      cy.get('input[type="radio"][value="Segment"]').should('be.checked');

      cy.get('#btnIntersiteFOPriceAmountPopUp').click();

      cy.get('tbody tr:first-child .btnSelect').click();

      cy.wait(2000)

      // handle HTML

      // cy.get('.slsBatchSLD').eq(0)
      //   .select('180', { force: true });

      // cy.get('.slsBatchSLD').eq(1)
      //   .select('180', { force: true });


      // cy.get('.slsBatchSLD').eq(2)
      //   .select('180', { force: true });

      // cy.get('.slsBatchSLD').eq(3)
      //   .select('180', { force: true });

      // cy.get('.slsBatchSLD').eq(4)
      //   .select('180', { force: true });


      // cy.get('.slsBatchSLD').eq(5)
      //   .select('180', { force: true });

      // cy.get('.slsBatchSLD').eq(6)
      //   .select('180', { force: true });

      // cy.get('.slsBatchSLD').eq(7)
      //   .select('180', { force: true });

      // cy.get('.slsBatchSLD').eq(8)
      //   .select('180', { force: true });
      // handle HTML

      cy.get('#tbxIntersiteFOSiteName').type('Site_' + unique + RangerandomValue);
      cy.get('#tbxIntersiteFOCustomerSiteID').type('Cust_' + unique + RangerandomValue);
      cy.get('#tbxIntersiteFOSPKWOLOINumber').type('WO_' + unique + RangerandomValue);

      cy.get('#slsIntersiteFODocumentOrder').then(($select) => {
        cy.wrap($select).select('7', { force: true })
      })

      cy.get('#tbxIntersiteFODocumentName').type('DoctName_' + unique + RangerandomValue);

      cy.get('#fleIntersiteFODocument').attachFile(filePath);

      cy.get('#slsIntersiteFOProvince').then(($select) => {
        cy.wrap($select).select('11', { force: true })
      })
      cy.wait(2000)
      // Near ENd
      cy.get('#slsIntersiteFOResidence').then(($select) => {
        cy.wrap($select).select('178', { force: true })
      })
      cy.get('#slsIntersiteFOTowerProviderNearEnd').then(($select) => {
        cy.wrap($select).select('10', { force: true })
      })
      cy.get('#tbxIntersiteFODocumentName').type('DoctName_' + unique + RangerandomValue);
      cy.get('#tbxIntersiteFOTowerIDNearEnd').type('Site_' + unique + RangerandomValue);
      cy.get('#tbxIntersiteFOTowerNameNearEnd').type('SiteName_' + unique + RangerandomValue);
      cy.get('#tbxLatitudeA').type(lat);
      cy.get('#tbxLongitudeA').type(long);

      // Far end
      cy.get('#slsIntersiteFOResidence').then(($select) => {
        cy.wrap($select).select('178', { force: true })
      })
      cy.get('#slsIntersiteFOTowerProviderFarEnd').then(($select) => {
        cy.wrap($select).select('11', { force: true })
      })
      cy.get('#tbxIntersiteFOTowerIDFarEnd').type('Site_' + unique + RangerandomValue);
      cy.get('#tbxIntersiteFOTowerNameFarEnd').type('SiteName_' + unique + RangerandomValue);
      cy.get('#tbxLatitudeB').type(lat);
      cy.get('#tbxLongitudeB').type(long);

      cy.get('#slsIntersiteFOLeadProjectManager').then(($select) => {
        cy.wrap($select).select(userLeadPM, { force: true })
      })
      cy.wait(2000)

      cy.get('#slsIntersiteFOAccountManager').then(($select) => {
        cy.wrap($select).select(userLeadAM, { force: true })
      })

      cy.get('#tbxFOCore').type(RangerandomValue);
      cy.get('#tbxFOLengthSPK').type(RangerandomValue);
      cy.get('#tbxIntersiteFOSegment').type('Segment_' + unique + RangerandomValue);

      // Debug: Log all labels inside the container


      // // cy.get('input[name=""][value="0"]').check({ force: true });
      // cy.get('input[type="Radio"][value="0"]').check().should('be.checked')

      cy.wait(1000); // Ensure dropdown selection is applied
      /// radio button with regex
      // cy.get('input[type="radio"]').filter((index, el) => /No\s*/i.test(el.value)).check()

      // input value No


      cy.get('input[name="rdoOverlapping"][value="0"]').parent().click()

      cy.get('#dpkIntersiteFORFSTarget')
        .invoke('val', date)
        .trigger('change');

      cy.get('#tbxIntersiteFOLeasePeriod').type(RangerandomValue);

      cy.get('#dpkIntersiteFOMLADate')
        .invoke('val', date)
        .trigger('change');

      cy.get('#slsIntersiteFOMLANumber').then(($select) => {
        cy.wrap($select).select('0010-14-F07-39033', { force: true })
      })


      cy.get('#dpkIntersiteFOMLADate')
        .invoke('val', date)
        .trigger('change');

      cy.get('#tarIntersiteFORemark').type('Remark' + unique);






      cy.get("#btnSubmitIntersiteFO").click();

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
            const filePath = Cypress.config('fileServerFolder') + '/cypress/e2e/STIP_1/INTERSITE_FO/soDataIntersiteFo.json';
            cy.writeFile(filePath, { soNumber, siteId });

          });
        });
        // Add your logic here using the Site ID
      });

      cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout')
      cy.then(() => {
        exportToExcel(testResults);
      });
    })
  }
})