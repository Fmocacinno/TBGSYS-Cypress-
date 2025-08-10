import 'cypress-file-upload';


// Daftar indeks baris yang ingin diubah
const worktypeRows = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20];
const XLSX = require('xlsx');
const fs = require('fs');
const randomStr = (length) => {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
};

const randomRangeValue = (min, max) =>
  (Math.floor(Math.random() * (max - min + 1)) + min).toString();
const randomNumber = randomRangeValue(1000000000000000, 9999999999999999);
// Example: fill all inputs with class .my-input with different random values
const randomNum = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min.toString();
// Function to export test results to Excel
const randnumber = (length) => {
  let result = '';
  for (let i = 0; i < length; i++) {
    result += Math.floor(Math.random() * 10); // 0–9
  }
  return result;
};
function exportToExcel(testResults) {
  const filePath = 'test-StipinputNewBuildMacroresults.xlsx'; // Path to the Excel file

  // Create a worksheet from the test results
  const worksheet = XLSX.utils.json_to_sheet(testResults);

  // Create a workbook and add the worksheet
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Test Results');

  // Write the workbook to a file
  XLSX.writeFile(workbook, filePath);
}


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
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, PICVendor, baseUrlVP, baseUrlTBGSYS, login, dashboard, menu1, menu2, menu3, menu4, logout, PICVendorMobile1, PICVendorMobile2, dateleasestart, dateleaseend;
  const baseId = 24; // Base ID
  const index = 1; // Increment index for unique IDs
  before(() => {
    testResults = []; // Reset results before all tests
  });




  after(() => {
    exportToExcel(testResults); // Export after all tests complete
  });
  const sorFilePath = "documents/SOR/1a0863_TO_0492_1550_14_20.sor"; // .sor file path
  const photoFilePath = "documents/IMAGE/adopt.png"; // Photo file path\
  const excelfilepath = "documents/EXCEL/.xlsx/EXCEL_(1).xlsx"; // Photo file path
  const KMLfilepath = "documents/KML/KML/KML_BAGUS1.kml"; // Photo file path
  const PDFFilepath = "documents/PDF/C (1).pdf"; // Photo file path



  beforeEach(() => {
    const testResults = []; // Array to store test results
    const user = "555504220025";
    cy.readFile('cypress/e2e/STIP_1/TBG/NEW_BUILD_MACRO/soDataNewBuild.json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;
    });
    cy.readFile('cypress/e2e/STIP_1/TBG/NEW_BUILD_MACRO/DataVariable.json').then((values) => {
      cy.log(values);
      unique = values.unique;
      userAM = values.userAM;
      userInputStip = values.userInputStip;
      userLeadAM = values.userLeadAM;
      userLeadPM = values.userLeadPM;
      userPMFO = values.userPMFO;
      userARO = values.userARO;
      PICVendor = values.PICVendor;
      date = values.date;
      dateleasestart = values.dateleasestart;
      dateleaseend = values.dateleaseend;
      pass = values.pass;
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

  it('STIP Input by vendor', () => {
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
        Test: 'Passworhasbeeninputed',
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

    cy.visit(`${baseUrlTBGSYS}/STIP/Input`);
    cy.url().should('include', `${baseUrlTBGSYS}/STIP/Input`); // Ensure the page changes or some result occurs
    // Ensure the page changes or some result occurs
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
      cy.wrap($select).select('1', { force: true })
    })

    cy.get('#slsNewCompany').then(($select) => {
      cy.wrap($select).select('TB', { force: true })
    })

    cy.get('#slsNewCustomer').then(($select) => {
      cy.wrap($select).select('TSEL', { force: true })
    })

    cy.get('#slsNewRegion').then(($select) => {
      cy.wrap($select).select('1', { force: true })
    })

    cy.wait(2000)
    cy.get('#btnNewPriceAmountPopUp').click();

    cy.get('tbody > tr:nth-child(3) .btnSelect').click();

    cy.wait(2000)

    cy.get('.slsBatchSLD').eq(0)
      .select('1', { force: true });

    cy.get('.slsBatchSLD').eq(1)
      .select('1', { force: true });


    cy.get('.slsBatchSLD').eq(2)
      .select('1', { force: true });

    cy.get('.slsBatchSLD').eq(3)
      .select('1', { force: true });

    cy.get('.slsBatchSLD').eq(4)
      .select('1', { force: true });

    cy.get('.slsBatchSLD').eq(5)
      .select('1', { force: true });

    cy.get('.slsBatchSLD').eq(6)
      .select('1', { force: true });

    cy.get('.slsBatchSLD').eq(7)
      .select('1', { force: true });

    cy.get('.slsBatchSLD').eq(8)
      .select('1', { force: true });

    cy.get('#tbxNewSiteName').type('Site_' + unique + randomNum(5));
    cy.get('#tbxNewCustomerSiteID').type('Cust_' + unique + randomNum(5));

    cy.get('#slsNewDocumentOrder').then(($select) => {
      cy.wrap($select).select('7', { force: true })
    })

    cy.get('#tbxNewDocumentName').type('DocName_' + unique + randomNum(5));

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

    cy.wait(2000)
    cy.get('#slsNewLeadProjectManager').first().then(($select) => {
      cy.wrap($select).select(userLeadPM, { force: true })
    })
    cy.get('#slsNewAccountManager').then(($select) => {
      cy.wrap($select).select(userLeadAM, { force: true })
    })

    cy.get('#slsNewTowerHeight').then(($select) => {
      cy.wrap($select).select('72', { force: true })
    })

    cy.get('#slsNewShelterType').then(($select) => {
      cy.wrap($select).select('7', { force: true })
    })
    cy.get('#tbxNewPLNPowerKVA').type(5);

    cy.get('#dpkNewRFITarget')
      .invoke('val', date)
      .trigger('change');

    cy.get('#slsNewMLANumber').then(($select) => {
      cy.wrap($select).select('060/BC/PROC-01/LOG/2010', { force: true })
    })

    cy.get('#tbxNewLeasePeriod').type(5);

    cy.get('#dpkNewMLADate')
      .invoke('val', date)
      .trigger('change');

    cy.get('#tarNewRemark').type('Remark' + unique + randomNum(5));

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
          const filePath = Cypress.config('fileServerFolder') + '/cypress/e2e/STIP_1/TBG/NEW_BUILD_MACRO/soDataNewBuild.json';
          cy.writeFile(filePath, { soNumber, siteId });

        });
      });
      // Add your logic here using the Site ID
    });

    cy.contains('a', 'Log Out').click({ force: true });
    cy.then(() => {
      exportToExcel(testResults);
    });
  });
});
