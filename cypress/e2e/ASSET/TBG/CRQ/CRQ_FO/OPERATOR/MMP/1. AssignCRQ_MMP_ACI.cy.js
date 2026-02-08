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

const filePath = 'documents/pdf/C (1).pdf';
const KMLfilepath = "documents/KML/KML/KML_BAGUS1.kml"; // Photo file path

const latMin = -11.0; // Southernmost point
const latMax = 6.5;   // Northernmost point
const longMin = 94.0; // Westernmost point
const longMax = 141.0; // Easternmost point

// Generate random latitude and longitude within bounds
const lat = (Math.random() * (latMax - latMin) + latMin).toFixed(6);
const long = (Math.random() * (longMax - longMin) + longMin).toFixed(6);

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


//Batas

describe('template spec', () => {


  let testResults = []; // Shared results array
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputCRQ, baseUrlVP, baseUrlTBGSYS, login, dashboard, menu1, menu2, menu3, menu4, logout, PICVendorMobile1, PICVendorMobile2, baseUrlCompass, Company;

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
    cy.readFile('cypress/e2e/ASSET/TBG/CRQ/CRQ_FO/DataCRQ.json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;
    });
    cy.readFile('cypress/e2e/ASSET/TBG/CRQ/CRQ_FO/DataVariableCRQ.json').then((values) => {
      cy.log(values);
      unique = values.unique;
      userAM = values.userAM;
      userInputCRQ = values.userInputCRQ;
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

    it(`THIS FLOW SUBMIT CRQ OPERATOR WITH REQUEST TYPE Additional Core Intersite ${i + 1}`, () => {

      cy.visit(`${baseUrlTBGSYS}${login}`);

      cy.get('#tbxUserID').type(userInputCRQ).should('have.value', userInputCRQ).then(() => {
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
      cy.visit(`${baseUrlTBGSYS}/Asset/CRQFO/Operator/Create`);
      cy.url().should('include', `${baseUrlTBGSYS}/Asset/CRQFO/Operator/Create`);
      // Ensure the page changes or some result occurs
      testResults.push({
        Test: 'User masuk ke Page PV Master List Maintenance',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
      // cy.get('#tblNYAssign_processing', { timeout: 20000 })
      //   .should('not.be.visible');




      cy.document().then((doc) => {
        const errorPopup = doc.querySelector('h2');

        if (errorPopup && errorPopup.innerText.includes('Error on System')) {
          cy.log('🚨 Error pop-up detected! Clicking OK.');

          // Click the "OK" button if the pop-up is present
          cy.get('.confirm.btn-error').should('be.visible').click();
        } else {
          cy.log('✅ No error pop-up detected, continuing...');
        }
      });
      cy.get('#slTenant').then(($select) => {
        cy.wrap($select).select('TSEL', { force: true })
      })
      cy.wait(2000)

      cy.get('#txtProjectName').type('Project Name' + unique + randnumber(5), { force: true });

      cy.get('#txtProjectName').type('Tenant' + unique + randnumber(5), { force: true });

      //-- Ganti ke pproduct type
      cy.get('#slProductType').then(($select) => {
        cy.wrap($select).select('2', { force: true })
      })
      cy.wait(2000)
      // EDIT VALUE TO 3 FOR Additional Core Intersite
      // EDIT VALUE TO 4 FOR Insert
      cy.get('#slRequestType').then(($select) => {
        cy.wrap($select).select('5', { force: true })
      })
      cy.wait(2000)




      cy.get('#slEquipmentIN1').then(($select) => {
        cy.wrap($select).select('1', { force: true })
      })
      cy.wait(2000)

      cy.get('#txtLongitudeIN1')
        .clear()
        .type(long, { force: true });
      cy.wait(2000)
      cy.get('#txtLatitudeIN1')
        .clear()
        .type(lat, { force: true });


      // cy.get('#btnAddSiteIDIN1').click();
      // cy.wait(2000)
      cy.get('#btnAddSiteIDIN1')
        .wait(200)
        .eq(0)
        .should('be.visible')
        .click();

      cy.get('#slFilterProduct').then(($select) => {
        cy.wrap($select).select('45', { force: true })
      })
      cy.wait(2000)
      cy.get('ins.iCheck-helper')
        .eq(0)
        .click();

      cy.get('.modal-dialog')   // atau container popup-mu
        .find('input.txtExistingCableUse')
        .type(randnumber(5) + '.22');

      cy.contains('Select Type Cable')
        .closest('.select2')
        .find('.select2-selection--single')
        .click()

      cy.get('.select2-results__options')
        .should('be.visible')
        .contains('.select2-results__option', '24')
        .click()
      cy.get('#btnSelectSite').click();
      cy.wait(5000)

      cy.get('#txtCoreProposeIN1')
        .clear()
        .type(randnumber(5), { force: true });
      cy.wait(2000)

      cy.get('#btnSubmit').click();
      cy.wait(5000)


      cy.get('.sweet-alert.showSweetAlert.visible', { timeout: 20000 })
        .should('be.visible')
        .within(() => {
          // Verifikasi isi teks popup
          cy.contains('CRQ success submitted').should('be.visible');


          // Klik tombol "OK"
          cy.get('.confirm.btn-success').should('be.visible').click();
        });
      cy.contains('a', 'Log Out').click({ force: true });
      cy.then(() => {
        exportToExcel(testResults);
      });
    })
  }
})

