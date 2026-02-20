import { timeStamp } from 'console';

import 'cypress-file-upload';
import { validateHeaderValue } from 'http';
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

const randnumber = (length) => {
  let result = '';
  for (let i = 0; i < length; i++) {
    result += Math.floor(Math.random() * 10); // 0–9
  }
  return result;
};
const randomYear = (min = 1980, max = new Date().getFullYear()) => {
  return Math.floor(Math.random() * (max - min + 1)) + min;
};

const year = randomYear();

const filePath = 'documents/pdf/C (1).pdf';
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

let sonumb, siteId, pass, siteName, userInputStip;
const minLength = 5;
const maxLength = 15;
const randomString = generateRandomString(minLength, maxLength);

// const date = "2-Jan-2025";
// const user = "555504220025"
// const pass = "123456"


//Batas

describe('template spec', () => {


  let testResults = []; // Shared results array
  let sonumb, siteId, SIMNUM, unique, date, custom, baseUrlCICO, company, operator, PIC_A, PIC_A_ID, Email, JenisPekerjaan;

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
    cy.readFile('cypress/e2e/ASSET/TBG/CICO/DataCICO.json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;
      SIMNUM = values.SIMNUM;
      siteName = values.siteName;
      company = values.company;
      operator = values.operator;
      PIC_A = values.PIC_A;
      PIC_A_ID = values.PIC_A_ID;
      Email = values.Email;
      JenisPekerjaan = values.JenisPekerjaan;
    });
    cy.readFile('cypress/e2e/ASSET/TBG/CICO/DataVariable.json').then((values) => {
      cy.log(values);
      unique = values.unique;
      custom = values.custom;
      date = values.date;
      baseUrlCICO = values.baseUrlCICO;
    });


    Cypress.on('uncaught:exception', (err, runnable) => {
      return false;
    });
  });
  const loopCount = 1; // Jumlah iterasi loop

  for (let i = 0; i < loopCount; i++) {

    it(`THIS FLOW FOR SUBMIT CICO ${i + 1}`, () => {

      cy.visit(`${baseUrlCICO}`);
      cy.contains('button', 'Next').click()

      cy.get('#heyform-date-day').type(randnumber(2), { force: true });
      cy.get('#heyform-date-month').type(randnumber(2), { force: true });
      cy.get('#heyform-date-year').type(String(year));
      cy.get('#heyform-date-hour').type(randnumber(2), { force: true });
      cy.get('#heyform-date-minute').type(randnumber(2), { force: true });
      cy.contains('button', 'Next').click()
      cy.wait(500)
      cy.get('.heyform-input')
        .should('be.visible')
        .type(SIMNUM)
      cy.contains('button', 'Next').click()
      cy.get('.heyform-input')
        .should('be.visible')
        .type(siteId)
      cy.contains('button', 'Next').click()

      cy.wait(500)
      cy.get('.heyform-input')
        .should('be.visible')
        .type(siteName)
      cy.contains('button', 'Next').click()
      cy.wait(500)
      cy.get('.heyform-input')
        .should('be.visible')
        .type(company)
      cy.wait(500)
      cy.contains('button', 'Next').click()

      cy.contains('.heyform-radio', operator)
        .should('be.visible')
        .click()
      cy.contains('button', 'Next').click()
      cy.wait(500)
      cy.contains('button', 'Next').click()
      cy.get('.heyform-input')
        .should('be.visible')
        .type(unique + ' ' + randomString)
      cy.wait(500)
      cy.contains('button', 'Next').click()
      cy.get('.heyform-input')
        .should('be.visible')
        .type(Email)
      cy.wait(500)
      cy.contains('button', 'Next').click()
      cy.get('.heyform-input')
        .should('be.visible')
        .type(randnumber(11))
      cy.wait(500)
      cy.contains('button', 'Next').click()
      cy.get('input[type="file"]').attachFile(filePath);
      cy.wait(500)
      cy.contains('button', 'Next').click()
      cy.get('.heyform-input')
        .should('be.visible')
        .type(PIC_A_ID)
      cy.wait(500)
      cy.contains('button', 'Next').click()
      cy.wait(500)
      cy.get('input[type="file"]').attachFile(photoFilePath);
      cy.contains('button', 'Next').click()
      cy.get('input[type="file"]').attachFile(photoFilePath);
      cy.contains('button', 'Next').click()
      cy.wait(500)

      cy.contains('.heyform-radio', 'No')
        .should('be.visible')
        .click()
      cy.wait(500)
      cy.contains('button', 'Next').click()
      cy.wait(500)
      cy.contains('.heyform-radio', 'Yes')
        .should('be.visible')
        .click()
      cy.contains('button', 'Next').click()
      cy.wait(500)
      cy.contains('.heyform-radio', JenisPekerjaan)
        .should('be.visible')
        .click()
      cy.wait(500)
      cy.get('button[type="submit"]').click()

    })
  }
})

