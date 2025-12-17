import 'cypress-file-upload';

const minLength = 5;
const maxLength = 15;
const randomString = generateRandomString(minLength, maxLength);
// Fungsi untuk menghasilkan nilai acak dalam rentang tertentu
const randomRangeValue = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
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
// Daftar indeks baris yang ingin diubah
const worktypeRows = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20];
const XLSX = require('xlsx');
const fs = require('fs');

// Function to export test results to Excel
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
describe('template spec', () => {
  let testResults = []; // Shared results array
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, PICVendor, baseUrlVP, baseUrlTBGSYS, login, dashboard, menu1, menu2, menu3, menu4, logout;

  before(() => {
    testResults = []; // Reset results before all tests
  });

  after(() => {
    exportToExcel(testResults); // Export after all tests complete
  });

  beforeEach(() => {
    const testResults = []; // Array to store test results
    const randomValue = Math.floor(Math.random() * 1000) + 1; // Random number between 1 and 1000

    const user = "555504220025";
    const filePath = 'documents/pdf/receipt.pdf';
    cy.readFile('cypress/e2e/STIP_1/TBG/INTERSITE_FO/soDataIntersiteFO.json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;
    });
    cy.readFile('cypress/e2e/STIP_1/TBG/INTERSITE_FO/DataVariable.json').then((values) => {
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
    });

    Cypress.on('uncaught:exception', (err, runnable) => {
      return false;
    });
  });

  //AM
  it('OTDR Input by vendor', () => {

    cy.visit(`${baseUrlTBGSYS}${login}`);

    cy.get('#tbxUserID').type(userPMFO);
    cy.get('#tbxPassword').type(pass);


    cy.window().its('rightCode').then((rightCode) => {
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get("#btnSubmit").click();
    cy.wait(2000);

    cy.visit(`${baseUrlTBGSYS}/ProjectActivity/ProjectActivityHeader`)
      .url().should('include', `${baseUrlTBGSYS}/ProjectActivity/ProjectActivityHeader`);
    testResults.push({
      Test: 'User AM melakukan akses ke menu Project activity Header',
      Status: 'Pass',
      timeStamp: new Date().toISOString(),
    });
    cy.get('.blockUI', { timeout: 300000 }).should('not.exist');

    // Check if the error pop-up is visible
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

    cy.get('#tbxSearchSONumber').type(sonumb).should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User AM melakukan klik tombol Search di Stip approval',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    }); // << Search Filter SONumber  disable it if u dont need



    cy.get('.btnSearch').first().click().should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User AM melakukan klik tombol Search di Stip approval',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    });

    cy.wait(2000);
    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.log("📌 Status Found:", text);
      cy.wait(2000);

      if (text.includes(sonumb)) {  // ✅ Checks if "Lead PM" is in the status
        cy.log("✅ SONumber find, proceeding with approval...");

        cy.get('tbody tr:first-child td:nth-child(1) .btnSelect').invoke('removeAttr', 'target').click();
      } else {
        cy.log("⚠️ Status does not match, skipping approval step.");
      }
    });
    cy.wait(2000);

    cy.get('tr')
      .filter((index, element) => Cypress.$(element).find('td').first().text().trim() === '5') // Find the row where the first column contains '6'    cy.wait(2000);

      .find('td:nth-child(2) .btnSelect') // Find the button in the second column
      .click(); // Click the button

    cy.wait(4000);
    cy.get('#tarMaterialDeliveryApprovalRemark').type('Remark FROM AUTOMATION' + unique + randomString);
    cy.wait(2000);

    cy.get('#btnApprove').click();
    cy.wait(7000);
    cy.get('.sa-confirm-button-container button.confirm').click();

    cy.wait(2000);

    // cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');
    cy.contains('a', 'Log Out').click({ force: true });
    cy.then(() => {
      exportToExcel(testResults);
    });
  });


});
