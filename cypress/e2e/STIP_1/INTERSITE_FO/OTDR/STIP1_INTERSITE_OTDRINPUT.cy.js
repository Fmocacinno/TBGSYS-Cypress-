// Fungsi untuk menghasilkan nilai acak dalam rentang tertentu
import 'cypress-file-upload';
const randomRangeValue = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;

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
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, PICVendor;
  const baseId = 24; // Base ID
  const index = 1; // Increment index for unique IDs
  before(() => {
    testResults = []; // Reset results before all tests
  });
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
  const minLength = 5;
  const maxLength = 15;
  const randomString = generateRandomString(minLength, maxLength);
  const randomValue = Math.floor(Math.random() * 1000) + 1; // Random number between 1 and 1000

  after(() => {
    exportToExcel(testResults); // Export after all tests complete
  });
  const sorFilePath = "documents/SOR/1a0863_TO_0492_1550_14_20.sor"; // .sor file path
  const photoFilePath = "documents/IMAGE/adopt.png"; // Photo file path\
  const excelfilepath = "documents/EXCEL/.xlsx/EXCEL_(1).xlsx"; // Photo file path



  beforeEach(() => {
    const testResults = []; // Array to store test results




    const user = "555504220025";


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
      PICVendor = values.PICVendor;
      date = values.date;
      pass = values.pass;
    });

    Cypress.on('uncaught:exception', (err, runnable) => {
      return false;
    });
  });

  //AM
  it('OTDR Input by vendor', () => {

    cy.visit('http://tbgappdev111.tbg.local:8128/Login');

    cy.get('#tbUserID').type(PICVendor);
    cy.get('#tbPassword').type(pass);


    cy.window().its('rightCode').then((rightCode) => {
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get("#btnsubmit").click();
    cy.wait(2000);

    cy.visit('http://tbgappdev111.tbg.local:8128/ProjectActivity/ProjectActivityHeader')
      .url().should('include', 'http://tbgappdev111.tbg.local:8128/ProjectActivity/ProjectActivityHeader');
    testResults.push({
      Test: 'User AM melakukan akses ke menu Project activity Header',
      Status: 'Pass',
      timeStamp: new Date().toISOString(),
    });
    // Wait for any loading overlay to disappear
    // Wait for any loading overlay to disappear
    cy.get('.blockUI', { timeout: 300000 }).should('not.exist');

    // Check if the error pop-up exists without failing the test
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
    cy.wait(5000);

    cy.get('tr')

      .filter((index, element) => Cypress.$(element).find('td').first().text().trim() === '8') // Find the row where the first column contains '6'    cy.wait(2000);

      .find('td:nth-child(2) .btnSelect') // Find the button in the second column
      .click(); // Click the button
    cy.wait(3000);
    cy.get('#fleOTDRCoreManagementDocument').attachFile(excelfilepath);

    cy.get('#dpkOTDRInstallationDate')
      .invoke('val', date)
      .trigger('change');


    for (let i = 1; i <= 1; i++) {
      // Click "No" radio button
      cy.get(`input[name="rdoIsCoreActiveFarEndOTDR${i}"][value="0"]`).parent().click({ force: true });

      // Upload .sor file and wait for the upload to be confirmed
      cy.get(`#fleSORFarEndOTDR${i}`).selectFile(`cypress/fixtures/${sorFilePath}`, { force: true });
      cy.wait(2000);

      // Upload photo file and wait for the upload to be confirmed
      cy.get(`#flePhotoFarEndOTDR${i}`).selectFile(`cypress/fixtures/${photoFilePath}`, { force: true });
      cy.wait(2000);
    }
    cy.get('.nav-tabs a[href="#tabOTDRNearEnd"]').click();
    cy.wait(3000);

    for (let i = 1; i <= 1; i++) {
      // Click "No" radio button
      cy.get(`input[name="rdoIsCoreActiveNearEndOTDR${i}"][value="0"]`).parent().click({ force: true });

      // Upload .sor file and wait for the upload to be confirmed
      cy.get(`#fleSORNearEndOTDR${i}`).selectFile(`cypress/fixtures/${sorFilePath}`, { force: true });
      cy.wait(2000);

      // Upload photo file and wait for the upload to be confirmed
      cy.get(`#flePhotoNearEndOTDR${i}`).selectFile(`cypress/fixtures/${photoFilePath}`, { force: true });
      cy.wait(2000);
    }
    //open accrodion
    cy.contains('.portlet-title', 'OPM') // Find the "OPM" accordion
      .parent() // Move to the parent container
      .find('.tools .expand') // Locate the expand button
      .click({ force: true }); // Click to expand
    cy.get('#dpkOPMInstallationDate')
      .invoke('val', date)
      .trigger('change');

    cy.get('#fleOPMCalibrationPhotoFarEnd').attachFile(photoFilePath);

    for (let i = 1; i <= 1; i++) {
      // Click "No" radio button
      cy.get(`input[name="rdoIsCoreActiveFarEndOPM${i}"][value="0"]`).parent().click({ force: true });
      // Upload photo file and wait for the upload to be confirmed
      cy.get(`#flePhotoFarEndOPM${i}`).selectFile(`cypress/fixtures/${photoFilePath}`, { force: true });
      cy.wait(2000);
      cy.get(`#flePhotoFarEndOPM${i}`).selectFile(`cypress/fixtures/${photoFilePath}`, { force: true });
      cy.wait(2000);
      cy.get(`#tbxCalibrationInputFarEndOPM${i}`).type(randomValue, { force: true });
      cy.wait(2000);
      cy.get(`#tbxTotalLossInputFarEndOPM${i}`).type(randomValue, { force: true });
      cy.wait(2000);
    }
    cy.get('.nav-tabs a[href="#tabOPMNearEnd"]').click();
    cy.wait(3000);
    cy.get('#fleOPMCalibrationPhotoNearEnd').attachFile(photoFilePath);

    for (let i = 1; i <= 1; i++) {
      // Click "No" radio button
      cy.get(`input[name="rdoIsCoreActiveNearEndOPM${i}"][value="0"]`).parent().click({ force: true });
      // Upload photo file and wait for the upload to be confirmed
      cy.get(`#flePhotoNearEndOPM${i}`).selectFile(`cypress/fixtures/${photoFilePath}`, { force: true });
      cy.wait(2000);
      cy.get(`#flePhotoNearEndOPM${i}`).selectFile(`cypress/fixtures/${photoFilePath}`, { force: true });
      cy.wait(2000);
      cy.get(`#tbxCalibrationInputNearEndOPM${i}`).type(randomValue, { force: true });
      cy.wait(2000);
      cy.get(`#tbxTotalLossInputNearEndOPM${i}`).type(randomValue, { force: true });
      cy.wait(2000);
    }
    cy.get('#tarOTDRInstallationRemark').type('Remark FROM AUTOMATION' + unique + randomRangeValue);
    cy.wait(2000);
    cy.get("#btnSubmit").click();
    cy.wait(25000);

    cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');
    cy.then(() => {
      exportToExcel(testResults);
    });
  });


});
