import 'cypress-file-upload';
const minLength = 5;
const maxLength = 15;
const randomString = generateRandomString(minLength, maxLength);
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
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, PICVendor, baseUrlVP, baseUrlTBGSYS, login, dashboard, menu1, menu2, menu3, menu4, logout;
  const baseId = 24; // Base ID
  const index = 1; // Increment index for unique IDs
  before(() => {
    testResults = []; // Reset results before all tests
  });

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
  const KMLfilepath = "documents/KML/KML/KML_BAGUS1.kml"; // Photo file path
  const PDFFilepath = "documents/PDF/C (1).pdf"; // Photo file path



  beforeEach(() => {
    const testResults = []; // Array to store test results




    const user = "555504220025";


    cy.readFile('cypress/e2e/STIP_1/MMP_FIBERIZATION/soDataMMP_FIBERIZATION.json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;
    });

    cy.readFile('cypress/e2e/STIP_1/MMP_FIBERIZATION/DataVariable.json').then((values) => {
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

    cy.visit(`${baseUrlVP}${login}`);

    cy.get('#tbUserID').type(PICVendor);
    cy.get('#tbPassword').type(pass);


    cy.window().its('rightCode').then((rightCode) => {
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get("#btnsubmit").click();
    cy.wait(2000);

    cy.visit(`${baseUrlVP}/ProjectActivity/ProjectActivityHeader`)
      .url().should('include', `${baseUrlVP}/ProjectActivity/ProjectActivityHeader`);
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

    cy.wait(5000);
    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.log("📌 Status Found:", text);
      cy.wait(2000);

      if (text.includes(sonumb)) {  // ✅ Checks if "Lead PM" is in the status
        cy.log("✅ Status contains 'AM', proceeding with approval...");

        cy.get('tbody tr:first-child td:nth-child(1) .btnSelect').invoke('removeAttr', 'target').click();
      } else {
        cy.log("⚠️ Status does not match, skipping approval step.");
      }
    });
    cy.wait(4000);

    cy.get('tr')
      .filter((index, element) => Cypress.$(element).find('td').first().text().trim() === '3') // Find the row where the first column contains '6'
      .find('td:nth-child(2) .btnSelect') // Find the button in the second column
      .click(); // Click the button
    cy.get('.blockUI', { timeout: 300000 }).should('not.exist');
    cy.wait(1000);

    cy.get('#tbxFOLengthKabupaten').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxFOLengthProvinsi').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxFOLengthNasional').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxFOLengthInlineKAI').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxFOLengthCrossFOKAI').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxFOLengthTol').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxFOLengthHutan').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxFOLengthKorporasi').type(randomValue);
    cy.wait(1000);
    //----
    cy.get('#dpkKickOffMeetingDate')
      .invoke('val', date)
      .trigger('change');

    cy.get('#fleKickOffMeetingDocument').attachFile(PDFFilepath);
    cy.wait(1000);

    cy.get('#tbxSSRLatitudeA').type(lat);
    cy.wait(1000);
    cy.get('#tbxSSRLongitudeA').type(randomValue);
    cy.wait(1000);
    cy.get('#tarMMPAddressNearEnd').type('address FROM AUTOMATION' + unique + randomString);
    cy.wait(1000);
    cy.get('#slsProviderNE').then(($select) => {
      cy.wrap($select).select('10', { force: true })
    })


    cy.get('#tbxSSRLatitudeB').type(lat);
    cy.wait(1000);
    cy.get('#tbxSSRLongitudeB').type(randomValue);
    cy.wait(1000);
    cy.get('#tarMMPAddress').type('address FROM AUTOMATION' + unique + randomString);
    cy.wait(1000);
    cy.get('#slsProviderFE').then(($select) => {
      cy.wrap($select).select('10', { force: true })
    })
    cy.get('#fleDocumentPolePositionPhoto').attachFile(photoFilePath);
    cy.wait(1000);
    //SUMMARY PLAN LENGTH FO (M)

    cy.get('#tbxKabupaten').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxProvinsi').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxNasional').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxAerial').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxBurial').type(randomValue);
    cy.wait(1000);
    //SUMMARY MATERIAL
    // Plan(DRM)
    cy.get('#tbxPlanCable').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxPlanPole7m').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxPlanPole9m').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxPlanODP').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxPlanOTB').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxPlanHDPE').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxPlanPatchCore').type(randomValue);
    cy.wait(1000);
    //survey
    cy.get('#tbxSurveyCable').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxSurveyPole7m').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxSurveyPole9m').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxSurveyODP').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxSurveyOTB').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxSurveyHDPE').type(randomValue);
    cy.wait(1000);
    cy.get('#tbxSurveyPatchCore').type(randomValue);
    cy.wait(1000);



    cy.get('#flePhotoAccessNE').attachFile(photoFilePath);
    cy.wait(1000);
    cy.get('#flePhotoAccessFE').attachFile(photoFilePath);
    cy.wait(1000);
    cy.get('#flePhotoSiteNE').attachFile(photoFilePath);
    cy.wait(1000);
    cy.get('#flePhotoSiteFE').attachFile(photoFilePath);
    cy.wait(1000);

    // RESULT
    cy.get('#slsIssue').then(($select) => {
      cy.wrap($select).select('2', { force: true })
    })
    cy.get('#tbxDetailIssue').type('Issue' + randomString + randomValue);
    cy.wait(1000);
    cy.get('#tbxAlternatifAction').type('Action' + randomString);
    cy.wait(1000);
    cy.get('#tbxPIC').type('PIC' + randomString + randomString);
    cy.wait(1000);
    cy.get('#dpkEstimateClose')
      .invoke('val', date)
      .trigger('change');
    cy.get("#btnAddResult").click();
    cy.wait(5000);
    cy.get('#tarSSRRemark').type('Remark FROM AUTOMATION' + unique + randomString);
    cy.wait(2000);


    // cy.get("#btnProcess").click();
    // cy.wait(5000);
    // cy.get('.confirm.btn-success').click({ force: true });
    // cy.wait(5000)

    cy.get("#btnSubmit").click();
    cy.wait(5000);

    cy.get('.confirm.btn-success').click({ force: true });
    cy.wait(5000)

    cy.contains('a', 'Log Out').click({ force: true });
    cy.then(() => {
      exportToExcel(testResults);
    });
  });


});
