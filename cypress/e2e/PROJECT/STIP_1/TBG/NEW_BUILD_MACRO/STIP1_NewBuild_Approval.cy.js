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
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, PICVendor, baseUrlVP, baseUrlTBGSYS, login, dashboard, menu1, menu2, menu3, menu4, logout, PICVendorMobile1, PICVendorMobile2, userPMSitac, userPMCME;
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
      userPMSitac = values.userPMSitac;
      userPMCME = values.userPMCME;
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
      PICVendorMobile1 = values.PICVendorMobile1;
      PICVendorMobile2 = values.PICVendorMobile2;
    });

    Cypress.on('uncaught:exception', (err, runnable) => {
      return false;
    });
  });
  //AM
  it('AM Test Case', () => {

    cy.visit(`${baseUrlTBGSYS}${login}`);

    cy.get('#tbxUserID').type(userAM);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get("#btnSubmit").click();
    cy.wait(3000);
    cy.visit(`${baseUrlTBGSYS}/STIP/Approval`);
    cy.url().should('include', `${baseUrlTBGSYS}/STIP/Approval`);
    // Ensure the page changes or some result occurs
    testResults.push({
      Test: 'User AM melakukan akses ke menu Stip Approval',
      Status: 'Pass',
      timeStamp: new Date().toISOString(),
    });
    cy.wait(3000);

    // cy.get('#tbxSearchSONumber').type(sonumb).should(() => {
    //   // Log the test result if button click is successful
    //   testResults.push({
    //     Test: 'User AM melakukan klik tombol Search di Stip approval',
    //     Status: 'Pass',
    //     Timestamp: new Date().toISOString(),
    //   });
    // }); // << Search Filter SONumber  disable it if u dont need

    cy.contains('label', /^\s *By SO Number\s*$/)
      .click(); // search By Radio Button SONumber
    cy.get('#tbxApprovalSONumber').type(sonumb).should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User AM melakukan klik tombol Search di Stip approval',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    }); // << Search Filter


    cy.get('.btnSearch').first().click().should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User AM melakukan klik tombol Search di Stip approval',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    });

    cy.wait(3000);
    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.log("📌 Status Found:", text);
      cy.wait(3000);

      if (text.includes("Waiting for Approval AM")) {  // ✅ Checks if "Lead AM" is in the status
        cy.log("✅ Status contains 'AM', proceeding with approval...");

        cy.get('tbody tr:first-child td:nth-child(1) .btnApprovalDetail').click();
        cy.get('#tarApprovalRemark').type('Remark_' + unique, { force: true });
        // SKIPPING 'tarApprovalRemark' input field
        cy.log("⚠️ Skipping remark input...");

        // Attempt to click the approval button only if it's visible and enabled
        cy.get("#btnApprove").then(($btn) => {
          if ($btn.is(':visible') && !$btn.is(':disabled')) {
            cy.wrap($btn).click();
            cy.log("✅ Button clicked successfully");
            cy.wait(3000);
          } else {
            cy.wait(3000);
            cy.log("⚠️ Button not clickable, skipping...");
          }
        });

      } else {
        cy.log("⚠️ Status does not match, skipping approval step.");
      }
    });

    cy.wait(3000);
    cy.contains('a', 'Log Out').click({ force: true });
    cy.wait(4000);

  });

  //LEAD AM
  it('Lead AM Test Case', () => {
    const testResults = [];
    // Lead PM
    cy.visit(`${baseUrlTBGSYS}${login}`);
    cy.get('#tbxUserID').type(userLeadAM);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();
    cy.wait(3000);
    cy.visit(`${baseUrlTBGSYS}/STIP/Approval`);
    cy.url().should('include', `${baseUrlTBGSYS}/STIP/Approval`);
    testResults.push({
      Test: 'User Lead AM melakukan akses ke menu Stip Approval',
      Status: 'Pass',
      timeStamp: new Date().toISOString(),
    });
    cy.wait(3000);


    // cy.get('#tbxSearchSONumber').type(sonumb).should(() => {
    //   // Log the test result if button click is successful
    //   testResults.push({
    //     Test: 'User AM melakukan klik tombol Search di Stip approval',
    //     Status: 'Pass',
    //     Timestamp: new Date().toISOString(),
    //   });
    // }); // << Search Filter SONumber  disable it if u dont need

    cy.contains('label', /^\s *By SO Number\s*$/)
      .click(); // search By Radio Button SONumber
    cy.get('#tbxApprovalSONumber').type(sonumb).should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User AM melakukan klik tombol Search di Stip approval',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    }); // << Search Filter


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

      if (text.includes("Waiting for Approval Lead AM")) {  // ✅ Checks if "Lead PM" is in the status
        cy.log("✅ Status contains 'Lead AM', proceeding with approval...");
        cy.get('tbody tr:first-child .btnApprovalDetail').click();
        cy.get('#tarApprovalRemark').type('Remark_' + unique, { force: true });
        // SKIPPING 'tarApprovalRemark' input field
        cy.log("⚠️ Skipping remark input...");

        // Attempt to click the approval button only if it's visible and enabled
        cy.get("#btnApprove").then(($btn) => {
          if ($btn.is(':visible') && !$btn.is(':disabled')) {
            cy.wrap($btn).click();
            cy.log("✅ Button clicked successfully");
          } else {
            cy.log("⚠️ Button not clickable, skipping...");
          }
        });

      } else {
        cy.log("⚠️ Status does not match, skipping approval step.");
      }
    });

    cy.wait(3000);
    cy.contains('a', 'Log Out').click({ force: true });
    cy.wait(4000);

  });
  //LEAD PM


  it('Lead PM Test Case', () => {
    // Lead PM
    cy.visit(`${baseUrlTBGSYS}${login}`);
    cy.get('#tbxUserID').type(userLeadPM);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();

    cy.wait(3000);
    cy.visit(`${baseUrlTBGSYS}/STIP/Approval`);
    cy.url().should('include', `${baseUrlTBGSYS}/STIP/Approval`);
    testResults.push({
      Test: 'User LEAD PM melakukan akses ke menu Stip Approval',
      Status: 'Pass',
      TimeStamp: new Date().toISOString(),
    });
    cy.wait(3000);
    cy.contains('label', /^\s *By SO Number\s*$/)
      .click(); // search By Radio Button SONumber
    cy.get('#tbxApprovalSONumber').type(sonumb).should('have.value', sonumb).then(() => {
      // Log the test result if input is successful
      testResults.push({
        Test: 'User LEAD PM melakukan input SONumber di Stip approval',
        Status: 'Pass',
        TimeStamp: new Date().toISOString(),
      });
    });

    cy.get('.btnSearch').first().click().should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User LEAD PM melakukan klik tombol Search di Stip approval',
        Status: 'Pass',
        TimeStamp: new Date().toISOString(),
      });
    });



    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.log("📌 Status Found:", text);

      if (text.includes("Lead PM")) {  // ✅ Flexible condition to match variations
        cy.log("✅ Status contains 'Lead PM', proceeding with approval...");

        cy.get('tbody tr:first-child td:nth-child(1) .btnApprovalDetail').click();
        cy.get('#tarApprovalRemark').type('Remark_' + unique, { force: true });
        cy.log("⚠️ Skipping remark input...");
        cy.wait(10000);
        cy.get("#btnConfirm").then(($btn) => {
          if ($btn.is(':visible') && !$btn.is(':disabled')) {
            cy.wait(3000);
            cy.get('.nav-tabs a[href="#tabPMAssignment"]').click();
            cy.wait(3000);
            cy.get('#slsPMSitac').select(userPMSitac, { force: true });
            cy.get('#slsPMCME').select(userPMCME, { force: true });
            cy.get('.nav-tabs a[href="#tabApprovalDetail"]').click();
            cy.wait(3000);
            cy.wrap($btn).click();
            cy.wait(10000);
            cy.log("✅ Button clicked successfully");
          } else {
            cy.log("⚠️ Button not clickable, skipping...");
          }
          // if ($btn.is(':visible')) {
          //   // Use dynamic values in select fields
          //   if (userPMSitac) {
          //     cy.get('#slsPMSitac').select(userPMSitac, { force: true });
          //   } else {
          //     cy.log('userPMSitac is missing or invalid');
          //   }

          //   if (userPMCME) {
          //     cy.get('#slsPMCME').select(userPMCME, { force: true });
          //   } else {
          //     cy.log('userPMCME is missing or invalid');
          //   } // Use dynamic values in 
          //   cy.wrap($btn).click();
          //   cy.log("✅ Button clicked successfully");
          //   cy.wait(6000);
          // } else {
          //   cy.log("⚠️ Button not clickable, skipping...");
          // }
        });

      } else {
        cy.log("⚠️ Status does not match, skipping approval step.");
      }
    });

    cy.wait(4000);
    cy.contains('a', 'Log Out').click({ force: true });
    cy.wait(4000);

  });
  //ARO
  it('ARO Test Case', () => {
    // ARO
    cy.visit(`${baseUrlTBGSYS}${login}`);
    cy.get('#tbxUserID').type(userARO);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();
    cy.wait(3000);
    cy.visit(`${baseUrlTBGSYS}/STIP/Approval`);
    cy.url().should('include', `${baseUrlTBGSYS}/STIP/Approval`);
    cy.wait(3000);
    cy.contains('label', /^\s *By SO Number\s*$/)
      .click(); // search By Radio Button SONumber
    cy.get('#tbxApprovalSONumber').type(sonumb).should('have.value', sonumb).then(() => {
      // Log the test result if input is successful
      testResults.push({
        Test: 'User ARO melakukan input SONumber di Stip approval',
        Status: 'Pass',
        TimeStamp: new Date().toISOString(),
      });
    });

    cy.get('.btnSearch').first().click().should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User ARO melakukan klik tombol Search di Stip approval',
        Status: 'Pass',
        TimeStamp: new Date().toISOString(),
      });
    });
    cy.wait(3000);
    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.log("📌 Status Found:", text);

      if (text === "Waiting for Confirmation ARO") {
        cy.log("✅ Status matches, proceeding with approval...");

        cy.get('tbody tr:first-child .btnApprovalDetail').click();
        cy.get('#tarApprovalRemark').type('Remark_' + unique, { force: true });
        // SKIPPING 'tarApprovalRemark' input field
        cy.log("⚠️ Skipping remark input...");

        // Attempt to click the approval button only if it's visible and enabled
        cy.get("#btnConfirm").then(($btn) => {
          if ($btn.is(':visible') && !$btn.is(':disabled')) {
            cy.wrap($btn).click();

            cy.log("✅ Button clicked successfully");
          } else {
            cy.log("⚠️ Button not clickable, skipping...");
          }
        });

      } else {
        cy.log("⚠️ Status does not match, skipping approval step.");
      }
      cy.wait(3000);
    });
    cy.wait(3000);
    cy.contains('a', 'Log Out').click({ force: true });
    cy.wait(4000);

    cy.then(() => {
      exportToExcel(testResults);
    });

  });

});
