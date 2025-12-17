import 'cypress-file-upload';
// Fungsi untuk menghasilkan nilai acak dalam rentang tertentu
const minLength = 5;
const maxLength = 15;
const randomRangeValue = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
const randomValue = Math.floor(Math.random() * 1000) + 1; // Random number between 1 and 1000

// Daftar indeks baris yang ingin diubah
const worktypeRows = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20];
const GDLRows = [0, 1, 2, 3, 4, 5, 6, 7, 8];
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
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, PICVendor, baseUrlVP, baseUrlTBGSYS, login, dashboard, menu1, menu2, menu3, menu4, logout, PICVendorManageServiceMapped;

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


    const randomString = generateRandomString(minLength, maxLength);
    const user = "555504220025";
    const filePath = 'documents/pdf/C (1).pdf';

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
      PICVendorManageServiceMapped = values.PICVendorManageServiceMapped;
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


    cy.get('#tbxSearchSONumber').type(sonumb).should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User AM melakukan klik tombol Search di Stip approval',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    }); // << Search Filter SONumber  disable it if u dont need

    // cy.contains('label', /^\s *By SO Number\s*$/)
    //   .click(); // search By Radio Button SONumber
    // cy.get('#tbxApprovalSONumber').type(sonumb).should(() => {
    //   // Log the test result if button click is successful
    //   testResults.push({
    //     Test: 'User AM melakukan klik tombol Search di Stip approval',
    //     Status: 'Pass',
    //     Timestamp: new Date().toISOString(),
    //   });
    // }); // << Search Filter


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
    cy.wait(5000);

    cy.get('tr')
      .filter((index, element) => Cypress.$(element).find('td').first().text().trim() === '11') // Find the row where the first column contains '6'
      .find('td:nth-child(2) .btnSelect') // Find the button in the second column
      .click(); // Click the button


    // Wait for the table to be visible before proceeding
    cy.get('#tblOSPFOInstallation', { timeout: 10000 }).should('be.visible');

    // Wait for at least one row with the edit button to appear
    cy.get('tr .btnEditWorktype', { timeout: 10000 }).should('be.visible');

    worktypeRows.forEach((index) => {
      cy.wait(2000);
      cy.get('#tblOSPFOInstallation tbody tr')
        .eq(index) // Pilih baris sesuai indeks
        .find('.btnEditWorktype') // Temukan tombol edit
        .click({ force: true });

      // Input nilai acak ke dalam "tbxTotal"

      // cy.get('#tbxTotal')
      //   .should('be.enabled') // Pastikan input bisa diedit
      //   .focus() // Fokus ke input
      //   .clear()
      //   .type(randomRangeValue(10, 100));
      cy.get('#flePhotoOK').attachFile(photoFilePath);
      cy.get('#tarRemarkFOInspector').clear().type('Remark_FOInspector' + unique + randomRangeValue(10, 100));
      cy.wait(1000);
      cy.get('#btnSaveWorktypeData').click();
      cy.wait(3000);
      cy.get('.sa-confirm-button-container button.confirm').click();

    });
    cy.wait(2000);


    // cy.get('#dpkATPPlanDate')
    //   .invoke('val', date)
    //   .trigger('change');
    // cy.get('#tblOSPFOInstallation', { timeout: 10000 }).should('be.visible');

    // // Wait for at least one row with the edit button to appear
    // cy.get('tr .btnEditWorktype', { timeout: 10000 }).should('be.visible');

    // cy.get('#tbxVendorProjectManager').type('PICVENDOR' + randomRangeValue);
    // cy.wait(1000);
    // cy.get('#tbxCableLength').type(randomValue);
    // cy.wait(1000);
    // Check if the submit button is disabled
    cy.get('#btnSubmitInputATP').should('be.disabled').then(($btn) => {
      if ($btn.is(':disabled')) {
        // Click btnBaso
        // cy.get('#btnBaso').click();

        // Type quantity


        // Type GDL number, ensuring max is 8
        // for (let i = 1; i <= 8; i++) {
        //   cy.get(`#txtGDLNumber${i}`).type(`GDL-${i}`);
        //   cy.get(`#txtQuantity${i}`).should('be.enabled') // Pastikan input bisa diedit
        //     .focus() // Fokus ke input
        //     .clear()
        //     .type(randomRangeValue(250, 500));
        // }

        // // // Type warehouse location and notes
        // // cy.get('#tarWarehouseLocation').type('Main Warehouse');
        // // cy.get('#tarGDLNotes').type('Handled with care');


        // // Select the "Yes" radio button for AvailabilityAdditionalGDL
        // // Try different approaches to ensure the radio button is checked
        // cy.get('input[name="AvailabilityAdditionalGDL"][value="Yes"]').parent().click(); // Preferred for iCheck
        // cy.get('input[name="AvailabilityAdditionalGDL"][value="Yes"]').check({ force: true }); // Backup solution

        // // Upload document
        // cy.get('#fleAdditionalGDLDocument').attachFile(PDFFilepath);
        // cy.wait(1000);
        // cy.get('#fleJustificationDocument').attachFile(PDFFilepath);
        // cy.wait(1000);

        // // Select the "Yes" radio button for AvailabilityAdditionalGDL
        // cy.get('#btnSaveGDL').click();
        // Wait for the modal to close and verify btnSubmitScheduling is enabled
        cy.get('.sweet-alert.showSweetAlert.visible', { timeout: 10000 }) // Menunggu hingga elemen muncul (maks 10 detik)
          .should('be.visible') // Memastikan elemen terlihat
          .within(() => {
            cy.get('button.confirm.btn.btn-lg.btn-success').click(); // Klik tombol "OK"
          });

        cy.get('#tblOSPFOInstallation', { timeout: 10000 }).should('be.visible');
        // Wait for at least one row with the edit button to appear
        cy.get('tr .btnEditWorktype', { timeout: 10000 }).should('be.visible');


        cy.get('#btnSubmitInputATP').click();
        // Wait for the modal to close and verify btnSubmitScheduling is enabled
        cy.get('#btnSubmitInputATP').should('not.be.disabled').click();

      }
    });
    // Add wait to ensure processing
    cy.wait(10000);


    cy.contains('a', 'Log Out').click({ force: true });
    cy.then(() => {
      exportToExcel(testResults);
    });
  });
});
