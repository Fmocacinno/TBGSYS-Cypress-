// Fungsi untuk menghasilkan nilai acak dalam rentang tertentu
const randomRangeValue = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
const randomValue = Math.floor(Math.random() * 1000) + 1; // Random number between 1 and 1000

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

  before(() => {
    testResults = []; // Reset results before all tests
  });

  after(() => {
    exportToExcel(testResults); // Export after all tests complete
  });

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

    const minLength = 5;
    const maxLength = 15;
    const randomString = generateRandomString(minLength, maxLength);
    const user = "555504220025";
    const filePath = 'documents/pdf/C (1).pdf';

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

    cy.wait(2000);
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
    cy.wait(2000);

    cy.get('tr')
      .filter((index, element) => Cypress.$(element).find('td').first().text().trim() === '11') // Find the row where the first column contains '6'
      .find('td:nth-child(2) .btnSelect') // Find the button in the second column
      .click(); // Click the button

    worktypeRows.forEach((index) => {
      cy.get('#tblOSPFOInstallation tbody tr')
        .eq(index) // Pilih baris sesuai indeks
        .find('.btnEditWorktype') // Temukan tombol edit
        .click({ force: true });

      // Input nilai acak ke dalam "tbxTotal"
      cy.get('#tbxTotal').clear().type(randomRangeValue(10, 100)); // Rentang bisa disesuaikan
      cy.get('#tarRemarkVendor').type('Remark' + unique);
      cy.get('#btnSaveWorktypeData').click();
      cy.wait(2000);
      cy.get('.sa-confirm-button-container button.confirm').click();
    });
    cy.wait(2000);


    cy.get('#dpkATPPlanDate')
      .invoke('val', date)
      .trigger('change');
    cy.wait(5000);

    cy.get('#tbxVendorProjectManager').type('PICVENDOR' + randomRangeValue);
    cy.wait(1000);
    cy.get('#tbxCableLength').type(randomValue);
    cy.wait(1000);
    cy.get('#btnSubmitScheduling').click();
    cy.wait(10000);


    cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');
    cy.then(() => {
      exportToExcel(testResults);
    });
  });




  //LEAD AM
  // it('Lead AM Test Case', () => {
  //   const testResults = [];
  //   // Lead PM
  //   cy.visit('http://tbgappdev111.tbg.local:8042');
  //   cy.get('#tbxUserID').type(userLeadAM);
  //   cy.get('#tbxPassword').type(pass);
  //   cy.get('#RefreshButton').click();

  //   cy.window().then((window) => {
  //     const rightCode = window.rightCode;
  //     cy.log('Right Code:', rightCode);
  //     cy.get('#captchaInsert').type(rightCode);
  //   });

  //   cy.get('#btnSubmit').click();
  //   cy.wait(2000);
  //   cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval')
  //     .url().should('include', 'http://tbgappdev111.tbg.local:8042/STIP/Approval');
  //   testResults.push({
  //     Test: 'User Lead AM melakukan akses ke menu Stip Approval',
  //     Status: 'Pass',
  //     timeStamp: new Date().toISOString(),
  //   });
  //   cy.wait(2000);


  //   // cy.get('#tbxSearchSONumber').type(sonumb).should(() => {
  //   //   // Log the test result if button click is successful
  //   //   testResults.push({
  //   //     Test: 'User AM melakukan klik tombol Search di Stip approval',
  //   //     Status: 'Pass',
  //   //     Timestamp: new Date().toISOString(),
  //   //   });
  //   // }); // << Search Filter SONumber  disable it if u dont need

  //   cy.contains('label', /^\s *By SO Number\s*$/)
  //     .click(); // search By Radio Button SONumber
  //   cy.get('#tbxApprovalSONumber').type(sonumb).should(() => {
  //     // Log the test result if button click is successful
  //     testResults.push({
  //       Test: 'User AM melakukan klik tombol Search di Stip approval',
  //       Status: 'Pass',
  //       Timestamp: new Date().toISOString(),
  //     });
  //   }); // << Search Filter


  //   cy.get('.btnSearch').first().click().should(() => {
  //     // Log the test result if button click is successful
  //     testResults.push({
  //       Test: 'User AM melakukan klik tombol Search di Stip approval',
  //       Status: 'Pass',
  //       Timestamp: new Date().toISOString(),
  //     });
  //   });
  //   cy.wait(2000);

  //   cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
  //     const text = $cell.text().trim();
  //     cy.log("📌 Status Found:", text);

  //     if (text.includes("Waiting for Approval Lead AM")) {  // ✅ Checks if "Lead PM" is in the status
  //       cy.log("✅ Status contains 'Lead AM', proceeding with approval...");
  //       cy.get('tbody tr:first-child .btnApprovalDetail').click();
  //       cy.get('#tarApprovalRemark').type('Remark_' + unique, { force: true });
  //       // SKIPPING 'tarApprovalRemark' input field
  //       cy.log("⚠️ Skipping remark input...");

  //       // Attempt to click the approval button only if it's visible and enabled
  //       cy.get("#btnApprove").then(($btn) => {
  //         if ($btn.is(':visible') && !$btn.is(':disabled')) {
  //           cy.wrap($btn).click();
  //           cy.log("✅ Button clicked successfully");
  //         } else {
  //           cy.log("⚠️ Button not clickable, skipping...");
  //         }
  //       });

  //     } else {
  //       cy.log("⚠️ Status does not match, skipping approval step.");
  //     }
  //   });

  //   cy.wait(2000);
  //   cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');

  // });
  // //LEAD PM
  // it('Lead PM Test Case', () => {
  //   // Lead PM
  //   cy.visit('http://tbgappdev111.tbg.local:8042');
  //   cy.get('#tbxUserID').type(userLeadPM);
  //   cy.get('#tbxPassword').type(pass);
  //   cy.get('#RefreshButton').click();

  //   cy.window().then((window) => {
  //     const rightCode = window.rightCode;
  //     cy.log('Right Code:', rightCode);
  //     cy.get('#captchaInsert').type(rightCode);
  //   });

  //   cy.get('#btnSubmit').click();

  //   cy.wait(2000);
  //   cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval')
  //     .url().should('include', 'http://tbgappdev111.tbg.local:8042/STIP/Approval');
  //   testResults.push({
  //     Test: 'User LEAD PM melakukan akses ke menu Stip Approval',
  //     Status: 'Pass',
  //     TimeStamp: new Date().toISOString(),
  //   });
  //   cy.wait(2000);

  //   cy.get('#tbxSearchSONumber').type(sonumb).should('have.value', sonumb).then(() => {
  //     // Log the test result if input is successful
  //     testResults.push({
  //       Test: 'User LEAD PM melakukan input SONumber di Stip approval',
  //       Status: 'Pass',
  //       TimeStamp: new Date().toISOString(),
  //     });
  //   });


  //   // cy.get('#tbxSearchSONumber').type(sonumb).should(() => {
  //   //   // Log the test result if button click is successful
  //   //   testResults.push({
  //   //     Test: 'User AM melakukan klik tombol Search di Stip approval',
  //   //     Status: 'Pass',
  //   //     Timestamp: new Date().toISOString(),
  //   //   });
  //   // }); // << Search Filter SONumber  disable it if u dont need

  //   cy.contains('label', /^\s *By SO Number\s*$/)
  //     .click(); // search By Radio Button SONumber
  //   cy.get('#tbxApprovalSONumber').type(sonumb).should(() => {
  //     // Log the test result if button click is successful
  //     testResults.push({
  //       Test: 'User AM melakukan klik tombol Search di Stip approval',
  //       Status: 'Pass',
  //       Timestamp: new Date().toISOString(),
  //     });
  //   }); // << Search Filter


  //   cy.get('.btnSearch').first().click().should(() => {
  //     // Log the test result if button click is successful
  //     testResults.push({
  //       Test: 'User AM melakukan klik tombol Search di Stip approval',
  //       Status: 'Pass',
  //       Timestamp: new Date().toISOString(),
  //     });
  //   });
  //   cy.wait(2000);


  //   cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
  //     const text = $cell.text().trim();
  //     cy.log("📌 Status Found:", text);

  //     if (text.includes("Lead PM")) {  // ✅ Flexible condition to match variations
  //       cy.log("✅ Status contains 'Lead PM', proceeding with approval...");

  //       cy.get('tbody tr:first-child td:nth-child(1) .btnApprovalDetail').click();
  //       cy.get('#tarApprovalRemark').type('Remark_' + unique, { force: true });
  //       cy.log("⚠️ Skipping remark input...");

  //       cy.get("#btnConfirm").then(($btn) => {
  //         if ($btn.is(':visible') && !$btn.is(':disabled')) {
  //           cy.get('#slsAssignedPM').select(userPMFO, { force: true });
  //           cy.wrap($btn).click();
  //           cy.log("✅ Button clicked successfully");
  //           cy.wait(6000);
  //         } else {
  //           cy.log("⚠️ Button not clickable, skipping...");
  //         }
  //       });

  //     } else {
  //       cy.log("⚠️ Status does not match, skipping approval step.");
  //     }
  //   });

  //   cy.wait(4000);
  //   cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');

  // });

  // //ARO
  // it('ARO Test Case', () => {
  //   // Lead PM
  //   cy.visit('http://tbgappdev111.tbg.local:8042');
  //   cy.get('#tbxUserID').type(userARO);
  //   cy.get('#tbxPassword').type(pass);
  //   cy.get('#RefreshButton').click();

  //   cy.window().then((window) => {
  //     const rightCode = window.rightCode;
  //     cy.log('Right Code:', rightCode);
  //     cy.get('#captchaInsert').type(rightCode);
  //   });

  //   cy.get('#btnSubmit').click();
  //   cy.wait(2000);
  //   cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval');
  //   cy.wait(2000);
  //   // cy.get('#tbxSearchSONumber').type(sonumb).should(() => {
  //   //   // Log the test result if button click is successful
  //   //   testResults.push({
  //   //     Test: 'User AM melakukan klik tombol Search di Stip approval',
  //   //     Status: 'Pass',
  //   //     Timestamp: new Date().toISOString(),
  //   //   });
  //   // }); // << Search Filter SONumber  disable it if u dont need

  //   cy.contains('label', /^\s *By SO Number\s*$/)
  //     .click(); // search By Radio Button SONumber
  //   cy.get('#tbxApprovalSONumber').type(sonumb).should(() => {
  //     // Log the test result if button click is successful
  //     testResults.push({
  //       Test: 'User AM melakukan klik tombol Search di Stip approval',
  //       Status: 'Pass',
  //       Timestamp: new Date().toISOString(),
  //     });
  //   }); // << Search Filter


  //   cy.get('.btnSearch').first().click().should(() => {
  //     // Log the test result if button click is successful
  //     testResults.push({
  //       Test: 'User AM melakukan klik tombol Search di Stip approval',
  //       Status: 'Pass',
  //       Timestamp: new Date().toISOString(),
  //     });
  //   });
  //   cy.wait(2000);

  //   cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
  //     const text = $cell.text().trim();
  //     cy.log("📌 Status Found:", text);

  //     if (text === "Waiting for Confirmation ARO") {
  //       cy.log("✅ Status matches, proceeding with approval...");

  //       cy.get('tbody tr:first-child .btnApprovalDetail').click();
  //       cy.get('#tarApprovalRemark').type('Remark_' + unique, { force: true });
  //       // SKIPPING 'tarApprovalRemark' input field
  //       cy.log("⚠️ Skipping remark input...");

  //       // Attempt to click the approval button only if it's visible and enabled
  //       cy.get("#btnConfirm").then(($btn) => {
  //         if ($btn.is(':visible') && !$btn.is(':disabled')) {
  //           cy.wrap($btn).click();

  //           cy.log("✅ Button clicked successfully");
  //         } else {
  //           cy.log("⚠️ Button not clickable, skipping...");
  //         }
  //       });

  //     } else {
  //       cy.log("⚠️ Status does not match, skipping approval step.");
  //     }
  //     cy.wait(2000);
  //   });
  //   cy.wait(2000);
  //   cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');

  //   cy.then(() => {
  //     exportToExcel(testResults);
  //   });

  // });
});
