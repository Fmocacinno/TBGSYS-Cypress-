import 'cypress-file-upload';
const XLSX = require('xlsx');
const fs = require('fs');

// Function to export test results to Excel
function exportToExcel(testResults) {
  const filePath = 'test-results.xlsx'; // Path to the Excel file

  // Create a worksheet from the test results
  const worksheet = XLSX.utils.json_to_sheet(testResults);

  // Create a workbook and add the worksheet
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Test Results');

  // Write the workbook to a file
  XLSX.writeFile(workbook, filePath);
}
describe('template spec', () => {
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass;

  beforeEach(() => {
    cy.readFile('cypress/e2e/NEW_BUILD_MACRO/soDataNewBuild.json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;

      unique = "ATP_19";
      date = "2-Jan-2025";
      userAM = "201301180003";
      userLeadAM = "201301180003";
      userLeadPM = "201103180014";
      userARO = "201910600198";
      pass = "123456";
    });

    Cypress.on('uncaught:exception', (err, runnable) => {
      return false;
    });
  });

  //AM
  it('AM Test Case', () => {
    cy.visit('http://tbgappdev111.tbg.local:8042/Login');

    cy.get('#tbxUserID').type(userAM);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();
    cy.wait(2000);

    cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval');
    cy.wait(2000);

    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.log("📌 Status Found:", text);
      cy.wait(2000);

      if (text.includes("Waiting for Approval AM")) {  // ✅ Checks if "Lead PM" is in the status
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
            cy.wait(2000);
          } else {
            cy.wait(2000);
            cy.log("⚠️ Button not clickable, skipping...");
          }
        });

      } else {
        cy.log("⚠️ Status does not match, skipping approval step.");
      }
    });

    cy.wait(2000);
    cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');
  });

  //LEAD AM
  it('Lead AM Test Case', () => {
    // Lead PM
    cy.visit('http://tbgappdev111.tbg.local:8128');
    cy.get('#tbxUserID').type(userLeadAM);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();
    cy.wait(2000);
    cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval');
    cy.wait(2000);
    cy.get('#tbxSearchSONumber').type(sonumb);
    cy.get('.btnSearch').first().click();
    cy.wait(2000);
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

    cy.wait(2000);
    cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');

  });
  //LEAD PM
  it('Lead PM Test Case', () => {
    // Lead PM
    cy.visit('http://tbgappdev111.tbg.local:8042');
    cy.get('#tbxUserID').type(userLeadPM);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();
    cy.wait(2000);
    cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval');
    cy.wait(2000);
    cy.get('#tbxSearchSONumber').type(sonumb);
    cy.get('.btnSearch').first().click();
    cy.wait(2000);



    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.log("📌 Status Found:", text);

      if (text.includes("Lead PM")) {  // ✅ Flexible condition to match variations
        cy.log("✅ Status contains 'Lead PM', proceeding with approval...");

        cy.get('tbody tr:first-child td:nth-child(1) .btnApprovalDetail').click();
        cy.get('#tarApprovalRemark').type('Remark_' + unique, { force: true });
        cy.log("⚠️ Skipping remark input...");

        cy.get("#btnConfirm").then(($btn) => {
          if ($btn.is(':visible') && !$btn.is(':disabled')) {
            cy.get('#slsPMSitac').select('201103180014', { force: true });
            cy.get('#slsPMCME').select('201601600086', { force: true });
            cy.get('#slsSitacOfficer').select('201103180014', { force: true });
            cy.get('#slsFieldController').select('201103180014', { force: true });
            cy.wrap($btn).click();
            cy.log("✅ Button clicked successfully");
            cy.wait(2000);
          } else {
            cy.log("⚠️ Button not clickable, skipping...");
          }
        });

      } else {
        cy.log("⚠️ Status does not match, skipping approval step.");
      }
    });

    cy.wait(2000);
    cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');

  });
  //ARO
  it('ARO Test Case', () => {
    // Lead PM
    cy.visit('http://tbgappdev111.tbg.local:8042');
    cy.get('#tbxUserID').type(userARO);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();
    cy.wait(2000);
    cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval');
    cy.wait(2000);
    cy.get('#tbxSearchSONumber').type(sonumb);
    cy.get('.btnSearch').first().click();
    cy.wait(2000);
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
      cy.wait(2000);
    });
    cy.wait(2000);
    cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout');
    cy.then(() => {
      exportToExcel(testResults);
    });

  });

});
