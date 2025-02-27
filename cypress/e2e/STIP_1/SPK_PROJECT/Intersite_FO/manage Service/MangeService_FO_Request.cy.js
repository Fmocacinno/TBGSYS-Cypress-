import { timeStamp } from 'console';
import 'cypress-file-upload';
const XLSX = require('xlsx');
const fs = require('fs');

// Function to export test results to Excel
function exportToExcel(testResults) {
  const filePath = 'resultsApproval_NewBuildMacro.xlsx'; // Path to the Excel file

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
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, UserRequestSPKProject, Uservendor, UservendorManageService;

  before(() => {
    testResults = []; // Reset results before all tests
  });

  after(() => {
    exportToExcel(testResults); // Export after all tests complete
  });
  beforeEach(() => {
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
      UserRequestSPKProject = values.UserRequestSPKProject;
      Uservendor = values.Uservendor;
      pass = values.pass;
      UservendorManageService = values.UservendorManageService;
    });


    Cypress.on('uncaught:exception', (err, runnable) => {
      return false;
    });
  });

  //AM
  it('Material ON Site Approval PM FO', () => {

    cy.visit('http://tbgappdev111.tbg.local:8127/Login');

    cy.get('#tbxUserID').type(UserRequestSPKProject);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get("#btnSubmit").click();
    cy.wait(2000);

    cy.visit('http://tbgappdev111.tbg.local:8127/BusinessSupport/SPKProject/List')
      .url().should('include', 'http://tbgappdev111.tbg.local:8127/BusinessSupport/SPKProject/List');
    testResults.push({
      Test: 'User PM FO melakukan akses ke menu Project activity Header',
      Status: 'Pass',
      timeStamp: new Date().toISOString(),
    });
    cy.wait(5000);

    cy.get('#btnRequest').invoke('removeAttr', 'target').click();

    cy.wait(5000);

    cy.get('#slType').then(($select) => {
      cy.wrap($select).select('9', { force: true })
    })
    cy.get('#slSubType').then(($select) => {
      cy.wrap($select).select('91', { force: true })
    })
    cy.get('#btnSearch').type(sonumb).should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User AM melakukan klik tombol Search di SPK Project',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    }); //


    // cy.get('#tbxSearchSONumber').type(sonumb).should(() => {
    //   // Log the test result if button click is successful
    //   testResults.push({
    //     Test: 'User PM FO melakukan klik tombol Search di Stip approval',
    //     Status: 'Pass',
    //     Timestamp: new Date().toISOString(),
    //   });
    // }); // << Search Filter SONumber  disable it if u dont need

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


    cy.get('.blockUI', { timeout: 300000 }).should('not.exist');


    cy.get('#txtSONumber').type(sonumb).should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User AM melakukan klik tombol Search di SPK Project',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    }); // << Search Filter SONumber  disable it if u dont need
    cy.get('td[rowspan="1"][colspan="1"]')
      .eq(1)
      .find('.btnSearch')
      .click();

    cy.wait(6000);
    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.get('.blockUI', { timeout: 300000 }).should('not.exist');
      cy.log("📌 Status Found:", sonumb);
      cy.get('.blockUI', { timeout: 300000 }).should('not.exist');
    });


    cy.window().then((win) => {
      cy.stub(win, 'open').callsFake((url) => {
        const fullUrl = `http://tbgappdev111.tbg.local:8127${url}`; // Ensure full URL
        cy.visit(fullUrl);
      });
    });

    cy.get('tbody tr:first-child td:nth-child(1) .btnSelect').click();

    cy.wait(6000);




    cy.get('#slCore').then(($select) => {
      cy.wrap($select).select('54', { force: true })
    })

    cy.wait(6000);


    cy.get('#btnSearchVendor').click();


    cy.wait(2000)
    cy.get('#tbxSearchVendorName').type(UservendorManageService);
    cy.wait(2000)
    cy.get('td[rowspan="1"][colspan="1"]')
      .eq(1)
      .find('.btnSearch')
      .click();
    cy.wait(2000)
    cy.get('tbody tr:first-child td:nth-child(1) .btnSelect').click();
    cy.wait(2000);
    cy.get('#txtRemark').type('Remark' + unique);
    cy.get('#btnAssign').click();
    cy.wait(2000);
    cy.contains('.sa-confirm-button-container button', 'Ok').click();

    cy.wait(2000)
    cy.visit('http://tbgappdev111.tbg.local:8127/Login/Logout');
  });

  //LEAD AM 
  // it('Lead AM Test Case', () => {
  //   const testResults = [];
  //   // Lead PM
  //   cy.visit('http://tbgappdev111.tbg.local:8127');
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
  //   cy.visit('http://tbgappdev111.tbg.local:8127/STIP/Approval')
  //     .url().should('include', 'http://tbgappdev111.tbg.local:8127/STIP/Approval');
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
  //   cy.visit('http://tbgappdev111.tbg.local:8127/Login/Logout');

  // });
  // //LEAD PM
  // it('Lead PM Test Case', () => {
  //   // Lead PM
  //   cy.visit('http://tbgappdev111.tbg.local:8127');
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
  //   cy.visit('http://tbgappdev111.tbg.local:8127/STIP/Approval')
  //     .url().should('include', 'http://tbgappdev111.tbg.local:8127/STIP/Approval');
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
  //   cy.visit('http://tbgappdev111.tbg.local:8127/Login/Logout');

  // });

  // //ARO
  // it('ARO Test Case', () => {
  //   // Lead PM
  //   cy.visit('http://tbgappdev111.tbg.local:8127');
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
  //   cy.visit('http://tbgappdev111.tbg.local:8127/STIP/Approval');
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
  //   cy.visit('http://tbgappdev111.tbg.local:8127/Login/Logout');

  //   cy.then(() => {
  //     exportToExcel(testResults);
  //   });

  // });
});
