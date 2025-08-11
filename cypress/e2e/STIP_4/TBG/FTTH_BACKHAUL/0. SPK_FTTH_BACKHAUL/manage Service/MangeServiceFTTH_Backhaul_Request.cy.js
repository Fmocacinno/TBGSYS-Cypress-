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
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, UserRequestSPKProject, Uservendor, UservendorManageService, baseUrlVP, baseUrlTBGSYS, login, dashboard, menu1, menu2, menu3, menu4, logout, UserRequestSPKFTTH;

  before(() => {
    testResults = []; // Reset results before all tests
  });

  after(() => {
    exportToExcel(testResults); // Export after all tests complete
  });
  beforeEach(() => {
    cy.readFile('cypress/e2e/STIP_4/TBG/FTTH_BACKHAUL/soDataBackhaul.json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;
    });

    cy.readFile('cypress/e2e/STIP_4/TBG/FTTH_BACKHAUL/DataVariable.json').then((values) => {
      cy.log(values);
      unique = values.unique;
      userAM = values.userAM;
      userInputStip = values.userInputStip;
      UserRequestSPKFTTH = values.UserRequestSPKFTTH;
      userLeadAM = values.userLeadAM;
      userLeadPM = values.userLeadPM;
      userPMFO = values.userPMFO;
      userARO = values.userARO;
      UserRequestSPKProject = values.UserRequestSPKProject;
      Uservendor = values.Uservendor;
      pass = values.pass;
      UservendorManageService = values.UservendorManageService;
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
  it('Material ON Site Approval PM FO', () => {

    cy.visit(`${baseUrlTBGSYS}${login}`);

    cy.get('#tbxUserID').type(UserRequestSPKFTTH);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get("#btnSubmit").click();
    cy.wait(2000);

    cy.visit(`${baseUrlTBGSYS}/BusinessSupport/SPKProject/List`)
      .url().should('include', `${baseUrlTBGSYS}/BusinessSupport/SPKProject/List`);
    testResults.push({
      Test: 'User PM FO melakukan akses ke menu Project activity Header',
      Status: 'Pass',
      timeStamp: new Date().toISOString(),
    });
    cy.wait(5000);

    cy.get('#btnRequest').invoke('removeAttr', 'target').click();

    cy.wait(5000);


    cy.get('#slType').then(($select) => {
      cy.wrap($select).select('12', { force: true })
    })
    cy.wait(2000);
    cy.get('#slSubType').then(($select) => {
      cy.wrap($select).select('54', { force: true })
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
        const fullUrl = `${baseUrlTBGSYS}${url}`; // Ensure full URL
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
    cy.wait(5000);
    cy.contains('.sa-confirm-button-container button', 'Ok').click();

    cy.wait(2000)
    cy.contains('a', 'Log Out').click({ force: true });
  });

});
