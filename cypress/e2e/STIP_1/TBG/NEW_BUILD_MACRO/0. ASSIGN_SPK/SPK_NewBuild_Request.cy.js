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
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, UserRequestSPKProject, Uservendor, baseUrlVP, baseUrlTBGSYS, login, dashboard, menu1, menu2, menu3, menu4, logout;

  before(() => {
    testResults = []; // Reset results before all tests
  });

  after(() => {
    exportToExcel(testResults); // Export after all tests complete
  });
  beforeEach(() => {
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
      UserRequestSPKProject = values.UserRequestSPKProject;
      Uservendor = values.Uservendor;
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
  it('Material ON Site Approval PM FO', () => {

    cy.visit(`${baseUrlTBGSYS}${login}`);

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
      cy.wrap($select).select('1', { force: true })
    })
    cy.get('#slSubType').then(($select) => {
      cy.wrap($select).select('1', { force: true })
    })

    cy.get('#btnSearch').type(sonumb).should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User AM melakukan klik tombol Search di SPK Project',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    }); //



    cy.get('.blockUI', { timeout: 300000 }).should('not.exist');
    cy.wait(2000)

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

    cy.wait(10000);




    cy.get('#slCore').then(($select) => {
      cy.wrap($select).select('23', { force: true })
    })


    cy.wait(10000);
    cy.get('#slSubCore').then(($select) => {
      cy.wrap($select).select('45', { force: true })
    })
    cy.wait(6000);


    cy.get('#btnSearchVendor').click();


    cy.wait(2000)
    cy.get('#tbxSearchVendorName').type(Uservendor);
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

    cy.get('.sweet-alert.showSweetAlert.visible', { timeout: 10000 }).should('be.visible');

    cy.get('.sweet-alert h2').should('have.text', 'Success'); // Verify success message

    cy.get('.sweet-alert .confirm').click(); // Click the "OK" button
    cy.wait(2000)
    cy.contains('a', 'Log Out').click({ force: true });

  });


});
