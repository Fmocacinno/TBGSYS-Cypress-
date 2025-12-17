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
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, UserSPKProject, Uservendor, UserRequestSPKApproval, PICVendor, baseUrlVP, baseUrlTBGSYS, login, dashboard, menu1, menu2, menu3, menu4, logout;
  before(() => {
    testResults = []; // Reset results before all tests
  });

  after(() => {
    exportToExcel(testResults); // Export after all tests complete
  });
  beforeEach(() => {
    cy.readFile('cypress/e2e/STIP_1/TBG/COLLOCATION_MACRO_FWA/SURGE/CollocationMacroFWASURGE/soDataCOLLOCATION_MACRO_FWA(SURGE).json').then((values) => {
      cy.log(values);
      sonumb = values.soNumber;
      siteId = values.siteId;
    });
    cy.readFile('cypress/e2e/STIP_1/TBG/COLLOCATION_MACRO_FWA/SURGE/CollocationMacroFWASURGE/DataVariable.json').then((values) => {
      cy.log(values);
      unique = values.unique;
      userAM = values.userAM;
      userInputStip = values.userInputStip;
      userLeadAM = values.userLeadAM;
      userLeadPM = values.userLeadPM;
      userPMFO = values.userPMFO;
      userARO = values.userARO;
      UserSPKProject = values.UserSPKProject;
      UserRequestSPKApproval = values.UserRequestSPKApproval;
      Uservendor = values.Uservendor;
      pass = values.pass;
      PICVendor = values.PICVendor;
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
    cy.visit(`${baseUrlVP}${login}`);
    cy.get('#tbUserID').type(PICVendor);
    cy.get('#tbPassword').type(pass);


    cy.window().its('rightCode').then((rightCode) => {
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });


    cy.get("#btnsubmit").click();
    cy.wait(2000);

    cy.visit(`${baseUrlVP}/BusinessSupport/SPKProject/List`)
      .url().should('include', `${baseUrlVP}/BusinessSupport/SPKProject/List`);
    testResults.push({
      Test: 'User PM FO melakukan akses ke menu Project activity Header',
      Status: 'Pass',
      timeStamp: new Date().toISOString(),
    });
    cy.wait(5000);


    cy.get('#slType').then(($select) => {
      cy.wrap($select).select('2', { force: true })
    })
    cy.get('#slSubType').then(($select) => {
      cy.wrap($select).select('11', { force: true })
    })
    cy.get('#btnSearch').type(sonumb).should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User AM melakukan klik tombol Search di Stip approval',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    }); // 




    cy.get('.blockUI', { timeout: 300000 }).should('not.exist');


    cy.get('#txtSONumber').type(sonumb).should(() => {
      // Log the test result if button click is successful
      testResults.push({
        Test: 'User AM melakukan klik tombol Search di Stip approval',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    }); // << Search Filter SONumber  disable it if u dont need
    cy.get('td[rowspan="1"][colspan="1"]')
      .eq(0)
      .find('.btnSearch')
      .click();

    cy.wait(6000);
    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.get('.blockUI', { timeout: 300000 }).should('not.exist');
      cy.log("📌 Status Found:", sonumb);
      cy.get('.blockUI', { timeout: 300000 }).should('not.exist');
    });


    // cy.get('a.btnSelect')
    //   .should('have.attr', 'href') // Ensure the element has an href
    //   .then((href) => {
    //     const baseUrl = "http://tbgappdev111.tbg.local:8128/BusinessSupport/SPKProject/"; // Base URL
    //     const fullUrl = new URL(href, baseUrl).href; // Correctly construct the full URL
    //     cy.visit(fullUrl); // Visit the page
    //   });
    // // cy.get('a.btnSelect', { timeout: 10000 })
    // //   .should('have.attr', 'href')
    // //   .then((href) => {
    // //     const baseUrl = "http://tbgappdev111.tbg.local:8128";
    // //     cy.visit(`${ baseUrl }${ href }`);
    // //   });
    // cy.get('.btnSelect').first().click();
    // cy.get('.btnSelect').first().should('have.attr', 'href').then((href) => {
    //   cy.visit(`http://tbgappdev111.tbg.local:8128${href}`);
    // });
    // Prevent new tabs by stubbing window.open

    cy.get('.btnSelect').first().trigger('click', { force: true });
    cy.get('.btnSelect').first().should('have.attr', 'href').then((href) => {
      cy.visit(`${baseUrlVP}${href}`);
    });

    cy.wait(6000);


    cy.get('#txtRemarkApproval').type('Remark' + unique);

    cy.get('input[name="radAction"][value="Approve"]').parent().click()

    cy.get('#btnSubmit').click();
    cy.wait(2000);
    cy.contains('.sa-confirm-button-container button', 'Ok').click();

    cy.wait(5000)
    cy.get('.sa-confirm-button-container .confirm').click();
    cy.wait(5000)

    cy.contains('a', 'Log Out').click({ force: true });
    cy.then(() => {
      exportToExcel(testResults);
    });

  });

});
