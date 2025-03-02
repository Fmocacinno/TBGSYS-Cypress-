import 'cypress-file-upload';

const minLength = 5;
const maxLength = 15;
const randomString = generateRandomString(minLength, maxLength);
// Fungsi untuk menghasilkan nilai acak dalam rentang tertentu
const randomRangeValue = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
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
// Daftar indeks baris yang ingin diubah
const worktypeRows = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20];
const XLSX = require('xlsx');
const fs = require('fs');
const randomValue = Math.floor(Math.random() * 1000) + 1; // Random number between 1 and 1000

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
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, PICVendor, baseUrlVP, baseUrlTBGSYS, login, dashboard, menu1, menu2, menu3, menu4, logout;

  before(() => {
    testResults = []; // Reset results before all tests
  });

  after(() => {
    exportToExcel(testResults); // Export after all tests complete
  });

  beforeEach(() => {
    const testResults = []; // Array to store test results

    const user = "555504220025";
    const filePath = 'documents/pdf/receipt.pdf';

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

    // Check if the error pop-up is visible
    cy.get('h2').then(($h2) => {
      if ($h2.text().includes('Error on System')) {
        cy.log('🚨 Error pop-up detected! Clicking OK.');

        // Click the "OK" button
        cy.get('.confirm.btn-error').click();
      } else {
        cy.log('✅ No error pop-up detected.');
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
      .filter((index, element) => Cypress.$(element).find('td').first().text().trim() === '9') // Find the row where the first column contains '6'    cy.wait(2000);

      .find('td:nth-child(2) .btnSelect') // Find the button in the second column
      .click(); // Click the button
    cy.wait(5000);

    // cy.get('input[type="radio"][value="1"]')
    //   .filter('[name^="RFI"]')
    //   .first()
    //   .parent()
    //   .find('.iCheck-helper')
    //   .click();

    // cy.log('✅ Clicked the first radio button with value=1');

    // Loop through dynamically incremented names (adjust range as needed)
    cy.get(`input[name="RFI263"][value="1"]`).should('exist');

    cy.wait(2000); // Ensure page loads

    for (let i = 263; i <= 271; i++) {
      cy.get(`input[name="RFI${i}"][value="1"]`)
        .should('exist')
        .invoke('attr', 'style', 'opacity: 1; position: static;') // Make sure it's visible
        .siblings('.iCheck-helper')
        .click({ force: true });

      // Type a note in the corresponding text box
      cy.get(`#Notes${i}`).type(`Remark for RFI${i}`, { force: true });
    }
    cy.wait(3000);

    cy.get('input[name="rdoPICVendorType"][value="0"]')
      .parent('div.iradio_flat-blue')
      .click();
    // Generate a random number and type in tbxContactNo
    // cy.get('input[name="rdoIsPICNew"][value="1"]').check({ force: true });
    cy.wait(1000);

    cy.get('input[name="rdoIsPICNew"][value="1"]')
      .parent('div.iradio_flat-blue')
      .click();
    // Click radio button rdoPICVendorType with value 0
    // Type in tbxPICNonRegistered
    cy.wait(1000);

    cy.get('#tbxPICNonRegistered').type('PIC' + randomString + randomValue);
    cy.wait(1000);
    cy.get('#tbxContactNo').type('089' + randomValue + '928282923');
    cy.wait(1000);
    cy.get('#tarAddress')
      .should('be.visible')
      .type('This is a sample address', { force: true });



    // Select value 11 on dropdown slsProvince

    cy.get('#slsProvince').then(($select) => {
      cy.wrap($select).select('11', { force: true })
    })
    cy.wait(1000);
    // Select value 178 on dropdown slsResidence

    cy.get('#slsResidence').then(($select) => {
      cy.wrap($select).select('178', { force: true })
    })

    cy.get('input[name="rdoIsPICNew"][value="1"]')
      .parent('div.iradio_flat-blue')
      .click();
    cy.wait(1000);

    cy.get('input[name="rdoQuality"][value="1"]')
      .parent('div.iradio_flat-blue')
      .click();
    cy.wait(1000);
    cy.get('#tarRemarkQuality')
      .should('be.visible')
      .type('This is a sample address', { force: true });

    cy.get('input[name="rdoAccuracy"][value="1"]')
      .parent('div.iradio_flat-blue')
      .click();
    cy.wait(1000);
    cy.get('#tarRemarkAccuracy')
      .should('be.visible')
      .type('This is a sample address', { force: true });
    cy.get('#btnApprove').click();
    cy.wait(7000);
    // cy.get('.sa-confirm-button-container button.confirm').click();
    // cy.wait(2000);





    cy.contains('a', 'Log Out').click({ force: true });
    cy.then(() => {
      exportToExcel(testResults);
    });
  });


});
