import 'cypress-file-upload';


// Daftar indeks baris yang ingin diubah
const worktypeRows = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20];
const XLSX = require('xlsx');
const fs = require('fs');
const randomStr = (length) => {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
};

const randomRangeValue = (min, max) =>
  (Math.floor(Math.random() * (max - min + 1)) + min).toString();
const randomNumber = randomRangeValue(1000000000000000, 9999999999999999);
// Example: fill all inputs with class .my-input with different random values
const randomNum = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min.toString();
// Function to export test results to Excel
const randnumber = (length) => {
  let result = '';
  for (let i = 0; i < length; i++) {
    result += Math.floor(Math.random() * 10); // 0–9
  }
  return result;
};
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
  let sonumb, siteId, unique, date, userAM, userLeadAM, userLeadPM, userARO, pass, userPMFO, userInputStip, PICVendor, baseUrlVP, baseUrlTBGSYS, login, dashboard, menu1, menu2, menu3, menu4, logout, PICVendorMobile1, PICVendorMobile2, dateleasestart, dateleaseend, userPMSitac;
  const baseId = 24; // Base ID
  const index = 1; // Increment index for unique IDs
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
      date = values.date;
      dateleasestart = values.dateleasestart;
      dateleaseend = values.dateleaseend;
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

  it('BAN BAK Input by vendor', () => {

    cy.visit(`${baseUrlTBGSYS}${login}`);


    cy.get('#tbxUserID').type(userPMSitac);
    cy.get('#tbxPassword').type(pass);


    cy.window().its('rightCode').then((rightCode) => {
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get("#btnSubmit").click();
    cy.wait(2000)

    cy.visit(`${baseUrlTBGSYS}/ProjectActivity/ProjectActivityHeader`)
      .url().should('include', `${baseUrlTBGSYS}/ProjectActivity/ProjectActivityHeader`);
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
        Test: 'User AM melakukan klik tombol Search di PROJECT ACTIVITY',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    });

    cy.wait(5000);
    cy.get('tbody tr:first-child td:nth-child(2)').then(($cell) => {
      const text = $cell.text().trim();
      cy.log("📌 Status Found:", text);
      cy.wait(4000);

      if (text.includes(sonumb)) {  // ✅ Checks if "Lead PM" is in the status
        cy.log("✅ Status contains 'SONUUMBER', proceeding with approval...");

        cy.get('tbody tr:first-child td:nth-child(1) .btnSelect').invoke('removeAttr', 'target').click();
      } else {
        cy.log("⚠️ Status does not match, skipping approval step.");
      }
    });
    cy.wait(5000);





    cy.get('body').then(($body) => {
      if ($body.find('#slsMobilePIC').length > 0 && $body.find('#slsMobilePIC').is(':visible')) {
        // Run the first set of actions
        cy.get('#slsMobilePIC').then(($select) => {
          cy.wrap($select).select(PICVendorMobile1, { force: true });
        });

        cy.get('#slsMobileCoPIC').then(($select) => {
          cy.wrap($select).select(PICVendorMobile2, { force: true });
        });

        cy.get("#btnMobilePICSubmit").click({ force: true });
        cy.wait(5000);
      }

      // Run the second block of code (always executed)
      cy.get('tr')
        .filter((index, element) => Cypress.$(element).find('td').first().text().trim() === '4') // Find the row where the first column contains '6'
        .find('td:nth-child(2) .btnSelect') // Find the button in the second column
        .click(); // Click the button
      cy.wait(1000);
    });


    // Always execute this part
    // cy.get('tr')
    //   .filter((index, element) => Cypress.$(element).find('td').first().text().trim() === '6')
    //   .find('td:nth-child(2) .btnSelect')
    //   .click();
    // cy.wait(1000);


    cy.get('td.dt-center .btnSelect').then(($btns) => {
      if ($btns.length > 0) {
        cy.wrap($btns.first()).click();
      } else {
        cy.log('No Select button found');
      }
    });

    cy.wait(6000);

    // cy.get('#btnCreate', { timeout: 10000 })
    //   .should('exist')
    //   .click({ force: true });
    // cy.wait(1000);

    // cy.get('#slsBAKType').then(($select) => {
    //   cy.wrap($select).select('1', { force: true })
    // })
    // cy.get('#slsAcquisitionStatus').then(($select) => {
    //   cy.wrap($select).select('0001', { force: true })
    // })
    // cy.get('#tbxOwnerIDCardNumber').type(`${randomNum(1000999, 9999999999999999)}`);

    // cy.get('#tbxOwnerName').clear().type(randomStr(8));
    // cy.wait(200);

    // cy.get('#tbxOwnerBehalf').type(randomStr(8) + `${randomNum(1000, 9999)}`); // 4-digit random
    // cy.get('#tarAddress').type(randomStr(50) + `${randomNum(1000, 9999)}`); // 4-digit random
    // cy.wait(200);
    // cy.get('#tbxOwnerPhone').type(randnumber(11)); // 4-digit random
    // cy.wait(200);
    // cy.get('#tbxOwnerOccupation').type(randnumber(11)); // 4-digit random
    // cy.wait(200);
    // cy.get('#tbxBankNumber').type(randnumber(11)); // 4-digit random
    // cy.wait(200);
    // cy.get('#slsBankName').then(($select) => {
    //   cy.wrap($select).select('60', { force: true })
    // })
    // cy.get('#tbxBankOwner').type(`${randomNum(1000, 99999999999)}`); // 4-digit random
    // cy.wait(200);
    // cy.get('#tbxLandLength').type(randnumber(2)); // 4-digit random

    // cy.wait(200);
    // cy.get('#tbxLandWidth').type(randnumber(2)); // 4-digit random

    // cy.wait(200);
    // cy.get('#tbxAccessLength').type(randnumber(2)); // 4-digit random

    // cy.wait(200);
    // cy.get('#tbxAccessWidth').type(randnumber(2)); // 4-digit random

    // cy.wait(200);
    // cy.get('#btnModalAddBAKPayment', { timeout: 10000 })
    //   .should('exist')
    //   .click({ force: true });

    // cy.wait(1000);
    // cy.get('#tbxPeriodYear').clear().type('21'); // 4-digit random

    // cy.wait(200);
    // cy.get('#dpkStartLease')
    //   .invoke('val', dateleasestart)
    //   .trigger('change');
    // cy.get('#slsLandStatus').then(($select) => {
    //   cy.wrap($select).select('0009', { force: true })
    // })
    // cy.get('#dpkEndLease')
    //   .invoke('val', dateleaseend)
    //   .trigger('change');

    // cy.get('#slsLandStatus').then(($select) => {
    //   cy.wrap($select).select('0009', { force: true })
    // })

    // cy.wait(200);
    // cy.get('#tbxNettTotal').type(randnumber(6)); // 4-digit random

    // cy.wait(200);
    // cy.get('#tbxDPPercent').type(randnumber(2)); // 4-digit random

    // cy.wait(200);
    // cy.get('#btnAddBakPayment', { timeout: 10000 })
    //   .should('exist')
    //   .click({ force: true });

    // cy.wait(1000);
    // cy.get('#slsOwnershipDocument').then(($select) => {
    //   cy.wrap($select).select('1', { force: true })
    // })

    // cy.wait(1000);
    // cy.get('#tbxOwnershipNumber').type(`${randomNum(1000, 999999)}`); // 4-digit random

    // cy.wait(200);
    // cy.get('#tbxNPWPNumber').type(`${randomNum(1000, 999999)}`); // 4-digit random

    // cy.wait(200);
    // cy.get('#tbxNPWPName').clear().type(randomStr(8));

    // cy.wait(200);
    // cy.get('#tarNPWPAddress').type(randomStr(50) + `${randomNum(1000, 9999)}`); // 4-digit random

    // cy.wait(200);
    // cy.get('#slsElectricity').then(($select) => {
    //   cy.wrap($select).select('1', { force: true })
    // })

    // cy.wait(200);
    // cy.get('#fleBAKDocument').attachFile(filePath);

    // cy.get('.nav-tabs a[href="#tabBAKUpload"]').click();
    // cy.wait(3000);

    // cy.wait(200);
    // cy.get('body').then($body => {
    //   if ($body.find('#fleOwnershipDocRFL').length > 0) {
    //     cy.get('#fleOwnershipDocRFL').attachFile(filePath);
    //   } else {
    //     cy.log('Ownership document upload field not found, skipping upload');
    //   }
    // });

    // cy.wait(200);
    // cy.get('body').then($body => {
    //   if ($body.find('#fleSPPT').length > 0) {
    //     cy.get('#fleSPPT').attachFile(filePath);
    //   } else {
    //     cy.log('Ownership document upload field not found, skipping upload');
    //   }
    // });

    // cy.wait(200);
    // cy.get('body').then($body => {
    //   if ($body.find('#fleOwnerID').length > 0) {
    //     cy.get('#fleOwnerID').attachFile(filePath);
    //   } else {
    //     cy.log('Ownership document upload field not found, skipping upload');
    //   }
    // });

    // cy.wait(200);
    // cy.get('body').then($body => {
    //   if ($body.find('#fleSpouseID').length > 0) {
    //     cy.get('#fleSpouseID').attachFile(filePath);
    //   } else {
    //     cy.log('Ownership document upload field not found, skipping upload');
    //   }
    // });

    // cy.wait(200);
    // cy.get('body').then($body => {
    //   if ($body.find('#fleKK').length > 0) {
    //     cy.get('#fleKK').attachFile(filePath);
    //   } else {
    //     cy.log('Ownership document upload field not found, skipping upload');
    //   }
    // });

    // cy.get('.nav-tabs a[href="#tabBAKInput"]').click();

    // cy.wait(2000);

    cy.wait(200);
    cy.get('#tarRemarkPM').type(randomStr(50) + `${randomNum(1000, 9999)}`); // 4-digit random


    cy.get("#btnApprove").click();
    cy.wait(5000);


    cy.wait(10000);
    cy.contains('a', 'Log Out').click({ force: true });
    cy.then(() => {
      exportToExcel(testResults);
    });
  });


});
