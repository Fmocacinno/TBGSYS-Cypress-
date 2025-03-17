// Fungsi untuk menghasilkan nilai acak dalam rentang tertentu


const randomValue = Math.floor(Math.random() * 1000) + 1; // Random number between 1 and 1000
const randomRangeValue = (min, max) => Math.floor(Math.random() * (max - min + 1)) + min;
const minLength = 5;
const maxLength = 15;
const randomString = generateRandomString(minLength, maxLength);
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
  const sorFilePath = "documents/SOR/1a0863_TO_0492_1550_14_20.sor"; // .sor file path
  const photoFilePath = "documents/IMAGE/adopt.png"; // Photo file path\
  const excelfilepath = "documents/EXCEL/.xlsx/EXCEL_(1).xlsx"; // Photo file path


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

    cy.get('#tbxUserID').type(userLeadPM);
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
    cy.document().then((doc) => {
      const h2Element = doc.querySelector('h2'); // Cek apakah elemen h2 ada
      if (h2Element && h2Element.textContent.includes('Error on System')) {
        cy.log('🚨 Error pop-up detected! Clicking OK.');
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
    cy.wait(2000);

    cy.get('tr')
      .filter((index, element) => Cypress.$(element).find('td').first().text().trim() === '8') // Find the row where the first column contains '6'    cy.wait(2000);

      .find('td:nth-child(2) .btnSelect') // Find the button in the second column
      .click(); // Click the button
    cy.wait(5000);

    // cy.get('#divGenerateInputsOTDRFarEnd')
    //   .find('[id^="divApprovalAllFarEndOTDR"]') // Select all approval divs dynamically
    //   .each(($div, index) => {
    //     const approvalDivId = $div.attr('id'); // Get the ID
    //     const radioName = `rdoApprovalFarEndOTDR${index + 1}`; // Increment radio button name

    //     cy.wrap($div).then(($el) => {
    //       if ($el.is(':visible') && $el.css('display') !== 'none' && !$el.prop('disabled')) {
    //         cy.log(`✅ Approval section ${approvalDivId} is active`);

    //         cy.get(`input[name="${radioName}"][value="2"]`)
    //           .parent()
    //           .click({ force: true });


    //         cy.log(`🎯 Clicked "OK (All)" radio button for: ${radioName}`);
    //       } else {
    //         cy.log(`⚠️ Skipping ${approvalDivId}, it is disabled or hidden`);
    //       }
    //     });
    //   });
    //-- generate approval 
    cy.get('#divGenerateInputsOTDRFarEnd', { timeout: 10000 })
      .should('exist')
      .should('be.visible')
      .then(() => {
        cy.wait(4000); // Tunggu 4 detik sebelum mencari elemen jika perlu
        cy.get('#divGenerateInputsOTDRFarEnd')
          .find('[id^="divApprovalAllFarEndOTDR"]', { timeout: 10000 })
          .should('exist')
          .and('be.visible')
          .each(($div, index) => {
            const approvalDivId = $div.attr('id'); // Get the ID
            const radioName = `rdoApprovalFarEndOTDR${index + 1}`; // Increment radio button name
            const remarkInputId = `tarRemarkRejectFarEndOTDR${index + 1}`; // Dynamic remark input ID
            const randomValue = Math.floor(Math.random() * 1000); // Generate a random remark value

            cy.wrap($div).then(($el) => {
              if ($el.is(':visible') && $el.css('display') !== 'none' && !$el.prop('disabled')) {
                cy.log(`✅ Approval section ${approvalDivId} is active`);

                // Tunggu radio button sebelum klik
                cy.get(`input[name="${radioName}"][value="2"]`, { timeout: 5000 })
                  .should('exist')
                  .and('not.be.disabled')
                  .parent()
                  .click({ force: true });
                cy.log(`🎯 Clicked "OK (All)" radio button for: ${radioName}`);

                // Tunggu remark input sebelum mengetik
                cy.get(`#${remarkInputId}`, { timeout: 5000 })
                  .should('exist')
                  .and('be.visible')
                  .and('not.be.disabled')
                  .clear()
                  .type(`Remark ${randomValue}`);
                cy.log(`📝 Entered remark in ${remarkInputId}: Remark ${randomValue}`);
              } else {
                cy.log(`⚠️ Skipping ${approvalDivId}, it is disabled or hidden`);
              }
            });
          });

        cy.wait(3000);


        cy.get('.nav-tabs a[href="#tabOTDRNearEnd"]').click();
        cy.wait(3000);

        cy.get('#divGenerateInputsOTDRNearEnd')
          .find('[id^="divApprovalAllNearEndOTDR"]') // Select all approval divs dynamically
          .each(($div, index) => {
            const approvalDivId = $div.attr('id'); // Get the ID dynamically
            const radioName = `rdoApprovalNearEndOTDR${index + 1}`; // Increment radio button name
            const remarkInputId = `tarRemarkRejectNearEndOTDR${index + 1}`; // Dynamic remark input ID


            cy.wrap($div).then(($el) => {
              if ($el.is(':visible') && $el.css('display') !== 'none' && !$el.prop('disabled')) {
                cy.log(`✅ Approval section ${approvalDivId} is active`);

                cy.get(`input[name="${radioName}"][value="2"]`)
                  .should('exist')
                  .and('not.be.disabled')
                  .parent() // Click the styled container div if needed
                  .click({ force: true });
                cy.log(`🎯 Clicked "OK (All)" radio button for: ${radioName}`);

                // Ensure remark input is visible before typing
                cy.get(`#${remarkInputId}`)
                  .should('exist')
                  .and('be.visible')
                  .and('not.be.disabled')
                  .clear()
                  .type(`Remark ${randomValue}`);
                cy.log(`📝 Entered remark in ${remarkInputId}: Remark ${randomValue}`);
              } else {
                cy.log(`⚠️ Skipping ${approvalDivId}, it is disabled or hidden`);
              }
            });
          });

        cy.get('#divGenerateInputsOPMFarEnd')
          .find('[id^="divApprovalAllFarEndOPM"]') // Select all approval divs dynamically
          .each(($div, index) => {
            const approvalDivId = $div.attr('id'); // Get the ID dynamically
            const radioName = `rdoApprovalFarEndOPM${index + 1}`; // Fix radio button name
            const remarkInputId = `tarRemarkRejectFarEndOPM${index + 1}`; // Dynamic remark input ID

            cy.wrap($div).then(($el) => {
              if ($el.is(':visible') && $el.css('display') !== 'none' && !$el.prop('disabled')) {
                cy.log(`✅ Approval section ${approvalDivId} is active`);

                cy.get(`input[name="${radioName}"][value="2"]`)
                  .should('exist')
                  .and('not.be.disabled')
                  .parent() // Click the styled container div if needed
                  .click({ force: true });
                cy.log(`🎯 Clicked "OK (All)" radio button for: ${radioName}`);

                // Ensure remark input is visible before typing
                cy.get(`#${remarkInputId}`)
                  .should('exist')
                  .and('be.visible')
                  .and('not.be.disabled')
                  .clear()
                  .type(`Remark ${randomValue}`);
                cy.log(`📝 Entered remark in ${remarkInputId}: Remark ${randomValue}`);
              } else {
                cy.log(`⚠️ Skipping ${approvalDivId}, it is disabled or hidden`);
              }
            });
          });



        // cy.contains('.portlet-title', 'OPM') // Find the "OPM" accordion
        //   .parent() // Move to the parent container
        //   .find('.tools .expand') // Locate the expand button
        //   .click({ force: true }); // Click to expand



        cy.get('.nav-tabs a[href="#tabOPMNearEnd"]').click();
        cy.wait(3000);

        cy.get('#divGenerateInputsOPMNearEnd') // Select the main container
          .find('[id^="divApprovalAllNearEndOPM"]') // Select all approval divs dynamically
          .each(($div, index) => {
            const approvalDivId = $div.attr('id'); // Get the ID dynamically
            const radioName = `rdoApprovalNearEndOPM${index + 1}`; // Increment radio button name
            const remarkInputId = `tarRemarkRejectNearEndOPM${index + 1}`; // Dynamic remark input ID
            const randomValue = Math.floor(Math.random() * 1000); // Generate a random remark value

            cy.wrap($div).then(($el) => {
              if ($el.is(':visible') && $el.css('display') !== 'none' && !$el.prop('disabled')) {
                cy.log(`✅ Approval section ${approvalDivId} is active`);

                cy.get(`input[name="${radioName}"][value="2"]`)
                  .should('exist')
                  .and('not.be.disabled')
                  .parent() // Click the styled container div if needed
                  .click({ force: true });
                cy.log(`🎯 Clicked "OK (All)" radio button for: ${radioName}`);

                // Ensure remark input is visible before typing
                cy.get(`#${remarkInputId}`)
                  .should('exist')
                  .and('be.visible')
                  .and('not.be.disabled')
                  .clear()
                  .type(`Remark ${randomValue}`);
                cy.log(`📝 Entered remark in ${remarkInputId}: Remark ${randomValue}`);
              } else {
                cy.log(`⚠️ Skipping ${approvalDivId}, it is disabled or hidden`);
              }
            });
          });
        cy.get('#tarOTDRInstallationApprovalRemark').type('Remark' + randomValue);
        cy.wait(2000);

        cy.get('#btnApprove').click();
        cy.wait(5000);
        // cy.get('.sa-confirm-button-container button.confirm').click();
        cy.get('.sweet-alert.showSweetAlert.visible', { timeout: 15000 })
          .should('be.visible')
          .contains('Success');

        cy.get('.sa-confirm-button-container .confirm') // Target tombol "OK"
          .should('be.visible')
          .click();

        cy.contains('a', 'Log Out').click({ force: true });
        cy.then(() => {
          exportToExcel(testResults);
        });
      });


  });

});
