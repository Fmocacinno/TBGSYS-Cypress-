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
  const loopCount = 1; // Jumlah iterasi loop

  for (let i = 0; i < loopCount; i++) {
    it(`passes iteration ${i + 1}`, () => {
      const testResults = []; // Array to store test results
      const randomValue = Math.floor(Math.random() * 1000) + 1; // Random number between 1 and 1000
      const RangerandomValue = Math.floor(Math.random() * 20) + 1; // Random number between 1 and 1000
      const unique = `APP_${randomValue}`;

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
      let maxPages = 10; // Batasi jumlah halaman yang dicek untuk menghindari loop tak terbatas
      let currentPage = 1;

      // Function to find and click the button
      // Function to find and click the button
      // Function to find and click the button
      function checkRowsSequentially(page = 1, maxPages = 10) {
        if (page > maxPages) return;

        cy.get('tbody tr .btnSelect:visible:not(:disabled)').first().then(($btn) => {
          if ($btn.length) {
            cy.wrap($btn).click();
          } else {
            // No available buttons, go to the next page
            cy.get('#tblSite_paginate a[title="Next"]:visible').then(($next) => {
              if ($next.length) {
                cy.wrap($next).click();
                cy.wait(2000);
                checkRowsSequentially(page + 1, maxPages);
              }
            });
          }
        });
      }
      const minLength = 5;
      const maxLength = 15;
      const randomString = generateRandomString(minLength, maxLength);
      const date = "10-Feb-2025";
      const user = "201301180003"
      const pass = "123456"
      const filePath = 'documents/pdf/receipt.pdf';
      const latMin = -11.0; // Southernmost point
      const latMax = 6.5;   // Northernmost point
      const longMin = 94.0; // Westernmost point
      const longMax = 141.0; // Easternmost point

      // Generate random latitude and longitude within bounds
      const lat = (Math.random() * (latMax - latMin) + latMin).toFixed(6);
      const long = (Math.random() * (longMax - longMin) + longMin).toFixed(6);

      Cypress.on('uncaught:exception', (err, runnable) => {
        return false
      })

      cy.visit('http://tbgappdev111.tbg.local:8042/Login')

      cy.get('#tbxUserID').type(user).should('have.value', user).then(() => {
        // Log the test result if input is successful
        testResults.push({
          Test: 'User ID Input',
          Status: 'Pass',
          Timestamp: new Date().toISOString(),
        });
      });

      cy.get('#tbxPassword').type(pass)
        .should('have.value', pass)
        .then(() => {
          // Log the test result if input is successful
          testResults.push({
            Test: 'Passworhasbeeninputed',
            Status: 'Pass',
            Timestamp: new Date().toISOString(),
          });
        });

      cy.get('#RefreshButton').click();

      cy.window().then((window) => {
        const rightCode = window.rightCode;
        cy.log('Right Code:', rightCode);

        cy.get('#captchaInsert').type(rightCode);
      });

      cy.get('#btnSubmit').click();
      cy.url().should('include', 'http://tbgappdev111.tbg.local:8042/Dashboard'); // Ensure the page changes or some result occurs
      testResults.push({
        Test: 'Button Clicked',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });

      // Export results to Excel after the test





      cy.wait(2000)

      cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Input')
      cy.url().should('include', 'http://tbgappdev111.tbg.local:8042/STIP/Input'); // Ensure the page changes or some result occurs
      testResults.push({
        Test: 'User masuk ke Page Stip Input',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
      cy.wait(2000)

      cy.get('#slsSTIPCategory').then(($select) => {
        cy.wrap($select).select('1', { force: true })
          .should('have.value', '1') // Menunggu hingga value benar-benar berubah
          .then(() => {
            // Log the test result if selection is successful
            testResults.push({
              Test: 'User memilih STIPCategory di dropdown Colocation STIPCategory',
              Status: 'Pass',
              Timestamp: new Date().toISOString(),
            });
          });
      })

      cy.get('#slsProduct').then(($select) => {
        cy.wrap($select).select('2', { force: true })
          .should('have.value', '2') // Menunggu hingga value benar-benar berubah
          .then(() => {
            // Log the test result if selection is successful
            testResults.push({
              Test: 'User memilih Product di dropdown Colocation Product',
              Status: 'Pass',
              Timestamp: new Date().toISOString(),
            });
          });
      })

      cy.get('#slsColocationCompany').then(($select) => {
        cy.wrap($select).select('TB', { force: true })
          .should('have.value', 'TB') // Menunggu hingga value benar-benar berubah
          .then(() => {
            // Log the test result if selection is successful
            testResults.push({
              Test: 'User memilih Customer di dropdown Colocation Customer',
              Status: 'Pass',
              Timestamp: new Date().toISOString(),
            });
          });
      })

      cy.get('#slsColocationCustomer').then(($select) => {
        cy.wrap($select).select('XL', { force: true })
          .should('have.value', 'XL') // Menunggu hingga value benar-benar berubah
          .then(() => {
            // Log the test result if selection is successful
            testResults.push({
              Test: 'User memilih Customer di dropdown Colocation Customer',
              Status: 'Pass',
              Timestamp: new Date().toISOString(),
            });
          });
      })

      cy.get('#slsColocationRegion').then(($select) => {
        cy.wrap($select).select('1', { force: true })
          .should('have.value', '1') // Menunggu hingga value benar-benar berubah
          .then(() => {
            // Log the test result if selection is successful
            testResults.push({
              Test: 'User memilih region di dropdown Colocation Region',
              Status: 'Pass',
              Timestamp: new Date().toISOString(),
            });
          });
      });


      cy.get('#btnColocationPriceAmountPopUp')
        .should('be.visible') // Pastikan tombol terlihat sebelum diklik
        .click()
        // .should('be.disabled') // Opsional: Pastikan tombol berubah status setelah diklik (jika ada perubahan status)
        .then(() => {
          // Log the test result if button click is successful
          testResults.push({
            Test: 'User melakukan click tombol Colocation Price Amount Popup',
            Status: 'Pass',
            Timestamp: new Date().toISOString(),
          });
        });
      cy.get('tbody > tr:nth-child(3) .btnSelect')
        .should('be.visible') // Pastikan tombol terlihat sebelum diklik
        .click()
        // .should('be.disabled') // Opsional: Pastikan tombol berubah status setelah diklik (jika ada perubahan status)
        .then(() => {
          // Log the test result if button click is successful
          testResults.push({
            Test: 'User melakukan click tombol Price Amount Popup',
            Status: 'Pass',
            Timestamp: new Date().toISOString(),
          });
        });
      cy.wait(2000)

      cy.get('.slsBatchSLD').eq(0)
        .select('29', { force: true });

      cy.get('.slsBatchSLD').eq(1)
        .select('29', { force: true });

      cy.get('.slsBatchSLD').eq(2)
        .select('29', { force: true });

      cy.get('.slsBatchSLD').eq(3)
        .select('29', { force: true });

      cy.get('.slsBatchSLD').eq(4)
        .select('29', { force: true });

      cy.get('.slsBatchSLD').eq(5)
        .select('29', { force: true });

      cy.get('.slsBatchSLD').eq(6)
        .select('29', { force: true });

      cy.get('.slsBatchSLD').eq(7)
        .select('29', { force: true });

      cy.get('.slsBatchSLD').eq(8)
        .select('29', { force: true });

      cy.get('#btnColocationSitePopUp')
        .should('be.visible') // Pastikan tombol terlihat sebelum diklik
        .click()
        // .should('be.disabled') // Opsional: Pastikan tombol berubah status setelah diklik (jika ada perubahan status)
        .then(() => {
          // Log the test result if button click is successful
          testResults.push({
            Test: 'User melakukan click tombol Colocation Price Amount Popup',
            Status: 'Pass',
            Timestamp: new Date().toISOString(),
          });
        });

      cy.get('.col-md-6.col-sm-6 select[name="tblSite_length"]')
        .should('be.visible') // Ensure it's visible
        .select('50')
        .should('have.value', '50'); // Confirm the selection worked
      // Step 2: Click all enabled buttons in the third row

      cy.wait(2000)
      // Find and click the button
      checkRowsSequentially();


      cy.get('#tbxColocationCustomerSiteID').type('Site_' + unique);
      cy.get('#tbxColocationCustomerSiteName').type('Cust_' + unique);

      cy.get('#slsColocationDocumentOrder').then(($select) => {
        cy.wrap($select).select('7', { force: true })
      })

      cy.get('#tbxColocationDocumentName').type('DoctName_' + unique);

      cy.get('#fleColocationDocument').attachFile(filePath);


      cy.get('#fleColocationColocationForm').attachFile(filePath);

      cy.get('#slsColocationLeadProjectManager').then(($select) => {
        cy.wrap($select).select('201102180019', { force: true })
      })
      cy.get('#slsColocationAccountManager').then(($select) => {
        cy.wrap($select).select('201301180003', { force: true })
      })

      cy.get('#slsNewTowerHeight').then(($select) => {
        cy.wrap($select).select('0', { force: true })
      })

      cy.get('#slsColocationShelterType').then(($select) => {
        cy.wrap($select).select('4', { force: true })
      })
      cy.get('#tbxColocationPLNPowerKVA').type(RangerandomValue);

      cy.get('#dpkColocationRFITarget')
        .invoke('val', date)
        .trigger('change');

      cy.get('#slsColocationMLANumber').then(($select) => {
        cy.wrap($select).select('0010-14-F07-39033', { force: true })
      })
      cy.get('#tbxColocationLeasePeriod').type(RangerandomValue);

      cy.get('#dpkColocationMLADate')
        .invoke('val', date)
        .trigger('change');

      cy.get('#tarColocationRemark').type('Remark' + unique);


      cy.get("#btnSubmitColocation").click();

      cy.wait(2000)

      cy.get('.sa-confirm-button-container button.confirm').click();

      cy.wait(15000)

      // Add this section to extract values from popup
      cy.get('p.lead.text-muted').should('be.visible').then(($el) => {
        const text = $el.text();

        // Extract SO Number using regex
        const soNumber = text.match(/\bSO Number = (\d+)\b/)[1];
        // Extract Site ID using regex
        const siteId = text.match(/\bSite ID = (\d+)\b/)[1];

        // Save to aliases for later use
        cy.wrap(soNumber).as('soNumber');
        cy.wrap(siteId).as('siteId');

        // Optional: log the values to Cypress console
        cy.log(`Captured SO Number: ${soNumber}`);
        cy.log(`Captured Site ID: ${siteId}`);
      });

      // To use these values later in the test or other tests:
      cy.get('@soNumber').then((soNumber) => {
        cy.log(`Using SO Number: ${soNumber}`);
        // Add your logic here using the SO Number
      });

      cy.get('@siteId').then((siteId) => {
        cy.log(`Using Site ID: ${siteId}`);

        cy.get('@soNumber').then((soNumber) => {
          cy.get('@siteId').then((siteId) => {
            const filePath = Cypress.config('fileServerFolder') + '/cypress/e2e/STIP_1/COLLOCATION_MACRO/soDataCOLLOCATION_MACRO.json';
            cy.writeFile(filePath, { soNumber, siteId });

          });
        });
        // Add your logic here using the Site ID
      });


      cy.contains('a', 'Log Out').click({ force: true });
      cy.then(() => {
        exportToExcel(testResults);
      });

    })
  }
})
