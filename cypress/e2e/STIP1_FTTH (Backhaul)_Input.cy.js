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
      const unique = `APP_PKP_${randomValue}`;

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
      const date = "25-Feb-2025";
      const user = "555504220025";
      const pass = "123456";
      const filePath = 'documents/pdf/receipt.pdf';
      const latMin = -11.0;
      const latMax = 6.5;
      const longMin = 94.0;
      const longMax = 141.0;

      const lat = (Math.random() * (latMax - latMin) + latMin).toFixed(6);
      const long = (Math.random() * (longMax - longMin) + longMin).toFixed(6);

      Cypress.on('uncaught:exception', (err, runnable) => false);

      cy.visit('http://tbgappdev111.tbg.local:8042/Login')

      cy.get('#tbxUserID').type(user).should('have.value', user).then(() => {
        // Log the test result if input is successful
        testResults.push({
          Test: 'User ID Input',
          Status: 'Pass',
          Timestamp: new Date().toISOString(),
        });
      });

      cy.get('#tbxPassword').type(pass).should('have.value', pass).then(() => {
        // Log the test result if input is successful
        testResults.push({
          Test: 'Passworhasben ',
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

      cy.wait(2000)

      cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Input')
      cy.url().should('include', 'http://tbgappdev111.tbg.local:8042/STIP/Input'); // Ensure the page changes or some result occurs
      testResults.push({
        Test: 'User masuk ke Page Stip Input',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
      cy.wait(2000);

      cy.get('#slsSTIPCategory').select('4', { force: true });
      cy.get('#slsProduct').select('93', { force: true });
      cy.get('#slsAssetSupportCompany').select('PT. PERMATA KARYA PERDANA', { force: true });
      cy.get('#slsAssetSupportCustomer').select('PKP', { force: true });
      cy.get('#slsAssetSupportRegion').select('1', { force: true });
      cy.get('#slsAssetSupportBatch').select('UMU - Q3 AOP 2024 Batch 2', { force: true });

      cy.get('#tbxAssetSupportSiteName').type('Site_' + unique);
      cy.get('#tbxAssetSupportCustomerSiteID').type('Cust_' + unique);

      cy.get('#slsNewDocumentOrder').select('7', { force: true });
      cy.get('#slsAssetSupportDocumentOrder').select('BAK', { force: true });

      cy.get('#tbxAssetSupportDocumentName').type('DoctName_' + unique);
      cy.get('#fleAssetSupportDocument').attachFile(filePath);

      cy.get('#slsAssetSupportProvince').select('11', { force: true });
      cy.wait(2000);
      cy.get('#slsAssetSupportResidence').select('178', { force: true });

      cy.get('#tbxAssetSupportNomLatitude').type(lat);
      cy.get('#tbxAssetSupportNomLongitude').type(long);
      cy.get('#tbxAssetSupportNomLatitudeEnd').type(lat);
      cy.get('#tbxAssetSupportNomLongitudeEnd').type(long);

      cy.get('#slsAssetSupportLeadProjectManager').select('201103180014', { force: true });
      cy.get('#slsAssetSupportAccountManager').select('201301180003', { force: true });

      cy.get('#tarAssetSupportRemark').type('REMARK NIH DARY ' + randomString);
      cy.get('#btnSubmitAssetSupport').click();

      cy.wait(2000);
      cy.get('.sa-confirm-button-container button.confirm').click();
      cy.wait(15000);
      // Add this section to extract values from popup

      // Extract values
      cy.get('p.lead.text-muted').should('be.visible').then(($el) => {
        const text = $el.text();
        const soNumber = text.match(/\bSO Number = (\d+)\b/)[1];
        const siteId = text.match(/\bSite ID = (\d+)\b/)[1];

        // Save values to alias
        cy.wrap(soNumber).as('soNumber');
        cy.wrap(siteId).as('siteId');
      });

      // Write to file outside of nested `then()`
      cy.then(() => {
        cy.get('@soNumber').then((soNumber) => {
          cy.get('@siteId').then((siteId) => {
            cy.writeFile('cypress/fixtures/soData.json', { soNumber, siteId });
          });
        });
      });

      cy.visit('http://tbgappdev111:8042/Login/Logout');

      cy.then(() => {
        exportToExcel(testResults);
      });

    });

  }
});
