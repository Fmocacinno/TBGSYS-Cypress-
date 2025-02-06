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
  it('passes', () => {
    // Array to store test results
    const testResults = []; // Array to store test results


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

    const randomValue = Math.floor(Math.random() * 1000) + 1; // Random number between 1 and 1000
    const RangerandomValue = Math.floor(Math.random() * 20) + 1; // Random number between 1 and 1000

    const unique = `APP_PKP_${randomValue}`;

    const date = "2-Jan-2025";
    const user = "201301180003"
    const pass = "123456"
    // const filePath = 'documents/pdf/receipt.pdf';
    const filePath = 'documents/pdf/receipt.pdf';
    // const latMin = -11.0; // Southernmost point
    // const latMax = 6.5;   // Northernmost point
    // const longMin = 94.0; // Westernmost point
    // const longMax = 141.0; // Easternmost point

    const latMin = -6.0; // Southernmost point
    const latMax = 6.5;   // Northernmost point
    const longMin = 105.0; // Westernmost point
    const longMax = 107.0; // Easternmost point


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

    cy.get('#tbxPassword').type(pass).should('have.value', pass).then(() => {
      // Log the test result if input is successful
      testResults.push({
        Test: 'Input Password',
        Status: 'Pass',
        Timestamp: new Date().toISOString(),
      });
    });

    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode).should('have.value', rightCode).then(() => {
        // Log the CAPTCHA input success
        testResults.push({
          Test: 'CAPTCHA Inserted',
          Status: 'Pass',
          Timestamp: new Date().toISOString(),
        });
      });
    });

    // reporting for button sukses click with link
    cy.get('#btnSubmit').click();
    cy.url().should('include', 'http://tbgappdev111.tbg.local:8042/Dashboard'); // Ensure the page changes or some result occurs
    testResults.push({
      Test: 'Button Clicked',
      Status: 'Pass',
      Timestamp: new Date().toISOString(),
    });


    cy.wait(2000)

    cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Input')

    cy.wait(2000)

    cy.get('#slsSTIPCategory').then(($select) => {
      cy.wrap($select).select('1', { force: true })
    })

    cy.get('#slsProduct').then(($select) => {
      cy.wrap($select).select('83', { force: true })
    })
    // if TBG -TB
    cy.get('#slsIntersiteFOCompany').then(($select) => {
      cy.wrap($select).select('PKP', { force: true })
    })

    cy.get('#slsIntersiteFOCustomer').then(($select) => {
      cy.wrap($select).select('XL', { force: true })
    })

    cy.get('#slsIntersiteFORegion').then(($select) => {
      cy.wrap($select).select('1', { force: true })
    })

    cy.wait(1000); // Ensure dropdown selection is applied
    /// radio button with regex
    cy.contains('label', /^\s *Segment\s*$/)
      .click(); // Click the label

    // Assert that the "Segment" radio button is selected
    cy.get('input[type="radio"][value="Segment"]').should('be.checked');

    cy.get('#btnIntersiteFOPriceAmountPopUp').click();

    cy.get('tbody tr:first-child .btnSelect').click();

    cy.wait(2000)

    // handle HTML

    // cy.get('.slsBatchSLD').eq(0)
    //   .select('180', { force: true });

    // cy.get('.slsBatchSLD').eq(1)
    //   .select('180', { force: true });


    // cy.get('.slsBatchSLD').eq(2)
    //   .select('180', { force: true });

    // cy.get('.slsBatchSLD').eq(3)
    //   .select('180', { force: true });

    // cy.get('.slsBatchSLD').eq(4)
    //   .select('180', { force: true });


    // cy.get('.slsBatchSLD').eq(5)
    //   .select('180', { force: true });

    // cy.get('.slsBatchSLD').eq(6)
    //   .select('180', { force: true });

    // cy.get('.slsBatchSLD').eq(7)
    //   .select('180', { force: true });

    // cy.get('.slsBatchSLD').eq(8)
    //   .select('180', { force: true });
    // handle HTML

    cy.get('#tbxIntersiteFOSiteName').type('Site_' + unique);
    cy.get('#tbxIntersiteFOCustomerSiteID').type('Cust_' + unique);
    cy.get('#tbxIntersiteFOSPKWOLOINumber').type('WO_' + unique);

    cy.get('#slsIntersiteFODocumentOrder').then(($select) => {
      cy.wrap($select).select('7', { force: true })
    })

    cy.get('#tbxIntersiteFODocumentName').type('DoctName_' + unique);

    cy.get('#fleIntersiteFODocument').attachFile(filePath);

    cy.get('#slsIntersiteFOProvince').then(($select) => {
      cy.wrap($select).select('11', { force: true })
    })
    cy.wait(2000)
    // Near ENd
    cy.get('#slsIntersiteFOResidence').then(($select) => {
      cy.wrap($select).select('178', { force: true })
    })
    cy.get('#slsIntersiteFOTowerProviderNearEnd').then(($select) => {
      cy.wrap($select).select('10', { force: true })
    })
    cy.get('#tbxIntersiteFOTowerIDNearEnd').type('Site_' + unique);
    cy.get('#tbxIntersiteFOTowerNameNearEnd').type('SiteName_' + unique);
    cy.get('#tbxLatitudeA').type(lat);
    cy.get('#tbxLongitudeA').type(long);

    // Far end
    cy.get('#slsIntersiteFOResidence').then(($select) => {
      cy.wrap($select).select('178', { force: true })
    })
    cy.get('#slsIntersiteFOTowerProviderFarEnd').then(($select) => {
      cy.wrap($select).select('11', { force: true })
    })
    cy.get('#tbxIntersiteFOTowerIDFarEnd').type('Site_' + unique);
    cy.get('#tbxIntersiteFOTowerNameFarEnd').type('SiteName_' + unique);
    cy.get('#tbxLatitudeB').type(lat);
    cy.get('#tbxLongitudeB').type(long);





    cy.get('#slsIntersiteFOLeadProjectManager').then(($select) => {
      cy.wrap($select).select('201103180014', { force: true })
    })
    cy.wait(2000)

    cy.get('#slsIntersiteFOAccountManager').then(($select) => {
      cy.wrap($select).select('201301180003', { force: true })
    })

    cy.get('#tbxFOCore').type(RangerandomValue);
    cy.get('#tbxFOLengthSPK').type(RangerandomValue);
    cy.get('#tbxIntersiteFOSegment').type('Segment_' + unique);

    // Debug: Log all labels inside the container
    cy.get('div.icheck-inline label').each(($label) => {
      cy.log($label.text().trim()); // Log the text of each label
    });

cy.get('div.form-group.divNewIntersiteFO')
  .should('be.visible');
  
    cy.get('div.icheck-inline label').each(($label) => {
      cy.log($label.text().trim()); // Log the text of each label
    });

    // Select the "No" radio button
    cy.get('div.icheck-inline')
      .contains('label',  /No/i)
      .click();

    // Assert that the "No" radio button is checked
    cy.get('div.icheck-inline')
      .contains('label', 'No')
      .find('input[type="radio"]')
      .should('be.checked');

    cy.get('#slsNewTowerHeight').then(($select) => {
      cy.wrap($select).select('32', { force: true })
    })

    cy.get('#slsNewShelterType').then(($select) => {
      cy.wrap($select).select('3', { force: true })
    })
    cy.get('#tbxNewPLNPowerKVA').type(5);

    cy.get('#tbxNewNumberOfAntenna').type(RangerandomValue);
    cy.get('#tbxNewNumberOfSectoral').type(RangerandomValue);



    cy.get('#dpkNewRFITarget')
      .invoke('val', date)
      .trigger('change');

    cy.get('#tbxNewLeasePeriod').type(RangerandomValue);



    cy.get('#slsNewMLANumber').then(($select) => {
      cy.wrap($select).select('0031-14-F07-121782', { force: true })
    })

    cy.get('#tbxNewLeasePeriod').type(5);

    cy.get('#dpkNewMLADate')
      .invoke('val', date)
      .trigger('change');

    cy.get('#tarNewRemark').type('Remark' + unique);






    cy.get("#btnSubmitNew").click();

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
          cy.writeFile('cypress/fixtures/soDataIntersite.json', { soNumber, siteId });
        });
      });
      // Add your logic here using the Site ID
    });

    cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout')

    // Export results to Excel after the test
    cy.then(() => {
      exportToExcel(testResults);
    });

  })
})
