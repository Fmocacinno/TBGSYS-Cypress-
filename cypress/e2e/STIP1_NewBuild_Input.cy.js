import 'cypress-file-upload';

describe('template spec', () => {
  const loopCount = 1; // Jumlah iterasi loop

  for (let i = 0; i < loopCount; i++) {
    it(`passes iteration ${i + 1}`, () => {
      const randomValue = Math.floor(Math.random() * 1000) + 1; // Random number between 1 and 1000
      const unique = `APP_PKP_${randomValue}`;

      const date = "2-Jan-2025";
      const user = "555504220025"
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

      cy.get('#tbxUserID').type(user)
      cy.get('#tbxPassword').type(pass)

      cy.get('#RefreshButton').click();

      cy.window().then((window) => {
        const rightCode = window.rightCode;
        cy.log('Right Code:', rightCode);

        cy.get('#captchaInsert').type(rightCode);
      });

      cy.get('#btnSubmit').click();

      cy.wait(2000)

      cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Input')

      cy.wait(2000)

      cy.get('#slsSTIPCategory').then(($select) => {
        cy.wrap($select).select('1', { force: true })
      })

      cy.get('#slsProduct').then(($select) => {
        cy.wrap($select).select('1', { force: true })
      })

      cy.get('#slsNewCompany').then(($select) => {
        cy.wrap($select).select('PKP', { force: true })
      })

      cy.get('#slsNewCustomer').then(($select) => {
        cy.wrap($select).select('PKP', { force: true })
      })

      cy.get('#slsNewRegion').then(($select) => {
        cy.wrap($select).select('1', { force: true })
      })

      cy.get('#tbxNewSiteName').type('Site_' + unique);
      cy.get('#tbxNewCustomerSiteID').type('Cust_' + unique);

      cy.get('#slsNewDocumentOrder').then(($select) => {
        cy.wrap($select).select('7', { force: true })
      })

      cy.get('#tbxNewDocumentName').type('DoctName_' + unique);

      cy.get('#fleNewDocument').attachFile(filePath);

      cy.get('#slsNewProvince').then(($select) => {
        cy.wrap($select).select('12', { force: true })
      })

      cy.get('#tbxNewNomLatitude').type(lat);
      cy.get('#tbxNewNomLongitude').type(long);

      cy.get('#slsNewLeadProjectManager').then(($select) => {
        cy.wrap($select).select('201103180014', { force: true })
      })
      cy.get('#slsNewAccountManager').then(($select) => {
        cy.wrap($select).select('201301180003', { force: true })
      })

      cy.get('#slsNewTowerHeight').then(($select) => {
        cy.wrap($select).select('0', { force: true })
      })

      cy.get('#slsNewShelterType').then(($select) => {
        cy.wrap($select).select('7', { force: true })
      })
      cy.get('#tbxNewPLNPowerKVA').type(5);

      cy.get('#dpkNewRFITarget')
        .invoke('val', date)
        .trigger('change');

      cy.get('#slsNewMLANumber').then(($select) => {
        cy.wrap($select).select('Risalah Rapat 9 Jul 2015', { force: true })
      })

      cy.get('#tbxNewLeasePeriod').type(5);

      cy.get('#dpkNewMLADate')
        .invoke('val', date)
        .trigger('change');

      cy.get('#tarNewRemark').type('Remark' + unique);

      cy.get('#slsNewResidence').then(($select) => {
        cy.wrap($select).select('205', { force: true })
      })

      cy.get('#btnNewPriceAmountPopUp').click();

      cy.get('tbody tr:first-child .btnSelect').click();

      cy.wait(2000)

      cy.get('.slsBatchSLD').eq(0)
        .select('967', { force: true });

      cy.get('.slsBatchSLD').eq(1)
        .select('967', { force: true });


      cy.get('.slsBatchSLD').eq(2)
        .select('967', { force: true });

      cy.get('.slsBatchSLD').eq(3)
        .select('967', { force: true });

      cy.get('.slsBatchSLD').eq(4)
        .select('967', { force: true });


      cy.get('.slsBatchSLD').eq(5)
        .select('967', { force: true });

      cy.get('.slsBatchSLD').eq(6)
        .select('967', { force: true });

      cy.get('.slsBatchSLD').eq(7)
        .select('967', { force: true });

      cy.get('.slsBatchSLD').eq(8)
        .select('967', { force: true });

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
            const filePath = Cypress.config('fileServerFolder') + '/cypress/e2e/NewBuildMacro/soDataNewBuild.json';
            cy.writeFile(filePath, { soNumber, siteId });
          });
        });
        // Add your logic here using the Site ID
      });


      cy.visit('http://tbgappdev111.tbg.local:8042/Login/Logout')


    })
  }
})
