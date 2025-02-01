import 'cypress-file-upload';

describe('template spec', () => {
  it('passes', () => {
    // ---random Random Value----//
    const randomValue = Math.floor(Math.random() * 1000) + 1; // Random number between 1 and 1000
    const unique = `APP_PKP_${randomValue}`;
    // ---random Character----//
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
    // Declare minLength and maxLength before use
    const minLength = 5;
    const maxLength = 15;

    // Ensure that the minLength and maxLength variables are initialized before calling generateRandomString
    
    const randomString = generateRandomString(minLength, maxLength);

    // ---random Character----//
    const date = "25-Feb-2025";
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

    cy.visit('http://tbgappdev111:8042/Login')

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

    cy.visit('http://tbgappdev111:8042/STIP/Input')

    cy.wait(2000)

    cy.get('#slsSTIPCategory').then(($select) => {
      cy.wrap($select).select('4', { force: true })
    })

    cy.get('#slsProduct').then(($select) => {
      cy.wrap($select).select('93', { force: true })
    })

    cy.get('#slsAssetSupportCompany').then(($select) => {
      cy.wrap($select).select('PT. PERMATA KARYA PERDANA', { force: true })
    })

    cy.get('#slsAssetSupportCustomer').then(($select) => {
      cy.wrap($select).select('PKP', { force: true })
    })

    cy.get('#slsAssetSupportRegion').then(($select) => {
      cy.wrap($select).select('1', { force: true })
    })
    cy.get('#slsAssetSupportBatch').then(($select) => {
      cy.wrap($select).select('UMU - Q3 AOP 2024 Batch 2', { force: true })
    })


    cy.get('#tbxAssetSupportSiteName').type('Site_' + unique);
    cy.get('#tbxAssetSupportCustomerSiteID').type('Cust_' + unique);

    cy.get('#slsNewDocumentOrder').then(($select) => {
      cy.wrap($select).select('7', { force: true })
    })
    cy.get('#slsAssetSupportDocumentOrder').then(($select) => {
      cy.wrap($select).select('BAK', { force: true })
    })

    cy.get('#tbxAssetSupportDocumentName').type('DoctName_' + unique);

    cy.get('#fleAssetSupportDocument').attachFile(filePath);



    cy.get('#slsAssetSupportProvince').then(($select) => {
      cy.wrap($select).select('11', { force: true })
    })
    cy.wait(2000)
    cy.get('#slsAssetSupportResidence').then(($select) => {
      cy.wrap($select).select('178', { force: true })
    })

    cy.get('#tbxAssetSupportNomLatitude').type(lat);
    cy.get('#tbxAssetSupportNomLongitude').type(long);

    cy.get('#tbxAssetSupportNomLatitudeEnd').type(lat);
    cy.get('#tbxAssetSupportNomLongitudeEnd').type(long);

    cy.get('#slsAssetSupportLeadProjectManager').then(($select) => {
      cy.wrap($select).select('201103180014', { force: true })
    })
    cy.get('#slsAssetSupportAccountManager').then(($select) => {
      cy.wrap($select).select('201301180003', { force: true })
    })

    cy.get('#tarAssetSupportRemark').type('REAMRK NIH DARY ' + randomString);

    cy.get("#btnSubmitAssetSupport").click();

    cy.wait(2000)

    cy.get('.sa-confirm-button-container button.confirm').click();

    cy.wait(15000)

    cy.visit('http://tbgappdev111:8042/Login/Logout')
  })
})
