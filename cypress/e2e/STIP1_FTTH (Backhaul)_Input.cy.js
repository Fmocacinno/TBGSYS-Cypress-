import 'cypress-file-upload';

describe('template spec', () => {
  const loopCount = 5; // Jumlah iterasi loop

  for (let i = 0; i < loopCount; i++) {
    it(`passes iteration ${i + 1}`, () => {
      const randomValue = Math.floor(Math.random() * 1000) + 1;
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

      cy.visit('http://tbgappdev111:8042/Login');

      cy.get('#tbxUserID').type(user);
      cy.get('#tbxPassword').type(pass);
      cy.get('#RefreshButton').click();

      cy.window().then((window) => {
        const rightCode = window.rightCode;
        cy.log('Right Code:', rightCode);
        cy.get('#captchaInsert').type(rightCode);
      });

      cy.get('#btnSubmit').click();
      cy.wait(2000);
      cy.visit('http://tbgappdev111:8042/STIP/Input');
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

      cy.visit('http://tbgappdev111:8042/Login/Logout');
    });
  }
});
