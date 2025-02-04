import 'cypress-file-upload';

describe('STIP Approval Workflow', () => {
  it('completes approval process using JSON data', () => {
    // Read data from fixture file first
    cy.readFile('cypress/fixtures/soData.json').then((values) => {
      // Extract values from JSON
      const sonumb = values.soNumber;
      const siteId = values.siteId;
      const unique = "ATP_19";
      const date = "2-Jan-2025";
      const userAM = "201301180003";
      const userLeadAM = "201301180003";
      const userLeadPM = "201106180055";
      const pass = "123456";
      const filePath = 'documents/pdf/receipt.pdf';

      Cypress.on('uncaught:exception', () => false);

      // AM Approval
      cy.visit('http://tbgappdev111.tbg.local:8042/Login');
      cy.get('#tbxUserID').type(userAM);
      cy.get('#tbxPassword').type(pass);
      cy.get('#RefreshButton').click();

      cy.window().then((window) => {
        cy.get('#captchaInsert').type(window.rightCode);
      });

      cy.get('#btnSubmit').click();
      cy.wait(2000);

      // Approval Process
      cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval');
      cy.wait(2000);

      // Using JSON-derived SO Number
      cy.get('#tbxSearchSONumber').type(sonumb);
      cy.get('.btnSearch').first().click();
      cy.get('tbody tr:first-child .btnApprovalDetail').click();

      cy.get('#tarApprovalRemark').type(`Remark_${unique}`, { force: true });
      cy.get("#btnApprove").click();
      cy.wait(2000);
      cy.visit('http://tbgappdev111:8042/Login/Logout');

      // Lead AM Approval
      cy.visit('http://tbgappdev111:8042');
      cy.get('#tbxUserID').type(userLeadAM);
      cy.get('#tbxPassword').type(pass);
      cy.get('#RefreshButton').click();

      cy.window().then((window) => {
        cy.get('#captchaInsert').type(window.rightCode);
      });

      cy.get('#btnSubmit').click();
      cy.wait(2000);

      cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval');
      cy.wait(2000);

      // Reuse the same SO Number
      cy.get('#tbxSearchSONumber').type(sonumb);
      cy.get('.btnSearch').first().click();
      cy.get('tbody tr:first-child .btnApprovalDetail').click();

      cy.get('#tarApprovalRemark').type(`Remark_${unique}`);
      cy.get("#btnApprove").click();
      cy.wait(2000);
      cy.visit('http://tbgappdev111:8042/Login/Logout');

      // Lead PM Approval
      cy.visit('http://tbgappdev111.tbg.local:8042');
      cy.get('#tbxUserID').type(userLeadPM);
      cy.get('#tbxPassword').type(pass);
      cy.get('#RefreshButton').click();

      cy.window().then((window) => {
        cy.get('#captchaInsert').type(window.rightCode);
      });

      cy.get('#btnSubmit').click();
      cy.wait(2000);

      cy.visit('http://tbgappdev111.tbg.local:8042/STIP/Approval');
      cy.wait(2000);

      // Final SO Number usage
      cy.get('#tbxSearchSONumber').type(sonumb);
      cy.get('.btnSearch').first().click();
      cy.get('tbody tr:first-child .btnApprovalDetail').click();

      cy.get('#tarApprovalRemark').type(`Remark_${unique}`);

      // Selection handlers
      cy.get('#slsPMSitac').select('201103180014', { force: true });
      cy.get('#slsPMCME').select('201601600086', { force: true });
      cy.get('#slsSitacOfficer').select('201103180014', { force: true });
      cy.get('#slsFieldController').select('201103180014', { force: true });

      cy.get("#btnConfirm").click();
      cy.wait(2000);
      cy.visit('http://tbgappdev111:8042/Login/Logout');
    });
  });
});