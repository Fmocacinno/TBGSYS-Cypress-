import 'cypress-file-upload';

describe('Template Spec', () => {
  let values;

  before(() => {
    cy.fixture('soData.json').then((data) => {
      values = data;
    });
  });

  it('passes', () => {
    const sonumb = values.soNumber;
    const siteId = values.siteId;

    const unique = "ATP_19";
    const date = "2-Jan-2025";
    const userAM = "201301180003";
    const userLeadAM = "201301180003";
    const userLeadPM = "201103180014";
    const pass = "123456";
    const SiteId = "1200271931201";
    const filePath = 'documents/pdf/receipt.pdf';

    Cypress.on('uncaught:exception', (err, runnable) => {
      return false;
    });

    // AM Login & Approval
    cy.visit('http://tbgappdev111.tbg.local:8042/Login');
    cy.get('#tbxUserID').type(userLeadPM);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();
    cy.wait(2000);
    cy.visit('http://tbgappdev111:8042/ProjectActivity/ProjectActivityHeader');
    cy.wait(2000);


    cy.get('#tbxSearchSONumber').type(sonumb);
    cy.get('.btnSearch').first().click();
    cy.get('tbody tr:first-child .btnApprovalDetail').click();
    cy.get('#tarApprovalRemark').type('Remark_' + unique, { force: true });
    cy.get("#btnApprove").click();
    cy.wait(2000);
    cy.visit('http://tbgappdev111:8042/Login/Logout');

    // Lead AM Login & Approval
    cy.visit('http://tbgappdev111:8042');
    cy.get('#tbxUserID').type(userLeadAM);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();
    cy.wait(2000);
    cy.visit('http://tbgappdev111:8042/STIP/Approval');
    cy.wait(2000);
    cy.get('#tbxSearchSONumber').type(sonumb);
    cy.get('.btnSearch').first().click();
    cy.get('tbody tr:first-child .btnApprovalDetail').click();
    cy.get('#tarApprovalRemark').type('Remark_' + unique);
    cy.get("#btnApprove").click();
    cy.wait(2000);
    cy.visit('http://tbgappdev111:8042/Login/Logout');

    // Lead PM Login & Approval
    cy.visit('http://tbgappdev111:8042');
    cy.get('#tbxUserID').type(userLeadPM);
    cy.get('#tbxPassword').type(pass);
    cy.get('#RefreshButton').click();

    cy.window().then((window) => {
      const rightCode = window.rightCode;
      cy.log('Right Code:', rightCode);
      cy.get('#captchaInsert').type(rightCode);
    });

    cy.get('#btnSubmit').click();
    cy.wait(2000);
    cy.visit('http://tbgappdev111:8042/STIP/Approval');
    cy.wait(2000);
    cy.get('#tbxSearchSONumber').type(sonumb);
    cy.get('.btnSearch').first().click();
    cy.get('tbody tr:first-child .btnApprovalDetail').click();
    cy.get('#tarApprovalRemark').type('Remark_' + unique);

    cy.get('#slsPMSitac').select('201103180014', { force: true });
    cy.get('#slsPMCME').select('201601600086', { force: true });
    cy.get('#slsSitacOfficer').select('201103180014', { force: true });
    cy.get('#slsFieldController').select('201103180014', { force: true });

    cy.get("#btnConfirm").click();
    cy.wait(2000);
    cy.visit('http://tbgappdev111:8042/Login/Logout');
  });
});
