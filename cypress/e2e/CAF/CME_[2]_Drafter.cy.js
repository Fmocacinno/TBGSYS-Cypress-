import 'cypress-file-upload';

describe('template spec', () => {
  it('passes', () => {
    const unique = "ATP_001"
    const random = 55
    const siteid = "010016109"
    const phone = "0851728319"
    const date = "2-Jan-2025";
    const user = "201110190095"
    const pass = "123456"
    const cafNumber = "00238192/CAF/API/032025"
    const filePath = 'documents/pdf/receipt.pdf';
    const lat = 5.4
    const long = 94


    const userAssignedDrafter = "201910150170"

    Cypress.on('uncaught:exception', (err, runnable) => {
      return false
    })

    cy.visit('http://tbgappdev111:8044/Marketing/CAF/Grid')

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

    cy.visit('http://tbgappdev111:8044/Marketing/CAF/Grid')

    cy.wait(5000)

    cy.get('#tbxSearchCAFNumber').type(cafNumber);
    // cy.get('#tbxSearchSiteIDNon').type(siteid); 


    cy.get('.btnSearch').first().click({ force: true });
    // cy.get('.btnSearchNon').first().click({ force: true });

    cy.wait(6000)

    cy.get('[title="Detail"]').click({ force: true }); // Clicks the button with title "Detail"

    cy.get('#slsDrafter').then(($select) => {
      cy.wrap($select).select('555508170046', { force: true })
    })
    cy.get('#btnSubmitDrafter').first().click({ force: true });
  })
})
