import 'cypress-file-upload';

describe('template spec', () => {
  it('passes', () => {
    const unique = "ATP_001"
    const random = 55
    const siteid = "010016109"
    const phone = "0851728319"
    const date = "2-Jan-2025";
    const user = "555508170046"
    const pass = "123456" 
    const cafNumber = "0219856/CAF/API/032025"
    const filePath = 'documents/pdf/receipt.pdf';
    const lat = 5.4
    const long = 94


    const userAssignedDrafter = "201910150170"

    Cypress.on('uncaught:exception', (err, runnable) => {
      return false
    })

    cy.visit('http://localhost:41699/Marketing/CAF/Grid')

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

    cy.visit('http://localhost:41699/Marketing/CAF/Grid')

    cy.wait(5000)

    cy.get('#tbxSearchCAFNumber').type(cafNumber); 

    cy.get('.btnSearch').first().click({ force: true });

    cy.wait(6000)

    cy.get('[title="Detail"]').click({ force: true }); // Clicks the button with title "Detail"

    cy.get('#slsGeneralSoW').then(($select) => {
      cy.wrap($select).select('4', { force: true }) 
    })

    cy.wait(2000)

    cy.contains('.ms-elem-selectable', 'NEW COLO').click();


    cy.get('#flSiteLayout')
    .attachFile(filePath)
    .trigger('change', { force: true });

    
    cy.get('#txtRemarkDrawing').type('Remark_' + unique); 
    
    // cy.get('#flSiteLayout')
    // .attachFile(filePath)
    // .trigger('change', { force: true });

    
    // cy.get('#txtRemarkDrawing').type('Remark_' + unique); 
    
    cy.get('#btnSubmitDrawing').first().click({ force: true });

  })
})
