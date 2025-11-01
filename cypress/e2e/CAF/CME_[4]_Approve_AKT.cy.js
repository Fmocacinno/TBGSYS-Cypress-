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
    const cafNumber = "0219856/CAF/API/032025"
    const filePath = 'documents/pdf/receipt.pdf';
    const lat = 5.4
    const long = 94


    const userAssignedDrafter = "201910150170"

    Cypress.on('uncaught:exception', (err, runnable) => {
      return false
    })

    cy.visit('http://tbgappdev111.tbg.local:8045/Marketing/CAF/Grid')

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

    cy.visit('http://tbgappdev111.tbg.local:8045/Marketing/CAF/Grid')

    cy.wait(5000)

    cy.get('#tbxSearchCAFNumber').type(cafNumber); 

    cy.get('.btnSearch').first().click({ force: true });

    cy.wait(6000)

    cy.get('[title="Detail"]').click({ force: true }); // Clicks the button with title "Detail"

    cy.get('#slsGeneralSoW').then(($select) => {
      cy.wrap($select).select('4', { force: true }) 
    })

    cy.wait(2000)
    
    cy.get('#slStrengtheningAnalysis').then(($select) => {
      cy.wrap($select).select('0', { force: true }) 
    })

    cy.get('#txtRemarkStrengthening').type('RemarkDrawing_' + unique); 
    
    
  })
})
