import 'cypress-file-upload';

describe('template spec', () => {
  it('passes', () => {
    const unique = "ATP_19"
    const date = "2-Jan-2025";
    const user = "0002012"
    const pass = "123456" 
    const filePath = 'documents/pdf/receipt.pdf';
    const lat = 5.4
    const long = 94

    Cypress.on('uncaught:exception', (err, runnable) => {
      return false
    })

    cy.visit('http://tbgappdev111:8024/Login')

    cy.get('#tbUserID').type(user)
    cy.get('#tbPassword').type(pass)

    cy.window().then((window) => {
      const rightCode = window.rightCode;  
      cy.log('Right Code:', rightCode);

      cy.get('#captchaInsert').type(rightCode); 
    });

    cy.get('#btnsubmit').click();
    
    cy.wait(2000)

    cy.visit('http://tbgappdev111:8024/ProjectActivity/ProjectActivityTransaction/1334311')

    cy.wait(6000)

    cy.get('#tblOSPFOInstallation tbody tr')
      .eq(0) // Select the second row (index 1)
      .find('.btnEditWorktype') // Find the button in the row
      .click({force: true});

    cy.get('#btnSaveWorktypeData').click()
    
    cy.wait(2000)

    cy.get('.sa-confirm-button-container button.confirm').click();
    
    cy.get('#tblOSPFOInstallation tbody tr')
      .eq(1) // Select the second row (index 1)
      .find('.btnEditWorktype') // Find the button in the row
      .click({force: true});

    cy.get('#btnSaveWorktypeData').click()

    cy.wait(2000)
    
    cy.get('.sa-confirm-button-container button.confirm').click();
    
    cy.get('#tblOSPFOInstallation tbody tr')
      .eq(2) // Select the second row (index 1)
      .find('.btnEditWorktype') // Find the button in the row
      .click({force: true});

    cy.get('#btnSaveWorktypeData').click()
    
    cy.wait(2000)
    
    cy.get('.sa-confirm-button-container button.confirm').click();

    // cy.get('#tblGDL tbody tr')
    //   .eq(0) // Select the second row (index 1)
    //   .find('.btnAddSONumberAdditional') // Find the button in the row
    //   .click({force: true});

    // cy.get('#tblGDL tbody tr').eq(1).click();

    // // Perform any assertions or actions after clicking row 2
    // cy.get('#tblGDL tbody tr').eq(1).should('have.class', 'expected-class'); // Example assertion

    // // Click on the fourth row
    // cy.get('#tblGDL tbody tr').eq(3).click();

    // // Perform any assertions or actions after clicking row 4
    // cy.get('#tblGDL tbody tr').eq(3).should('have.class', 'expected-class'); // Example assertion
  

    // cy.get('#slsSTIPCategory').then(($select) => {
    //   cy.wrap($select).select('1', { force: true }) 
    // })

    // cy.get('#slsProduct').then(($select) => {
    //   cy.wrap($select).select('1', { force: true }) 
    // })

    // cy.get('#slsNewCompany').then(($select) => {
    //   cy.wrap($select).select('PKP', { force: true }) 
    // })
    
    // cy.get('#slsNewCustomer').then(($select) => {
    //   cy.wrap($select).select('PKP', { force: true }) 
    // })
    
    // cy.get('#slsNewRegion').then(($select) => {
    //   cy.wrap($select).select('1', { force: true }) 
    // })

    // cy.get('#tbxNewSiteName').type('Site_' + unique); 
    // cy.get('#tbxNewCustomerSiteID').type('Cust_' + unique); 
    
    // cy.get('#slsNewDocumentOrder').then(($select) => {
    //   cy.wrap($select).select('7', { force: true }) 
    // })
    
    // cy.get('#tbxNewDocumentName').type('DoctName_' + unique); 
    
    // cy.get('#fleNewDocument').attachFile(filePath);

    // cy.get('#slsNewProvince').then(($select) => {
    //   cy.wrap($select).select('12', { force: true }) 
    // })
    
    // cy.get('#tbxNewNomLatitude').type(lat); 
    // cy.get('#tbxNewNomLongitude').type(long); 

    // cy.get('#slsNewLeadProjectManager').then(($select) => {
    //   cy.wrap($select).select('201106180055', { force: true }) 
    // })
    // cy.get('#slsNewAccountManager').then(($select) => {
    //   cy.wrap($select).select('201301180003', { force: true }) 
    // })
    
    // cy.get('#slsNewTowerHeight').then(($select) => {
    //   cy.wrap($select).select('0', { force: true }) 
    // })

    // cy.get('#slsNewShelterType').then(($select) => {
    //   cy.wrap($select).select('7', { force: true }) 
    // })
    // cy.get('#tbxNewPLNPowerKVA').type(5); 
    
    // cy.get('#dpkNewRFITarget')
    //   .invoke('val', date)  
    //   .trigger('change'); 

    // cy.get('#slsNewMLANumber').then(($select) => {
    //   cy.wrap($select).select('Risalah Rapat 9 Jul 2015', { force: true }) 
    // })
      
    // cy.get('#tbxNewLeasePeriod').type(5); 

    // cy.get('#dpkNewMLADate')
    // .invoke('val', date)  
    // .trigger('change');  

    // cy.get('#tarNewRemark').type('Remark' + unique); 
    
    // cy.get('#slsNewResidence').then(($select) => {
    //   cy.wrap($select).select('205', { force: true }) 
    // })

    // cy.get('#btnNewPriceAmountPopUp').click();

    // cy.get('tbody tr:first-child .btnSelect').click();

    // cy.wait(2000)

    // cy.get('.slsBatchSLD').eq(0)  
    //   .select('967', { force: true });  

    // cy.get('.slsBatchSLD').eq(1)  
    //   .select('967', { force: true });


    // cy.get('.slsBatchSLD').eq(2)  
    //   .select('967', { force: true });

    // cy.get('.slsBatchSLD').eq(3)  
    //   .select('967', { force: true });

    // cy.get('.slsBatchSLD').eq(4)  
    //   .select('967', { force: true });

    
    // cy.get('.slsBatchSLD').eq(5)  
    //   .select('967', { force: true });

    // cy.get('.slsBatchSLD').eq(6)  
    //   .select('967', { force: true });

    // cy.get('.slsBatchSLD').eq(7)  
    //   .select('967', { force: true });

    // cy.get('.slsBatchSLD').eq(8)  
    //   .select('967', { force: true });

    // cy.get("#btnSubmitNew").click();

    // cy.wait(2000)

    // cy.get('.sa-confirm-button-container button.confirm').click();
    
    // cy.wait(15000)

    // cy.visit('http://tbgappdev111:8042/Login/Logout')
  })
})
