import 'cypress-file-upload';

// script
// SELECT RegionName, PICVendorID, * FROM dbo.vwPRJReadyToPay WHERE ID = '0044/0030739470031/00000/RTP/03/2024'
// SELECT * FROM dbo.trxPRJReadyToPay WHERE ID = '0044/0030739470031/00000/RTP/03/2024'
// UPDATE dbo.trxPRJReadyToPay SET ApprovalStatusIDDP = 335, mstApprovalStatusID = null WHERE ID = '0044/0030739470031/00000/RTP/03/2024'
// SELECT * FROM dbo.vwApproval WHERE UniqueID = '0044/0030739470031/00000/RTP/03/2024' ORDER BY CreatedDate desc

describe('template spec', () => {
  it('passes', () => {
    const unique = "ATP_19"
    const date = "2-Jan-2025";
    const user = "555510230156"
    const userCSA = "0033003"
    const userVendor = "0033008"
    const userPMSitac = "202307020084"
    const userSAS = "555504140256"
    const userSASSection = "201911020179"
    const sonumb = "0030739470031"
    const pass = "123456" 
    const filePath = 'documents/pdf/receipt.pdf';
    const lat = 5.4
    const long = 94

    const linkTBiGSys = "http://tbgappdev111:8050/"
    const linkVendor = "http://tbgappdev111:8024/"

    Cypress.on('uncaught:exception', (err, runnable) => {
      return false
    })

    // #region Submit Vendor

    cy.visit(linkVendor + 'Login')

    cy.get('#tbUserID').type(userVendor)
    cy.get('#tbPassword').type(pass)

    cy.window().then((window) => {
      const rightCode = window.rightCode;  
      cy.log('Right Code:', rightCode);

      cy.get('#captchaInsert').type(rightCode); 
    });
    
    cy.get('#btnsubmit').click();
    
    cy.wait(2000)

    cy.visit(linkVendor + 'Project/ProjectActivity/ReadyToPay')

    cy.wait(2000)
    
    cy.get('#txtSONumber').type(sonumb); 
    
    cy.get('#btnFilterGrid').first().click();
    
    cy.wait(3000)

    cy.get('table.tblSummaryData tbody tr td .btnView').first().click({force: true});

    // cy.get('#rtpDoc52864').trigger('click');
    // cy.get('#rtpDoc52864').attachFile(filePath);
    
    // cy.get('#rtpDoc52868').trigger('click');
    // cy.get('#rtpDoc52868').attachFile(filePath);

    // cy.get('#rtpDoc52869').trigger('click');
    // cy.get('#rtpDoc52869').attachFile(filePath);

    cy.get('#btnSubmitDP').click();

    cy.wait(2000)

    cy.get('.sa-confirm-button-container button.confirm').click();

    cy.wait(5000)

    cy.visit(linkVendor + 'Login/Logout')

    // #endregion

    // #region Sitac Admin

    cy.clearCookies();
    cy.clearLocalStorage();

    cy.visit(linkVendor + 'Login')

    cy.get('#tbUserID').type(userCSA)
    cy.get('#tbPassword').type(pass)

    cy.window().then((window) => {
      const rightCode = window.rightCode;  
      cy.log('Right Code:', rightCode);

      cy.get('#captchaInsert').type(rightCode); 
    });

    cy.get('#btnsubmit').click();
    
    cy.wait(2000)

    cy.visit(linkVendor + 'Project/ProjectActivity/ReadyToPay')

    cy.wait(2000)
    
    cy.get('#txtSONumber').type(sonumb); 
    
    cy.get('#btnFilterGrid').first().click();
    
    cy.wait(2000)

    cy.get('table.tblSummaryData tbody tr td .btnView').first().click({force: true});

    cy.wait(2000)

    cy.get('#slBulkyValidationSitacDP').then(($select) => {
      cy.wrap($select).select('1', { force: true }) 
    })

    cy.get('#slBulkyValidationLandOwnerDP').then(($select) => {
      cy.wrap($select).select('1', { force: true }) 
    })

    cy.get('#slBulkyValidationLandDP').then(($select) => {
      cy.wrap($select).select('1', { force: true }) 
    })

    cy.wait(2000)

    cy.get('#btnValidateDP').click();

    cy.wait(2000)

    cy.get('.sa-confirm-button-container button.confirm').click();

    cy.wait(5000)

    cy.visit(linkVendor + 'Login/Logout')

    // #endregion
    
    // #region PM Sitac
    
    cy.clearCookies();
    cy.clearLocalStorage();

    cy.visit(linkTBiGSys + 'Login')
    
    cy.wait(5000)

    cy.get('#tbxUserID').type(userPMSitac)
    cy.get('#tbxPassword').type(pass)

    cy.window().then((window) => {
      const rightCode = window.rightCode;  
      cy.log('Right Code:', rightCode);

      cy.get('#captchaInsert').type(rightCode); 
    });

    cy.get('#btnSubmit').click();
    
    cy.wait(2000)

    cy.visit(linkTBiGSys + 'Project/ProjectActivity/ReadyToPay')

    cy.wait(2000)
    
    cy.get('#txtSONumber').type(sonumb); 
    
    cy.get('#btnFilterGrid').first().click();
    
    cy.wait(2000)

    cy.get('table.tblSummaryData tbody tr td .btnView').first().click({force: true});

    cy.wait(2000)

    cy.get('#slBulkyValidationSitacDP').then(($select) => {
      cy.wrap($select).select('1', { force: true }) 
    })

    cy.get('#slBulkyValidationLandOwnerDP').then(($select) => {
      cy.wrap($select).select('1', { force: true }) 
    })

    cy.get('#slBulkyValidationLandDP').then(($select) => {
      cy.wrap($select).select('1', { force: true }) 
    })

    cy.wait(2000)
    
    cy.get('#btnValidateDP').click();

    cy.wait(2000)

    cy.get('.sa-confirm-button-container button.confirm').click();

    cy.wait(5000)

    cy.visit(linkTBiGSys + 'Login/Logout')
    
    // #endregion

    // #region SAS

    cy.clearCookies();
    cy.clearLocalStorage();

    cy.visit(linkTBiGSys + 'Login')

    cy.wait(5000)

    cy.get('#tbxUserID').type(userSAS)
    cy.get('#tbxPassword').type(pass)

    cy.window().then((window) => {
      const rightCode = window.rightCode;  
      cy.log('Right Code:', rightCode);

      cy.get('#captchaInsert').type(rightCode); 
    });

    cy.get('#btnSubmit').click();
    
    cy.wait(2000)

    cy.visit(linkTBiGSys + 'Project/ProjectActivity/ReadyToPay')

    cy.wait(2000)
    
    cy.get('#txtSONumber').type(sonumb); 
    
    cy.get('#btnFilterGrid').first().click();
    
    cy.wait(2000)

    cy.get('table.tblSummaryData tbody tr td .btnView').first().click({force: true});

    cy.wait(2000)

    cy.get('#slBulkyValidationSitacDP').then(($select) => {
      cy.wrap($select).select('1', { force: true }) 
    })

    cy.get('#slBulkyValidationLandOwnerDP').then(($select) => {
      cy.wrap($select).select('1', { force: true }) 
    })

    cy.get('#slBulkyValidationLandDP').then(($select) => {
      cy.wrap($select).select('1', { force: true }) 
    })

    cy.wait(2000)
    
    cy.get('#btnValidateDP').click();

    cy.wait(2000)

    cy.get('.sa-confirm-button-container button.confirm').click();

    cy.wait(5000)

    cy.visit(linkTBiGSys + 'Login/Logout')
    
    // #endregion

    // #region SAS Section

    cy.clearCookies();
    cy.clearLocalStorage();

    cy.visit(linkTBiGSys + 'Login')

    cy.wait(5000)

    cy.get('#tbxUserID').type(userSASSection)
    cy.get('#tbxPassword').type(pass)

    cy.window().then((window) => {
      const rightCode = window.rightCode;  
      cy.log('Right Code:', rightCode);

      cy.get('#captchaInsert').type(rightCode); 
    });

    cy.get('#btnSubmit').click();
    
    cy.wait(2000)

    cy.visit(linkTBiGSys + 'Project/ProjectActivity/ReadyToPay')

    cy.wait(2000)
    
    cy.get('#txtSONumber').type(sonumb); 
    
    cy.get('#btnFilterGrid').first().click();
    
    cy.wait(2000)

    cy.get('table.tblSummaryData tbody tr td .btnView').first().click({force: true});

    cy.wait(2000)

    cy.get('#slBulkyValidationSitacDP').then(($select) => {
      cy.wrap($select).select('1', { force: true }) 
    })

    cy.get('#slBulkyValidationLandOwnerDP').then(($select) => {
      cy.wrap($select).select('1', { force: true }) 
    })

    cy.get('#slBulkyValidationLandDP').then(($select) => {
      cy.wrap($select).select('1', { force: true }) 
    })

    cy.wait(2000)
    
    cy.get('#btnApproveDP').click();

    cy.wait(2000)

    cy.get('.sa-confirm-button-container button.confirm').click();

    cy.wait(5000)

    cy.visit(linkTBiGSys + 'Login/Logout')

    // #endregion

  })
})
