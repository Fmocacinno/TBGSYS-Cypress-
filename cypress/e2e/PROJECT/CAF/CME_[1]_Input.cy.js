import 'cypress-file-upload';

// # Site for Colocation Additional Form (CAF)
// SELECT * FROM TBiGSys.dbo.vwMKTCAFSiteRequestNonNewColo

// # Validation There's CAF Request still on going for this site! CAF Number:
// SELECT * FROM TBiGSys.dbo.vwMKTCAFSiteRequestNonNewColo A
// INNER JOIN TBiGSys.dbo.trxMKTCAFHeader B ON A.SiteID <> B.SiteID
// WHERE 1 = 1
//       AND B.StatusID NOT IN ( 245, 246, 250, 253, 254, 361 )
//       AND A.CustomerID IN ( 'SMART8', 'SMART', 'M8' )


describe('template spec', () => {
  it('passes', () => {
    const unique = "APP_002"
    const random = 55
    const siteid = "110160110"
    const phone = "0851728319"
    const date = "2-Jan-2025";
    const user = "555504220025"
    const pass = "123456"
    const filePath = 'documents/pdf/receipt.pdf';
    const lat = 5.4
    const long = 94


    const userAssignedDrafter = "201910150170"

    Cypress.on('uncaught:exception', (err, runnable) => {
      return false
    })

    cy.visit('http://tbgappdev111.tbg.local:8044/Login')

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

    cy.visit('http://tbgappdev111.tbg.local:8044/Marketing/CAF/Create')

    cy.wait(2000)

    cy.get('input[name="IsNewColo"][value="1"]').parent().find('.iCheck-helper').click();

    cy.get('#txtSiteID').click();

    cy.wait(5000)

    cy.get('#tbxSearchSiteID').type(siteid);
    // cy.get('#tbxSearchSiteIDNon').type(siteid); 


    cy.get('.btnSearch').first().click({ force: true });
    // cy.get('.btnSearchNon').first().click({ force: true });

    cy.wait(5000)

    cy.get('tbody tr:first-child .btnSelectSiteNew').click();
    // cy.get('tbody tr:first-child .btnSelectSiteNon').click();


    cy.get('#slCustomerID').then(($select) => {
      cy.wrap($select).select('TELKOMSEL', { force: true })
    })

    cy.get('#txtSiteNameCustomer').type('SiteName_' + unique);
    cy.get('#txtSiteIDCustomer').type('SiteID_' + unique);
    // 4 atau 1
    cy.get('#slsGeneralSoW').then(($select) => {
      cy.wrap($select).select('1', { force: true })
    })

    cy.get('#txtContactClient').type('ContactClient_' + unique);
    cy.get('#txtPhoneClient').type(phone);

    cy.get('#txtShelterLength').type(random);
    cy.get('#txtShelterWidth').type(random);
    cy.get('#txtShelterHeight').type(random);
    cy.get('#txtSpaceLength').type(random);
    cy.get('#txtSpaceWidth').type(random);

    cy.get('#btnSubmit').first().click({ force: true });
  })
})
