var JSZip = require('jszip');

var fs = require('fs');
var path = require('path');
var invoice1 = require('./modules/invoice1');
var invoice2 = require('./modules/invoice2');
var invoice3 = require('./modules/invoice3');

const docx = require("@nativedocuments/docx-wasm");

docx.init({
    ND_DEV_ID: "42MTM7UUPD62SFQI4EE1NKQL0B",    // goto https://developers.nativedocuments.com/ to get a dev-id/dev-secret
    ND_DEV_SECRET: "145NKAITU88NVDMT75A3E4B4RP", // you can also set the credentials in the enviroment variables
    ENVIRONMENT: "NODE", // required
    LAZY_INIT: true      // if set to false the WASM engine will be initialized right now, usefull pre-caching (like e.g. for AWS lambda)
}).catch( function(e) {
    console.error(e);
});

async function convertHelper(document, exportFct) {
    const api = await docx.engine();
    await api.load(document);
    const arrayBuffer = await api[exportFct]();
    await api.close();
    return arrayBuffer;
}
    var goods = [
        ['Apple carplay','ACPL', 28803.95],
        ['Adaptive front lights','ADS2', 23043.50],
        ['Leather Upholst','BAC1', 10369.15],
    ];
    
    var sup = {
        name: 'Aquila Africa Limited',
        regNo: 'C123101 C1/GBL',
        licenceNo: 'C114013067',
        vatNo: '',
        building:'',
        addressNum:'33',
        addressStreet: 'Edith Street',
        addressCity: 'Port-Louis',
        addressCountry:'Mauritius',
        backAcc: 'AfrAsia Bank',
        corBank: 'Firstrand Bank Ltd',
        contactPerson:'',
        contactNo:'',
        emailAddr: '',
        invoiceNote: 'Exporter note that should be inserted here'
    }
    var cust = {
        name: 'Scuderia South Afica (PTY) LTD',
        regNo: '',
        licenceNo: '00652837',
        vatNo: '',
        building:'',
        addressNum:'1',
        addressStreet: 'Bruton Road',
        addressCity: 'Braynston',
        addressCountry:'South-Africa',
        backAcc: '',
        corBank: '',
        contactPerson:'Scuderia Contact person',
        contactNo:'',
        emailAddr: '',
        invoiceNote: ''
    }
    var ba = {
        name: 'Datasmart hong kong limited',
        regNo: '2337214',
        licenceNo: '',
        vatNo: '',
        building:'11th Floor, Tern Centre',
        addressNum:'251',
        addressStreet: 'Queens Road Central',
        addressCity: 'Sheung Wan',
        addressCountry:'Hong Kong',
        backAcc: '',
        corBank: '',
        contactPerson:'Sharma Busgeeth',
        contactNo:'+230 5445 1005',
        emailAddr: 'admin@datasmart.hk',
        invoiceNote: ''
    }


    invoice1.invoice1(goods,sup,cust);
    invoice2.invoice2(412154.40,sup,ba);
    invoice3.invoice3(412154.40,ba,cust);
    
    convertHelper("./invoices/invoice1.docx", "exportPDF").then((arrayBuffer) => {
        fs.writeFileSync("./pdf/invoice1.pdf", new Uint8Array(arrayBuffer));
    }).catch((e) => {
        console.error(e);
    });
    convertHelper("./invoices/invoice2.docx", "exportPDF").then((arrayBuffer) => {
        fs.writeFileSync("./pdf/invoice2.pdf", new Uint8Array(arrayBuffer));
    }).catch((e) => {
        console.error(e);
    });
    convertHelper("./invoices/invoice3.docx", "exportPDF").then((arrayBuffer) => {
        fs.writeFileSync("./pdf/invoice3.pdf", new Uint8Array(arrayBuffer));
    }).catch((e) => {
        console.error(e);
    });
// buf is a nodejs buffer, you can either write it to a file or do anything else with it.
// ND_DEV_ID: "42MTM7UUPD62SFQI4EE1NKQL0B",
// ND_DEV_SECRET: "145NKAITU88NVDMT75A3E4B4RP"