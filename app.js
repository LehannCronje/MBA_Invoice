var JSZip = require('jszip');
var fs = require('fs');
var path = require('path');
var invoice1 = require('./modules/invoice1');
var invoice2 = require('./modules/invoice2');
var invoice3 = require('./modules/invoice3');
var zip = require('express-zip');
const bodyParser = require('body-parser');
var mammoth = require('mammoth');


const docx = require("@nativedocuments/docx-wasm");

var express = require('express');
var exphbs = require('express-handlebars');

const app = express();

app.use('/static', express.static(__dirname + '/public'));

app.use(bodyParser.urlencoded({ extended: true}));
app.engine('hbs',exphbs({defaultLayout:'main',extname:'.hbs'}));
app.set('view engine', 'hbs');

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
var optionsData = [
    'scuderia',
    'nda',
    'test',
    'hello'
];
var goodsData = [
    'Apple carplay',
    'Adaptive front lights',
    'Leather Upholst',
]
app.get('/',(req,res) => {
    res.render('home', {options: optionsData, goodsD : goodsData});
    
})

    app.post('/generateInvoice',(req,res) => {
        var data = req.body.data;
        var goodsDataSet = req.body.goodsD;
        var goods1 = [];
        for(var i=0;i<data.length;i++){
            goods1.push([goodsDataSet[i],goodsDataSet[i],data[i]]);
            // goods1[i][0] = goodsDataSet[i];
            // goods1[i][1] = goodsDataSet[i];
            // goods1[i][2] = data[i];
        }
        if(req.body.client = "scuderia"){
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
        }
        invoice1.invoice1(goods1,sup,cust);
        invoice2.invoice2(412154.40,sup,ba);
        invoice3.invoice3(412154.40,ba,cust);
        setTimeout(function(){
            convertHelper("./invoices/invoice1.docx", "exportPDF").then((arrayBuffer) => {
                fs.writeFileSync("./public/pdf/invoice1.pdf", new Uint8Array(arrayBuffer));
            }).catch((e) => {
                console.error(e);
            });
            convertHelper("./invoices/invoice2.docx", "exportPDF").then((arrayBuffer) => {
                fs.writeFileSync("./public/pdf/invoice2.pdf", new Uint8Array(arrayBuffer));
            }).catch((e) => {
                console.error(e);
            });
            convertHelper("./invoices/invoice3.docx", "exportPDF").then((arrayBuffer) => {
                fs.writeFileSync("./public/pdf/invoice3.pdf", new Uint8Array(arrayBuffer));
            }).then(res.redirect('/')).catch((e) => {
                console.error(e);
            });
        }, 1000);
    });

    app.get('/generatePDF',(req,res) => {
        convertHelper("./invoices/invoice1.docx", "exportPDF").then((arrayBuffer) => {
            fs.writeFileSync("./public/pdf/invoice1.pdf", new Uint8Array(arrayBuffer));
        }).catch((e) => {
            console.error(e);
        });
        convertHelper("./invoices/invoice2.docx", "exportPDF").then((arrayBuffer) => {
            fs.writeFileSync("./public/pdf/invoice2.pdf", new Uint8Array(arrayBuffer));
        }).catch((e) => {
            console.error(e);
        });
        convertHelper("./invoices/invoice3.docx", "exportPDF").then((arrayBuffer) => {
            fs.writeFileSync("./public/pdf/invoice3.pdf", new Uint8Array(arrayBuffer));
        }).catch((e) => {
            console.error(e);
        });
        
    });
    app.get('/preview',(req,res) => {
        
    });

    app.get('/download',(req,res)=> {
        res.zip([
            {path: './pdf/invoice1.pdf', name: 'invoice1.pdf'},
            {path: './pdf/invoice2.pdf', name: 'invoice2.pdf'},
            {path: './pdf/invoice3.pdf', name: 'invoice3.pdf'}
        ])
    });
// buf is a nodejs buffer, you can either write it to a file or do anything else with it.
// ND_DEV_ID: "42MTM7UUPD62SFQI4EE1NKQL0B",
// ND_DEV_SECRET: "145NKAITU88NVDMT75A3E4B4RP"

app.listen(3000, () => {
    console.log('App running on port 3000!')
});