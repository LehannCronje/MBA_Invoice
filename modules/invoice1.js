var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
var fs = require('fs');
var path = require('path');
var dateTime = require('node-datetime');
var dt = dateTime.create().format('Y-m-d');
exports.invoice1 = function(goods,sup,cust){
    var content = fs
        .readFileSync(path.resolve(__dirname, '../Inv_Template/DS_invoice.docx'), 'binary');
    
    var zip = new JSZip(content);
    var doc = new Docxtemplater();
    var d = new Date();
    
    
    var itemsSubT = 0;
    var quantity = 1;
    var invoice_items = [];
    for(var i=0; i<goods.length;i++){
        invoice_items[i] = [{
                pDescription: goods[i][0],
                pCode: goods[i][1],
                pQuan:quantity,
                pPrice: goods[i][2],
                pTotal: goods[i][2] * quantity
            },
        ];
        itemsSubT = itemsSubT + (goods[i][2] * quantity);
    }
    
    doc.loadZip(zip);
    //set the templateVariables
    doc.setData({
        sName: sup.name,
        sRegNo: sup.regNo,
        sAddressNum: sup.addressNum,
        sStreet: sup.addressStreet,
        sCity: sup.addressCity,
        invNo:'inv7007',
        date: dt,
        custAccNo: "SA75001Vxx",
        terms:'',
        rName: cust.name,
        rAddressNum:cust.addressNum,
        rStreet:cust.addressStreet,
        rCity: cust.addressCity,
        rCountry: cust.addressCountry,
        cName : cust.contactPerson,
        cNum: cust.contactNo,
        cEmail: cust.emailAddr,
        items: invoice_items,
        subT: itemsSubT,
        vat:'',
        shipHan:'',
        disc:'',
        total:itemsSubT,
        footNote:sup.invoiceNote
    });
    
    try {
        // render the document ie replace the variables
        doc.render()
    }
    catch (error) {
        var e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties,
        }
        console.log(JSON.stringify({error: e}));
        throw error;
    }
    
    var buf = doc.getZip()
                 .generate({type: 'nodebuffer'});
    fs.writeFileSync(path.resolve(__dirname, '../invoices/invoice1.docx'), buf);
}
