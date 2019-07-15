var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
var fs = require('fs');
var path = require('path');
var dateTime = require('node-datetime');
var dt = dateTime.create().format('Y-m-d');
exports.invoice3 = function(SavPrice,sup,cust){
    var content = fs
        .readFileSync(path.resolve(__dirname, '../Inv_Template/DS_invoice1.docx'), 'binary');
    
    var zip = new JSZip(content);
    var doc = new Docxtemplater();
    
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
        pDescription:'Buying Agent Services',
        pCode: 'N/A',
        pQuan: 'N/A',
        pPrice: 'N/A',
        pTotal: SavPrice,
        ordering: '(Incl. registration and management of Ferari Contact Club membership)',
        invNo: 'INV7007',
        subT: SavPrice,
        vat:'',
        shipHan:'',
        disc:'',
        total:SavPrice,
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
    fs.writeFileSync(path.resolve(__dirname, '../invoices/invoice3.docx'), buf);
}
