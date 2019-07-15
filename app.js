var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');

var fs = require('fs');
var path = require('path');

//Load the docx file as a binary
var content = fs
    .readFileSync(path.resolve(__dirname, 'DS_invoice.docx'), 'binary');

var zip = new JSZip(content);
var doc = new Docxtemplater();

var goods = [
    ['Apple carplay','ACPL', 28803.95],
    ['Adaptive front lights','ADS2', 23043.50],
    ['Leather Upholst','BAC1', 10369.15],
]
var invoice_items = [];
for(var i=0; i<goods.length;i++){
    invoice_items[i] = [{
            Variable18: goods[i][0],
            Variable19: goods[i][1],
            Variable20:2,
            Variable21: goods[i][2],
            Variable22: goods[i][2] * 2
        },
    ];
}
console.log(invoice_items);
doc.loadZip(zip);
//set the templateVariables
doc.setData({
    items : invoice_items,
    Variable1: '1',
    Variable2: '2',
    Variable3: '3',
    Variable4: '4',
    Variable5: '5',
    Variable6: '6',
    Variable7: '7',
    Variable8: '8',
    Variable9: '9',
    Variable10: '10',
    Variable11: '11',
    Variable12: '12',
    Variable13: '13',
    Variable14: '14',
    Variable15: '15',
    Variable16: '16',
    Variable17: '17',
    // Variable18: '18',
    // Variable19: '19',
    // Variable20: '20',
    // Variable21: '21',
    // Variable22: '22',
    Variable23: '23',
    Variable24: '24',
    Variable25: '25',
    Variable26: '26',
    Variable27: '27',
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

// buf is a nodejs buffer, you can either write it to a file or do anything else with it.
fs.writeFileSync(path.resolve(__dirname, 'invoice-instance.docx'), buf);