var PizZip = require('pizzip');
var Docxtemplater = require('docxtemplater');
var fs = require('fs');
var path = require('path');
const convert = require('xml-js');

const xmlFile = fs.readFileSync('DummyData.xml', 'utf8');
const jsonData = JSON.parse(convert.xml2json(xmlFile, {compact: true, spaces: 2}));
var content = fs
    .readFileSync(path.resolve(__dirname, 'Template.docx'), 'binary');

var zip = new PizZip(content);
var doc;
try {
    doc = new Docxtemplater(zip);
} catch(error) {
    // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
    errorHandler(error);
}

//set the templateVariables
doc.setData({
    vulnerabilityname: jsonData.vulnerability.name._text.toUpperCase(),
    severitylevel: jsonData.vulnerability.severitylevel._text.toUpperCase(),
    urlaffected: jsonData.vulnerability.URLaffected._text,
    analysis:jsonData.vulnerability.analysis._text,
    impact: jsonData.vulnerability.impact._text,
    recommendation: jsonData.vulnerability.recommendation._text,
    references: jsonData.vulnerability.references._text

});

try {
    // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
    doc.render()
}
catch (error) {
    // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
    errorHandler(error);
}

var buf = doc.getZip()
             .generate({type: 'nodebuffer'});

// buf is a nodejs buffer, you can either write it to a file or do anything else with it.
fs.writeFileSync(path.resolve(__dirname, 'Output Report.docx'), buf);
console.log("Report Generated Successfully!")