const convert = require('xml-js');
const fs = require('fs');
const officegen = require('officegen')
// read file
const xmlFile = fs.readFileSync('DummyData.xml', 'utf8');

// parse xml file as a json object
const jsonData = JSON.parse(convert.xml2json(xmlFile, {compact: true, spaces: 2}));
let docx = officegen('docx')
docx.on('finalize', function(written) {
    console.log(
      'Document Created Successfully'
    )
  })
  docx.on('error', function(err) {
    console.log(err)
  })
  let pObj = docx.createP({align:"center"})
  pObj.addText("Penetration Testing Report",{bold: true,font_size:14})
  pObj = docx.createP ({ backline: "000000" });  //Horizontal Line but not as expected
  pObj = docx.createP()
  pObj.addText ( '2 TECHNICAL REPORT', { font_face: 'Cambria', font_size: 16, color: "#2D4061" } );
  pObj.addLineBreak()
  pObj.addLineBreak()
  pObj.addText ( '2.1 GET IN TOUCH URL SECURITY TESTING', { font_face: 'Cambria', font_size: 14, color: "#2D4061" } );
  pObj.addLineBreak()
  pObj.addLineBreak()
  pObj.addText ( '2.1.1 '+jsonData.vulnerability.name._text.toUpperCase(), { font_face: 'Cambria', font_size: 14, color: "#2D4061" } );
  pObj.addLineBreak()
  pObj = docx.createP ({ backline: '80CEF7' });
    pObj.addText ( 'SEVERITY LEVEL', {font_face: "Cambria",bold:true})
    pObj = docx.createP()
    pObj.addText ( jsonData.vulnerability.severitylevel._text, { font_face: 'Cambria', font_size: 12, color: "#F8A854" , bold:true} );
    
    pObj = docx.createP ({ backline: '80CEF7' });
    pObj.addText ( 'EASE OF EXPLOITATION', {font_face: "Cambria",bold:true})
    pObj = docx.createP()
    pObj.addText ( "Medium", { font_face: 'Cambria', font_size: 12, color: "3E9638" , bold:true} );
    
    pObj = docx.createP ({ backline: '80CEF7' });
    pObj.addText ( 'IP ADDRESS/ URL  AFFECTED', {font_face: "Cambria",bold:true})
    pObj = docx.createP()
    pObj.addText ( jsonData.vulnerability.URLaffected._text, { font_face: 'Cambria', font_size: 12,} );
    
    pObj = docx.createP ({ backline: '80CEF7' });
    pObj.addText ( 'ANALYSIS', {font_face: "Cambria",bold:true})
    pObj = docx.createP()
    pObj.addText ( jsonData.vulnerability.analysis._text, { font_face: 'Cambria', font_size: 12,} );
    
    pObj = docx.createP ({ backline: '80CEF7' });
    pObj.addText ( 'IMPACT', {font_face: "Cambria",bold:true})
    pObj = docx.createP()
    pObj.addText ( jsonData.vulnerability.impact._text, { font_face: 'Cambria', font_size: 12,} );
    
    pObj = docx.createP ({ backline: '80CEF7' });
    pObj.addText ( 'REFERNCES', {font_face: "Cambria",bold:true})
    pObj = docx.createP()
    pObj.addText ( jsonData.vulnerability.references._text, { font_face: 'Cambria', font_size: 12,} );
    let out = fs.createWriteStream('Report.docx')
 
out.on('error', function(err) {
  console.log(err)
})
 
// Async call to generate the output file:
docx.generate(out)