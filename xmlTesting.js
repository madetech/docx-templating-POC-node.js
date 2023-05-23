const { DOMParser, XMLSerializer } = require('xmldom');
const XMLReader = require('xml-reader');
const XMLQuery = require('xml-query');
const fs = require("fs");


const xmlString = `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Dear Name,</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>From, Sender</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
`;

// converts XML string into a DOM object
const document = XMLReader.parseSync(xmlString)

// produces an array of w:t tags found
const wtTags = XMLQuery(document).find('w:t')


// change the value of the w:t tags
for(const wtTag of wtTags.ast){
    if(Array.isArray(wtTag.children) && wtTag.children.length > 0) {
        const child = wtTag.children[0];
        if(child.type === 'text') {
            child.value = 'succesfully changed'
        }
    }
}

const xmlSerializer = new XMLSerializer();
const str = xmlSerializer.serializeToString(document);
console.log(str)
// const modifiedXmlString = XMLReader.parseSync(parsedXML.toString())
// console.log(modifiedXmlString)
// console.log(document.children[0].children[0].children[0].children[0].children)




