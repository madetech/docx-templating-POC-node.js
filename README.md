## Introduction

Code to provide proof of concept for parsing MS Word Templates and automating letter production.

There are currently 2 files providing different functionality: `findAndReplaceMultipleTags.js` and `findAndReplaceTextFields.js`

 - `findAndReplaceMultipleTags.js` parses a docx  document containing tags (e.g {example}), locates instances of '{example}' and replaces the tags with corresponding data from json file containing dummy data (`data.json)`. An example of the template document is `tag-example.docx`
 - 
 - `findAndReplaceTextFields.js` converts a docx document to xml, finds and replaces instances of a given xml tag (e.g. `<w:t>example</w:t>` before converting the xml file back to docx.

### Next Steps:

The letter templates for the DVLA currently use textFields to highlight the text that is to be populated with customer information. `findAndReplaceTextFields.js` shows that it's possible to target specific xml tags and amend the content, so the next steps are to pass in a json object containing customer data, update the default textFields and produce multiple docx files containing different customer's information.

## Set up

Run: `npm install` to install the required dependencies

In the console, run `node <file to be tested>`. You should the output files in the the project folder.
