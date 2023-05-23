const DomParser = require('dom-parser');
const fs = require("fs");
const { type } = require('os');
const PizZip = require("pizzip");
const xml2js = require('xml2js');

// import xmlData from './NO-TEXT-FIELDS-EXAMPLE.xml'
// const xmlData = require('./NO-TEXT-FIELDS-EXAMPLE.xml')

// const xmlData = fs.readFileSync('./NO-TEXT-FIELDS-EXAMPLE.xml')
// console.log(xmlData[2])

const content = fs.readFileSync("no-text-fields.docx", "binary");
const zip = new PizZip(content);
const xmlContentAsString = zip.file("word/document.xml").asText();



const parser = new DomParser()


const parsedXML = parser.parseFromString(xmlContentAsString)

console.log(parsedXML);