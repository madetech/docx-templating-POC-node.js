const fs = require("fs");
const JSZip = require("jszip");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

// Read the DOCX file
const content = fs.readFileSync("templateletter.docx", "binary");

// Load the content into a Zip object
const zip = new PizZip(content);

// Extract the XML file from the Zip object
const xmlContent = zip.file("word/document.xml").asText();

const searchText = '<w:t>TEST</w:t>';
const replaceText = '<w:t>IT WORKED!</w:t>';

const updatedXmlContent = xmlContent.replace(searchText, replaceText);

fs.writeFileSync('FIND_AND_REPLACE.xml', updatedXmlContent);

// Convert the modified XML back to a buffer
const updatedXmlBuffer = Buffer.from(updatedXmlContent, "utf-8");

// Update the Zip object with the modified XML
zip.file("word/document.xml", updatedXmlBuffer);

// Generate the updated DOCX file
const updatedContent = zip.generate({ type: "nodebuffer" });

// Save the updated DOCX file
fs.writeFileSync("FIND_AND_REPLACE.docx", updatedContent);
