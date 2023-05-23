const fs = require("fs");
const JSZip = require("jszip");
const PizZip = require("pizzip");

// Read the DOCX file
const content = fs.readFileSync("./ExampleDOCX/text-field-example.docx", "binary");

// Load the content into a Zip object
const zip = new PizZip(content);

// Extract the XML file from the Zip object
const xmlContent = zip.file("word/document.xml").asText();
fs.writeFileSync('./Output/TEXT-FIELDS-BEFORE.xml', xmlContent);

const nameSearch = '<w:t>Name</w:t>';
const nameReplace = '<w:t>Mr John Smith</w:t>';
const senderSearch = '<w:t>Sender</w:t>'
const senderReplace = '<w:t>DVLA</w:t>';

const updatedXmlContentWithName = xmlContent.replace(nameSearch, nameReplace);
const updatedXmlContentWithSender = updatedXmlContentWithName.replace(senderSearch, senderReplace);

fs.writeFileSync('./Output/TEXT-FIELDS-AFTER.xml', updatedXmlContentWithSender);

// Convert the modified XML back to a buffer
const updatedXmlBuffer = Buffer.from(updatedXmlContentWithSender, "utf-8");

// Update the Zip object with the modified XML
zip.file("word/document.xml", updatedXmlBuffer);

// Generate the updated DOCX file
const updatedContent = zip.generate({ type: "nodebuffer" });

// Save the updated DOCX file
fs.writeFileSync("./Output/REPLACED_TEXT_FIELDS.docx", updatedContent);
