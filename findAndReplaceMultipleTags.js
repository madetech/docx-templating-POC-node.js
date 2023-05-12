const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const fs = require("fs");
const path = require("path");

const jsonData = fs.readFileSync('./data.json', 'utf8');

// Parse the JSON file contents into a JavaScript object
const customerDataObject = JSON.parse(jsonData);

const customerData = customerDataObject.data


customerData.forEach((customer) => {

    // Load the docx file as binary content
    const content = fs.readFileSync(
        path.resolve(__dirname, "tag-example.docx"),
        "binary"
    );
    
    const zip = new PizZip(content);
    
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });
    doc.render(customer);

    const buf = doc.getZip().generate({
        type: "nodebuffer",
        compression: "DEFLATE",
    });

    const outputFileName = `test_doc_10_${customer.name}.docx`; 
    
    fs.writeFileSync(path.resolve(__dirname, outputFileName), buf);
    
})
