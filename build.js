const fs = require('fs');
const path = require('path');
const Handlebars = require('handlebars');

// Register helper for adding index
Handlebars.registerHelper('add', function(a, b) {
  return a + b;
});

// Read the template file
const templatePath = path.join(__dirname, 'src', 'templates', 'newsletter.hbs');
const templateContent = fs.readFileSync(templatePath, 'utf8');

// Compile the template
const template = Handlebars.compile(templateContent);

// Read the data file
const dataPath = path.join(__dirname, 'src', 'data', 'newsletter.json');
const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));

// Generate the HTML
const html = template(data);

// Ensure build directory exists
const buildDir = path.join(__dirname, 'build');
if (!fs.existsSync(buildDir)) {
  fs.mkdirSync(buildDir);
}

// Write the output file
const outputPath = path.join(buildDir, 'newsletter.html');
fs.writeFileSync(outputPath, html);

console.log('Newsletter generated successfully!');