const fs = require('fs');
const vm = require('vm');
const path = require('path');

const html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
const match = html.match(/<script>(.*?)<\/script>/s);

if (match) {
    try {
        new vm.Script(match[1]);
        console.log("Syntax is OK!");
    } catch (e) {
        console.error("Syntax Error in inline script:");
        console.error(e);
    }
} else {
    console.log("No script tag found.");
}
