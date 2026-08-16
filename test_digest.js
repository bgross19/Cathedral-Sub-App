const fs = require('fs');

function testSetDefaultDate() {
    let html = fs.readFileSync('Index.html', 'utf8');
    if (!html.includes('id="principalsDigestDate"')) {
        console.error("Input missing");
    }
}
testSetDefaultDate();
