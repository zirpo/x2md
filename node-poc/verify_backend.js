const converter = require('./services/converter');
const fs = require('fs');
const path = require('path');

async function verify() {
    try {
        const csvPath = path.join(__dirname, 'test.csv');
        const buffer = fs.readFileSync(csvPath);
        const markdown = await converter.convert(buffer, 'text/csv', 'test.csv');

        console.log('--- CSV Conversion Result ---');
        console.log(markdown);

        if (markdown.includes('| Name | Age | City |') && markdown.includes('| Alice | 30 | New York |')) {
            console.log('SUCCESS: CSV conversion verified.');
        } else {
            console.error('FAILURE: CSV conversion output incorrect.');
            process.exit(1);
        }

    } catch (error) {
        console.error('Verification failed:', error);
        process.exit(1);
    }
}

verify();
