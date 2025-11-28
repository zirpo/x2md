const mammoth = require('mammoth');
const pdf = require('pdf-parse');
const xlsx = require('xlsx');
const { parse } = require('csv-parse/sync');

class ConversionService {
    async convert(fileBuffer, mimetype, originalName) {
        const extension = originalName.split('.').pop().toLowerCase();

        switch (extension) {
            case 'docx':
                return this.convertDocx(fileBuffer);
            case 'pdf':
                return this.convertPdf(fileBuffer);
            case 'xlsx':
                return this.convertXlsx(fileBuffer);
            case 'csv':
                return this.convertCsv(fileBuffer);
            default:
                throw new Error(`Unsupported file type: ${extension}`);
        }
    }

    async convertDocx(buffer) {
        try {
            const result = await mammoth.extractRawText({ buffer: buffer });
            return result.value; // The raw text
        } catch (error) {
            throw new Error(`DOCX conversion failed: ${error.message}`);
        }
    }

    async convertPdf(buffer) {
        try {
            const data = await pdf(buffer);
            return data.text;
        } catch (error) {
            throw new Error(`PDF conversion failed: ${error.message}`);
        }
    }

    async convertXlsx(buffer) {
        try {
            const workbook = xlsx.read(buffer, { type: 'buffer' });
            let markdown = '';

            workbook.SheetNames.forEach(sheetName => {
                const sheet = workbook.Sheets[sheetName];
                const json = xlsx.utils.sheet_to_json(sheet, { header: 1 });
                
                if (json.length > 0) {
                    markdown += `## Sheet: ${sheetName}\n\n`;
                    markdown += this.jsonToMarkdownTable(json);
                    markdown += '\n\n';
                }
            });

            return markdown;
        } catch (error) {
            throw new Error(`XLSX conversion failed: ${error.message}`);
        }
    }

    async convertCsv(buffer) {
        try {
            const records = parse(buffer, {
                skip_empty_lines: true
            });
            return this.jsonToMarkdownTable(records);
        } catch (error) {
            throw new Error(`CSV conversion failed: ${error.message}`);
        }
    }

    jsonToMarkdownTable(rows) {
        if (!rows || rows.length === 0) return '';

        const header = rows[0];
        const body = rows.slice(1);

        let markdown = '| ' + header.join(' | ') + ' |\n';
        markdown += '| ' + header.map(() => '---').join(' | ') + ' |\n';

        body.forEach(row => {
            // Ensure row has same number of columns as header
            const paddedRow = [...row];
            while (paddedRow.length < header.length) {
                paddedRow.push('');
            }
            markdown += '| ' + paddedRow.join(' | ') + ' |\n';
        });

        return markdown;
    }
}

module.exports = new ConversionService();
