const fs = require('fs');
const XLSX = require('xlsx');

/**
 * 
 *
 * 
 * Config Area starts
 * 
 * 
 * 
 */

const originalName = 'template.xlsx';
const sheetName = 'Sheet1';

// Upload products without ...
const uploadNoProdImg = 0;
const uploadNoDescImg = 1;
const uploadNoEngDesc = 1;
const uploadNoChnDesc = 1;

/**
 * 
 *
 * 
 * Config Area Ends
 * 
 * 
 * 
 */

// The map of columns between the original and the exported file. {col in exported: col in original}
const map = {
    D: 'A', // Product name, diff
    H: 'J', // Short description, diff
    O: 'K', // Inventory / Stock
    Z: 'M', // Regular price
    AN: 'H', // Reorder# / Barcode
    AO: 'G', // Brand
    AP: 'C', // Sub Dept, diff
    AQ: 'E' // 2nd Sub Dept, diff
};

/**
 * Reformat the excel to the exportable format (column names & positions)
 * Store the formatted excel in a file / js object
 * 
 * Select rows with empty values of prod_img / desc_img / desc_eng / desc_chn
 * Output each result to separate files with sku# and name
 * Delete (splice) the rows with empty values according to the config (config=1)
 * 
 * Output the final file without empty values
 * 
 * Issues:
 * Commas in the value (solution use db-quote, maybe supported by the package itself)
 * Non-ascii characters
 */

const workbook = XLSX.readFile(originalName);
const worksheet = workbook.Sheets[sheetName];

const numOfRows = worksheet['!ref'].match(/\d+$/)[0];

// Final data for export. Initialized by the values of the header. Do not change the structure.
const data = {
    "!ref": 'A1:AQ' + numOfRows,
    A1: { v: 'ID' },
    B1: { v: 'Type' },
    C1: { v: 'SKU' },
    D1: { v: 'Name' },
    E1: { v: 'Published' },
    F1: { v: 'Is featured?' },
    G1: { v: 'Visibility in catalog' },
    H1: { v: 'Short description' },
    I1: { v: 'Description' },
    J1: { v: 'Date sale price starts' },
    K1: { v: 'Date sale price ends' },
    L1: { v: 'Tax status' },
    M1: { v: 'Tax class' },
    N1: { v: 'In stock?' },
    O1: { v: 'Stock' },
    P1: { v: 'Low stock amount' },
    Q1: { v: 'Backorders allowed?' },
    R1: { v: 'Sold individually?' },
    S1: { v: 'Weight (kg)' },
    T1: { v: 'Length (cm)' },
    U1: { v: 'Width (cm)' },
    V1: { v: 'Height (cm)' },
    W1: { v: 'Allow customer reviews?' },
    X1: { v: 'Purchase note' },
    Y1: { v: 'Sale price' },
    Z1: { v: 'Regular price' },
    AA1: { v: 'Categories' },
    AB1: { v: 'Tags' },
    AC1: { v: 'Shipping class' },
    AD1: { v: 'Images' },
    AE1: { v: 'Download limit' },
    AF1: { v: 'Download expiry days' },
    AG1: { v: 'Parent' },
    AH1: { v: 'Grouped products' },
    AI1: { v: 'Upsells' },
    AJ1: { v: 'Cross-sells' },
    AK1: { v: 'External URL' },
    AL1: { v: 'Button text' },
    AM1: { v: 'Position' },
    AN1: { v: 'Meta: reorder_num' },
    AO1: { v: 'Meta: brand_catg' },
    AP1: { v: 'Meta: sub_catg_1' },
    AQ1: { v: 'Meta: sub_catg_2' },
    "!margins": {
        "left": 0.7,
        "right": 0.7,
        "top": 0.75,
        "bottom": 0.75,
        "header": 0.3,
        "footer": 0.3
    }
};

// Map the necessary data from the original to the exported
Object.entries(map).forEach(([exported, original]) => {
    for (let i = 2; i <= numOfRows; i++) {
        data[exported + i] = { v: worksheet[original + i].v };
    }
});

// Set the fixed values. All rows have the same value
const fixed = {
    // A1: { v: 'ID' },
    B: { v: 'simple' },
    C: { v: '' },
    E: { v: '1' },
    F: { v: '0' },
    G: { v: 'visible' },
    // I1: { v: 'Description' },
    J: { v: '' },
    K: { v: '' },
    L: { v: 'taxable' },
    M: { v: '' },
    // N1: { v: 'In stock?' },
    P: { v: '' },
    Q: { v: '0' },
    R: { v: '0' },
    S: { v: '' },
    T: { v: '' },
    U: { v: '' },
    V: { v: '' },
    W: { v: '1' },
    X: { v: '' },
    Y: { v: '' },
    AA: { v: 'Canadian Warehouse' },
    AB: { v: '' },
    AC: { v: '' },
    // AD1: { v: 'Images' },
    AE: { v: '' },
    AF: { v: '' },
    AG: { v: '' },
    AH: { v: '' },
    AI: { v: '' },
    AJ: { v: '' },
    AK: { v: '' },
    AL: { v: '' },
    AM: { v: '0' }
};
Object.entries(fixed).forEach(([col, value]) => {
    for (let i = 2; i <= numOfRows; i++) {
        data[col + i] = value;
    }
});







// Add property 't' to each cell (key)
Object.keys(data).forEach(cell => {
    if (cell !== '!ref' && cell !== '!margin') {
        data[cell].t = 's';
    }
});

XLSX.stream.to_csv(data)
    .pipe(fs.createWriteStream('test.csv'))
    .on('error', err => {
        console.log(err);
    });
// fs.writeFile('xlsx-worksheet.json', JSON.stringify(worksheet), () => {});
