const fs = require('fs');
const XLSX = require('xlsx');
const sizeOf = require('image-size');

/**
 * -------------------------------------------------------------------------------
 *
 * 
 * Config Area starts
 * 
 * 
 * -------------------------------------------------------------------------------
 */

// Must be the year/month of image being uploaded. Check and change every time.
const imgUploadMonth = '2020/12';

// Upload products without ... Check and change every time.
const configs = {
    uploadNoProdImg: 1,
    uploadNoDescImg: 1,
    uploadNoEngDesc: 0,
    uploadNoChnDesc: 1
};
// const uploadNoProdImg = 1;
// const uploadNoDescImg = 1;
// const uploadNoEngDesc = 0;
// const uploadNoChnDesc = 1;

// Config that normally you don't need to change.
const originalName = 'template.xlsx';
const sheetName = 'Sheet1';
const imgDir = './all products/';
const missing = {
    'product images': 'AD',
    'description images': 'I',
    'english description': 'H',
    'chinese description': 'H'
};

/**
 * -------------------------------------------------------------------------------
 *
 * 
 * Config Area Ends
 * 
 * 
 * -------------------------------------------------------------------------------
 */

/**
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 *
 * Do not need to change codes below
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 * 
 */

/**
* @desc add {t: 's'} and margins to the data matrix
* @param {object} data 
* @return {object}
*/
const setCellType = data => {
    Object.keys(data).forEach(cell => {
        if (cell !== '!ref' && cell !== '!margin') {
            data[cell].t = 's';
        }
    });
    data['!margins'] = {
        "left": 0.7,
        "right": 0.7,
        "top": 0.75,
        "bottom": 0.75,
        "header": 0.3,
        "footer": 0.3
    }
    return data;
};

/**
 * @desc Set image html for product & descriptio images. This function will loop through the image folder and check the existance of images named as the Reorder#
 * @param {object} data worksheet
 * @param {string} type 'prodct' | 'decrpt'
 * @param {number} row
 * @return {string} 
 */
const setImgHtml = (data, type, row) => {
    if (type !== 'prodct' && type !== 'decrpt') {
        throw new Error('Must be "prodct" or "decrpt" as the parameter!');
    }

    let img = 1;
    let imgName = `${data[`AN${row}`].v}-${type}_${img}.jpg`;
    let html = '';
    while (fs.existsSync(imgDir + imgName)) {
        const dimensions = sizeOf(imgDir + imgName);

        html += `<img class="alignnone size-full" src="/wp-content/uploads/${imgUploadMonth}/${imgName}" alt="" width="${dimensions.width}" height="${dimensions.height}" style="display: inline-block; margin: 0;" />`;

        img++;
        imgName = `${data[`AN${row}`]}-${type}_${img}.jpg`;
    }

    return html;
};

/**
 * @desc Safely delete rows
 * @param {object} ws worksheet
 * @param {number} row 0-based row index
 * @return {undefined}
 */
const deleteRow = (ws, row) => {
    const range = XLSX.utils.decode_range(ws["!ref"])
    for (let R = row; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            ws[XLSX.utils.encode_cell({ r: R, c: C })] = ws[XLSX.utils.encode_cell({ r: R + 1, c: C })]
        }
    }
    range.e.r--
    ws['!ref'] = XLSX.utils.encode_range(range.s, range.e)
};

/**
 * @desc Find the rows with empty value at certain column (type)
 * @param {string} type Within 'product images', 'description images', 'english description', 'chinese description'
 * @param {object} data worksheet
 * @return {array} rows without certain info (empty cell)
 */
const getRowsMissing = (type, data) => {
    if (!Object.keys(missing).includes(type)) {
        throw new Error('This type was not in the config!');
    }

    const rows = [];
    const oldRows = data['!ref'].match(/\d+$/)[0];
    for (let i = 2; i <= oldRows; i++) {
        if (data[missing[type] + i].v === '') {
            rows.push(i);
        }
    }

    return rows;
}

/**
 * @desc Export rows with missing info to separate excel files
 * @param {string} type Within 'product images', 'description images', 'english description', 'chinese description'
 * @param {array} data worksheet
 * @return {array} rows without certain info (empty cell)
 */
const exportMissing = (type, data) => {
    if (!Object.keys(missing).includes(type)) {
        throw new Error('This type was not in the config!');
    }

    // Final data for export. Initialized by the values of the header. Do not change the structure.
    const dataMissing = {
        A1: { v: 'REORDER#' },
        B1: { v: 'Name' }
    };

    // Find the rows with empty values of type and take the Reorder# and Name to the dataMissing
    const rows = getRowsMissing(type, data);
    let newRows = 2;
    for (let row of rows) {
        dataMissing['A' + newRows] = { v: data['AN' + row].v };
        dataMissing['B' + newRows] = { v: data['D' + row].v };
        newRows++;
    }
    dataMissing['!ref'] = 'A1:B' + newRows;

    // Add property 't' to each cell (key)
    const exportedData = setCellType(dataMissing);

    // Write the new file
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, exportedData, sheetName);
    try {
        XLSX.writeFile(workbook, `./missing/missing ${type}.xlsx`);
        console.log(`Successfully export 'missing ${type}.xlsx' to path: ./missing/`);
    } catch (e) {
        console.error(e);
    } finally {
        return rows;
    }
};

/**
 * @desc Main function to export
 * @param {string} lang 'english' | 'chinese'
 * @param {obejct} wb workbook
 * @return {undefined}
 */
const main = (lang, wb) => {
    if (lang !== 'english' && lang !== 'chinese') {
        throw new Error('Must be "english" or "chinese" as the parameter!');
    }

    const worksheet = wb.Sheets[sheetName];
    const numOfRows = worksheet['!ref'].match(/\d+$/)[0];

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
    if (lang === 'chinese') {
        map.D = 'B';
        map.AP = 'D';
        map.AQ = 'F';
        delete map.H; // Will add this back if our client gives us Chinese description
    }

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
        AQ1: { v: 'Meta: sub_catg_2' }
    };

    // Map the necessary data from the original to the exported
    Object.entries(map).forEach(([exported, original]) => {
        for (let i = 2; i <= numOfRows; i++) {
            data[exported + i] = { v: worksheet[original + i].v };
        }
    });

    // Set the fixed values. All rows have the same value, normally
    const fixed = {
        B: { v: 'simple' },
        C: { v: '' },
        E: { v: '1' },
        F: { v: '0' },
        G: { v: 'visible' },
        J: { v: '' },
        K: { v: '' },
        L: { v: 'taxable' },
        M: { v: '' },
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
        AA: { v: lang === 'english' ? 'Canadian Warehouse' : '加拿大仓' },
        AB: { v: '' },
        AC: { v: '' },
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
    if (lang === 'chinese') {
        fixed.H = { v: '' }; // Like just said, for now no chinese description
    }
    Object.entries(fixed).forEach(([col, value]) => {
        for (let i = 2; i <= numOfRows; i++) {
            data[col + i] = value;
        }
    });

    // Set other changing values
    for (let i = 2; i <= numOfRows; i++) {

        // Self-increasing id like InnoDB
        data['A' + i] = { v: i - 1 };

        // Has inventory
        data['N' + i] = { v: data['O' + i] ? '1' : '0' };

        // Get description images html according to reorder#, if exist
        data['I' + i] = { v: setImgHtml(data, 'decrpt', i) };

        // Get product images html according to reorder#, if exist
        data['AD' + i] = { v: setImgHtml(data, 'prodct', i) };

    }

    // Export rows with missing info according to the config
    const types = Object.keys(missing);
    const currentMissing = lang === 'english' ? types.slice(0, 3) : [types[3]];
    // const currentConfigs = Object.fromEntries(lang === 'english' ? Object.entries(configs).slice())
    // for (let i = 0; i < 3)
    for (let type of currentMissing) {
        exportMissing(type, data);
    }

    // Remove rows from being exported according to the config
    // if (!uploadNoProdImg) {
    //     deleteRow(data, row);
    // }



    // Add property 't' to each cell (key)
    const exportedData = setCellType(data);

    // Write the csv
    try {
        XLSX.stream
            .to_csv(exportedData)
            .pipe(fs.createWriteStream(`exported - ${lang} edition.csv`));
        console.log(`Successfully export 'exported - ${lang} edition.csv' to path: ./`);
    } catch (e) {
        console.error(err);
    }

}

// Must use original file for both export, i.e. doing identical tasks twice, since the data structure is matched strictly with the original file.
const workbook = XLSX.readFile(originalName);
main('english', workbook);
main('chinese', workbook);
