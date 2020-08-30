const moment = require('moment');
const XLSX = require("xlsx");

module.exports = {
    generateXLSFile: function (data, path) {
        // If agruments are not proper
        if (!(data && path)) {
            console.error("Please provide the arguments.");
            return;
        }

        // Key checkings
        if (!data.hasOwnProperty('nroProforma')) {
            console.error("Please set the nroProforma property.");
            return;
        }

        if (!data.hasOwnProperty('fechaEmision')) {
            console.error("Please set the fechaEmision property.");
            return;
        }

        if (!data.hasOwnProperty('cliente')) {
            console.error("Please set the cliente property.");
            return;
        }

        // create a new workbook
        const workBook = XLSX.utils.book_new();

        // setup a new worksheet
        let worksheet = {};

        // set styles and border
        let headerStyle = { font: { bold: true }, fill: { bgColor: { rgb: "93D057" } } };
        let border = { top: { style: 'thin', color: '000000' }, bottom: { style: 'thin', color: '000000' }, left: { style: 'thin', color: '000000' }, right: { style: 'thin', color: '000000' } };

        // Set headers
        worksheet["A1"] = { v: "Cliente" };
        worksheet["A2"] = { v: "Nro Proforma" };
        worksheet["A3"] = { v: "Fetcha Emision" };
        worksheet["A5"] = { v: "Tela", s: headerStyle };
        worksheet["B5"] = { v: "Color", s: headerStyle };
        worksheet["C5"] = { v: "suma rollos", s: headerStyle };
        worksheet["D5"] = { v: "peso", s: headerStyle };
        worksheet["E5"] = { v: "precio", s: headerStyle };
        worksheet["F5"] = { v: "total", s: headerStyle };

        // Set values
        worksheet["B1"] = { v: data.cliente.razonSocial || "", s: { font: { sz: 14 } } };
        worksheet["B2"] = { v: data.nroProforma || "", t: 's', s: { font: { sz: 14 } } };
        worksheet["B3"] = { v: data.fechaEmision ? moment(data.fechaEmision).format("DD/MM/YYYY") : '', s: { font: { sz: 14 } } };

        // group the data
        let detalleProformasObject = processDetalleProforma(data.detalleProforma);
        // Looped data to start from
        let cellNumber = 6;

        // parse the grouped data
        for (let index in detalleProformasObject) {
            // object
            let detalleProformaDataArray = detalleProformasObject[index];
            let detalleProformaDataLength = detalleProformaDataArray.length;

            // check if it has more than one item
            if (detalleProformaDataLength) {
                // final data to be written on worksheet
                let finalData = {
                    descripcionTela: detalleProformaDataArray[0].descripcionTela,
                    descripcionColor: detalleProformaDataArray[0].descripcionColor,
                    peso: 0,
                    precioVenta: 0,
                    codigoBarra: 0
                };
                let codigoBarra = [];

                // Find the sums
                for (let detalleProformaData of detalleProformaDataArray) {
                    // collect all codigoBarra values
                    if (detalleProformaData.codigoBarra) {
                        codigoBarra.push(detalleProformaData.codigoBarra.toLowerCase());
                    }

                    // Find the sums
                    finalData.peso += Number.parseFloat(detalleProformaData.peso) || 0;
                    finalData.precioVenta += Number.parseFloat(detalleProformaData.precioVenta) || 0;
                }

                // unque codigoBarra
                finalData.codigoBarra = [...new Set(codigoBarra)].length;
                // avg of precioVenta
                finalData.precioVenta = finalData.precioVenta / detalleProformaDataLength;

                // set the data in sheet
                worksheet[`A${cellNumber}`] = { v: finalData.descripcionTela, s: { border: border } };
                worksheet[`B${cellNumber}`] = { v: finalData.descripcionColor, s: { border: border } };
                worksheet[`C${cellNumber}`] = { v: `${finalData.codigoBarra}`, t: 'n', s: { border: border } };
                worksheet[`D${cellNumber}`] = { f: `ROUND(${finalData.peso}, 3)`, s: { border: border } };
                worksheet[`E${cellNumber}`] = { f: `ROUND(${finalData.precioVenta}, 3)`, t: 'n', s: { border: border } };
                worksheet[`F${cellNumber}`] = { f: `ROUND(D${cellNumber}*E${cellNumber}, 3)`, t: 'n', s: { border: border } };
            }
            cellNumber++;
        }

        // set the total rows
        worksheet[`A${cellNumber}`] = { v: "TOTAL GENERAL", s: headerStyle };
        worksheet[`C${cellNumber}`] = { f: `ROUND(SUM(C6:C${cellNumber - 1}), 3)`, t: 'n', s: headerStyle };
        worksheet[`D${cellNumber}`] = { f: `ROUND(SUM(D6:D${cellNumber - 1}), 3)`, t: 'n', s: headerStyle };
        worksheet[`E${cellNumber}`] = { v: ``, s: headerStyle };
        worksheet[`F${cellNumber}`] = { f: `ROUND(SUM(F6:F${cellNumber - 1}), 3)`, t: 'n', s: headerStyle };

        // References
        const range = {
            s: { r: 0, c: 0 },
            e: { r: cellNumber + 10, c: 8 }
        };
        // Row/Cols to merge
        const merge = [
            { s: { r: cellNumber - 1, c: 0 }, e: { r: cellNumber - 1, c: 1 } },
        ];
        // Column width
        const wscols = [
            { wch: 30 },
            { wch: 35 },
            { wch: 15 },
            { wch: 15 },
            { wch: 15 },
            { wch: 15 },
        ];

        worksheet['!cols'] = wscols;
        worksheet["!merges"] = merge;
        worksheet['!ref'] = XLSX.utils.encode_range(range);
        workBook.SheetNames.push('Sheet');
        workBook.Sheets.Sheet = worksheet;
        XLSX.writeFile(workBook, path);
    }
};

function processDetalleProforma(data) {
    let groupedData = {};
    data.forEach(item => {
        if (item.descripcionTela && item.descripcionColor) {
            let key = item.descripcionTela + "-" + item.descripcionColor;
            if (!groupedData.hasOwnProperty(key)) {
                groupedData[key] = [];
            }
            groupedData[key].push(item);
        }
    });

    return groupedData;
}
