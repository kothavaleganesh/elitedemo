const excel = require('node-excel-export');
var fs = require('fs');
var xlsx = require('node-xlsx').default;
var express = require("express");
var json2xls = require("json2xls");
var app = express();
var _ = require('lodash');
app.use(express.static("myApp")); // myApp will be the same folder name.
app.get("/", function (req, res, next) {
    res.redirect("/");
});
app.listen(8080, "0.0.0.0");
console.log("MyProject Server is Listening on port 8080");
// fs.readFile("./xyz.xlsx", function (err, excel) {
//     if (err) {
//         // res.callback(err, null);
//     } else {
//         var jsonExcel = xlsx.parse(excel);
//         var retVal = [];
//         var firstRow = _.slice(jsonExcel[0].data, 0, 1)[0];
//         var excelDataToExport = _.slice(jsonExcel[0].data, 1);
//         var dataObj = [];
//         _.each(excelDataToExport, function (val, key) {
//             dataObj.push({});
//             _.each(val, function (value, key2) {
//                 dataObj[key][firstRow[key2]] = value;
//             });
//         });
//         console.log(dataObj)
//         return dataObj;
//         // fs.unlink(finalPath);
//     }
// });

// node_xj = require("xls-to-json");
// node_xj({
//     input: "./xyz.xlsx", // input xls
//     output: "output.json", // output json
//     sheet: "Invoice", // specific sheetname
//     rowsToSkip: 5 // number of rows to skip at the top of the sheet; defaults to 0
// }, function (err, result) {
//     if (err) {
//         console.error(err);
//     } else {
//         // console.log(result);
//     }
// });

const readXlsxFile = require('read-excel-file/node');

// File path.
readXlsxFile('./xyz.xlsx').then((rows) => {
    // console.log(rows);
    var invoiceArray = [];
    for (i = 22; i <= 29; i++) {
        var invoiceData = {};
        if (rows[i][0] !== null) {
            invoiceData.invNo = rows[6][2];
            invoiceData.date = rows[7][2];
            invoiceData.company = rows[11][0];
            invoiceData.model = rows[i][1];
            invoiceData.quantity = rows[i][5];
            invoiceData.rate = rows[i][7];
            invoiceArray.push(invoiceData);
        }
    }

    console.log(invoiceArray);
    fs.readFile("./xyz1.xlsx", function (err, excel) {
        if (err) {
            // res.callback(err, null);
        } else {
            var jsonExcel = xlsx.parse(excel);
            console.log(jsonExcel);
            var firstRow = _.slice(jsonExcel[0].data, 0, 1)[0];
            var excelDataToExport = _.slice(jsonExcel[0].data, 1);
            var dataObj = [];
            _.each(excelDataToExport, function (val, key) {
                dataObj.push({});
                _.each(val, function (value, key2) {
                    dataObj[key][firstRow[key2]] = value;
                });
            });
            console.log(dataObj)
            found = _.concat(dataObj, invoiceArray);
            var excelData = [];
            _.each(found, function (singleData) {
                var singleExcel = {};
                _.each(singleData, function (n, key) {
                    singleExcel[key] = n;
                });
                excelData.push(singleExcel);
            });
            var xls = json2xls(excelData);
            var folder = "./";
            var path = "xyz1.xlsx";
            var finalPath = folder + path;
            fs.writeFile(finalPath, xls, "binary", function (err) {
                if (err) {
                    // res.callback(err, null);
                } else {
                    // fs.readFile(finalPath, function (err, excel) {
                    //     if (err) {
                    //         res.callback(err, null);
                    //     } else {
                    //         res.set("Content-Type", "application/octet-stream");
                    //         res.set("Content-Disposition", "attachment;filename=" + path);
                    //         res.send(excel);
                    //         fs.unlink(finalPath);
                    //     }
                    // });
                }
            });
            // fs.unlink(finalPath);
        }
    });

    // `rows` is an array of rows
    // each row being an array of cells.
})

// Readable Stream.
// readXlsxFile(fs.createReadStream('./xyz.xlsx')).then((rows) => {
//     console.log(rows);

// })

// You can define styles as json object
const styles = {
    headerDark: {
        fill: {
            fgColor: {
                rgb: 'FF000000'
            }
        },
        font: {
            color: {
                rgb: 'FFFFFFFF'
            },
            sz: 14,
            bold: true,
            underline: true
        }
    },
    cellPink: {
        fill: {
            fgColor: {
                rgb: 'FFFFCCFF'
            }
        }
    },
    cellGreen: {
        fill: {
            fgColor: {
                rgb: 'FF00FF00'
            }
        }
    }
};

//Array of objects representing heading rows (very top)
const heading = [
    [{
        value: 'a1',
        style: styles.headerDark
    }, {
        value: 'b1',
        style: styles.headerDark
    }, {
        value: 'c1',
        style: styles.headerDark
    }],
    ['a2', 'b2', 'c2'] // <-- It can be only values
];

//Here you specify the export structure
const specification = {
    customer_name: { // <- the key should match the actual data key
        displayName: 'Customer', // <- Here you specify the column header
        headerStyle: styles.headerDark, // <- Header style
        cellStyle: function (value, row) { // <- style renderer function
            // if the status is 1 then color in green else color in red
            // Notice how we use another cell value to style the current one
            return (row.status_id == 1) ? styles.cellGreen : {
                fill: {
                    fgColor: {
                        rgb: 'FFFF0000'
                    }
                }
            }; // <- Inline cell style is possible 
        },
        width: 120 // <- width in pixels
    },
    status_id: {
        displayName: 'Status',
        headerStyle: styles.headerDark,
        cellFormat: function (value, row) { // <- Renderer function, you can access also any row.property
            return (value == 1) ? 'Active' : 'Inactive';
        },
        width: '10' // <- width in chars (when the number is passed as string)
    },
    note: {
        displayName: 'Description',
        headerStyle: styles.headerDark,
        cellStyle: styles.cellPink, // <- Cell style
        width: 220 // <- width in pixels
    }
}

// The data set should have the following shape (Array of Objects)
// The order of the keys is irrelevant, it is also irrelevant if the
// dataset contains more fields as the report is build based on the
// specification provided above. But you should have all the fields
// that are listed in the report specification
const dataset = [{
        customer_name: 'IBM',
        status_id: 1,
        note: 'some note',
        misc: 'not shown'
    },
    {
        customer_name: 'HP',
        status_id: 0,
        note: 'some note'
    },
    {
        customer_name: 'MS',
        status_id: 0,
        note: 'some note',
        misc: 'not shown'
    }
]

// Define an array of merges. 1-1 = A:1
// The merges are independent of the data.
// A merge will overwrite all data _not_ in the top-left cell.
const merges = [{
        start: {
            row: 1,
            column: 1
        },
        end: {
            row: 1,
            column: 10
        }
    },
    {
        start: {
            row: 2,
            column: 1
        },
        end: {
            row: 2,
            column: 5
        }
    },
    {
        start: {
            row: 2,
            column: 6
        },
        end: {
            row: 2,
            column: 10
        }
    }
]

// Create the excel report.
// This function will return Buffer
const report = excel.buildExport(
    [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
        {
            name: 'Report', // <- Specify sheet name (optional)
            heading: heading, // <- Raw heading array (optional)
            merges: merges, // <- Merge cell ranges
            specification: specification, // <- Report specification
            data: dataset // <-- Report data
        }
    ]
);

// You can then return this straight
// res.attachment('report.xlsx'); // This is sails.js specific (in general you need to set headers)
// return res.send(report);

var folder = "./";
var path = "invoice.xlsx";
var finalPath = folder + path;
fs.writeFile(finalPath, report, "binary", function (err) {
    if (err) {
        // res.callback(err, null);
    } else {
        console.log("done");
        // fs.readFile(finalPath, function (err, excel) {
        //     if (err) {
        //         // res.callback(err, null);
        //     } else {
        //         // fs.unlink(finalPath);
        //     }
        // });
    }
});