// Require library
var excel = require('excel4node');

// Create a new instance of a Workbook class
var workbook = new excel.Workbook();

// Add Worksheets to the workbook
var worksheet = workbook.addWorksheet('Sheet1');

// Create a reusable style
var style = workbook.createStyle({
  font: {
    color: '#000000',
    size: 12
  },
  numberFormat: '$#,##0.00; ($#,##0.00); -'
});

// Set the header of cell A1 to 100 as a string type styled with paramaters of style
worksheet.cell(1, 1).string("Name").style(style).style({font: { bold: true }});
worksheet.cell(1, 2).string("Value 1").style(style).style({font: { bold: true }});
worksheet.cell(1, 3).string("Value 2").style(style).style({font: { bold: true }});
worksheet.cell(1, 4).string("Total").style(style).style({font: { bold: true }});

var myArray = [{
        "Name": "Allex",
        "Value1": 100,
        "Value2": 200
    }, {
        "Name": "Laura",
        "Value1": 100,
        "Value2": 100
    }
];
 
for (var i in myArray) {
    var row = Number(i) + 1;
    var item = myArray[i];
    var formula = "B" + (row + 1) + " + C" + (row + 1);
    
    console.log("Row: " + row + ", Name: " + item.Name);

    // Set value of cell A2 to 'item.Name' as a string type styled with paramaters of style
    // Add + 1 to skip the header
    worksheet.cell(row + 1, 1).string(item.Name).style(style);

    // Set value of cell B2 to 'item.Value1' as a number type styled with paramaters of style
    // Add + 1 to skip the header
    worksheet.cell(row + 1, 2).number(item.Value1).style(style);

    // Set value of cell C2 to 'item.Value2' as a number type styled with paramaters of style
    // Add + 1 to skip the header
    worksheet.cell(row + 1, 3).number(item.Value2).style(style);

    // Set value of cell D2 to a formula styled with paramaters of style
    // Add + 1 to skip the header
    worksheet.cell(row + 1, 4).formula(formula).style(style);
}

workbook.write('Excel.xlsx');