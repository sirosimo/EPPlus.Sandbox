# EPPlus.Sandbox
An exploratory project to tinker with EPPLus


Using:

- .Net Framework 4.6
- EPPlus 6.1.1
- Console

## Goal
To insert data in an existing spreadsheet containing formulas developped in Excel, run the calculations, extract the calculated data. 
No formulas to be developped in EPPlus, all the required calcutions are already defined in the spreadsheet.


## Current Issues

- Formulas not calculated properly (for instance: in a cell in Excel, I have =B1+B2. Using EPPlus and Calculate(), I'm getting the value of B2 only...)
- The formulas is being re-written... (for instance, in a cell in Excel, I have =B2-B1. Using EPPlus and Calculate(), also writting the formulas seen by EPPLus, I'm getting B1-B2 and the results are sometime B1-B2 but sometimes B2-B1....)
- Multiplication aren't even calculated at all. (for instance: in a cell in Excel, I have =B1B2. Using EPPlus and Calculate(), I'm always getting 0 and not error are being triggered)*

