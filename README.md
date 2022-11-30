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

- It seems that EPPlus reformat some of the existing formulas in the spreadsheet
- Some of the formulas aren't calculated correctly
- Multiplications doesnt seem to work at all (ie: B1*B2)

