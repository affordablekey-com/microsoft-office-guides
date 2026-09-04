# Microsoft Excel Useful Formulas Guide

Microsoft Excel formulas make it easier to calculate, organize, and analyze data.

This guide covers several useful formulas for everyday Excel tasks.

## 1. SUM

The `SUM` function adds numbers together.

Example:

`=SUM(B2:B10)`

This adds all values from cells B2 through B10.

## 2. AVERAGE

The `AVERAGE` function calculates the average of a range of numbers.

Example:

`=AVERAGE(B2:B10)`

This calculates the average value in the selected range.

## 3. COUNT

The `COUNT` function counts cells that contain numbers.

Example:

`=COUNT(B2:B10)`

This is useful when you want to know how many numeric values are in a range.

## 4. COUNTA

The `COUNTA` function counts cells that are not empty.

Example:

`=COUNTA(A2:A10)`

It can count text, numbers, and other values.

## 5. MAX

The `MAX` function returns the largest value in a range.

Example:

`=MAX(B2:B10)`

## 6. MIN

The `MIN` function returns the smallest value in a range.

Example:

`=MIN(B2:B10)`

## 7. IF

The `IF` function returns one result when a condition is true and another result when it is false.

Example:

`=IF(B2>=50,"Pass","Fail")`

If B2 is 50 or higher, the formula returns "Pass". Otherwise, it returns "Fail".

## 8. SUMIF

The `SUMIF` function adds values that meet a specific condition.

Example:

`=SUMIF(A2:A10,"Sales",B2:B10)`

This adds values in B2:B10 when the corresponding cell in A2:A10 contains "Sales".

## 9. COUNTIF

The `COUNTIF` function counts cells that meet a condition.

Example:

`=COUNTIF(B2:B10,">=50")`

This counts values that are 50 or higher.

## 10. AVERAGEIF

The `AVERAGEIF` function calculates the average of cells that meet a condition.

Example:

`=AVERAGEIF(B2:B10,">=50")`

This calculates the average of values that are 50 or higher.

## 11. ROUND

The `ROUND` function rounds a number to a specified number of decimal places.

Example:

`=ROUND(B2,2)`

This rounds the value in B2 to two decimal places.

## 12. TODAY

The `TODAY` function returns the current date.

Example:

`=TODAY()`

The result updates based on the current date.

## 13. NOW

The `NOW` function returns the current date and time.

Example:

`=NOW()`

The result can update when Excel recalculates the worksheet.

## 14. CONCAT

The `CONCAT` function combines text from multiple cells.

Example:

`=CONCAT(A2," ",B2)`

This can combine a first name and last name stored in separate cells.

## 15. TEXTJOIN

The `TEXTJOIN` function combines text from multiple cells using a delimiter.

Example:

`=TEXTJOIN(", ",TRUE,A2:A5)`

This combines the values in A2:A5 and separates them with commas.

## 16. LEN

The `LEN` function counts the number of characters in a text value.

Example:

`=LEN(A2)`

Spaces are also counted as characters.

## 17. LEFT

The `LEFT` function returns characters from the beginning of a text value.

Example:

`=LEFT(A2,5)`

This returns the first five characters from A2.

## 18. RIGHT

The `RIGHT` function returns characters from the end of a text value.

Example:

`=RIGHT(A2,4)`

This returns the last four characters from A2.

## 19. TRIM

The `TRIM` function removes extra spaces from text.

Example:

`=TRIM(A2)`

This is useful when cleaning imported or copied data.

## 20. XLOOKUP

`XLOOKUP` can find a value in one range and return a related value from another range.

Example:

`=XLOOKUP(E2,A2:A10,B2:B10,"Not Found")`

This searches for the value in E2 within A2:A10 and returns the corresponding value from B2:B10.

Availability of XLOOKUP depends on the Excel version and platform.

## Example Table

Suppose you have this data:

| Employee | Department | Sales |
|---|---|---:|
| Alex | Sales | 5000 |
| Maria | Marketing | 4200 |
| John | Sales | 6100 |
| Emma | Support | 3500 |

You could use:

`=SUM(C2:C5)`

to calculate total sales.

You could use:

`=MAX(C2:C5)`

to find the highest sales value.

You could use:

`=AVERAGE(C2:C5)`

to calculate average sales.

You could use:

`=SUMIF(B2:B5,"Sales",C2:C5)`

to calculate total sales for the Sales department.

## Formula Tips

### Start With an Equals Sign

Excel formulas normally begin with:

`=`

For example:

`=SUM(A1:A10)`

### Use Cell References

Instead of entering numbers directly into every formula, use cell references whenever possible.

Example:

`=A2+B2`

This makes formulas easier to update.

### Check Parentheses

Complex formulas can produce errors when parentheses are missing or placed incorrectly.

For example:

`=IF(A2>10,"Yes","No")`

Make sure every opening parenthesis has a matching closing parenthesis.

### Use Absolute References When Needed

A dollar sign can keep a row or column reference fixed when copying formulas.

Example:

`=$A$1*B2`

When the formula is copied, A1 remains fixed.

## Common Formula Errors

### #DIV/0!

This usually occurs when a formula attempts to divide by zero or an empty cell.

### #VALUE!

This can occur when a formula receives an unexpected type of value.

### #N/A

This often indicates that a lookup formula could not find the requested value.

### #REF!

This usually indicates that a formula refers to an invalid or deleted cell reference.

### #NAME?

This can occur when Excel does not recognize a function or name used in a formula.

## Quick Reference

| Function | Purpose |
|---|---|
| SUM | Add numbers |
| AVERAGE | Calculate average |
| COUNT | Count numeric cells |
| COUNTA | Count non-empty cells |
| MAX | Find largest value |
| MIN | Find smallest value |
| IF | Test a condition |
| SUMIF | Add values matching a condition |
| COUNTIF | Count values matching a condition |
| AVERAGEIF | Average values matching a condition |
| ROUND | Round numbers |
| TODAY | Return current date |
| NOW | Return current date and time |
| CONCAT | Combine text |
| TEXTJOIN | Combine text with a delimiter |
| LEN | Count characters |
| LEFT | Return characters from the left |
| RIGHT | Return characters from the right |
| TRIM | Remove extra spaces |
| XLOOKUP | Search and return related data |

## Final Tip

You do not need to memorize every Excel formula.

Start with a few frequently used functions such as `SUM`, `AVERAGE`, `IF`, `COUNTIF`, and `XLOOKUP`.

As your spreadsheet tasks become more complex, learn additional functions based on the problem you need to solve.

## Important Note

Formula availability and behavior can vary depending on the Excel version, platform, language settings, and workbook configuration.

Always verify formulas in your specific Excel environment.
