# xlsx-cvs-to-tex-converter
Simple excel table to tex table converter

## Important
**By now it is csv only!**

## Short description:

### Modes
- "Treat empty rows as new table sign" - empty rows becomes trigger to end current table and start a new one. The number of columns is calculated for each table separately;
- "Remove table borders" - no table border;
- "Use latex math detection" - words containing special characters and numbers are enclosed in "$" characters. List of special characters - `latexMathList`;
- "Table caption at bottom" - caption appears under corresponding table if checked, otherwise above the table;
- separator selection is available when selecting a [csv file](https://en.wikipedia.org/wiki/Comma-separated_values)

Latex special characters are shielded by "\". List of these characters - `latexEscapingCharacter`

By default [csv quoting character](https://en.wikipedia.org/wiki/Comma-separated_values#History) is "**"**". It can only be configured through the source code.

Result of  processing saves to the CONVERTER_RESILT.texCONVERTER_RESILT.tex file in the same place where the application or script lies

### To-do
- Add allowing of empty cell in table
- Add multicolumn processing
- Implement xlsx processing