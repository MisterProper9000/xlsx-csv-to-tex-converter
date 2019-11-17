# xlsx-cvs-to-tex-converter
Simple excel/csv table to tex table converter

## Short description:

### Modes
- "Treat empty rows as new table sign" - empty rows becomes trigger to end current table and start a new one. The number of columns is calculated for each table separately;
- "Remove table borders" - no table border;
- "Use latex math detection" - words containing special characters and numbers are enclosed in "$" characters. List of special characters - `latexMathList`;
- "Table caption at bottom" - caption appears under corresponding table if checked, otherwise above the table;
- delimiter and [quoting character](https://en.wikipedia.org/wiki/Comma-separated_values#History) selection is available when selecting a [csv file](https://en.wikipedia.org/wiki/Comma-separated_values)

Latex special characters are shielded by "\\". List of these characters - `latexEscapingCharacter`. 

When "Use latex math detection" mode is off, characters from `latexMathList` are also shielded by "\\".

Result of  processing saves to the **CONVERTER_RESILT.tex** file in the same place where the application or script lies

### To-do
- Add merged columns/rows processing
