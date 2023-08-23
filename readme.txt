This tool allows for automating searches in Scopus.
A description of its full functionality is pending publication.

INSTRUCTIONS
This tool consists of three files: key.txt, new template.xlsx, and permusearch.Rmd
To set up
1. Get a Scopus API key and write it into key.txt
2. Get new template.xlsx and specify your input here (see below)
3. Run Permusearch.Rmd after ensuring it refers to your saved and named excel file (assuming same work directory)
4. An Excel file will open after running. Convert text to numbers before visualising the results.

SPECIFY EXCEL INPUT
Specify input in five template sheets:
2.1 List of search terms in sheet 1(optional)
2.2 List cross search terms in sheet 2 (optional)
2.3 Select journals by specifying yes in column 1 of sheet 3 (optional)
2.4 Select years in sheet 4 (optional)
2.5 Select matrix dimensions X and Y in sheet 5
2.6 Define applicable constraints (e.g. years, journals, article type, search terms if not already defined as X or Y dimension) also in sheet 5.
2.7 Save file, and copy its name to your clipboard for insertion into R.
