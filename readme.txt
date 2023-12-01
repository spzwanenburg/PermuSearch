This tool allows for automating searches in Scopus at PermuSearch.com
A description of its full functionality is pending publication.

The service on PermuSearch is provided by a Shiny app (app.R), leveraging the Scopus API, and using XLSX input and output.
Get a Scopus API key before you use this service, and ensure you can read and write XLSX documents.
Log in to Scopus before clicking on the links in the output.

HOW TO SPECIFY EXCEL INPUT
Specify input in five template sheets:
1 List of search terms in sheet 1(optional)
2 List cross search terms in sheet 2 (optional)
3 Select journals by specifying yes in column 1 of sheet 3 (optional)
4 Select years in sheet 4 (optional)
5 Select matrix dimensions X and Y in sheet 5
6 Define applicable constraints (e.g. years, journals, article type, search terms if not already defined as X or Y dimension) also in sheet 5.
7 Save file, and copy its name to your clipboard for insertion into R.

As an alternative to using the online app, you can run the script on your desktop:
This desktop tool consists of three files: key.txt, new template.xlsx, and permusearch.Rmd
To set up
1. Get a Scopus API key and write it into key.txt
2. Get new template.xlsx and specify your input here (see above)
3. Run Permusearch.Rmd after ensuring it refers to your saved and named excel file (assuming same work directory)
4. An Excel file will open after running. Convert text to numbers before visualising the results.
