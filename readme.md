# Write_Digital_Object_Template
This application takes a spreadsheet listing digital objects and searches ArchivesSpace for matching archival objects, 
then writes data from the spreadsheet and matches from ArchivesSpace to the ArchivesSpace digital object import 
template. The purpose was to find the matching archival object URIs in ArchivesSpace for digital objects to use with
the ArchivesSpace spreadsheet importer feature (as found in v2.8.0 and above), but then expanded to include all the data
needed to pre-populate the spreadsheet template for importing digital objects into an ArchivesSpace resource.

## Digital Object Spreadsheet
The digital object spreadsheet should contain the ID, title, file version URI (digital object URL), date expression, and
publish status (TRUE or FALSE). These should exist in the following columns with these headers:
1. Column 0: digital_object_id
2. Column 2: digital_object_title
3. Column 3: file_version_file_uri
4. Column 5: date_1_expression
5. Column 8 digital_object_publish

In order to find the matching archival objects in ArchivesSpace, the digital object title and date have to match the 
archival object title and date. The application does an exact search for "title, date". If no matches are found, it 
returns NONE and highlights the row in the digital object template spreadsheet with an ERROR message in the archival 
object URI and resource URI columns.

## ArchivesSpace Digital Object Template Spreadsheet
Download the ArchivesSpace digital object template spreadsheet here: 
https://github.com/archivesspace/archivesspace/blob/master/templates/bulk_import_DO_template.xlsx
