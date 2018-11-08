Load and extract ms excel document properties metadata

Win 7 64bit Only for XML Excel file ie xlsx (post 2003 Excel).

As Chris Hemmidinger has shown you can parse the excel XML file for
the documents properties.

see github
https://tinyurl.com/y9dksdx8
https://github.com/rogerjdeangelis/utl-load-and-extract-ms-excel-document-properties-metadata

macros
https://tinyurl.com/y9nfugth
https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories

PROGRAM FLOW
 ============

      1. Create a workbook and programtically populate document properties
         Microsoft document properties. (python)
      2. Use SAS to add a sheet to the workbook.
      3. Remove the default sheet1 created by python..
      4. Save the workbook.
      5. Reopen the workbook
      6. Read the document properties (python)
      7. Create SAS table from the document properies

 OUTPUT  (populate document properties)
 ======

  Populate document properties programatically
  --------------------------------------------

  If you right click on the created workbook file in windows explorer and
  click on properties and then details you will see

  +-------------------------------------------+
  | have.xlsx Properties                      |
  +-------------------------------------------+
  | General | Security | Details | Previous   |
  +-------------------------------------------+
  |                                           |
  | Title:    This is an example spreadsheet  |
  | Subject:  With document properties        |
  | Author:   Roger Deangelis                 |
  | Manager:  John Smith                      |
  | Company:  CompuCraft                      |
  | Category: Example spreadsheets            |
  | Keywords: Sample Example Properties       |
  | Comments: Created with Python XlsxWriter  |
  | Status:   Quo                             |
  |                                           |
  | Created       :  2018; 11; 7; 23;         |
  | Modified      :  2018; 11; 7; 23;         |
  | Lastmodifiedby:  Roger Deangelis          |
  | Contentstatus :  Quo                      |
  | Version       :  None                     |
  | Revision      :  None                     |
  | Lastprinted   :  None                     |
  |                                           |
  +-------------------------------------------+


  Create a SAS table from the Document Properties
  -----------------------------------------------

  WORK.WANT

  Middle Observation(1 ) of want - Total Obs 1

   -- CHARACTER --

  Variable         Type      Value

  CREATOR           C15      Roger Deangelis
  TITLE             C30      This is an example spreadsheet
  DESCRIPTION       C34      Created with Python XlsxWrite
  SUBJECT           C24      With document properties
  IDENTIFIER        C4       None
  LANGUAGE          C4       None
  CREATED           C23      2018; 11; 7; 23;
  MODIFIED          C23      2018; 11; 7; 23;
  LASTMODIFIEDBY    C15      Roger Deangelis
  CATEGORY          C20      Example spreadsheets
  CONTENTSTATUS     C3       Quo
  VERSION           C4       None
  REVISION          C4       None
  KEYWORDS          C27      Sample; Example; Properties
  LASTPRINTED       C4       None


 Populate Document properties and add sashelp.class sheet
 ========================================================

  INPUT

  title:    This is an example spreadsheet,
  subject:  With document properties,
  author:   Roger Deangelis,
  manager:  John Smith,
  company:  CompuCraft,
  category: Example spreadsheets,
  keywords: Sample, Example, Properties,
  comments: Created with Python and XlsxWriter,
  status:   Quo


  PROCESS

  libname xel clear;
  %utlfkil(d:/xls/have.xlsx);

  * create empty woorkbook file and populate document properties;

  %utl_submit_py64("
  import xlsxwriter;
  workbook = xlsxwriter.Workbook('d:/xls/have.xlsx');
  worksheet = workbook.add_worksheet();
  workbook.set_properties({
      'title':    'This is an example spreadsheet',
      'subject':  'With document properties',
      'author':   'Roger Deangelis',
      'manager':  'John Smith',
      'company':  'CompuCraft',
      'category': 'Example spreadsheets',
      'keywords': 'Sample, Example, Properties',
      'comments': 'Created with Python and XlsxWriter',
      'status':   'Quo'
  });
  ");

  * add sashelp.class sheet to workbook;

  libname xel "d:/xls/have.xlsx";
  data xel.class;
    set sashelp.class;
  run;quit;
  libname xel clear;

  * drop empty sheet created above by python;

  %utl_submit_py64("
  import openpyxl;
  workbook=openpyxl.load_workbook('d:/xls/have.xlsx');
  workbook.get_sheet_names();
  std=workbook.get_sheet_by_name('Sheet1');
  workbook.remove_sheet(std);
  workbook.save('d:/xls/have.xlsx');
  ");


 Read Document properties and create sas table WANT
 ==================================================


  %utlfkil(d:/txt/properties.txt);
  %utlfkil(d:/txt/properties.sas);
  proc datasets lib=work;
    delete want;
  run;quit;

  %utl_submit_py64("
  import sys;
  from openpyxl import load_workbook;
  wb = load_workbook('d:/xls/properties.xlsx');
  sys.stdout = open('d:/txt/properties.txt', 'w');
  print(wb.properties);
  sys.stdout.close();
  ");


  data want;

     if _n_=0 then do; %let rc=%sysfunc(dosubl(%nrstr(

        data _null_;

          length cut $200;
          infile "d:/txt/properties.txt" lrecl=4096 ;
          input #3;
          putlog _infile_;

          file "d:/sas/properties.sas";

          _infile_=tranwrd(_infile_,"u'","'");
          _infile_=tranwrd(_infile_,", ","; ");
          _infile_=tranwrd(_infile_,"',","';");
          _infile_=tranwrd(_infile_,"None","'None'");
          _infile_=tranwrd(_infile_,"datetime.datetime(","'");
          _infile_=tranwrd(_infile_,")","'");
          _infile_=cats(_infile_,";");

          cw=countc(_infile_,"'")/2 +1;

          do idx=1 to cw;

             cut=cats(scan(strip(_infile_),idx,"="),"=");

             if idx eq cw then cut=substr(cut,1,length(cut)-1);
             putlog  cut=;
             put cut;

          end;

        run;quit;

       )));
     end;

     %include "d:/sas/properties.sas";

  run;quit;

/*
This is valid SAS code - just include it

d:/sas/properties.sas

CUT='Roger Deangelis'; title=
CUT='This is an example spreadsheet'; description=
CUT='Created with Python and XlsxWriter'; subject=
CUT='With document properties'; identifier=
CUT='None'; language=
CUT='None'; created=
CUT='2018; 11; 7; 23; 36; 36'; modified=
CUT='2018; 11; 7; 23; 36; 36'; lastModifiedBy=
CUT='Roger Deangelis'; category=
CUT='Example spreadsheets'; contentStatus=
CUT='Quo'; version=
CUT='None'; revision=
CUT='None'; keywords=
CUT='Sample; Example; Properties'; lastPrinted=
CUT='None';
*/


