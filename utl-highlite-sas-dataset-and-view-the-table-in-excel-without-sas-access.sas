%let pgm=utl-highlite-sas-dataset-and-view-the-table-in-excel-without-sas-access;

Highlite sas dataset and type xlh on command and view the SAS table in excel without sas access

github  (github is a code development site - user feedback and issues are welcome)
https://tinyurl.com/jjpxp3p5
https://github.com/rogerjdeangelis/utl-highlite-sas-dataset-and-view-the-table-in-excel-without-sas-access

This will not work with concatenated libraries like sashelp.

In the pre SAS enhanced editor highlite sd1.heart and type xlh on the old command line.
Excel will open up with sd1.heart converted to an excel workbook.

If you have sas access to excel you can use the xlsh macro. (see many macros utl_perpac)
https://tinyurl.com/ae4cszy7
https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories

libname sd1 "d:/sd1";

data sd1.heart; /* highlite sd1.heart and type xlh - part of my point and shoot performance macros */
  set sashelp.heart;
run;quit;

* put both  macros in member xlh.sas in your autocall library;

/*
 _ __ ___   __ _  ___ _ __ ___
| `_ ` _ \ / _` |/ __| `__/ _ \
| | | | | | (_| | (__| | | (_) |
|_| |_| |_|\__,_|\___|_|  \___/

*/

%macro xlh /cmd ;
   store;note;notesubmit '%xlha;';
   run;
%mend xlh;

%macro xlha/cmd;

    filename clp clipbrd ;

    data _null_;
       length fyl $500;
       infile clp;
       input;

       dsn=_infile_;

       wrk=translate("%sysfunc(getoption(work))",'/','\');

       if index(dsn,".")=0 then do;
           * get work directory fix slashes \ is R escape char;
           fyl=cats(wrk,'/',dsn);
           put fyl;
       end;
       else do;
          folder=translate(pathname(scan(dsn,1,'.')),'/','\');
          fyl=cats(folder,'/',scan(dsn,2,'.'));
          put fyl;
       end;
       call symputx('fyl',fyl);
       call symputx('wrk',wrk);
   run;quit;

    %utlfkil(&wrk/_xls.xlsx);

    %utl_submit_r64("
       library(haven);
       library(XLConnect);
       have<-read_sas('&fyl..sas7bdat');
       wb <- loadWorkbook('&wrk/_xls.xlsx', create = TRUE);
       createSheet(wb, name = 'have');
       writeWorksheet(wb, have, sheet = 'have');
       saveWorkbook(wb);
    ");

    options noxwait noxsync;
    /* Open Excel */
    x "'C:\Program Files\Microsoft Office\OFFICE14\excel.exe' &wrk/_xls.xlsx";
    run;quit;

%mend xlha;
