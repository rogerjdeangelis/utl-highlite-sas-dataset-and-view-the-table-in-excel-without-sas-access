# utl-highlite-sas-dataset-and-view-the-table-in-excel-without-sas-access
    %let pgm=utl-hilite-sas-dataset-and-view-the-table-in-excel-without-sas-access;

    Highlite sas dataset and type xlr on command and view the SAS table in excel without sas access

    github  (github is a code development site - user feedback and issues are welcome)
    https://tinyurl.com/jjpxp3p5
    https://github.com/rogerjdeangelis/utl-highlite-sas-dataset-and-view-the-table-in-excel-without-sas-access


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

    %macro xlr /cmd ;
       store;note;notesubmit '%xlra;';
       run;
    %mend xlr;

    %macro xlra/cmd;

        %local argx;

        filename clp clipbrd ;

        data _null_;
           infile clp;
           input;
           argx=_infile_;
           call symputx("argx",argx);
           putlog argx=;
        run;quit;

        /* %let argx=sashelp.class; */

        %utlfkil(%sysfunc(getoption(work))/_rpt.xlsx);

        ods listing close;

        ods excel file="%sysfunc(getoption(work))/_rpt.xlsx"
                options(
                   sheet_name                 = "&argx"
                   autofilter                 = "yes"
                   frozen_headers             = "1"
                   gridlines                  = "yes"
                   embedded_titles            = "yes"
                   embedded_footnoteS         = "yes"
                   );

        proc report data=&argx missing;
        run;quit;

        ods excel close;

        ods listing;

        options noxwait noxsync;
        /* Open Excel */
        x "'C:\Program Files\Microsoft Office\OFFICE14\excel.exe' %sysfunc(getoption(work))/_rpt.xlsx";
        run;quit;

    %mend xlra;
