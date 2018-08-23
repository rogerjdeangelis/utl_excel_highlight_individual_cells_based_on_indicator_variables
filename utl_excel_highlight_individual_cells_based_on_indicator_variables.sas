Excel highlight individual cells based on indicator variables

Problem

  Highlight Age and Height when indicator variables hilite1 and hilite2 contain 'AGE' and 'HEIGHT' respectively 

  I assume hilite1 and hilite2 have the same variable names for all observations.
  Can be made dynamic.

Using ods excel instead of tagsets

for output see
https://tinyurl.com/yadneyrj
https://github.com/rogerjdeangelis/utl_excel_highlight_individual_cells_based_on_indicator_variables/blob/master/utl_excel_highlight_individual_cells_based_on_indicator_variables.xlsx

github
https://tinyurl.com/ya5t8xvo
https://github.com/rogerjdeangelis/utl_excel_highlight_individual_cells_based_on_indicator_variables

SAS Forum
https://tinyurl.com/yb3ceq27
https://communities.sas.com/t5/SAS-Programming/Unable-to-highlight-cells-in-ODS-tagsets-excelxp-and-proc-report/m-p/489237


INPUT
=====

 WORK.HAVE total obs=3                                      | RULES (where to highlight)
                                                            |
  NAME     AGE    HEIGHT      FLAG      HILITE1    HILITE2  |    AGE     HEIGHT
                                                            |
 Joyce     11      51.3     UPDATED       AGE      HEIGHT   |    YELLOW   YELLOW
 Joyce     12      51       UPDATED       AGE      HEIGHT   |    YELLOW   YELLOW
 Thomas    11      57.5     NOCHANGE                        |    11      57.5      <- not yellow


EXAMPLE OUTPUT
==============

 WORK.LOG total obs=1

   RC                     STATUS

    0    Meta data extraction and report completed

PROCESS
=======

  options validvarname=upcase; * otherwise use upcase in code;

  %symdel cc_mac cc_rpt / nowarn;
  data log;

    if _n_=0 then do;
       %let rc=%sysfunc(dosubl('
           data _null_;
              set have(obs=1);
              call symputx("hilite1",hilite1);
              call symputx("hilite2",hilite2);
           run;quit;
           %let cc_mac=&syserr;
       '));
    end;

    rc=dosubl('
       ods excel file="d:/xls/utl_excel_highlight_individual_cells_based_on_indicator_variables.xlsx";

       * use list option to generate report code and then edit it;
       PROC REPORT DATA=WORK.HAVE LS=171 PS=65  SPLIT="/" NOCENTER MISSING ;
       COLUMN  HILITE1 HILITE2  NAME AGE HEIGHT FLAG;

       DEFINE  HILITE1/ DISPLAY FORMAT= $8. WIDTH=8 SPACING=2 LEFT NOPRINT;
       DEFINE  HILITE2/ DISPLAY FORMAT= $8. WIDTH=8 SPACING=2 LEFT NOPRINT;
       DEFINE  NAME   / DISPLAY FORMAT= $8. WIDTH=8 SPACING=2 LEFT        ;
       DEFINE  AGE    / DISPLAY FORMAT= $8. WIDTH=8 SPACING=2 LEFT        ;
       DEFINE  HEIGHT / DISPLAY FORMAT= $8. WIDTH=8 SPACING=2 LEFT        ;
       DEFINE  FLAG   / DISPLAY FORMAT= $8. WIDTH=8 SPACING=2 LEFT        ;

       compute &hilite1;
           if hilite1="&hilite1" then call define (_col_, "style", "style={background=yellow}");
       endcomp;

       compute &hilite2;
           if hilite2="&hilite2" then call define (_col_, "style", "style={background=yellow}");
       endcomp;
       run;quit;
       %let cc_rpt=&syserr;
       ods excel close;
    ');

    if symgetn('cc_mac')=0 and symgetn('cc_rpt')=0 then status="Meta data extraction and report completed";
    else status="Failed";
    output;
    stop;

run;quit;

OUTPUT
======

*                _              _       _
 _ __ ___   __ _| | _____    __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/ | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|

;


data have;
  input (Name Age Height FLAG Hilite1 hilite2) ($);
cards4;
Joyce 11 51.3 UPDATED AGE HEIGHT
Joyce 12 51 UPDATED AGE HEIGHT
Thomas 11 57.5 NOCHANGE . .
;;;;
run;quit;

/* see process above */

