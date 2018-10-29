/*****************************************************************************************
* fixpdf.sas version 1.0 -- Update a define.pdf with certain sets of requirements 
******************************************************************************************
* Project :  Statistical Computing Platform
* Study   :  Any
* Author  :  Jayant Solanki
* Creator :  Sy Truong
* Date    :  07/23/2018
* Updated :  07/30/2018
* Note    :  This is macro which can be included in any other SAS script
******************************************************************************************
* SAS Macro used to update a define.pdf removing TFL prefixes used for ordering 
******************************************************************************************
* Macro Parameters
*
* inpath     C (200 chars) path and file name of define.pdf file
* anaspec   C (200 chars) optional, The full path and file name of the Excel analysis result
*			specification file such as: pcyc_anlres_specs_1130.xlsx.  If this was left blank,
*			no updates will be applied to the define.pdf relating to analysis results
******************************************************************************************/

*** options mprint symbolgen;
/*options mprint merror symbolgen;*/
options mprint;


data _null_;
  workpath = pathname('work');
  call symput('workpath',trim(workpath));
run;


************************************************;
*** Start of the fixpdf Macro definition ***;
************************************************;
%macro fixpdf(inpath=, anaspec=);	
    %let pythonPath=J:\Biometrics\Statistical Programming\Intern\2018\pdftools\Miniconda2\python;
	%let mainFile=J:\Biometrics\Statistical Programming\Intern\2018\pdftools\pdfeditor\src\pdftools.py;
	%let inFile=define.pdf;
	%let outFile=postprocessed_define.pdf; *** subject to change ***;
	%let batchFile=&workpath\fixpdf.bat;
	%put &pythonPath;
    *** Intialize validation error findings ***;
    %let anyerror=no;
    %let _err=no;
	%let _success=no;

    *** Perform Error checking on parameters ***;
	data _null_;

		put "NOTE:  *** You are currently running the macro fixpdf version 1.0 ***";

        *** Verify that all required parameters are not missing ***;
        *** Handle default values ***;
		*** check if the inpath parameter is provided by the user or not ***;
        if compress("&inpath") = '' then do;
           put " ";
		   put " ";
           put "ER" "ROR: *** [fixpdf] is not able to find the path specified by the INPATH parameter ***";
           put "ER" "ROR: *** [fixpdf] was not able to find the DEFINE.PDF file in the path specified ***";
           put " ";
		   put " ";
           call symput('_err','yes');
        end;
		*** check whether the define.pdf is present in the directory provided the user ***;
		%if %sysfunc(fileexist(&inpath/&inFile))= 0 %then
			%do;
				put "ER" "ROR: *** [fixpdf] is not able to find the define.pdf file in the working directory ***";
				call symput('_err','yes');
			%end;
    run;
    *** Skip the logic to the end if there was an error encountered ***;
    %if "&_err" = "yes" %then %goto endcheck;

   	*** Document what parameters users specified ***;
   	%put *** User specified the following parameters for fixpdf macro. ***;
   	%put *** INPATH=&inpath;
	%put *** ANASPEC=&anaspec;
	*** Start preparing the batch for SAS ***;
	data pdf_tool;
					
/*		file "fixpdf.bat";*/
		file "&batchFile";
		if compress("&anaspec") = '' then 
			do;
				temppythonPath = quote("&pythonPath","");
				tempmainFile = quote("&mainFile","");
				tempinFile = quote("&inpath\&inFile", "");
				tempoutFile = quote("&inpath\&outFile", "");
				templogFile = quote("&inpath\define.txt", "");
				put temppythonPath tempmainFile "--inFile " tempinFile "--outFile " tempoutFile; *** " --logFile " templogFile***;
			end;
		else 
			do;
				temppythonPath = quote("&pythonPath","");
				tempmainFile = quote("&mainFile","");
				tempanaSpec = quote("&anaspec","");
				tempinFile = quote("&inpath\&inFile", "");
				tempoutFile = quote("&inpath\&outFile", "");
				templogFile = quote("&inpath\define.txt", "");
				put temppythonPath tempmainFile "--specFile " tempanaSpec "--inFile " tempinFile "--outFile " tempoutFile; *** " --logFile " templogFile***;
			end;
	run;
	*** Create default for OUTXML and make a backup before updating ***; 
   	%let incpcmd=no;
	%let outcpcmd=no;
	%let mvcmd=no;
   	%let backup=no;
	data _null_;
        curdate = compress(tranwrd(strip(tranwrd(put(date(),yymmdd10.) || '_' || put(time(),time5.),':','_')),'-',''));
/*        inbackup = tranwrd("&inFile",'.pdf','_preprocessed'||strip(curdate) || '.pdf');*/
		inbackup = strip(curdate) || '_preprocessed_'|| "&inFile";
		outbackup = strip(curdate) ||'_' || "&outFile";
        call symput('backup',strip(backup));
		cdcmd = "cd " || "&inpath";
        call symput('cdcmd',strip(cdcmd));
        incpcmd = "'if exist "|| strip("&outFile") || " copy /Y " || strip("&inFile") || " Archive\" || strip(inbackup) || "'";
        call symput('incpcmd',strip(incpcmd));
		outcpcmd = "'if exist "|| strip("&outFile") || " copy /Y " || strip("&outFile") || " Archive\" || strip(outbackup) || "'";
        call symput('outcpcmd',strip(outcpcmd));
		mvcmd = "'if exist "|| strip("&outFile") || " move /Y " || strip("&outFile") ||" "|| strip("&inFile") || "'";
		call symput('mvcmd',strip(mvcmd));
	run;
	%put &cdcmd;
	%put &incpcmd;
	%put &outcpcmd;
	%put &mvcmd;
	*** This step first creates Archive folder, then runs the fixpdf batch file, if define_good.pdf produced by the batch file then ***;
	*** it will archive the define.pdf, then rename define_good.pdf as define.pdf ***;
	options noxwait xsync;
	x if not exist "&inpath"/Archive md "&inpath"/Archive & exit; *** Creating Archive folder in the working directory, incase it is not present *** ;
	%put ;
	%put ;
	%put WARNING: *** Please do not close the DOS Window, it will exit itself ***;
	%put ;
	%put ;
	options noxwait xsync;
/*	x "&batchFile"; *** Running the batch file ***;*/
	%put ;
	%put ;
	%put WARNING: *** Log File generated by the Python pdftool displayed below ***;
	%put ;
	%put ;
	filename tasks pipe "&batchFile"; *** Piping the output of bat file to the sas log ***;
	data pythonLog;
		infile tasks lrecl=1500;
		LENGTH logLines $2000;
		input logLines $char2000.;
		put logLines;
	run;
	options noxwait xsync;
	x del "&inpath\define.txt";
	%put ;
	%put ;
	%put WARNING: *** End of Log File ***;
	%put ;
	%put ;
	data _null_;
	    %if "&incpcmd" ne "no" %then 
			%do;
		        options noxwait xsync;
					
				
				%if %sysfunc(fileexist(&inpath/&outFile))= 0 %then
					%do;
						call symput('_err','yes');
						put " ";
						put " ";
						put "ER" "ROR: *** [fixpdf] is not able to find the output file generated by the fixpdf tool ***";
						put " ";
						put " ";
					%end;
				%else
					%do;
						x &cdcmd; *** changing current directory to inpath directory ***;
						x &incpcmd; *** copying define.pdf into Archive folder ***;
						x &outcpcmd; *** copying define_postprocessed.pdf into Archive folder ***;
						x &mvcmd; *** renaming define_postprocessed.pdf to define.pdf ***;
						call symput('_success','yes');
						put " ";
						put " ";
						put "NOTE: *** SUCCESS, [fixpdf] successfully generated the define.pdf file ***";
						put " ";
						put " ";
					%end;
		 %end;

	run;
   	*** Mark the end of parameter checking and there were ERRORs found ***;
   	%endcheck:
	data _null_;
		%if "&_err" = "yes" %then 
			%do;
				%put ;
				%put ;
				put "ER" "ROR: *** Unsuccessful, please check log for more details ***";
				%put ;
				%put ;
			%end;
		%if "&_success" = "yes" %then 
			%do;
				%put ;
				%put ;
				put "NOTE: *** SUCCESS ***";
				%put ;
				%put ;
			%end;
	run;
%mend fixpdf;

