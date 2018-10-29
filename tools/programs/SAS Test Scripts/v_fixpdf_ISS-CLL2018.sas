/*****************************************************************************************
* v_fixpdf_ISS-CLL2018.sas version 1.0 -- Update a define.pdf with certain sets of requirements 
******************************************************************************************
* Project :  Statistical Computing Platform
* Study   :  Any
* Author  :  Jayant Solanki
* Creator :  Sy Truong
* Date    :  08/01/2018
* Updated :  08/02/2018
* Note    :  This script will be used to call the fixpdf macro
******************************************************************************************/

Filename fixpdf 'J:\Biometrics\Statistical Programming\Intern\2018\pdftools\macros\fixpdf.sas';
%include fixpdf;
%fixpdf(inpath = J:\Biometrics\Statistical Programming\Intern\2018\pdftools\esub\analysis\adam\datasets);
