Folder PATH listing for PDFeditor Folder
PDFeditor #Overall repository of PDFtools project
|   readMe.txt
|   
+---docs #containing literature survey
|   \---reference
|           itext_pdfabc.pdf #best and quickest book to understand PDF in easy manner
|           PDF32000_2008.pdf #PDF standard literature by Adobe
|           pdfmarkreference.pdf #Chapter 2 and 3 are useful
|           
+---report
|   \---Final Presentation #contains my presentation
|       |   Enhance PDF paper_20181022.pdf
|
|               
+---src #contains the Source code of the Python program 
|   |   config.xlsx # documents all the opeartion needed to fix the deviations, proper coding for each operations done here
|   |   jdcal.py # helper python library for openpyxl
|   |   jdcal.pyc # neglect it, compiler file
|   |   pdftools.py
|   |   
|   +---et_xmlfile #helper python library for the openpyxl
|   |   |   xmlfile.py # no concern
|   |   |   xmlfile.pyc # neglect it, compiler file
|   |   |   __init__.py # no concern
|   |   |   __init__.pyc # neglect it, compiler file
|   |   |   
|   |   \---tests
|   |           common_imports.py
|   |           helper.py
|   |           test_incremental_xmlfile.py
|   |           __init__.py
|   |           
|   +---openpyxl # helper python library for reading the Excel files, such as config.xlsx or ANASPEC fike
|   |   |   conftest.py
|   |   |   _constants.py
|   |   |   _constants.pyc
|   |   |   __init__.py
|   |   |   __init__.pyc
|   |   |   
|   |   +---cell
|   |   |       cell.py
|   |   |       cell.pyc
|   |   |       interface.py
|   |   |       read_only.py
|   |   |       read_only.pyc
|   |   |       text.py
|   |   |       text.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---chart
|   |   |       area_chart.py
|   |   |       area_chart.pyc
|   |   |       axis.py
|   |   |       axis.pyc
|   |   |       bar_chart.py
|   |   |       bar_chart.pyc
|   |   |       bubble_chart.py
|   |   |       bubble_chart.pyc
|   |   |       chartspace.py
|   |   |       chartspace.pyc
|   |   |       data_source.py
|   |   |       data_source.pyc
|   |   |       descriptors.py
|   |   |       descriptors.pyc
|   |   |       error_bar.py
|   |   |       error_bar.pyc
|   |   |       label.py
|   |   |       label.pyc
|   |   |       layout.py
|   |   |       layout.pyc
|   |   |       legend.py
|   |   |       legend.pyc
|   |   |       line_chart.py
|   |   |       line_chart.pyc
|   |   |       marker.py
|   |   |       marker.pyc
|   |   |       picture.py
|   |   |       picture.pyc
|   |   |       pie_chart.py
|   |   |       pie_chart.pyc
|   |   |       plotarea.py
|   |   |       plotarea.pyc
|   |   |       print_settings.py
|   |   |       print_settings.pyc
|   |   |       radar_chart.py
|   |   |       radar_chart.pyc
|   |   |       reader.py
|   |   |       reader.pyc
|   |   |       reference.py
|   |   |       reference.pyc
|   |   |       scatter_chart.py
|   |   |       scatter_chart.pyc
|   |   |       series.py
|   |   |       series.pyc
|   |   |       series_factory.py
|   |   |       series_factory.pyc
|   |   |       shapes.py
|   |   |       shapes.pyc
|   |   |       stock_chart.py
|   |   |       stock_chart.pyc
|   |   |       surface_chart.py
|   |   |       surface_chart.pyc
|   |   |       text.py
|   |   |       text.pyc
|   |   |       title.py
|   |   |       title.pyc
|   |   |       trendline.py
|   |   |       trendline.pyc
|   |   |       updown_bars.py
|   |   |       updown_bars.pyc
|   |   |       _3d.py
|   |   |       _3d.pyc
|   |   |       _chart.py
|   |   |       _chart.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---chartsheet
|   |   |       chartsheet.py
|   |   |       chartsheet.pyc
|   |   |       custom.py
|   |   |       custom.pyc
|   |   |       properties.py
|   |   |       properties.pyc
|   |   |       protection.py
|   |   |       protection.pyc
|   |   |       publish.py
|   |   |       publish.pyc
|   |   |       relation.py
|   |   |       relation.pyc
|   |   |       views.py
|   |   |       views.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---comments
|   |   |       author.py
|   |   |       author.pyc
|   |   |       comments.py
|   |   |       comments.pyc
|   |   |       comment_sheet.py
|   |   |       comment_sheet.pyc
|   |   |       shape_writer.py
|   |   |       shape_writer.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---compat
|   |   |       abc.py
|   |   |       accumulate.py
|   |   |       accumulate.pyc
|   |   |       numbers.py
|   |   |       numbers.pyc
|   |   |       singleton.py
|   |   |       strings.py
|   |   |       strings.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---descriptors
|   |   |       base.py
|   |   |       base.pyc
|   |   |       excel.py
|   |   |       excel.pyc
|   |   |       namespace.py
|   |   |       namespace.pyc
|   |   |       nested.py
|   |   |       nested.pyc
|   |   |       sequence.py
|   |   |       sequence.pyc
|   |   |       serialisable.py
|   |   |       serialisable.pyc
|   |   |       slots.py
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---drawing
|   |   |       colors.py
|   |   |       colors.pyc
|   |   |       drawing.py
|   |   |       drawing.pyc
|   |   |       effect.py
|   |   |       effect.pyc
|   |   |       fill.py
|   |   |       fill.pyc
|   |   |       graphic.py
|   |   |       graphic.pyc
|   |   |       image.py
|   |   |       image.pyc
|   |   |       line.py
|   |   |       line.pyc
|   |   |       shape.py
|   |   |       shapes.py
|   |   |       shapes.pyc
|   |   |       spreadsheet_drawing.py
|   |   |       spreadsheet_drawing.pyc
|   |   |       text.py
|   |   |       text.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---formatting
|   |   |       formatting.py
|   |   |       formatting.pyc
|   |   |       rule.py
|   |   |       rule.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---formula
|   |   |       tokenizer.py
|   |   |       tokenizer.pyc
|   |   |       translate.py
|   |   |       translate.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---packaging
|   |   |       core.py
|   |   |       core.pyc
|   |   |       extended.py
|   |   |       extended.pyc
|   |   |       interface.py
|   |   |       manifest.py
|   |   |       manifest.pyc
|   |   |       relationship.py
|   |   |       relationship.pyc
|   |   |       workbook.py
|   |   |       workbook.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---pivot
|   |   |       cache.py
|   |   |       cache.pyc
|   |   |       fields.py
|   |   |       fields.pyc
|   |   |       record.py
|   |   |       record.pyc
|   |   |       table.py
|   |   |       table.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---reader
|   |   |       excel.py
|   |   |       excel.pyc
|   |   |       strings.py
|   |   |       strings.pyc
|   |   |       worksheet.py
|   |   |       worksheet.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---styles
|   |   |       alignment.py
|   |   |       alignment.pyc
|   |   |       borders.py
|   |   |       borders.pyc
|   |   |       builtins.py
|   |   |       builtins.pyc
|   |   |       cell_style.py
|   |   |       cell_style.pyc
|   |   |       colors.py
|   |   |       colors.pyc
|   |   |       differential.py
|   |   |       differential.pyc
|   |   |       fills.py
|   |   |       fills.pyc
|   |   |       fonts.py
|   |   |       fonts.pyc
|   |   |       named_styles.py
|   |   |       named_styles.pyc
|   |   |       numbers.py
|   |   |       numbers.pyc
|   |   |       protection.py
|   |   |       protection.pyc
|   |   |       proxy.py
|   |   |       proxy.pyc
|   |   |       styleable.py
|   |   |       styleable.pyc
|   |   |       stylesheet.py
|   |   |       stylesheet.pyc
|   |   |       table.py
|   |   |       table.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---utils
|   |   |       bound_dictionary.py
|   |   |       bound_dictionary.pyc
|   |   |       cell.py
|   |   |       cell.pyc
|   |   |       dataframe.py
|   |   |       datetime.py
|   |   |       datetime.pyc
|   |   |       escape.py
|   |   |       escape.pyc
|   |   |       exceptions.py
|   |   |       exceptions.pyc
|   |   |       formulas.py
|   |   |       formulas.pyc
|   |   |       indexed_list.py
|   |   |       indexed_list.pyc
|   |   |       protection.py
|   |   |       protection.pyc
|   |   |       units.py
|   |   |       units.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---workbook
|   |   |   |   child.py
|   |   |   |   child.pyc
|   |   |   |   defined_name.py
|   |   |   |   defined_name.pyc
|   |   |   |   external_reference.py
|   |   |   |   external_reference.pyc
|   |   |   |   function_group.py
|   |   |   |   function_group.pyc
|   |   |   |   parser.py
|   |   |   |   parser.pyc
|   |   |   |   properties.py
|   |   |   |   properties.pyc
|   |   |   |   protection.py
|   |   |   |   protection.pyc
|   |   |   |   smart_tags.py
|   |   |   |   smart_tags.pyc
|   |   |   |   views.py
|   |   |   |   views.pyc
|   |   |   |   web.py
|   |   |   |   web.pyc
|   |   |   |   workbook.py
|   |   |   |   workbook.pyc
|   |   |   |   __init__.py
|   |   |   |   __init__.pyc
|   |   |   |   
|   |   |   \---external_link
|   |   |           external.py
|   |   |           external.pyc
|   |   |           __init__.py
|   |   |           __init__.pyc
|   |   |           
|   |   +---worksheet
|   |   |       cell_range.py
|   |   |       cell_range.pyc
|   |   |       copier.py
|   |   |       copier.pyc
|   |   |       datavalidation.py
|   |   |       datavalidation.pyc
|   |   |       dimensions.py
|   |   |       dimensions.pyc
|   |   |       drawing.py
|   |   |       drawing.pyc
|   |   |       filters.py
|   |   |       filters.pyc
|   |   |       header_footer.py
|   |   |       header_footer.pyc
|   |   |       hyperlink.py
|   |   |       hyperlink.pyc
|   |   |       merge.py
|   |   |       merge.pyc
|   |   |       page.py
|   |   |       page.pyc
|   |   |       pagebreak.py
|   |   |       pagebreak.pyc
|   |   |       properties.py
|   |   |       properties.pyc
|   |   |       protection.py
|   |   |       protection.pyc
|   |   |       read_only.py
|   |   |       read_only.pyc
|   |   |       related.py
|   |   |       related.pyc
|   |   |       table.py
|   |   |       table.pyc
|   |   |       views.py
|   |   |       views.pyc
|   |   |       worksheet.py
|   |   |       worksheet.pyc
|   |   |       write_only.py
|   |   |       write_only.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   +---writer
|   |   |       etree_worksheet.py
|   |   |       etree_worksheet.pyc
|   |   |       excel.py
|   |   |       excel.pyc
|   |   |       strings.py
|   |   |       strings.pyc
|   |   |       theme.py
|   |   |       theme.pyc
|   |   |       workbook.py
|   |   |       workbook.pyc
|   |   |       worksheet.py
|   |   |       worksheet.pyc
|   |   |       __init__.py
|   |   |       __init__.pyc
|   |   |       
|   |   \---xml
|   |           constants.py
|   |           constants.pyc
|   |           functions.py
|   |           functions.pyc
|   |           __init__.py
|   |           __init__.pyc
|   |           
|   +---pdfminer # helper library file for fetching the bookmarks from the Define.pdf
|   |       arcfour.py
|   |       arcfour.pyc
|   |       ascii85.py
|   |       ascii85.pyc
|   |       ccitt.py
|   |       ccitt.pyc
|   |       cmapdb.py
|   |       converter.py
|   |       encodingdb.py
|   |       fontmetrics.py
|   |       glyphlist.py
|   |       image.py
|   |       latin_enc.py
|   |       layout.py
|   |       lzw.py
|   |       lzw.pyc
|   |       Makefile
|   |       pdfcolor.py
|   |       pdfdevice.py
|   |       pdfdocument.py
|   |       pdfdocument.pyc
|   |       pdffont.py
|   |       pdfinterp.py
|   |       pdfpage.py
|   |       pdfparser.py
|   |       pdfparser.pyc
|   |       pdftypes.py
|   |       pdftypes.pyc
|   |       psparser.py
|   |       psparser.pyc
|   |       rijndael.py
|   |       runlength.py
|   |       runlength.pyc
|   |       utils.py
|   |       utils.pyc
|   |       __init__.py
|   |       __init__.pyc
|   |       
|   |       
|   \---PyPDF2 # the core of the PDFTools, responsible for reading the input Define.pdf and breaking it down into objects for further manupulations by the pdftools.py
|           filters.py
|           filters.pyc
|           generic.py # contains classes which defines the data structure in the PDF, such as dictionary, NameObject, NumberObject, ArrayObject etc
|           generic.pyc
|           merger.py
|           merger.pyc
|           pagerange.py
|           pagerange.pyc
|           pdf.py # contains PDFFileReader and PDFFileWriter classes for reading and writing the pdf
|           pdf.pyc
|           utils.py
|           utils.pyc
|           xmp.py
|           _version.py
|           _version.pyc
|           __init__.py
|           __init__.pyc
|           
\---tools # previously tried tools, Look for RUPS, it is cool, and helps in visualising the PDF
    |   cpdfdemo.zip #zipfiles
    |   itext-rups-5.5.9.zip #zipfiles
    |   pdfminer-master.zip #zipfiles
    |   pdftk_server-2.02-win-setup.exe #zipfiles
    |   
    +---cpdfdemo #command based tool, totally uneditable, lacks most of the features in handling the PDF
    |   |   cpdf.exe # refer to demo.cmd or cpdfmanual to learn about this tool
    |   |   cpdfmanual.pdf
    |   |   demo.cmd
    |   |   license.txt
    |   |   output.txt
    |   |   readme.txt
    |   |   ssed.exe
    |   |   
    |   \---linearize
    |           cpdflin.exe
    |           libgcc_s_dw2-1.dll
    |           libstdc++-6.dll
    |           qpdf13.dll
    |           
    +---itext-rups-5.5.9 # RUPS: read update PDF specs, good for visualising the whole pdf, one of the important tool
    |       BUILDING.md
    |       CODE_OF_CONDUCT.md
    |       CONTRIBUTING.md
    |       itext-rups-5.5.9-jar-with-dependencies.jar
    |       itext-rups-5.5.9-jar-with-dependencies.jar.asc
    |       itext-rups-5.5.9-jar-with-dependencies.jar.md5
    |       itext-rups-5.5.9-jar-with-dependencies.jar.sha1
    |       itext-rups-5.5.9-javadoc.jar
    |       itext-rups-5.5.9-javadoc.jar.asc
    |       itext-rups-5.5.9-javadoc.jar.md5
    |       itext-rups-5.5.9-javadoc.jar.sha1
    |       itext-rups-5.5.9-sources.jar
    |       itext-rups-5.5.9-sources.jar.asc
    |       itext-rups-5.5.9-sources.jar.md5
    |       itext-rups-5.5.9-sources.jar.sha1
    |       itext-rups-5.5.9.exe ## run this tool
    |       itext-rups-5.5.9.exe.asc
    |       itext-rups-5.5.9.exe.md5
    |       itext-rups-5.5.9.exe.sha1
    |       itext-rups-5.5.9.jar
    |       itext-rups-5.5.9.jar.asc
    |       itext-rups-5.5.9.jar.md5
    |       itext-rups-5.5.9.jar.sha1
    |       itext-rups-5.5.9.pom
    |       itext-rups-5.5.9.pom.asc
    |       itext-rups-5.5.9.pom.md5
    |       itext-rups-5.5.9.pom.sha1
    |       LICENSE.md
    |       README.md
    |       
    +---pdfminer-master # used for fetching the bookmarks from the any pdf, it is being used along with PYPDF2 to correctly repair the bookmarks
    |   |   .travis.yml
    |   |   LICENSE
    |   |   Makefile
    |   |   MANIFEST.in
    |   |   README.md
    |   |   setup.py
    |   |   
    |   +---cmaprsrc
    |   |       cid2code_Adobe_CNS1.txt
    |   |       cid2code_Adobe_GB1.txt
    |   |       cid2code_Adobe_Japan1.txt
    |   |       cid2code_Adobe_Korea1.txt
    |   |       README.txt
    |   |       
    |   +---docs
    |   |       cid.obj
    |   |       cid.png
    |   |       index.html
    |   |       layout.obj
    |   |       layout.png
    |   |       objrel.obj
    |   |       objrel.png
    |   |       programming.html
    |   |       style.css
    |   |       
    |   +---pdfminer
    |   |       arcfour.py
    |   |       ascii85.py
    |   |       ccitt.py
    |   |       cmapdb.py
    |   |       converter.py
    |   |       encodingdb.py
    |   |       fontmetrics.py
    |   |       glyphlist.py
    |   |       image.py
    |   |       latin_enc.py
    |   |       layout.py
    |   |       lzw.py
    |   |       Makefile
    |   |       pdfcolor.py
    |   |       pdfdevice.py
    |   |       pdfdocument.py
    |   |       pdffont.py
    |   |       pdfinterp.py
    |   |       pdfpage.py
    |   |       pdfparser.py
    |   |       pdftypes.py
    |   |       psparser.py
    |   |       rijndael.py
    |   |       runlength.py
    |   |       utils.py
    |   |       __init__.py
    |   |       
    |   +---samples
    |   |   |   jo.html.ref
    |   |   |   jo.pdf
    |   |   |   jo.tex
    |   |   |   jo.txt.ref
    |   |   |   jo.xml.ref
    |   |   |   Makefile
    |   |   |   README
    |   |   |   simple1.html.ref
    |   |   |   simple1.pdf
    |   |   |   simple1.txt.ref
    |   |   |   simple1.xml.ref
    |   |   |   simple2.html.ref
    |   |   |   simple2.pdf
    |   |   |   simple2.txt.ref
    |   |   |   simple2.xml.ref
    |   |   |   simple3.html.ref
    |   |   |   simple3.pdf
    |   |   |   simple3.txt.ref
    |   |   |   simple3.xml.ref
    |   |   |   
    |   |   +---encryption
    |   |   |       aes-128-m.pdf
    |   |   |       aes-128.pdf
    |   |   |       aes-256-m.pdf
    |   |   |       aes-256.pdf
    |   |   |       base.pdf
    |   |   |       base.xml
    |   |   |       Makefile
    |   |   |       rc4-128.pdf
    |   |   |       rc4-40.pdf
    |   |   |       
    |   |   \---nonfree
    |   |           dmca.html.ref
    |   |           dmca.pdf
    |   |           dmca.txt.ref
    |   |           dmca.xml.ref
    |   |           f1040nr.html.ref
    |   |           f1040nr.pdf
    |   |           f1040nr.txt.ref
    |   |           f1040nr.xml.ref
    |   |           i1040nr.html.ref
    |   |           i1040nr.pdf
    |   |           i1040nr.txt.ref
    |   |           i1040nr.xml.ref
    |   |           kampo.html.ref
    |   |           kampo.pdf
    |   |           kampo.txt.ref
    |   |           kampo.xml.ref
    |   |           naacl06-shinyama.html.ref
    |   |           naacl06-shinyama.pdf
    |   |           naacl06-shinyama.txt.ref
    |   |           naacl06-shinyama.xml.ref
    |   |           nlp2004slides.html.ref
    |   |           nlp2004slides.pdf
    |   |           nlp2004slides.txt.ref
    |   |           nlp2004slides.xml.ref
    |   |           
    |   \---tools
    |           conv_afm.py
    |           conv_cmap.py
    |           conv_glyphlist.py
    |           dumppdf.py
    |           latin2ascii.py
    |           Makefile
    |           pdf2html.cgi
    |           pdf2txt.py
    |           prof.py
    |           runapp.py
    |           
    \---programs
        +---BAT Test Scripts
        |       pdftools_1112.bat
        |       pdftools_1116.bat
        |       pdftools_1127.bat
        |       pdftools_1130.bat
        |       pdftools_ISS-CLL2018.bat
        |       pdftools_pmr.bat
        |       
        +---MACRO
        |       fixpdf.sas
        |       fixpdf_old.sas
        |       
        \---SAS Test Scripts
                v_fixpdf_1112.sas
                v_fixpdf_1116.sas
                v_fixpdf_1127.sas
                v_fixpdf_1130.sas
                v_fixpdf_ISS-CLL2018.sas
                v_fixpdf_pmr.sas
                
