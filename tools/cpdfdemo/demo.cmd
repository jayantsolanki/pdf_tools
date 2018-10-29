rem rem %demo.cmd
rem rem del compressed.pdf
rem rem del out.pdf
rem rem cpdf -prerotate  -add-rectangle "60 2000" -pos-right "620 0"  -color white define_bad.pdf -o tempout.pdf
rem rem cpdf -prerotate -bottom 20 -font-size 10 -add-text "%%Page of %%EndPage" tempout.pdf -o out.pdf
rem rem del tempout.pdf
rem cpdf -list-bookmarks out.pdf > output.txt
rem cpdf -decompress out.pdf -o uncompressed.pdf
rem ssed -e "s/001-Table/Table/g" < uncompressed.pdf > modified1.pdf
rem ssed -e "s/002-Table/Table/g" < modified1.pdf > modified.pdf
rem cpdf -compress modified.pdf -o compressed.pdf
rem del modified1.pdf
rem del modified.pdf
rem del uncompressed.pdf
rem rem cpdf -list-bookmarks compressed.pdf
rem compressed.pdf

rem rem cpdf -add-bookmarks output.txt uncompress.pdf > output.pdf


rem ssed -e "s/GoToR//F/Table/g" < uncompressed.pdf > modified1.pdf

cpdf -decompress define.pdf -o define_uncompressed.pdf