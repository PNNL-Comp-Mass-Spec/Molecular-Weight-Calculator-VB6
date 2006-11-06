echo Re-creating MWT2_87.ZIP and .EXE files
del mwt2_87*.*
attrib mwt*.* -r
fd mwt*.* /d1-1-97 /t2:08a
attrib mwt_abbr.bkk +r
pkzip mwt2_87.zip mwt*.* -ex -z < comment.txt
zip2exe mwt2_87.zip
rem uuencode -s mwt2_85.exe
fd mwt2_87.* /d1-1-97 /t2:08a
echo.
