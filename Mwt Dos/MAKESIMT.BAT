echo Re-creating MWT2_87s.ZIP
del mwt2_87*.*
attrib mwt*.* -r
fd mwt*.* /d1-1-97 /t2:08a
fd readme.txt /d1-1-97 /t2:08a
attrib mwt_abbr.bkk +r
pkzip mwt2_87s.zip mwt*.* -ex -z < comment2.txt
pkzip mwt2_87s.zip readme.txt -a -ex
fd mwt2_87s.* /d1-1-97 /t2:08a
