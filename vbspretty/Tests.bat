call del /f *-js*.vbs
call node index.js
call cscript //nologo index-pretty-js.vbs

call del /f *-cls*.vbs
call npx vbspretty ./index-unpretty.vbs --level 0 --indentChar "\t" --breakLineChar "\r\n" --breakOnSeperator --removeComments --output ./index-pretty-cls-break-nocomments.vbs
call npx vbspretty ./index-unpretty.vbs --level 1 --indentChar "\t" --breakLineChar "\r\n" --output ./index-pretty-cls-level1.vbs
REM call npx vbspretty ./index-unpretty.vbs --level 0 --indentChar "  " --output ./index-pretty-cls-spaced.vbs
call cscript //nologo ./index-pretty-cls-break-nocomments.vbs
call cscript //nologo ./index-pretty-cls-level1.vbs
REM call cscript //nologo ./index-pretty-cls-spaced.vbs