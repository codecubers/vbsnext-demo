
call del /f build\*-bundle.vbs

call npx vbsnext index.vbs
call cscript //nologo build\index-bundle.vbs