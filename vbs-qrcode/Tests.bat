
call del /f test\test*.bmp
call del /f test\test*.svg
call rm -f test/test*.bmp
call rm -f test/test*.svg

call npx qrcode /data:"Hello World" /out:"test\test-qrcode1.bmp"
call npx qrcode /data:"Hello World" /out:"test\test-qrcode2.bmp" /forecolor:#0000FF /backcolor:#E0FFFF /ecr:L /scale:5 /colordepth:1
call npx qrcode /data:"Hello World" /out:"test\test-qrcode3.svg"
REM call npx qrcode "test\test.txt" /out:"test\test-qrcode4.bmp"
