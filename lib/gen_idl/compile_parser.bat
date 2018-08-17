@echo off
setlocal
set path=..\..\..\vbpeg;%PATH%
set vbpeg=vbpeg.exe
set infile=%~dp0gen_idl.peg
set outfile=%~dp0mdParser.bas

%vbpeg% %infile% -o %outfile%
