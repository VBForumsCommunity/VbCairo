@echo off
setlocal
set prj_root=%~dp0..\..
set gen_idl=%prj_root%\lib\gen_idl\gen_idl.exe
set idl_file=%~dp0VbCairo.idl

if [%1]==[tlb] goto make_tlb
%gen_idl% %prj_root%\lib\rcairo-1.14.12 -def %prj_root%\src\VbCairo.def -o %idl_file%
:make_tlb
mktyplib %idl_file% /tlb %prj_root%\bin\typelib\VbCairo.tlb
