@echo off
setlocal
call "C:\Program Files (x86)\Microsoft Visual Studio 14.0\VC\bin\vcvars32.bat"
set cl_exe=cl.exe /DZLIB_STDCALL=__stdcall /DZLIB_STANDALONE /DZLIB_STDCALL=__stdcall
set cl_opt=-arch:IA32 /MD /LD /O2 /Ox /DNDEBUG -GS- -Gs9999999 -EHa- -Oi
:: -c -FAs
set link_opt=/link /def:debug_sshzlib.def
set bin_dir=bin
set bin_file=debug_sshzlib.dll

pushd %~dp0

if [%1]==[debug] (
    set cl_opt=/MDd /LDd /Zi
    set link_opt=%link_opt% /debug /incremental:no
    shift /1
)
if [%1]==[test] (
    set cl_opt=/MD /O2 /Ox /DNDEBUG /DZLIB_TEST_MAIN
    set bin_file=debug_sshzlib.exe
)

%cl_exe% %cl_opt% sshzlib.c /Fe%bin_file% %link_opt%
if errorlevel 1 goto :eof

copy *.dll %bin_dir% > nul || exit /b 1
copy *.pdb %bin_dir% > nul 2> nul
copy *.obj %bin_dir% > nul 2> nul
dumpbin /relocations %bin_dir%\sshzlib.obj > %bin_dir%\sshzlib.txt

..\codegen\codegen.exe

:cleanup
del /q *.exp *.lib *.obj *.dll *.pdb *.ilk ~$* 2> nul

:eof
popd
