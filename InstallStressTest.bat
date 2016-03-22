@echo off
pushd %~dp0
cscript %~n0.js %0 %1 %2 %3 %4 %5
popd
