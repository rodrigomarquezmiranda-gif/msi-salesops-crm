@echo off
echo Eliminando git lock files...
if exist ".git\index.lock" (
    del /f ".git\index.lock"
    echo index.lock eliminado
) else (
    echo No habia index.lock
)
if exist ".git\HEAD.lock" (
    del /f ".git\HEAD.lock"
    echo HEAD.lock eliminado
) else (
    echo No habia HEAD.lock
)
if exist ".git\ORIG_HEAD.lock" (
    del /f ".git\ORIG_HEAD.lock"
    echo ORIG_HEAD.lock eliminado
) else (
    echo No habia ORIG_HEAD.lock
)
echo Listo. Podes hacer commit en GitHub Desktop.
pause
