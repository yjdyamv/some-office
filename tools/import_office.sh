export GHIDRA_INSTALL_DIR=/c/ghidra_12.2_PUBLIC
G=~/re/.venv/Scripts/ghlab.exe
O="/c/Program Files/Microsoft Office/root/Office16"
"$G" import "$O/POWERPNT.EXE" > /tmp/imp_ppt.log 2>&1
"$G" import "$O/wwlib.dll"    > /tmp/imp_wwlib.log 2>&1
"$G" import "$O/EXCEL.EXE"    > /tmp/imp_excel.log 2>&1
echo DONE > /tmp/imp_done
