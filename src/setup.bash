#!/bin/bash

xlsxwritter="pip3 install xlsxwriter --user"
echo "installing" $xlsxwritter
eval $xlsxwritter
xlrd="pip3 install xlrd --user"
echo "installing" $xlrd
eval $xlrd