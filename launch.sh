#!/bin/bash

fullpath=$(realpath "$0")
directory=$(dirname "$fullpath")

echo -e "cd $directory"
cd $directory

"C:\Program Files\LibreOffice\program\python.exe" main.py