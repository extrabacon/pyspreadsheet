#!/bin/bash

rm -rf .deps
mkdir .deps
cd .deps

echo Downloading xlrd dependency from Github
URL=https://github.com/python-excel/xlrd/tarball/master
(wget -O - $URL || curl -s -L $URL) | tar zx
rm -rf ../python/xlrd
cp -R python-excel-xlrd-*/xlrd ../python/xlrd

echo Downloading xlwt dependency from Github
URL=https://github.com/python-excel/xlwt/tarball/master
(wget -O - $URL || curl -s -L $URL) | tar zx
rm -rf ../python/xlwt
cp -R python-excel-xlwt-*/xlwt ../python/xlwt

echo Downloading XlsxWriter dependency from Github
URL=https://github.com/jmcnamara/XlsxWriter/tarball/master
(wget -O - $URL || curl -s -L $URL) | tar zx
rm -rf ../python/xlsxwriter
cp -R jmcnamara-XlsxWriter-*/xlsxwriter ../python/xlsxwriter

cd ..
rm -rf .deps
