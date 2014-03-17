#!/bin/bash

rm -rf .deps
mkdir .deps
cd .deps

echo Downloading xlrd dependency from Github
git clone https://github.com/python-excel/xlrd xlrd
rm -rf ../python/xlrd
cp -R xlrd/xlrd ../python/xlrd

echo Downloading xlwt dependency from Github
git clone https://github.com/python-excel/xlwt xlwt
rm -rf ../python/xlwt
cp -R xlwt/xlwt ../python/xlwt

echo Downloading XlsxWriter dependency from Github
git clone https://github.com/jmcnamara/XlsxWriter xlsxwriter
rm -rf ../python/xlsxwriter
cp -R xlsxwriter/xlsxwriter ../python/xlsxwriter

cd ..
rm -rf .deps
