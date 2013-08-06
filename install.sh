rm -rf deps
mkdir deps
cd deps

echo dateutil...
curl -s -L https://pypi.python.org/packages/source/p/python-dateutil/python-dateutil-2.1.tar.gz#md5=1534bb15cf311f07afaa3aacba1c028b | tar zx
cp -R python-dateutil-2.1/dateutil ../python/dateutil

echo Downloading xlrd dependency from Github...
rm -rf ../python/xlrd
curl -s -L https://github.com/python-excel/xlrd/tarball/master | tar zx
mv python-excel-xlrd-* python-excel-xlrd
cp -R python-excel-xlrd/xlrd ../python/xlrd

echo Downloading xlwt dependency from Github...
rm -rf ../python/xlwt
curl -s -L https://github.com/python-excel/xlwt/tarball/master | tar zx
mv python-excel-xlwt-* python-excel-xlwt
cp -R python-excel-xlwt/xlwt ../python/xlwt

echo Downloading XlsxWriter dependency from Github...
rm -rf ../python/xlsxwriter
curl -s -L https://github.com/jmcnamara/XlsxWriter/tarball/master | tar zx
mv jmcnamara-XlsxWriter-* XlsxWriter
cp -R XlsxWriter/xlsxwriter ../python/xlsxwriter

echo Downloading openpyxl dependency from Bitbucket...
rm -rf ../python/openpyxl
curl -s -L https://bitbucket.org/ericgazoni/openpyxl/get/default.tar.gz | tar zx
mv ericgazoni-openpyxl-* ericgazoni-openpyxl
cp -R ericgazoni-openpyxl/openpyxl ../python/openpyxl
