import sys, json, datetime, xlwt

def create_workbook(self, filename, options):
  self.filename = filename
  self.workbook = xlwt.Workbook()
  #if options and options["properties"]:
  #  self.workbook.set_properties(options["properties"])

def add_sheet(self, name = None):
  self.current_sheet = self.workbook.add_sheet(name)

def write(self, row, col, data):
  self.current_sheet.write(row, col, data)

def close(self):
  self.workbook.save(self.filename)
