import sys, json, datetime, xlwt
from xlwt import *

def create_workbook(self, options = None):
  self.workbook = Workbook()
  self.sheet_count = 0
  self.dump_record("open", self.filename)

  if options and "properties" in options:
    # TODO: add support for more properties, xlwt has a ton of them
    prop = options["properties"]
    if "owner" in prop:
      self.workbook.owner = prop["owner"]

def add_sheet(self, name = None):
  if name == None:
    name = "Sheet" + str(self.sheet_count + 1)
  self.current_sheet = self.workbook.add_sheet(name)
  self.sheet_count += 1

def write(self, row, col, data, format_name = None):
  style = self.formats[format_name] if format_name != None else None

  def write_one(row, col, val):
    if style != None:
      self.current_sheet.write(row, col, val, style)
    else:
      self.current_sheet.write(row, col, val)

  if isinstance(data, list):
    row_index = row
    col_index = col
    for v1 in data:
      if isinstance(v1, list):
        col_index = col
        for v2 in v1:
           write_one(row_index, col_index, v2)
           col_index += 1
        row_index += 1
      else:
        write_one(row_index, col_index, v1)
        col_index += 1
  else:
    write_one(row, col, data)

def format(self, name, properties):

  style = XFStyle()

  if "font" in properties:
    style.font = Font()
    font = properties["font"]
    if "name" in font:
      style.font.name = font["name"]
    if "size" in font:
      style.font.size = font["size"]
    if "color" in font:
      # TODO: need to convert color codes
      style.font.colour_index = font["color"]
    if font.get("bold", False):
      style.font.bold = True
    if font.get("italic", False):
      style.font.italic = True
    if "underline" in font:
      if font["underline"] == True or font["underline"] == "single":
        style.font.underline = Font.UNDERLINE_SINGLE
      elif font["underline"] == "double":
        style.font.underline = Font.UNDERLINE_DOUBLE
      elif font["underline"] == "single accounting":
        style.font.underline = Font.UNDERLINE_SINGLE_ACC
      elif font["underline"] == "double accounting":
        style.font.underline = Font.UNDERLINE_DOUBLE_ACC
    if font.get("strikeout", False):
      style.font.struck_out = True
    if font.get("superscript", False):
      style.font.escapement = Font.ESCAPEMENT_SUPERSCRIPT
    elif font.get("subscript", False):
      style.font.escapement = Font.ESCAPEMENT_SUBSCRIPT

  if "numberFormat" in properties:
    style.num_format_str = properties["numberFormat"]

  # TODO: locked
  # TODO: hidden
  # TODO: alignment
  # TODO: textWrap
  # TODO: rotation
  # TODO: indent
  # TODO: shrinkToFit
  # TODO: justifyLastText
  # TODO: fill
  # TODO: borders

  self.formats[name] = style

def activate_sheet(self, id):
  # TODO: implement
  raise Exception("not implemented")

def set_sheet_settings(self, id, settings = None):
  # TODO: implement
  raise Exception("not implemented")

def set_row(self, index, settings):
  # TODO: implement
  raise Exception("not implemented")

def set_column(self, index, settings):
  # TODO: implement
  raise Exception("not implemented")

def close(self):
  self.workbook.save(self.filename)
