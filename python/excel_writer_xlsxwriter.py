import sys, json, datetime, xlsxwriter
from xlsxwriter.workbook import Workbook

def create_workbook(self, options = None):

  if options and "defaultDateFormat" in options:
    default_date_format = options["defaultDateFormat"]
  else:
    default_date_format = "yyyy-mm-dd"

  self.workbook = Workbook(self.filename, { "default_date_format": default_date_format, "constant_memory": True })
  self.dump_record("open", self.filename)

  if options and "properties" in options:
    self.workbook.set_properties(options["properties"])

def add_sheet(self, name = None):
  self.current_sheet = self.workbook.add_worksheet(name)

def write(self, row, col, data, format_name = None):
  format = self.formats[format_name] if format_name != None else None

  if isinstance(data, list):
    row_index = row
    col_index = col
    for v1 in data:
      if isinstance(v1, list):
        col_index = col
        for v2 in v1:
           self.current_sheet.write(row_index, col_index, v2, format)
           col_index += 1
        row_index += 1
      else:
        self.current_sheet.write(row_index, col_index, v1, format)
        col_index += 1
  else:
    self.current_sheet.write(row, col, data, format)

def format(self, name, properties):

  format = self.workbook.add_format()

  if "font" in properties:
    font = properties["font"]
    if "name" in font:
      format.set_font_name(font["name"])
    if "size" in font:
      format.set_font_size(int(font["size"]))
    if "color" in font:
      format.set_font_color(font["color"])
    if font.get("bold", False):
      format.set_bold()
    if font.get("italic", False):
      format.set_italic()
    if "underline" in font:
      if font["underline"] == True or font["underline"] == "single":
        format.set_underline = 1
      elif font["underline"] == "double":
        format.set_underline = 2
      elif font["underline"] == "single accounting":
        format.set_underline = 33
      elif font["underline"] == "double accounting":
        format.set_underline = 34
    if font.get("strikeout", False):
      format.set_strikeout()
    if font.get("superscript", False):
      format.set_font_script(1)
    elif font.get("subscript", False):
      format.set_font_script(2)

  if "numberFormat" in properties:
    format.set_num_format(properties["numberFormat"])
  if "locked" in properties:
    format.set_locked(properties["locked"])
  if properties.get("hidden", False):
    format.set_hidden();
  if "alignment" in properties:
    aligns = properties["alignment"].split()
    for alignment in aligns:
      alignment = "vcenter" if alignment == "middle" else alignment
      format.set_align(alignment)
  if properties.get("textWrap", False):
    format.set_text_wrap()
  if "rotation" in properties:
    format.set_rotation(int(properties["rotation"]))
  if "indent" in properties:
    format.set_indent(int(properties["indent"]))
  if properties.get("shrinkToFit", False):
    format.set_shrink()
  if properties.get("justifyLastText", False):
    format.set_text_justlast()

  if "fill" in properties:
    fill = properties["fill"]
    if isinstance(fill, basestring):
      format.set_pattern(1)
      format.set_bg_color(fill)
    else:
      if "pattern" in fill:
        format.set_pattern(int(fill["pattern"]))
      else:
        format.set_pattern(1)
      if "color" in fill:
        fill["backgroundColor"] = fill["color"]
      if "backgroundColor" in fill:
        format.set_bg_color(fill["backgroundColor"])
      if "foregroundColor" in fill:
        format.set_fg_color(fill["foregroundColor"])

  if "borders" in properties:
    borders = properties["borders"]
    if isinstance(borders, int):
      format.set_border(borders)
    else:
      if "style" in borders:
        format.set_border(borders["style"])
      if "color" in borders:
        format.set_border_color(borders["color"])
      for border_pos in ["top", "left", "right", "bottom"]:
        if border_pos in borders:
          value = borders[border_pos]
          set_style = getattr(format, "set_" + border_pos)
          if isinstance(value, int):
            set_style(value)
          else:
            if "style" in value:
              set_style(value["style"])
            if "color" in value:
              set_color = getattr(format, "set_" + border_pos + "_color")
              set_color(value["color"])

  self.formats[name] = format

def merge_range(self, range, data, format_name = None):
  self.current_sheet.merge_range(range, data, format_name)

def activate_sheet(self, id):
  for index, sheet in enumerate(self.workbook.worksheets()):
    if index == id or sheet.get_name() == id:
      self.current_sheet = sheet
      break

def set_sheet_settings(self, id, settings = None):

  for index, sheet in enumerate(self.workbook.worksheets()):
    if index == id or sheet.get_name() == id:
      self.current_sheet = sheet
      break

  if settings != None:
    if settings.get("hidden", False):
      self.current_sheet.hide()
    if settings.get("activated", False):
      self.current_sheet.activate()
    if settings.get("selected", False):
      self.current_sheet.select()
    if settings.get("rightToLeft", False):
      self.current_sheet.right_to_left()
    if settings.get("hideZeroValues", False):
      self.current_sheet.hide_zero()
    if "selection" in settings:
      selection = settings["selection"]
      if isinstance(selection, basestring):
        self.current_sheet.set_selection(selection)
      else:
        self.current_sheet.set_selection(int(selection["top"]), int(selection["left"]), int(selection["bottom"]), int(selection["right"]))

def set_row(self, index, settings):
  self.current_sheet.set_row(index, settings.get("height"), settings.get("format"), settings.get("options"))

def set_column(self, index, settings):
  self.current_sheet.set_column(index, index, settings.get("width"), settings.get("format"), settings.get("options"))

def close(self):
  self.workbook.close()
  self.dump_record("close")
