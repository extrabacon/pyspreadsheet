import sys, json, traceback, datetime, glob, xlrd
from xlrd import open_workbook, cellname, xldate_as_tuple, error_text_from_code

def dump_record(record_type, values):
  print(json.dumps([record_type, values]));

def parse_cell_value(sheet, cell):
  if cell.ctype == xlrd.XL_CELL_DATE:
    year, month, day, hour, minute, second = xldate_as_tuple(cell.value, sheet.book.datemode)
    return ['date', year, month, day, hour, minute, second]
  elif cell.ctype == xlrd.XL_CELL_ERROR:
    return ['error', error_text_from_code[cell.value]]
  elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
    return False if cell.value == 0 else True
  elif cell.ctype == xlrd.XL_CELL_EMPTY:
    return None
  return cell.value

def dump_sheet(sheet, sheet_index, max_rows):
  dump_record("s", {
    "index": sheet_index,
    "name": sheet.name,
    "rows": sheet.nrows,
    "columns": sheet.ncols,
    "visibility": sheet.visibility
  })
  for rowx in range(max_rows or sheet.nrows):
    for colx in range(sheet.ncols):
      cell = sheet.cell(rowx, colx)
      dump_record("c", [rowx, colx, cellname(rowx, colx), parse_cell_value(sheet, cell)])

def main(cmd_args):
  import optparse
  usage = "\n%prog [options] [file1] [file2] ..."
  oparser = optparse.OptionParser(usage)
  oparser.add_option(
    "-m", "--meta",
    dest = "iterate_sheets",
    action = "store_false",
    default = True,
    help = "dumps only the workbook record, does not load any worksheet")
  oparser.add_option(
    "-s", "--sheet",
    dest = "sheets",
    action = "append",
    help = "names of the sheets to load - if omitted, all sheets are loaded")
  oparser.add_option(
    "-r", "--rows",
    dest = "max_rows",
    default = None,
    action = "store",
    type = "int",
    help = "maximum number of rows to load")
  options, args = oparser.parse_args(cmd_args)

  # loop on all input files
  for file in args:
    try:
      wb = open_workbook(filename=file, on_demand=True)
      sheet_names = wb.sheet_names()

      dump_record("w", {
        "file": file,
        "sheets": sheet_names,
        "user": wb.user_name
      })

      if options.iterate_sheets:
        if options.sheets:
          for sheet_to_load in options.sheets:
            try:
              sheet_name = sheet_to_load
              if sheet_to_load.isdigit():
                sheet = wb.sheet_by_index(int(sheet_to_load))
                sheet_name = sheet.name
              else:
                sheet = wb.sheet_by_name(sheet_to_load)
              dump_sheet(sheet, sheet_names.index(sheet_name.decode('utf-8')), options.max_rows)
              wb.unload_sheet(sheet_name)
            except:
              dump_record("error", {
                "id": "load_sheet_failed",
                "file": file,
                "sheet": sheet_name,
                "traceback": traceback.format_exc()
              })
        else:
          for sheet_index in range(len(sheet_names)):
            try:
              sheet = wb.sheet_by_index(sheet_index)
              dump_sheet(sheet, sheet_index, options.max_rows)
              wb.unload_sheet(sheet_index)
            except:
              dump_record("error", {
                "id": "load_sheet_failed",
                "file": file,
                "sheet": sheet_index,
                "traceback": traceback.format_exc()
              })
    except:
      dump_record("error", {
        "id": "open_workbook_failed",
        "file": file,
        "traceback": traceback.format_exc()
      })

  sys.exit()

main(sys.argv[1:])
