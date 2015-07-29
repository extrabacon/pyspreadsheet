import sys, datetime, json, uuid, traceback
import excel_writer_xlsxwriter, excel_writer_xlwt

def dump_record(record_type, values = None):
  if values != None:
    print(json.dumps([record_type, values], default = json_extended_handler))
  else:
    print(json.dumps([record_type], default = json_extended_handler))

def json_extended_handler(obj):
  if hasattr(obj, 'isoformat'):
    return obj.isoformat()
  return obj

# extended JSON parser for handling dates
def json_extended_parser(dct):
  for k, v in dct.items():
    if k == "$date":
      return datetime.datetime.fromtimestamp(v / 1000)
  return dct

class SpreadsheetWriter:
  def __init__(self, filename, module):
    self.filename = filename
    self.workbook = None
    self.current_sheet = None
    self.current_cell = None
    self.formats = dict()
    # inject module functions into writer instance
    self.__dict__.update(module.__dict__)
    self.dump_record = dump_record

def main(cmd_args):
  import optparse
  usage = "\n%prog -m [module] -o [output]"
  oparser = optparse.OptionParser(usage)
  oparser.add_option(
    "-m", "--module",
    dest = "module_name",
    action = "store",
    default = "xlsxwriter",
    help = "the name of the writer module to load")
  oparser.add_option(
    "-o", "--output",
    dest = "output_path",
    action = "store",
    help = "the path of the output file to write")
  options, args = oparser.parse_args(cmd_args)

  # load the specified module and use it to create a writer
  module = sys.modules["excel_writer_" + options.module_name]
  writer = SpreadsheetWriter(options.output_path, module)

  # read JSON commands from stdin and forward them to the writer
  for line in sys.stdin:
    try:
      command = json.loads(line, object_hook = json_extended_parser)
      method_name = command[0]
      args = command[1:]

      if getattr(writer, method_name, None)!=None:
        # Try find method at writer class
        getattr(writer, method_name)(writer, *args)
      elif getattr(writer.current_sheet, method_name, None)!=None:
        # try find original method at active sheet
        getattr(writer.current_sheet, method_name)(*args)
      else:
        dump_record("unknown method", {
          "method_name": method_name,
          "args": args
        })
    except:
      dump_record("error", {
        "method_name": method_name,
        "args": args,
        "traceback": traceback.format_exc()
      })

  writer.close(writer)
  sys.exit()

main(sys.argv[1:])
