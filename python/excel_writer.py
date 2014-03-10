import sys, datetime, json, uuid, traceback, writer_module_xlsxwriter, writer_module_xlwt

def dump_record(record_type, values):
  print(json.dumps([record_type, values], default = json_extended_handler))

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
    self.__dict__.update(module.__dict__)
    self.dump_record = dump_record

writer = None
module_name = ""
filename = ""

for line in sys.stdin:
  directive = json.loads(line, object_hook = json_extended_parser)
  cmd = directive[0]
  args = directive[1:]

  try:
    if cmd == "open":
      module_name = args[0]
      filename = args[1]
      module = sys.modules["writer_module_" + module_name]
      writer = SpreadsheetWriter(filename, module)
    else:
      getattr(writer, cmd)(writer, *args)
  except:
    e0, e1 = sys.exc_info()[:2]
    dump_record("err", {
      "command": cmd,
      "arguments": args,
      "module": module_name,
      "file": filename,
      "error": traceback.format_exc()
    })

writer.close(writer)
dump_record("close", writer.filename)
