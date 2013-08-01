import sys, json, dateutil.parser, uuid, traceback, xlsxwriter_handlers, xlwt_handlers
from dateutil import tz

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
      return dateutil.parser.parse(v).replace(tzinfo = None)
  return dct

class SpreadsheetWriter:
  def __init__(self, handlers):
    self.id = uuid.uuid4();
    self.filename = ""
    self.workbook = None
    self.current_sheet = None
    self.current_cell = None
    self.formats = dict()
    self.__dict__.update(handlers.__dict__)
    self.dump_record = dump_record

writer = SpreadsheetWriter(xlsxwriter_handlers)
module_name = "xlsxwriter"

for line in sys.stdin:

  directive = json.loads(line, object_hook = json_extended_parser)
  cmd = directive[0]
  args = directive[1:]

  try:
    if cmd == "setup":
      module_name = args[0]
      handler_module = sys.modules[module_name + "_handlers"]
      writer = SpreadsheetWriter(handler_module)
    else:
      getattr(writer, cmd)(writer, *args)
  except:
    e0, e1 = sys.exc_info()[:2]
    dump_record("err", {
      "command": cmd,
      "arguments": args,
      "module": module_name,
      "file": writer.filename,
      "error": traceback.format_exc()
    })

writer.close(writer)
dump_record("close", writer.filename)
