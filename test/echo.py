import sys, json

for line in sys.stdin:
  message = json.loads(line)
  print json.dumps(message)
