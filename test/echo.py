import sys, json

# simple JSON echo script
for line in sys.stdin:
  message = json.loads(line)
  print json.dumps(message)
