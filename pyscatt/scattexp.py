import sys
import comtypes.client

def print_use():
  print("Usage: scattexp <filename.scatt>")

def main(filename):
  if len(sys.argv) < 2:
    print_use()
  print("Loading file " + filename)
  scattdoc = comtypes.client.CreateObject("ScattDoc.ScattDocument")
  scattdoc.FileName = filename
  scattdoc.Load()
  if not scattdoc.Valid:
    print("Problem loading file")
    return 1
  return 0

if __name__ == '__main__':
  # Scatt Professional is only 32-bit and so only can be accessed from a 32-bit process.
  if sys.platform != "win32" or sys.maxsize > 2**32:
    print("This application needs to be run with a 32-bit python on Windows")
    sys.exit(1)
  if len(sys.argv) < 2:
    print_use()
    sys.exit(1)
  sys.exit(main(sys.argv[1]))
