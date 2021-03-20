import sys
import comtypes.client


def print_use():
    print("Usage: scattexp <filename.scatt>")


def main(filename):
    print("Loading file " + filename)
    try:
        scattdoc = comtypes.client.CreateObject("ScattDoc.ScattDocument")
    except OSError:
        print("Cannot find Scatt Professional, please check it is installed")
        return 1
    scattdoc.FileName = filename
    try:
        scattdoc.Load()
    except comtypes.COMError:
        print("Problem loading file")
        return 1
    print(scattdoc.ShooterName)
    return 0


if __name__ == '__main__':
    # Scatt Professional is only 32-bit and so only can be accessed from a 32-bit process.
    if sys.platform != "win32" or sys.maxsize > 2 ** 32:
        print("This application needs to be run with a 32-bit python on Windows")
        sys.exit(1)
    if len(sys.argv) < 2:
        print_use()
        sys.exit(1)
    sys.exit(main(sys.argv[1]))
