import sys
import time
import os

def runFusion():
    """global path"""
    projectDir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(
        os.path.abspath(__file__))

    print("start sleeping")
    time.sleep(10)
    with open(os.path.join(projectDir, "lollol.txt"),'w') as myfile:
        myfile.write("")

    print("end slipping")
    for i in range(1, len(sys.argv)):
        print(argv[i])


if __name__ == '__main__':
    runFusion()
