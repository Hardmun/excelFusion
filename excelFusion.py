from sys import argv

def runFusion():
    for i in range(1, len(argv)):
        print(argv[i])

if __name__ == '__main__':
    runFusion()