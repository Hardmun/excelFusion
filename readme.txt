1) Read from file
# with open(os.path.join(projectDir, "file.json"), "r", encoding="UTF-8") as filereader:
#     files_to_read = filereader.read()

2) parameter
"{ 'settings': { 'files': [ 'Ярославль', 'Челябинск', 'сеть', 'Client Summary' ], 'uuid': 'test_files' } }"

3) compile
pyinstaller -F -w --icon=excel.ico excelFusion.py