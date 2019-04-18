# 2excel
Программа создавалась для инженеров, чтобы создавать опись чертежей, хранящихся в текущем каталоге и подкаталогах в виде файлов .pdf.
Опись сохраняется в таблице Excel. 
Эта программа - пример парсинга, поиска файлов, работы с модулями Python, такими как openpyxl и PyPDF2.
Необходимые библиотеки: openpyxl, PyPDF2.
------------------------------------------------------------------------------------------------------------
The program was created for engineers to create an inventory of drawings stored in the current directory and subdirectories in the
form of .pdf files. The inventory is saved in an Excel spreadsheet.
This program is an example of parsing, file searching, working with Python modules such as openpyxl and PyPDF2.
Required libraries: openpyxl, PyPDF2.
------------------------------------------------------------------------------------------------------------
Сборка exe файла с помощью pyinstaller:/Building an exe file using pyinstaller:
pyinstaller -F 2excel.py

Рекомендуется также:
--exclude matplotlib --exclude pandas --exclude PyQt5 --exclude numpy --exclude PySide2