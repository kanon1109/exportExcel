# exportExcel
调用某个目录下所有*.xlsm文件的宏函数。用于批量导出xlsm成xml格式。

exportExcel.py ———— 代码都在这里。

setup.py ———— py2exe的代码在这里 用于打包exe文件。

/cfg/ ———— 存放excel的文件夹，必须在cfg目录下，除非修改exportExcel.py代码 此目录下的excel文件（必须是xlsm文件）都是已经创建过宏"export()"的文件。
(此目录下export.xlsm文件功能和exportExcel.py一样是一个脚本调用除自己以外所有的excel中的宏)

     
