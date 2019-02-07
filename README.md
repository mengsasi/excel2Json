# excel2Json
excel转json，转换工具

### 使用bat批处理执行，需要安装python
excel使用的是xlrd库
import xlrd #http://pypi.python.org/pypi/xlrd

### excel 文件格式
第一行为json数据的key  
第二行为数据类型  
第三行是数据描述  
从第四行开始转换数据，转换之后为json数组

数组键为sheet的名称  
sheet名称中有‘noexport’时，不导出json数据  
excel列类型中为‘ignore’时，不导出此列

### simple and generatorAll
excel2Json导出文件夹下所有excel文件，每个sheet表单独一个json文件  
excel2Json_All将所有sheet导出到一个json文件中
