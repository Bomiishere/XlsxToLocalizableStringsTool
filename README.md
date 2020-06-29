#### Features

- Convert Xlsx into Output.strings

#### Requirements

* Python 3
* pip install

#### Modules
| Xlsx  | Regular Expression  | Date
| :---: |:---------------:|:----:|
| xlsxwriter | re | datetime

### Usage
##### Required
###### 1. Edit Setup in xlsx_convert.py 
```
# 讀取檔案
read_xlsx_file_name = 'result.xlsx'
# 專案內表示文字之欄位
column_expression = 0 
# 欲轉換語系之欄位
column_translate = 1
```
###### 2. Execute convert_xlsx.py, generate result.xlsx
```
python3 xlsx_convert.py
```
##### Optional
###### 修改 output file 註解
```
header_1_export_file_name = 'Output.strings'
header_2_from_where = 'from xlsx_convert.py' 
header_3_created_by = 'auto generated'
header_4_create_date = datetime.datetime.now().strftime('%Y/%-m/%-d')
header_5_create_year = datetime.datetime
header_6_copyright = 'JohnsonTechInc.'
```

#### Tips
* python3 install, [click me for reference](https://stringpiggy.hpd.io/mac-osx-python3-dual-install/)

* install modules
```bash
python3 -m pip install --upgrade xlsxwriter
```

* install jupyter
```
pip3 install jupyter
```

