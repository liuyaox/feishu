# feishu

飞书文档API: 目前主要是SpreadSheet，读写飞书表格

## 环境配置

```shell
cd <YOUR_DIRECTORY>
git clone https://github.com/liuyaox/feishu.git
cd feishu
pip install -r requirements.txt
```

环境变量配置： 

```shell
export PYTHONPATH=$PYTHONPATH:<YOUR_DIRECTORY>/feishu

# 用于认证(详见identification.py)
export FEISHU_APP_ID="xxx"
export FEISHU_APP_SECRET="xxx"
export FEISHU_REDIRECT_URI="xxx"
export FEISHU_CONFIG_KEY="yao.liu"    # 修改为你自己的名字
```

以上xxx具体取值，详见文档 [feishu环境变量配置](https://rg975ojk5z.feishu.cn/docx/BLXrdah64oylNBxHieYcXKndnkd)

## 初始化

主要是获取自己飞书账号的授权码code，并且以此code获取自己的一系列xxx_token。

以下2种情况，必须初始化:
- 第1次使用本项目
- 上次使用距今时间大于30天（注意，不是第1次距今。每使用1次，会自动更新token，有效期是30天）

请在交互式界面（比如ipython或notebook ）执行以下命令：
```python
from feishu import Identification
idt = Identification(get_new_code=True)
```
此时会输出1个url，比如：

`请在浏览器里打开这个URL，飞书授权后获得code，随后使用code进行初始化：https://open.feishu.cn/open-apis/authen/v1/index?redirect_uri=https%3A%2F%2Fopen.feishu.cn%2Fdocument%2Fliuyao&app_id=xxx&state=test`

在浏览器打开这个URL，飞书授权后，浏览器链接里可获得code(`<YOUR_CODE>`)，比如：

`https://open.feishu.cn/document/liuyao?code=<YOUR_CODE>&state=test`

继续在交互式界面执行以下命令：

```python
code = '<YOUR_CODE>'
idt.init_with_code(code)
```

输出 `Get User Info Successfully!` 表示初始化成功了。

## 正式使用

正式使用前要有一个飞书表格，如下所示，红框里分别是**spreadsheet_token**和**sheet_id**

![img.png](./images/img.png)

### 简单demo

```python
from feishu import SpreadSheet

spsh = SpreadSheet()
# 注意，以下2行中，sheet可以是sheet_id（建议这个）, sheet_title 或 sheet_index
df = spsh.read_sheet(spreadsheet_token='xxx1', sheet='xxx', cell_start='B1', cell_end='F501') # 读取sheet，范围是B1:F501
spsh.write_df(df, spreadsheet_token='xxx2', sheet='xxx', cell_start='D1')                     # 写入sheet，从D1开始写，若cell_start若是A1，可省略

# 操作别的文档，只需修改spreadsheet_token即可，不需要重新实例化SpreadSheet
df = spsh.read_sheet(spreadsheet_token='xxx3', sheet='xxx', cell_start='B1', cell_end='F501')

# 也可在SpreadSheet实例化时指定spreadsheet_token，尤其适用于对同一文档多次操作时
spsh = SpreadSheet(spreadsheet_token='xxx4')
df = spsh.read_sheet(sheet='xxx', cell_start='B1', cell_end='F501')      # 读取sheet
spsh.write_df(df, sheet='xxx', cell_start='D1')                          # 写入sheet
```

### 进阶demo

```python
from feishu import SpreadSheet
spsh = SpreadSheet()


# demo1: 读取sheet的范围，第一行不是列名，需要指定列名
df = spsh.read_sheet(spreadsheet_token='xxx', sheet='xxx', cell_start='B2', cell_end='C501', has_cols=False, col_names=['col1', 'col2'])     # 读取范围内(B2:C501)第1行不是列名，需要指定列名col_names


# demo2: 连续写入同一个sheet
cell_start = 'A1'
for df in [df1, df2, df3]:
    cell_start = spsh.write_df(df, spreadsheet_token='xxx', sheet='xxx', cell_start=cell_start)
    # 若有需要，可修改cell_start，比如每隔1行写入一份数据，则修改cell_start: A200 -> A201
    # 其实下面这行也行（API会自行判断可以写入的第1个空行），但返回值一直是第1次写入时的cell_start，不优雅，无法准确得知下一行可以写入的行号
    # spsh.write_df(df, spreadsheet_token='xxx', sheet='xxx')

    
# demo3: 图片写入sheet
image_paths = ['test1.png', 'test2.png', 'test3.png']
spsh.write_image(image_paths, sheet='dzwtzZ', cell_start='B2')              # 写入一列：B2到B4
spsh.write_image(image_paths, sheet='dzwtzZ', cell_start='F5', axis='row')  # 写入一行：F5到F7
```

### 注意事项
- 写入sheet时，df必须是DataFrame类型，若只有一列，不要写`df['col1']`，而是写`df[['col1']]`
- 写入sheet时，df的cell数值类型不能是dict, list等复杂数据类型，若想写入，可以转化为str，比如`df['dic']=df['dic'].map(str)`
- 读写文档和用户认证过程中，都会输出详细信息，若想控制，可以配置环境变量`FEISHU_VERBOSE`
    > `export FEISHU_VERBOSE='all'`表示都输出详细信息
    >
    > `export FEISHU_VERBOSE='spreadsheet'`表示只输出读写文档过程中的详细信息
    >
    > `export FEISHU_VERBOSE='identification'`表示只输出用户认证过程中的详细信息
    >
    > `export FEISHU_VERBOSE='none'`表示都不输出详细信息


## 写在最后

有任何问题，可随时联系作者。
