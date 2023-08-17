import os
import requests
import logging
import re
from tqdm import tqdm
import pandas as pd
import cv2

from .identification import Identification
from .feishu_util import get_headers

logger = logging.getLogger(__name__)


# 0. Common
ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
PATTERN = re.compile(r'([a-zA-Z]+)(\d+)')   # 拆分字母和数字
FEISHU_VERBOSE = os.environ.get('FEISHU_VERBOSE', 'spreadsheet')


def xy_to_cell(row_index, col_index):
    """
    坐标(x, y)转化为Cell编号，比如：(5, 1) -> 'B6'
    :param row_index: 数组的行，从0开始
    :param col_index: 数组的列，从0开始
    :return:
    """
    temp = col_index + 1
    res = []
    while temp > 0:
        temp -= 1
        res.append(ALPHABET[temp % 26])
        temp //= 26
    res.reverse()
    col = ''.join(res)
    return f'{col}{row_index + 1}'


def cell_to_xy(cell):
    """
    cell编号转化为坐标(x, y)，比如：'B6' -> (5, 1)
    :param cell:
    :return:
    """
    try:
        col, row = re.findall(PATTERN, cell)[0]
        col_index = -1
        for i, letter in enumerate(col[::-1]):
            col_index += (ALPHABET.find(letter.upper()) + 1) * (26 ** i)
        row_index = int(row) - 1
        return row_index, col_index
    except Exception as e:
        logger.error(f'解析单元格错误，cell={cell}')
        raise e


class SpreadSheet(object):
    """
    操作SpreadSheet，暂时只需关注write_df(df写入sheet)和read_sheet(读取sheet为df)这2个API
    """
    def __init__(self, spreadsheet_token=None, user_access_token=None):
        """
        若1个文档要操作多次，建议为其专门初始化一个实例(在初始化时指定spreadsheet_token)
        若有多个文档，每个文档只操作一两次，建议先初始化1个公共实例，在操作具体每个文档时再指定spreadsheet_token
        :param spreadsheet_token:
        :param user_access_token:
        """
        if user_access_token is None:
            self.idt = Identification()
            user_access_token = self.idt.user_access_token
        self.user_access_token = user_access_token
        self.headers = get_headers(self.user_access_token)
        if spreadsheet_token:
            self._set_spreadsheet_token(spreadsheet_token)

    def _set_spreadsheet_token(self, spreadsheet_token):
        """
        设置或修改spreadsheet_token
        :param spreadsheet_token:
        :return:
        """
        self.spreadsheet_token = spreadsheet_token
        self.api_url_v2 = f'https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}'
        self.api_url_v3 = f'https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}'
        self._update_meta_info()

    def _update_meta_info(self):
        """
        获取并更新表格元数据
        获取表格信息:   https://open.feishu.cn/document/server-docs/docs/sheets-v3/spreadsheet/get
        获取sheet信息: https://open.feishu.cn/document/server-docs/docs/sheets-v3/spreadsheet-sheet/query
        update: 20230725
        """
        url = self.api_url_v3
        resp = requests.get(url, headers=self.headers).json()
        if resp['code'] == 0:
            spreadsheet = resp['data']['spreadsheet']
            self.title = spreadsheet['title']
            self.owner_id = spreadsheet['owner_id']
            self.spreadsheet_url = spreadsheet['url']
        else:
            logger.error(f'Get SpreadSheet Meta Info Failed: {resp}')
            return

        url = f'{self.api_url_v3}/sheets/query'
        resp = requests.get(url, headers=self.headers).json()
        if resp['code'] == 0:
            sheets = resp['data']['sheets']
            self.sheets = {x['index']: x for x in sheets}
            self.sheet_index2id = {val['index']: val['sheet_id'] for key, val in self.sheets.items()}
            self.sheet_title2id = {val['title']: val['sheet_id'] for key, val in self.sheets.items()}
            self.sheet_id2index = {val: key for key, val in self.sheet_index2id.items()}
        else:
            logger.error(f'Get SpreadSheet Sheets Meta Info Failed: {resp}')

    def _change_title(self, title):
        """
        修改表格title
        doc: https://open.feishu.cn/document/server-docs/docs/sheets-v3/spreadsheet/patch
        update: 20230726
        :param title:
        :return:
        """
        url = self.api_url_v3
        body = {
            'title': title
        }
        resp = requests.patch(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            self.title = title
        return resp

    def create_spreadsheet(self, folder_token, title=None):
        """
        创建表格
        doc: https://open.feishu.cn/document/server-docs/docs/sheets-v3/spreadsheet/create
        update: 20230726
        :param folder_token:
        :param title:
        :return:
        """
        url = 'https://open.feishu.cn/open-apis/sheets/v3/spreadsheets'
        body = {
            'title': title,
            'folder_token': folder_token
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            spreadsheet = resp['data']['spreadsheet']
            self.folder_token = spreadsheet['folder_token']
            spreadsheet_token = spreadsheet['spreadsheet_token']
            spreadsheet_url = spreadsheet['url']
            self._set_spreadsheet_token(spreadsheet_token)
            return {
                'spreadsheet_token': spreadsheet_token,
                'spreadsheet_url': spreadsheet_url,
                'sheet_token': self.sheet_index2id[0]
            }
        else:
            logger.error(f'Add SpreadSheet Failed: {resp}')

    def _query_sheet(self, sheet):
        """
        查询sheet
        doc: https://open.feishu.cn/document/server-docs/docs/sheets-v3/spreadsheet-sheet/get
        update: 20230726
        :param sheet:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        return self.sheets[sheet_id]

    def _add_sheet(self, title, index=-1):
        """
        添加sheet
        doc: https://open.feishu.cn/document/server-docs/docs/sheets-v3/spreadsheet-sheet/operate-sheets
        update: 20230726
        :param title:
        :param index: 0表示添加为第1个，1表示添加为第2个，……，-2表示添加为倒数第2个，-1表示添加为最后1个
        :return:
        """
        url = f'{self.api_url_v2}/sheets_batch_update'      # 其他url是v3，它是v2
        body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': title,
                        'index': len(self.sheets) + 1 + index if index < 0 else index
                    }
                }
            }]
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            self._update_meta_info()
            properties = list(resp['data']['replies'][0].values())[0]['properties']
            sheet_id, title, index = properties['sheetId'], properties['title'], properties['index']
            logger.info(f'Add Sheet Successfully: index={index}, sheet_id={sheet_id}, title={title}')
            return sheet_id, index
        else:
            logger.error(f'Add Sheet Failed: {resp}')
            return None, None

    def _copy_sheet(self, title, sheet=0):
        """
        复制sheet
        doc: https://open.feishu.cn/document/server-docs/docs/sheets-v3/spreadsheet-sheet/operate-sheets
        update: 20230726
        :param title:
        :param sheet:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url_v2}/sheets_batch_update'      # 其他url是v3，它是v2
        body = {
            'requests': [{
                'copySheet': {
                    'source': {
                        'sheetId': sheet_id
                    },
                    'destination': {
                        'title': title
                    }
                }
            }]
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            self._update_meta_info()
            properties = list(resp['data']['replies'][0].values())[0]['properties']
            sheet_id, title, index = properties['sheetId'], properties['title'], properties['index']
            logger.info(f'Copy Sheet Successfully: index={index}, sheet_id={sheet_id}, title={title}')
            return sheet_id, index
        else:
            logger.error(f'Copy Sheet Failed: {resp}')
            return None, None

    def _delete_sheet(self, sheet=0):
        """
        删除sheet
        doc: https://open.feishu.cn/document/server-docs/docs/sheets-v3/spreadsheet-sheet/operate-sheets
        update: 20230726
        :param sheet:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url_v2}/sheets_batch_update'      # 其他url是v3，它是v2
        body = {
            'requests': [{
                'deleteSheet': {
                    'sheetId': sheet_id
                }
            }]
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            info = list(resp['data']['replies'][0].values())[0]
            result, sheet_id = info['result'], info['sheetId']
            if result:
                self._update_meta_info()
                logger.info(f'Delete Sheet Successfully: sheet_id={sheet_id}')
        else:
            logger.error(f'Delete Sheet Failed: {resp}')

    def _change_sheet(self, sheet, title=None, index=None, hidden=None, lock=None, users=None):
        """
        更新sheet属性：更新标题，移动index，隐藏sheet，锁定sheet
        doc: https://open.feishu.cn/document/server-docs/docs/sheets-v3/spreadsheet-sheet/update-sheet-properties
        update: 20230726
        :param sheet:
        :param title:
        :param index:
        :param hidden:
        :param lock:
        :param users:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url_v2}/sheets_batch_update'      # 其他url是v3，它是v2
        properties = {'sheetId': sheet_id}
        if title:
            properties.update({'title': title})
        if index:
            properties.update({'index': index})
        if hidden:
            properties.update({'hidden': hidden})
        if lock:
            properties.update({'protect': {'lock': lock}})
            if users:
                properties.update({'protect': {'lock': lock, 'userIDs': users}})
        body = {
            'requests': [{
                'updateSheet': {
                    'properties': properties
                }
            }]
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            self._update_meta_info()
            logger.info(f'Change Sheet Meta Info Successfully: sheet_id={sheet_id}, title={title}, index={index}, '
                        f'hidden={hidden}, lock={lock}, users={users}')
        else:
            logger.error(f'Change Sheet Meta Info Failed: {resp}')

    def _prepend_data(self, cell_start, cell_end, values, sheet=0, update=True):
        """
        在范围range(cell_start到cell_end)内，前插数据：其他数据下移(并非右移)，相当于excel中在range上界处向上插入n行
        doc: https://open.feishu.cn/document/server-docs/docs/sheets-v3/data-operation/prepend-data
        update: 20230727
        :param cell_start:
        :param cell_end:
        :param values:
        :param sheet:
        :param update:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url_v2}/values_prepend'
        body = {
            'valueRange': {
                'range': f'{sheet_id}!{cell_start}:{cell_end}',
                'values': values
            }
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            if update:
                self._update_meta_info()
            data = resp['data']
            table_range, revision, updates = data['tableRange'], data['revision'], data['updates']
            range, cells = updates['updatedRange'], updates['updatedCells']
            rows, columns = updates['updatedRows'], updates['updatedColumns']
            logger.info(f'Prepend Data Successfully: range={range}, {rows} rows, {columns} columns, {cells} cells')
            return range
        else:
            logger.error(f'Prepend Data Failed: {resp}')

    def _append_data(self, cell_start, cell_end, values, sheet=0, option='OVERWRITE', update=True):
        """
        在范围range内或后，追加数据：从range起始行列开始向下寻找第1个空白位置，向下写入数据，相当于excel中在第1个空行处粘贴n行（覆盖或插入）
        doc: https://open.feishu.cn/document/server-docs/docs/sheets-v3/data-operation/append-data
        update: 20230727
        :param cell_start:
        :param cell_end:
        :param values:
        :param sheet:
        :param option: OVERWRITE会直接覆盖下面的数据，INSERT_ROWS会先在第1个空白位置后插入足够行后再写入数据，不会覆盖下面已有的数据。
        :param update:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url_v2}/values_append'
        params = {
            'insertDataOption': option,
        }
        body = {
            'valueRange': {
                'range': f'{sheet_id}!{cell_start}:{cell_end}',
                'values': values
            }
        }
        resp = requests.post(url, params=params, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            if update:
                self._update_meta_info()
            data = resp['data']
            table_range, revision, updates = data['tableRange'], data['revision'], data['updates']
            range, cells = updates['updatedRange'], updates['updatedCells']
            rows, columns = updates['updatedRows'], updates['updatedColumns']
            if FEISHU_VERBOSE in ['spreadsheet', 'all']:
                print(f'Append Data Successfully: range={range}, {rows} rows, {columns} columns, {cells} cells')
            logger.info(f'Append Data Successfully: range={range}, {rows} rows, {columns} columns, {cells} cells')
            return range
        else:
            logger.error(f'Append Data Failed: {resp}')

    def _write_range(self, cell_start, cell_end, values, sheet=0, update=True):
        """
        向单个range范围写入数据，若范围内有数据，会被更新覆写
        doc: https://open.feishu.cn/document/server-docs/docs/sheets-v3/data-operation/write-data-to-a-single-range
        update: 20230727
        :param cell_start:
        :param cell_end:
        :param values:
        :param sheet:
        :param update:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url_v2}/values'
        body = {
            'valueRange': {
                'range': f'{sheet_id}!{cell_start}:{cell_end}',
                'values': values
            }
        }
        resp = requests.put(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            if update:
                self._update_meta_info()
            data = resp['data']
            range, cells = data['updatedRange'], data['updatedCells']
            rows, columns = data['updatedRows'], data['updatedColumns']
            logger.info(f'Write Data Successfully: range={range}, {rows} rows, {columns} columns, {cells} cells')
            return range
        else:
            logger.error(f'Write Data Failed: {resp}')

    def _read_range(self, cell_start, cell_end, sheet=0):
        """
        读取单个range范围：返回数据限制为10M
        doc: https://open.feishu.cn/document/server-docs/docs/sheets-v3/data-operation/reading-a-single-range
        update: 20230727
        :param cell_start:
        :param cell_end:
        :param sheet:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url_v2}/values/{sheet_id}!{cell_start}:{cell_end}'
        params = {
            'valueRenderOption': 'ToString',            # 先ToString再读取，否则对于包含url的cell，会按FormattedValue来读取?
            'dateTimeRenderOption': 'FormattedString'
        }
        resp = requests.get(url, params=params, headers=self.headers).json()
        if resp['code'] == 0:
            values = resp['data']['valueRange']['values']
            return values
        else:
            logger.error(f'Read Range Failed: {resp}')
            return None

    def _read_ranges(self, cells, sheet=0):
        """
        读取多个range范围
        doc: https://open.feishu.cn/document/server-docs/docs/sheets-v3/data-operation/reading-multiple-ranges
        update: 20230727
        :param cells: 形如[{cell_start, cell_end), (), ...]
        :param sheet:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url_v2}/values_batch_get'
        params = {
            'ranges': ','.join([f'{sheet_id}!{x[0]}:{x[1]}' for x in cells]),
            'valueRenderOption': 'ToString',
            'dateTimeRenderOption': 'FormattedString'
        }
        resp = requests.get(url, params=params, headers=self.headers).json()
        if resp['code'] == 0:
            data = resp['data']
            value_ranges, total_cells = data['valueRange'], data['totalCells']
            range2values = {x['range']: x['values'] for x in value_ranges}
            return range2values
        else:
            logger.error(f'Read Ranges Failed: {resp}')
            return None

    def _write_image(self, cell, image=None, image_path=None, image_type=None, name=None, sheet=0, update=True):
        """
        向一个cell写入一张图片，可以直接指定image（优先），也可以指定image文件和类型（以生成image）
        doc: https://open.feishu.cn/document/server-docs/docs/sheets-v3/data-operation/write-images
        update: 20230801
        :param cell: 一个cell，比如'B3'
        :param image: array，图片二进制流，可通过cv2等来生成，若指定，则优先级高于image_path
        :param image_path: 图片文件，若为None，必须指定image和image_type
        :param image_type: 支持png, jpeg, jpg, gif, bmp, jfif, exif, tiff, bpg, webp, heic等格式
        :param name:
        :param update:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url_v2}/values_image'
        if image_path:
            name = name if name else image_path.split('/')[-1]
            image_type = image_type if image_type else name.split('.')[-1]
        if image is None:
            image = cv2.imencode(f'.{image_type}', cv2.imread(image_path))[1].tolist()
        body = {
            'range': f'{sheet_id}!{cell}:{cell}',
            'image': image,
            'name': name if name else f'test.{image_type}'
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            if update:
                self._update_meta_info()
            range = resp['data']['updateRange']
            return range
        else:
            logger.error(f'Write Image Failed: {resp}')

    def write_image(self, image_paths, spreadsheet_token=None, sheet=0, cell_start='A1', axis='column', update=True):
        """
        调用_write_image写入图片，目前只支持写入一列或一行
        :param image_paths: 暂时先只支持_write_image中的image_path
        :param spreadsheet_token:
        :param sheet:
        :param cell_start: 从cell_start向下写入一列，或向右写入一行
        :param axis: column表示写入一列，row表示写入一行
        :param update:
        :return:
        """
        if spreadsheet_token:
            self._set_spreadsheet_token(spreadsheet_token)
        else:
            assert self.spreadsheet_token is not None, '没有spreadsheet_token，需要指定！'
            self._update_meta_info()        # 写之前先更新并获取最新信息，因为sheet可能刚更新，比如新增sheet、写入数据等

        x_start, y_start = cell_to_xy(cell_start)
        if axis == 'column':
            cells = [xy_to_cell(x_start + i, y_start) for i in range(len(image_paths))]
            next_cell_start = xy_to_cell(x_start + len(image_paths), y_start)
        elif axis == 'row':
            cells = [xy_to_cell(x_start, y_start + i) for i in range(len(image_paths))]
            next_cell_start = xy_to_cell(x_start, y_start + len(image_paths))

        for cell, image_path in tqdm(zip(cells, image_paths)):
            self._write_image(cell, image_path=image_path, sheet=sheet, update=False)
        if update:
            self._update_meta_info()  # 写完所有数据后再update

        if FEISHU_VERBOSE in ['spreadsheet', 'all']:
            print(f'下次write_image，请从cell_start={next_cell_start}开始')
        logger.info(f'下次write_image，请从cell_start={next_cell_start}开始')
        return next_cell_start


    def write_df(self, df, spreadsheet_token=None, sheet=0, cell_start='A1', xy_start=None, max_num=1000, update=True):
        """
        调用_append_data，把DataFrame写入sheet，从cell_start或xy_start开始写(append)，返回下一个可用的cell
        支持的cell类型：数值、字符串、日期、None、URL(按字符串写入)
        不支持的cell类型：List, Dict等，若想写入，先转化为str类型；对于图片，会单独处理，此API不处理图片
        update: 20230801
        :param df:
        :param spreadsheet_token:
        :param sheet:
        :param cell_start:
        :param xy_start:
        :param max_num:
        :param update:
        :return:
        """
        if spreadsheet_token:
            self._set_spreadsheet_token(spreadsheet_token)
        else:
            assert self.spreadsheet_token is not None, '没有spreadsheet_token，需要指定！'
            self._update_meta_info()        # 写之前先更新并获取最新信息，因为sheet可能刚更新，比如新增sheet、写入数据等

        if xy_start:
            cell_start = xy_to_cell(xy_start[0], xy_start[1])
        if cell_start:
            letter_start, _ = re.findall(PATTERN, cell_start)[0]
            x_start, y_start = cell_to_xy(cell_start)
            cell_end = xy_to_cell(x_start + max_num - 1, y_start + df.shape[1] - 1)
        # cell_start = xy_to_cell(line_start, 0)
        # cell_end = xy_to_cell(line_start + max_num - 1, df.shape[1] - 1)

        letter_end, row_end = re.findall(PATTERN, cell_end)[0]
        row_end = int(row_end)
        values = [list(df.columns)]
        for _, se in tqdm(df.iterrows()):
            values.append(se.to_list())
            if len(values) == max_num:      # 每次只写max_num行
                logger.info(f'Range: {cell_start}, {cell_end}')
                self._append_data(cell_start, cell_end, values, sheet, update=False)    # 单次写入，先不update
                # 更新下一次的values, cell_start, cell_end, row_end
                values = []
                cell_start = letter_start + str(row_end + 1)
                cell_end = letter_end + str(row_end + max_num)
                row_end += max_num
        if len(values) > 0:
            logger.info(f'Range: {cell_start}:{cell_end}')
            self._append_data(cell_start, cell_end, values, sheet, update=False)
        if update:
            self._update_meta_info()        # 写完所有数据后再update

        cell_start = letter_start + str(row_end - max_num + len(values) + 1)
        if FEISHU_VERBOSE in ['spreadsheet', 'all']:
            print(f'下次write_df，请从cell_start={cell_start}开始')
        logger.info(f'下次write_df，请从cell_start={cell_start}开始')
        return cell_start

    def read_sheet(self, spreadsheet_token=None, sheet=0, cell_start='A1', cell_end=None,
                   xy_start=(0, 0), xy_end=None, has_cols=True, col_names=None, max_num=1000):
        """
        调用read_range，读取某sheet中某区域的数据，可指定cell_start到cell_end，或xy_start到xy_end
        没指定区域的话，可自行判断所有有效区域，建议明确指定起始cell，尤其是cell_end
        update: 20230817
        :param spreadsheet_token:
        :param sheet:
        :param cell_start:
        :param cell_end:
        :param xy_start:
        :param xy_end:
        :param has_cols: range内第1行是不是列名
        :param col_names: 若range第1行不是列名，指定列名为col_names
        :param max_num:
        :return:
        """
        if spreadsheet_token:
            self._set_spreadsheet_token(spreadsheet_token)
        else:
            assert self.spreadsheet_token is not None, '暂无spreadsheet_token，需要指定！'
            self._update_meta_info()        # 读之前先更新一下最新信息，因为sheet可能刚更新，如写入数据、新增sheet等

        if cell_end:                        # 若指定了cell_end，则优先使用cell_start和cell_end
            xy_start = cell_to_xy(cell_start)
            xy_end = cell_to_xy(cell_end)

        if xy_end is None:                  # 若使用xy坐标，且没指定xy_end，则自行判断xy_end
            sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
            sheet_index = self.sheet_id2index[sheet_id]
            grid_properties = self.sheets[sheet_index]['grid_properties']
            xy_end = (grid_properties['row_count'] - 1, grid_properties['column_count'] - 1)

        x_start, y_start = xy_start
        x_end, y_end = xy_end
        values = []
        for x in range(x_start, x_end + 1, max_num):    # 每次只读取max_num行
            cell_start = xy_to_cell(x, y_start)
            cell_end = xy_to_cell(min(x_end, x + max_num - 1), y_end)
            value = self._read_range(cell_start, cell_end, sheet)
            values.extend(value)

        if has_cols:
            col_names = col_names if col_names else values[0]
            return pd.DataFrame(values[1:], columns=col_names)
        else:
            return pd.DataFrame(values, columns=col_names)


if __name__ == '__main__':

    # 首次使用，需要先初始化
    idt = Identification(get_new_code=True)
    code = 'xxx'
    idt.init_with_code(code)

    # 简单demo
    spsh = SpreadSheet()
    df = spsh.read_sheet(spreadsheet_token='xxx', sheet='Sheet1', cell_start='B1', cell_end='F501')     # 读取sheet
    spsh.write_df(df, spreadsheet_token='xxx', sheet='sheet2', cell_start='D1')                         # 写入sheet

    # 复杂demo:
    # demo1: 连续写入同一个sheet
    spsh = SpreadSheet()
    res = {'model1': None, 'model2': None}
    cell_start = 'A1'
    for col, metric in res.items():
        metric['model'] = [col] * metric.shape[0]
        cell_start = spsh.write_df(metric, spreadsheet_token='xxx', sheet='xxx', cell_start=cell_start)
        # 若有需要，可更新cell_start，比如每隔2行写入一份数据，则修改cell_start: A200 -> A202
        # 其实下面这样也行，但返回值一直是第1次写入时的cell_start(???)，不优雅，无法准确得知下一行可以写入的行号
        # spsh.write_df(metric, spreadsheet_token='xxx', sheet='xxx')

    # demo2: 在现有sheet数据上，单独添加一列或多列，比如从N1位置向右添加一列或多列   注意：只有1列时，df一定要使用2层中括号！
    spsh.write_df(df['score1', 'score2'], spreadsheet_token='xxx', sheet='xxx', cell_start='N1')
