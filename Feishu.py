from urllib.parse import urlencode
import requests
import logging
import re
import time
from tqdm import tqdm
import pandas as pd

logger = logging.getLogger(__name__)


# 0. Common
APP_ID = 'xxx'
APP_SECRET = 'xxx'
REDIRECT_URI = 'xxx'


ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
PATTERN = re.compile(r'([a-zA-Z]+)(\d+)')   # 拆分字母和数字


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


def get_headers(user_access_token):
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {user_access_token}'
    }
    return headers


# 1. 身份验证
class Identification(object):

    def __init__(self, app_id=None, app_secret=None, redirect_uri=None, get_new_code=False):
        """
        要么扫码获得code以重新获得相关token并保存XXX，要么从XXX获得user_refresh_token，以refresh所有变量
        初次初始化要飞书授权获得code，然后在XXX保存user_refresh_token(有效期30天，过期要也要重新扫码获得code)，之后每次初始化建议直接读XXX
        :param app_id:
        :param app_secret:
        :param redirect_uri:
        :param get_new_code:
        """
        self.app_id = app_id if app_id else APP_ID
        self.app_secret = app_secret if app_secret else APP_SECRET
        self.api_url = 'https://open.feishu.cn/open-apis'
        self.redirect_uri = redirect_uri if redirect_uri else REDIRECT_URI
        self.headers = {
            'Content-Type': 'application/json'
        }
        self._get_app_access_token()

        if get_new_code:
            self._get_id_url()
        else:
            # 读取TCC，获得xxx
            config = 'xxx'
            self.user_refresh_token = 'xxx'
            self.user_refresh_token_expiration = int('xxx')
            code_times = int('xxx') + 1
            self.refresh_user_access_token(code_times)  # 重新生成user_access_token等user_xxx变量

    def _get_id_url(self):
        """
        1.1 请求身份验证：获得用户登录预授权码code，有效期为5分钟，且只能使用一次
        该方法会返回url，用浏览器打开后返回结果，形如：{REDIRECT_URI}?code=xxx&state=test
        :param self: 
        :return: 
        """
        body = {
            'redirect_uri': self.redirect_uri,
            'app_id': self.app_id,
            'state': 'test'
        }
        self.id_url = f'{self.api_url}/authen/v1/index?{urlencode(body)}'
        print(f'请在浏览器里打开这个URL，飞书授权后获得code，随后使用code进行初始化：{self.id_url}')

    def init_with_code(self, code):
        """
        使用用户登录预授权码code进行初始化
        :param code:
        :return:
        """
        self.code = code
        self._get_user_info()

    def _get_app_access_token(self):
        """
        1.2 获取app_access_token（企业自建应用）
        :return:
        """
        url = f'{self.api_url}/auth/v3/app_access_token/internal/'
        body = {
            'app_id': self.app_id,
            'app_secret': self.app_secret
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            self.app_access_token = resp['app_access_token']    # app_access_token每隔一段时间会变
        else:
            logger.error(f'Get App Access Token Failed: {resp}')

    def _update_token(self, code_times=0):
        """
        修改XXX里保存的token和expiration
        :param code_times:
        :return:
        """
        values = {
            'app_access_token': self.app_access_token,
            'user_access_token': self.user_access_token,
            'user_access_token_expiration': self.user_access_token_expiration,
            'user_refresh_token': self.user_refresh_token,
            'user_refresh_token_expiration': self.user_refresh_token_expiration,
            'user_open_id': self.user_open_id,
            'user_name': self.user_name,
            'time': time.strftime('%Y-%m-%d %H:%M:%S', time.gmtime()),  # 北京时区 - 8个小时
            'code_times': code_times        # 基于最近1次code，refresh tokens的次数
        }
        # 保存到XXX里   TODO
        print(values)

    def _get_user_info(self):
        """
        1.3 获取登录用户身份 1.1 + 1.2 -> 1.3
        :return:
        """
        url = f'{self.api_url}/authen/v1/access_token'
        body = {
            'app_access_token': self.app_access_token,
            'grant_type': 'authorization_code',
            'code': self.code
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            data = resp['data']
            self.user_access_token = data['access_token']
            self.user_access_token_expiration = int(time.time()) + data['expires_in'] - 120     # 提前2分钟，下同
            self.user_refresh_token = data['refresh_token']     # 有效期30天
            self.user_refresh_token_expiration = int(time.time()) + data['refresh_expires_in'] - 120
            self.user_open_id = data['open_id']
            self.user_name = data['name']
            self._update_token(0)
        else:
            logger.error(f'Get User Info Failed: {resp}')

    def _if_user_access_token_expired(self):
        """
        判断是否需要更新user_access_token  user_access_token有效期约6900秒（约1.9小时），user_refresh_token有效期约30天
        :return:
        """
        return int(time.time()) >= self.user_access_token_expiration

    def refresh_user_access_token(self, code_times):
        """
        1.4 刷新user_access_token   基于user_refresh_token获得新的user_access_token和user_refresh_token
        :param code_times:
        :return:
        """
        assert int(time.time()) < self.user_refresh_token_expiration, 'user_refresh_token已过期，无法refresh，需要重新生成'
        url = f'{self.api_url}/authen/v1/refresh_access_token'
        body = {
            'app_access_token': self.app_access_token,
            'grant_type': 'refresh_token',
            'refresh_token': self.user_refresh_token
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            data = resp['data']
            self.user_access_token = data['access_token']
            self.user_access_token_expiration = int(time.time()) + data['expires_in'] - 120  # 提前2分钟，下同
            self.user_refresh_token = data['refresh_token']  # 有效期30天
            self.user_refresh_token_expiration = int(time.time()) + data['refresh_expires_in'] - 120
            self.user_open_id = data['open_id']
            self.user_name = data['name']
            self._update_token(code_times)
        else:
            logger.error(f'Refresh User Access Token Failed: {resp}')

    # 错误码code含义：
    # 20007: 生成access_token失败，请确保code没有重复消费或过期消费

    def get_user_info_identification(self):
        """
        1.5 获取用户信息（身份验证）
        :return:
        """
        headers = get_headers(self.user_access_token)
        url = f'{self.api_url}/authen/v1/user_info'
        resp = requests.get(url, headers=headers).json()
        return resp


# 2. Folder
# 3. File
# 4. Permission
# 4.1 增加权限
def add_permission(file_token, file_type, user_access_token, openids=[], emails=[]):
    headers = get_headers(user_access_token)
    body = {
        'token': file_token,
        'type': file_type,
        'members': [
            {
                'member_type': 'openid',
                'member_id': x,
                'perm': 'view'
            } for x in openids
        ] + [
            {
                'member_type': 'email',
                'member_id': x,
                'perm': 'view'
            } for x in emails
        ]
    }
    url = 'https://open.feishu.cn/open-apis/drive/permission/member/create'
    resp = requests.post(url, json=body, headers=headers).json()
    return resp


# 5. Doc
# 6. Sheets
class SpreedSheet(object):
    """
    操作SpreedSheet，暂时只需关注write_df(df写入sheet)和read_sheet(读取sheet为df)这2个API
    """
    def __init__(self, spreedsheet_token=None, user_access_token=None):
        """
        若1个文档要操作多次，建议为其专门初始化一个实例(在初始化时指定spreedsheet_token)
        若有多个文档，每个文档只操作一两次，建议先初始化1个公共实例，在操作具体每个文档时再指定spreedsheet_token
        :param spreedsheet_token:
        :param user_access_token:
        """
        if user_access_token is None:
            self.idt = Identification()
            user_access_token = self.idt.user_access_token
        self._refresh_user_access_token(user_access_token)
        if spreedsheet_token:
            self._set_spreedsheet_token(spreedsheet_token)

    def _refresh_user_access_token(self, user_access_token):
        """
        使用user_access_token初始化或更新headers
        :param user_access_token:
        :return:
        """
        self.user_access_token = user_access_token
        self.headers = get_headers(self.user_access_token)

    def _set_spreedsheet_token(self, spreedsheet_token):
        """
        设置或修改spreedsheet_token
        :param spreedsheet_token:
        :return:
        """
        self.spreedsheet_token = spreedsheet_token
        self.api_url = f'https://open.feishu.cn/open-apis/sheet/v2/spreedsheets/{spreedsheet_token}'
        self._update_meta_info()


    def _update_meta_info(self):
        """
        更新表格元数据
        sheets = {
            0: {
                "sheetId": "***",
                "title": "***",
                "index": 0,
                "rowCount": 0,
                'columnCount": 0
            },
            1: {}
        }
        :return:
        """
        url = f'{self.api_url}/metainfo'
        resp = requests.get(url, headers=self.headers).json()
        if resp['code'] == 0:
            data = resp['data']
            self.title = data['properties']['title']
            self.sheets = {x['index']: x for x in data['sheets']}
            self.sheet_index2id = {val['index']: val['sheetId'] for key, val in self.sheets.items()}
            self.sheet_title2id = {val['title']: val['sheetId'] for key, val in self.sheets.items()}
            self.sheet_id2index = {val: key for key, val in self.sheet_index2id.items()}
        else:
            logger.error(f'Get Meta Info Failed: {resp}')

    def _change_title(self, title):
        """
        修改表格title
        :param title:
        :return:
        """
        url = f'{self.api_url}/properties'
        body = {
            'properties': {
                'title': title
            }
        }
        resp = requests.put(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            self.title = title
        return resp

    def _add_sheet(self, title, index=0):
        """
        添加sheet
        :param title:
        :param index:
        :return:
        """
        url = f'{self.api_url}/sheets_batch_update'
        body = {
            'requests': [{
                'addSheet': {
                    'properties': {
                        'title': title,
                        'index': index
                    }
                }
            }]
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            self._update_meta_info()
            properties = list(resp['data']['replies'][0].values())[0]['properties']
            sheet_id, index = properties['sheetId'], properties['index']
            logger.info(f'Add Sheet Successfully: index={index}, sheet_id={sheet_id}, title={title}')
            return sheet_id, index
        else:
            logger.error(f'Add Sheet Failed: {resp}')
            return None, None

    def _copy_sheet(self, title, sheet=0):
        """
        复制sheet
        :param title:
        :param sheet:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url}/sheets_batch_update'
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
        :param sheet:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url}/sheets_batch_update'
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

    def _change_sheet_metainfo(self, sheet, title=None, index=None, hidden=None, lock=None, users=None):
        """
        更新sheet属性：更新标题，移动index，隐藏sheet，锁定sheet
        :param sheet:
        :param title:
        :param index:
        :param hidden:
        :param lock:
        :param users:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url}/sheets_batch_update'
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
                properties.update({'protect': {'lock': lock, 'userIds': users}})
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

    def _merge_cells(self, cell_start, cell_end, sheet=0, merge_type='MERGE_ALL'):
        """
        合并cells，取值为最左上角cell的值
        :param cell_start:
        :param cell_end:
        :param sheet:
        :param merge_type: MERGE_ALL=区域直接合并，MERGE_ROWS=区域内按行合并，MERGE_COLUMNS=区域内按列合并
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url}/merge_cells'
        body = {
            'range': f'{sheet_id}|{cell_start}:{cell_end}',
            'mergeType': merge_type
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            self._update_meta_info()
            logger.info(f'Merge Cells Successfully: sheet_range={sheet_id}|{cell_start}:{cell_end}')
        else:
            logger.error(f'Merge Cells Failed: {resp}')

    def _prepend_data(self, cell_start, cell_end, values, sheet=0, update=True):
        """
        在范围range(cell_start到cell_end)内，前插数据：其他数据下移(注意并非右移)，相当于excel中在range上界处向上插入n行
        前插的数据的长和宽应该<=range的长和宽，单次写入不超过5000行和100列
        :param cell_start:
        :param cell_end:
        :param values:
        :param sheet:
        :param update:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url}/values_prepend'
        body = {
            'valueRange': {
                'range': f'{sheet_id}|{cell_start}:{cell_end}',
                'values': values
            }
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            if update:
                self._update_meta_info()
            data = resp['data']
            table_range, updates = data['tableRange'], data['updates']
            range, cells = updates['updatedRange'], updates['updatedCells']
            rows, columns = updates['updatedRows'], updates['updatedColumns']
            logger.info(f'Prepend Data Successfully: range={range}, rows={rows}, columns={columns}, cells={cells}')
        else:
            logger.error(f'Prepend Data Failed: {resp}')

    def _append_data(self, cell_start, cell_end, values, sheet=0, option='OVERWRITE', update=True):
        """
        在范围range内或后面，追加数据：从range内或后面第1个空行，开始覆写各行，相当于excel中在第1个空行处粘贴n行（会覆盖后面已有数据？）
        单次写入不超过5000行和100列
        :param cell_start:
        :param cell_end:
        :param values:
        :param sheet:
        :param option:
        :param update:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url}/values_append'
        body = {
            'insertDataOption': option,
            'valueRange': {
                'range': f'{sheet_id}|{cell_start}:{cell_end}',
                'values': values
            }
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            if update:
                self._update_meta_info()
            data = resp['data']
            table_range, updates = data['tableRange'], data['updates']
            range, cells = updates['updatedRange'], updates['updatedCells']
            rows, columns = updates['updatedRows'], updates['updatedColumns']
            logger.info(f'Append Data Successfully: range={range}, rows={rows}, columns={columns}, cells={cells}')
        else:
            logger.error(f'Append Data Failed: {resp}')

    def _write_range(self, cell_start, cell_end, values, sheet=0, update=True):
        """
        向单个range范围写入或覆写数据
        :param cell_start:
        :param cell_end:
        :param values:
        :param sheet:
        :param update:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url}/values'
        body = {
            'valueRange': {
                'range': f'{sheet_id}|{cell_start}:{cell_end}',
                'values': values
            }
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            if update:
                self._update_meta_info()
            data = resp['data']
            range, cells = data['updatedRange'], data['updatedCells']
            rows, columns = data['updatedRows'], data['updatedColumns']
            logger.info(f'Write Data Successfully: range={range}, rows={rows}, columns={columns}, cells={cells}')
        else:
            logger.error(f'Write Data Failed: {resp}')

    def _write_ranges(self):
        """
        向多个range范围写入或覆写数据
        :return:
        """
        raise NotImplementedError

    def _read_range(self, cell_start, cell_end, sheet=0):
        """
        读取单个range范围：返回数据限制为10M
        :param cell_start:
        :param cell_end:
        :param sheet:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url}/values/{sheet_id}|{cell_start}:{cell_end}'
        params = {
            'valueRenderOption': 'ToString'     # 先ToString再读取，否则对于包含url的cell，会按FormattedValue来读取 TODO
        }
        resp = requests.post(url, params=params, headers=self.headers).json()
        if resp['code'] == 0:
            values = resp['data']['valueRange']['values']
            return values
        else:
            logger.error(f'Read Range Failed: {resp}')
            return None

    def _read_ranges(self, cells, sheet=0):
        """
        读取多个range范围  cells=[{cell_start, cell_end), (), ...]
        :param cells:
        :param sheet:
        :return:
        """
        sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
        url = f'{self.api_url}/values_batch_get'
        url += '?ranges=' + ','.join([f'{sheet_id}|{x[0]}:{x[1]}' for x in cells])
        params = {
            'valueRenderOption': 'ToString'
        }
        resp = requests.post(url, params=params, headers=self.headers).json()
        if resp['code'] == 0:
            value_range = resp['data']['valueRange']
            range2values = {x['range']: x['values'] for x in value_range}
            return range2values
        else:
            logger.error(f'Read Ranges Failed: {resp}')
            return None

    def write_df(self, df, spreedsheet_token=None, sheet=0, cell_start='A1', xy_start=None, max_num=1000, update=True):
        """
        调用append_data，把DataFrame写入sheet，从cell_start开始写(append)，返回下一个可用的cell
        TODO 暂时不支持包含nan、List的df，需要提前处理nan！
        TODO email等复杂类型的df
        TODO 图片呢？
        :param df:
        :param spreedsheet_token:
        :param sheet:
        :param cell_start:
        :param xy_start:
        :param max_num:
        :param update:
        :return:
        """
        if spreedsheet_token:
            self._set_spreedsheet_token(spreedsheet_token)
        else:
            assert self.spreedsheet_token is not None, '暂无spreedsheet_token，需要指定！'
            self._update_meta_info()        # 写之前先更新一下最新信息，因为sheet可能刚更新，如新增sheet等

        if xy_start:
            cell_start = xy_to_cell(xy_start[0], xy_start[1])
        if cell_start:
            letter_start, _ = re.findall(PATTERN, cell_start)[0]
            x_start, y_start = cell_to_xy(cell_start)
            cell_end = xy_to_cell(x_start + max_num - 1, y_start + df.shape[1] - 1)
        # cell_start = xy_to_cell(line_start, 0)
        # cell_end = xy_to_cell(line_start + max_num - 1, df.shape[1] - 1)

        letter_end, row_end =re.findall(PATTERN, cell_end)[0]
        row_end = int(row_end)
        values = [list(df.columns)]
        for _, se in tqdm(df.iterrows()):
            values.append(se.to_list())
            if len(values) == max_num:      # 每次只写max_num行
                logger.info(f'Range: {cell_start}, {cell_end}')
                self._append_data(cell_start, cell_end, values, sheet, update=False)
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
        logger.info(f'下次write_df，请从cell_start={cell_start}开始')
        return cell_start

    def read_sheet(self, spreedsheet_token=None, sheet=0, cell_start='A1', cell_end=None,
                   xy_start=(0, 0), xy_end=None, has_cols=True, col_names=None, max_num=1000):
        """
        调用read_range，读取某sheet中某区域的数据，可指定cell_start到cell_end，或xy_start到xy_end
        没指定区域的话，可自行判断所有有效区域，建议明确指定起始cell，尤其是cell_end
        :param spreedsheet_token:
        :param sheet:
        :param cell_start:
        :param cell_end:
        :param xy_start:
        :param xy_end:
        :param has_cols: 原始数据第1行是不是列名
        :param col_names: 原始数据第1行不是列名，指定列名为col_names
        :param max_num:
        :return:
        """
        if spreedsheet_token:
            self._set_spreedsheet_token(spreedsheet_token)
        else:
            assert self.spreedsheet_token is not None, '暂无spreedsheet_token，需要指定！'
            self._update_meta_info()        # 读之前先更新一下最新信息，因为sheet可能刚更新，如写入数据、新增sheet等

        if cell_end:                        # 若指定了cell_end，则优先使用cell_start和cell_end
            xy_start = cell_to_xy(cell_start)
            xy_end = cell_to_xy(cell_end)

        if xy_end is None:                  # 若使用xy坐标，且没指定xy_end，则自行判断xy_end
            sheet_id = self.sheet_index2id.get(sheet, self.sheet_title2id.get(sheet, sheet))
            sheet_index = self.sheet_id2index[sheet_id]
            sheet_info = self.sheets[sheet_index]
            xy_end = (sheet_info['rowCount'] - 1, sheet_index['columnCount'] - 1)

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


class Message(object):
    """Feishu发消息"""
    def __init__(self, tenant_access_token=None):
        self.api_url = 'https://open.feishu.cn/open-apis/message/v4'
        if tenant_access_token is None:
            self.idt = Identification()     # TODO 目前Identification不支持tenant_access_token！！！
            tenant_access_token = self.idt.tenant_access_token
        self.tenant_access_token = tenant_access_token
        self.headers = get_headers(self.tenant_access_token)

    def _send_text(self, text, open_id=None, user_id=None, email=None, chat_id=None, root_id=None, at_user_id=None):
        """
        发送文本消息
        :param text:
        :param open_id:
        :param user_id:
        :param email:
        :param chat_id:
        :param root_id:
        :param at_user_id:
        :return:
        """
        url = f'{self.api_url}/send/'
        if at_user_id:
            text += f' <at user_id=\"{at_user_id}\">test</at>'
        body = {
            'open_id': open_id,
            'root_id': root_id,
            'chat_id': chat_id,
            'user_id': user_id,
            'email': email,
            'msg_type': 'text',
            'content': {
                'text': text
            }
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            message_id = resp['data']['message_id']
            logger.info(f'Send Text Successfully, message_id={message_id}')
        else:
            logger.error(f'Send Text Failed: {resp}')


if __name__ == '__main__':

    # 首次使用，需要先初始化
    idt = Identification(get_new_code=True)
    code = 'd97lf74209014bfdac779b881cd6f879'
    idt.init_with_code(code)

    # 简单demo
    spsh = SpreedSheet()
    df = spsh.read_sheet(spreedsheet_token='xxx', sheet='Sheet1', cell_start='B1', cell_end='F501')     # 读取sheet
    spsh.write_df(df, spreedsheet_token='xxx', sheet='sheet2', cell_start='D1')                         # 写入sheet

    # 复杂demo:
    # demo1: 连续写入同一个sheet
    spsh = SpreedSheet()
    res = {'model1': None, 'model2': None}
    cell_start = 'A1'
    for col, metric in res.items():
        metric['model'] = [col] * metric.shape[0]
        cell_start = spsh.write_df(metric, spreedsheet_token='xxx', sheet='xxx', cell_start=cell_start)
        # 若有需要，可更新cell_start，比如每隔2行写入一份数据，则修改cell_start: A200 -> A202
        # 其实下面这样也行，但返回值一直是第1次写入时的cell_start(???)，不优雅，无法准确得知下一行可以写入的行号
        # spsh.write_df(metric, spreedsheet_token='xxx', sheet='xxx')

    # demo2: 在现有sheet数据上，单独添加一列或多列，比如从N1位置向右添加一列或多列   注意：只有1列时，df一定要使用2层中括号！
    spsh.write_df(df['score1', 'score2'], spreedsheet_token='xxx', sheet='xxx', cell_start='N1')
