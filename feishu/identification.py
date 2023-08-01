import os
from urllib.parse import urlencode
import requests
import logging
import time

from .feishu_util import get_headers
from .config_util import config_feishu_read, config_feishu_write

logger = logging.getLogger(__name__)


# 0. Common
APP_ID = os.environ.get('FEISHU_APP_ID', None)
APP_SECRET = os.environ.get('FEISHU_APP_SECRET', None)
REDIRECT_URI = os.environ.get('FEISHU_REDIRECT_URI', None)
CONFIG_KEY = os.environ.get('FEISHU_CONFIG_KEY', 'yao.liu')


def write_config(key, value):
    """
    在配置中心对指定key写入value  TODO 因为没有专门的配置中心，这只是临时方案，后续需要修改为读取专门的配置中心
    :param key:
    :param value:
    :return:
    """
    return config_feishu_write({key: value})


def read_config(key):
    """
    从配置中心对指定key读取value  TODO 同write_config
    :param key:
    :return:
    """
    return config_feishu_read(key)[key]


# 1. 身份验证
class Identification(object):

    def __init__(self, app_id=None, app_secret=None, redirect_uri=None, get_new_code=False, config_key=None):
        """
        要么飞书授权获得code以重新获得相关token并保存配置中心，要么从配置中心获得user_refresh_token，以refresh所有变量
        初次初始化要飞书授权获得code，然后在配置中心保存user_refresh_token(有效期30天，过期要重新获得code)，之后每次初始化建议直接读配置中心
        update: 20230725
        :param app_id:
        :param app_secret:
        :param redirect_uri:
        :param get_new_code: 是否重新获得用户登录预授权码code，以重新
        :param config_key: 配置中心的key，用于保存feishu相关的token，默认是yao.liu，也可以创建自己的key（建议是自己的名字）
                        初始化时指定config_key，或者直接修改本地环境变量FEISHU_CONFIG_KEY
        """
        self.app_id = app_id if app_id else APP_ID
        self.app_secret = app_secret if app_secret else APP_SECRET
        self.api_url = 'https://open.feishu.cn/open-apis'
        self.redirect_uri = redirect_uri if redirect_uri else REDIRECT_URI
        self.headers = {
            'Content-Type': 'application/json; charset=utf-8'
        }
        self._get_app_access_token()
        self.headers.update({
            'Authorization': f'Bearer {self.app_access_token}'
        })
        self.config_key = config_key if config_key else CONFIG_KEY

        if get_new_code:
            self._get_id_url()
        else:
            config = read_config(self.config_key)
            self.code = config['code']
            self.user_refresh_token = config['user_refresh_token']
            self.user_refresh_token_expire = int(config['user_refresh_token_expire'])
            code_times = int(config['code_times']) + 1
            self.refresh_user_access_token(code_times)        # 重新生成user_access_token等user_xxx变量

    def _get_app_access_token(self):
        """
        1.1 获取app_access_token（企业自建应用）
        doc: https://open.feishu.cn/document/server-docs/authentication-management/access-token/app_access_token_internal
        update: 20230725
        :return:
        """
        url = f'{self.api_url}/auth/v3/app_access_token/internal'
        body = {
            'app_id': self.app_id,
            'app_secret': self.app_secret
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            self.app_access_token = resp['app_access_token']    # app_access_token每隔一段时间会变
            self.app_access_token_expire = int(time.time()) + resp['expire'] - 120  # 最大有效期2小时  TODO 会不会影响使用？不会，及时用app_access_token获得user_access_xxx即可
        else:
            logger.error(f'Get App Access Token Failed: {resp}')

    def _get_id_url(self):
        """
        1.2 获得用户登录预授权码code: 有效期为5分钟，且只能使用一次
        该方法会返回url，用浏览器打开后返回结果，形如：{REDIRECT_URI}?code=xxx&state=test
        doc: https://open.feishu.cn/document/server-docs/authentication-management/login-state-management/obtain-code
        update: 20230725
        """
        body = {
            'redirect_uri': self.redirect_uri,
            'app_id': self.app_id,
            'state': 'test'
        }
        self.id_url = f'{self.api_url}/authen/v1/index?{urlencode(body)}'
        print(f'请在浏览器里打开这个URL，飞书授权后获得code，随后使用code进行初始化：{self.id_url}')
        logger.info(f'请在浏览器里打开这个URL，飞书授权后获得code，随后使用code进行初始化：{self.id_url}')

    def init_with_code(self, code):
        """
        使用用户登录预授权码code进行初始化
        update: 20230725
        :param code:
        :return:
        """
        self.code = code
        self._get_user_info()

    def _get_user_info(self):
        """
        1.3 获取user_access_token
        doc: https://open.feishu.cn/document/server-docs/authentication-management/access-token/create-2
        update: 20230725
        :return:
        """
        url = f'{self.api_url}/authen/v1/access_token'
        body = {
            'grant_type': 'authorization_code',
            'code': self.code
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            data = resp['data']
            self.user_access_token = data['access_token']
            self.user_access_token_expire = int(time.time()) + data['expires_in'] - 120     # 提前2分钟，下同
            self.user_refresh_token = data['refresh_token']                                 # 有效期30天
            self.user_refresh_token_expire = int(time.time()) + data['refresh_expires_in'] - 120
            self.user_open_id = data['open_id']
            # self.user_id = data['user_id']
            self.user_name = data['name']
            self.user_en_name = data['en_name']
            # self.user_email = data['email']
            self._update_token(0)
        else:
            print(f'Get User Info Failed: {resp}')

    def _update_token(self, code_times=0):
        """
        修改XXX里保存的token和expiration
        update: 20230725
        :param code_times:
        :return:
        """
        values = {
            'app_access_token': self.app_access_token,
            'app_access_token_expire': self.app_access_token_expire,
            'user_access_token': self.user_access_token,
            'user_access_token_expire': self.user_access_token_expire,
            'user_refresh_token': self.user_refresh_token,
            'user_refresh_token_expire': self.user_refresh_token_expire,
            'user_open_id': self.user_open_id,
            # 'user_id': self.user_id,
            'user_name': self.user_name,
            'user_en_name': self.user_en_name,
            # 'user_email': self.user_email,
            'time': time.strftime('%Y%m%d %H:%M:%S', time.localtime()),  # 北京时区
            'code': self.code,
            'code_times': code_times        # 基于最近1次code，refresh tokens的次数
        }
        print(f'往配置中心写入配置：config_key={self.config_key}, config_value=\n{values}')
        logger.info(f'往配置中心写入配置：config_key={self.config_key}, config_value=\n{values}')
        write_config(self.config_key, values)

    def refresh_user_access_token(self, code_times):
        """
        1.4 刷新user_access_token   基于user_refresh_token获得新的user_access_token和user_refresh_token
        doc: https://open.feishu.cn/document/server-docs/authentication-management/access-token/create
        update: 20230725
        :param code_times:
        :return:
        """
        assert int(time.time()) < self.user_refresh_token_expire, 'user_refresh_token已过期，无法refresh，需要重新生成'
        url = f'{self.api_url}/authen/v1/refresh_access_token'
        body = {
            'grant_type': 'refresh_token',
            'refresh_token': self.user_refresh_token
        }
        resp = requests.post(url, json=body, headers=self.headers).json()
        if resp['code'] == 0:
            data = resp['data']
            self.user_access_token = data['access_token']
            self.user_access_token_expire = int(time.time()) + data['expires_in'] - 120     # 提前2分钟，下同
            self.user_refresh_token = data['refresh_token']                                 # 有效期30天
            self.user_refresh_token_expire = int(time.time()) + data['refresh_expires_in'] - 120
            self.user_open_id = data['open_id']
            # self.user_id = data['user_id']
            self.user_name = data['name']
            self.user_en_name = data['en_name']
            # self.user_email = data['email']
            self._update_token(code_times)
        else:
            logger.error(f'Refresh User Access Token Failed: {resp}')

    def get_user_info_identification(self):
        """
        1.5 获取登录用户信息
        doc: https://open.feishu.cn/document/server-docs/authentication-management/login-state-management/get
        update: 20230725
        :return:
        """
        headers = get_headers(self.user_access_token)
        url = f'{self.api_url}/authen/v1/user_info'
        resp = requests.get(url, headers=headers).json()
        if resp['code'] == 0:
            return resp['data']
        else:
            logger.error(f'Get User Info Failed: {resp}')


if __name__ == '__main__':

    # 首次使用，需要先初始化
    idt = Identification(get_new_code=True)
    code = 'xxx'
    idt.init_with_code(code)

    # 错误码code含义：
    # 20007: 生成access_token失败，请确保code没有重复消费或过期消费
