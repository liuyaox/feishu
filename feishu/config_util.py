import requests
import json

HEADERS = {
    'Content-Type': 'application/json'
}
URL_FEISHU = 'http://47.253.94.159:8000/config/feishu'


def config_feishu_read(keys=None):
    """
    读取feishu配置，当keys为None时，读取所有key的配置
    """
    if keys is None:
        resp = requests.get(url=f'{URL_FEISHU}/read', headers=HEADERS).json()
    else:
        keys = [keys] if isinstance(keys, str) else list(keys)
        resp = requests.post(url=f'{URL_FEISHU}/read', data=json.dumps(keys), headers=HEADERS).json()
    if resp['code'] == 0:
        return resp['data']
    else:
        message = resp['message']
        print(f'Config Feishu Read Failed: message={message}')
    return None


def config_feishu_write(kvs):
    """
    写入feishu配置，kvs是key-value字典形式，若key已存在，则覆写
    """
    resp = requests.post(url=f'{URL_FEISHU}/write', data=json.dumps(kvs), headers=HEADERS).json()
    if resp['code'] == 0:
        return True
    else:
        message = resp['message']
        print(f'Config Feishu Write Failed: message={message}')
    return False


if __name__ == '__main__':

    config_feishu_read()
    config_feishu_read('key2')
    config_feishu_read(['key7', 'key3'])
    config_feishu_write({'key2': '888', 'key5': (44, 55)})
