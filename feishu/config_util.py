import requests
import json
import os


# 简单配置中心：我是在公司内网一个公共服务器上用FastAPI开个服务，用字典和pickle来读写配置数据，后续要修改。可以使用任意配置中心
CONFIG_SERVICE_IP = os.environ['CONFIG_SERVICE_IP']
URL_FEISHU = f'http://{CONFIG_SERVICE_IP}:8000/config/feishu'
HEADERS = {
    'Content-Type': 'application/json'
}


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
