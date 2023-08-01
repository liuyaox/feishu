import requests
import logging

from .identification import Identification
from .feishu_util import get_headers

logger = logging.getLogger(__name__)


class Message(object):
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

    pass
