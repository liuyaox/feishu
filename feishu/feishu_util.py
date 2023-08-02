import requests


def get_headers(access_token):
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {access_token}'
    }
    return headers


def add_permission(file_token, file_type, user_access_token, openids=[], emails=[]):
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
    resp = requests.post(url, json=body, headers=get_headers(user_access_token)).json()
    return resp


if __name__ == '__main__':
    pass

