import urllib.request
import urllib.error
import json
import sys
import os

token = os.environ.get('GITHUB_TOKEN', '')  # 请通过环境变量传入，不要硬编码

headers = {
    'Authorization': f'token {token}',
    'Accept': 'application/vnd.github.v3+json',
    'Content-Type': 'application/json'
}

body = json.dumps({
    'name': 'ei-dianjing',
    'private': False,
    'description': 'EI期刊智能筛选工具 - Web版 v2.0'
})

try:
    req = urllib.request.Request(
        'https://api.github.com/user/repos',
        data=body.encode('utf-8'),
        headers=headers,
        method='POST'
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        result = json.loads(resp.read())
        print(f"Repo created: {result['html_url']}")
        print(f"Clone URL: {result['clone_url']}")
except urllib.error.HTTPError as e:
    error_body = e.read().decode('utf-8')
    print(f"HTTP Error {e.code}: {error_body}")
except Exception as e:
    print(f"Error: {e}")
