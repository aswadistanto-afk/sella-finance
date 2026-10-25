# -*- coding: utf-8 -*-
import json
import urllib.request
import urllib.parse
import os

SCRIPT_ID = '1loT49vsoKGHq5eBybIvfAgg3_Aw3QZpDkKB27EuXaqwyzpkl4Jqc34HT'
DEPLOYMENT_ID = 'AKfycbyBXYSMp-7ZXvBflVJeSpNi-D35QPOfz9Tow9NdxKgIgvQQLtdcwuFT2qux7YvTRXN2bw'
DIR = r'C:\Users\HYPE\.gemini\antigravity\scratch\sella-finance'

def get_access_token():
    with open(r'C:\Users\HYPE\.clasprc.json', encoding='utf-8') as f:
        creds = json.load(f)['tokens']['default']
    payload = urllib.parse.urlencode({
        'client_id': creds['client_id'],
        'client_secret': creds['client_secret'],
        'refresh_token': creds['refresh_token'],
        'grant_type': 'refresh_token'
    }).encode()
    req = urllib.request.Request('https://oauth2.googleapis.com/token', data=payload, method='POST')
    with urllib.request.urlopen(req) as resp:
        return json.load(resp)['access_token']

def api_call(url, method='GET', payload=None, token=None):
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    data = json.dumps(payload).encode('utf-8') if payload is not None else None
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    with urllib.request.urlopen(req) as resp:
        return json.load(resp)

def deploy():
    print("1. Refreshing OAuth Token...")
    token = get_access_token()
    print("Token obtained successfully.")

    # Read files
    with open(os.path.join(DIR, 'appsscript.json'), encoding='utf-8') as f:
        appsscript_source = f.read()
    with open(os.path.join(DIR, 'Code.js'), encoding='utf-8') as f:
        code_source = f.read()
    with open(os.path.join(DIR, 'Index.html'), encoding='utf-8') as f:
        index_source = f.read()

    files = [
        {
            "name": "appsscript",
            "type": "JSON",
            "source": appsscript_source
        },
        {
            "name": "Code",
            "type": "SERVER_JS",
            "source": code_source
        },
        {
            "name": "Index",
            "type": "HTML",
            "source": index_source
        }
    ]

    print(f"2. Uploading content to script {SCRIPT_ID}...")
    content_url = f"https://script.googleapis.com/v1/projects/{SCRIPT_ID}/content"
    update_res = api_call(content_url, method='PUT', payload={"files": files}, token=token)
    print("Content uploaded successfully. Files in project:", [f['name'] for f in update_res.get('files', [])])

    print("3. Creating a new project version...")
    version_url = f"https://script.googleapis.com/v1/projects/{SCRIPT_ID}/versions"
    version_res = api_call(version_url, method='POST', payload={"description": "v3 - AI Nota Tracker via Gemini Vision"}, token=token)
    new_version = version_res['versionNumber']
    print(f"Version {new_version} created successfully!")

    print(f"4. Updating deployment {DEPLOYMENT_ID} to version {new_version}...")
    deploy_url = f"https://script.googleapis.com/v1/projects/{SCRIPT_ID}/deployments/{DEPLOYMENT_ID}"
    deploy_payload = {
        "deploymentConfig": {
            "scriptId": SCRIPT_ID,
            "versionNumber": new_version,
            "manifestFileName": "appsscript",
            "description": f"v{new_version} - AI Nota Tracker (Gemini Vision)"
        }
    }
    deploy_res = api_call(deploy_url, method='PUT', payload=deploy_payload, token=token)
    print("Deployment updated successfully!")

    print("\n5. Checking all entry points...")
    deployments_url = f"https://script.googleapis.com/v1/projects/{SCRIPT_ID}/deployments"
    all_deps = api_call(deployments_url, method='GET', token=token)
    for d in all_deps.get('deployments', []):
        d_id = d.get('deploymentId')
        v_num = d.get('deploymentConfig', {}).get('versionNumber', 'HEAD')
        entry_points = d.get('entryPoints', [])
        for ep in entry_points:
            if ep.get('entryPointType') == 'WEB_APP':
                url = ep.get('webApp', {}).get('url')
                print(f"Deployment [{d_id}] (v{v_num}): {url}")

if __name__ == '__main__':
    deploy()
