"""Save a verified quote to Windows Desktop without overwriting existing files."""
import argparse
import hashlib
import json
import os
import re
from pathlib import Path
import winreg

def save(result_file):
    record=Path(result_file)
    result=json.loads(record.read_text(encoding='utf-8-sig'))
    source=Path(result['xlsx']); payload=source.read_bytes()
    digest=hashlib.sha256(payload).hexdigest()
    if digest!=result['sha256'] or not result.get('money_and_time_blank'):
        raise ValueError('Quote changed after verification; verify it again before saving')
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders') as key:
        desktop=Path(os.path.expandvars(winreg.QueryValueEx(key,'Desktop')[0]))
    if not desktop.is_dir(): raise FileNotFoundError('Windows Desktop folder is unavailable')
    title=re.sub(r'[<>:"/\\|?*\x00-\x1f]','_',result['title']).rstrip(' .')[:150] or '报价表'
    for index in range(10000):
        name=title+(f' ({index})' if index else '')+'.xlsx'
        target=desktop/name
        try:
            with target.open('xb') as f: f.write(payload)
            break
        except FileExistsError:
            if target.is_file() and hashlib.sha256(target.read_bytes()).hexdigest()==digest: break
    else: raise FileExistsError('No unused desktop filename found')
    if hashlib.sha256(target.read_bytes()).hexdigest()!=digest: raise IOError('Desktop copy verification failed')
    result['desktop_file']=str(target)
    record.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    return str(target)

if __name__=='__main__':
    parser=argparse.ArgumentParser(); parser.add_argument('result_json')
    print(json.dumps({'desktop_file':save(parser.parse_args().result_json)},ensure_ascii=False))
