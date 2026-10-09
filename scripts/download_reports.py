"""Download public financial reports; never overwrite existing local PDFs."""
import argparse,hashlib,json
from pathlib import Path
import requests

def run(root):
    root=Path(root);items=json.loads((root/'data/source_manifest.json').read_text(encoding='utf-8'))
    dest=root/'data/source_pdf';dest.mkdir(parents=True,exist_ok=True);log=[]
    for item in items:
        folder=dest/('年报' if item['report_type']=='年度报告' else '中度报告');folder.mkdir(exist_ok=True)
        target=folder/item['source_file'];url=item.get('official_download_url')
        if target.exists():print('已存在，跳过：',target.name);continue
        if not url:raise ValueError('缺少已查到的官网下载链接：'+item['source_file'])
        r=requests.get(url,timeout=120);r.raise_for_status()
        if not r.content.startswith(b'%PDF-'):raise ValueError('下载内容不是PDF：'+url)
        actual=hashlib.sha256(r.content).hexdigest()
        target.write_bytes(r.content)
        same=actual==item['original_local_sha256']
        log.append({'source_file':item['source_file'],'url':url,'sha256':actual,'same_bytes_as_audited_local_pdf':same})
        print(target.name,'与抽检文件字节一致' if same else '官网版本与用户文件字节不同；重新抽取并核对页码后再替换语料')
    (root/'output').mkdir(exist_ok=True)
    (root/'output/download_log.json').write_text(json.dumps(log,ensure_ascii=False,indent=2),encoding='utf-8')

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--root',type=Path,default=Path(__file__).resolve().parents[1]);a=p.parse_args();run(a.root)
