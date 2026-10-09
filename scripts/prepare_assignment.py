"""Prepare B assignment without calling APIs, training, or changing the source QA."""
import argparse,collections,copy,csv,gzip,hashlib,json,random,re
from pathlib import Path

SYSTEM = '你是太古地产的投资者关系问答助手。按公司财报披露口径简洁作答，明确期间、指标及单位，并注明报告名称和报告印刷页码。未印页码时明确使用PDF物理页。资料不足就说明无法确认，不编造数字或作投资建议。'

def read_jsonl(path):
    opener=gzip.open if str(path).endswith('.gz') else open
    with opener(path,'rt',encoding='utf-8-sig') as f:
        return [json.loads(line) for line in f if line.strip()]

def write_jsonl(path,rows):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    with path.open('w',encoding='utf-8') as f:
        for r in rows:f.write(json.dumps(r,ensure_ascii=False)+'\n')

def write_csv(path,rows):
    if not rows:return
    with Path(path).open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=list(rows[0]));w.writeheader();w.writerows(rows)

def citation(s):
    printed=s.get('printed_pages',[])
    pp='、'.join(str(p) for p in printed if p is not None)
    physical='、'.join(map(str,s['pdf_pages']))
    report=f"太古地产{s['year']}年{s['report_type']}"
    return f"《{report}》印刷第{pp}页（PDF第{physical}页）" if pp else f"《{report}》{s.get('section','指定章节')}（{'该页未印页码' if s.get('printed_page_verified') else '印刷页码未核实'}；PDF第{physical}页）"

def prepare(root):
    root=Path(root);data=root/'data';out=root/'output';out.mkdir(exist_ok=True)
    qas=read_jsonl(data/'qa/qa_raw.jsonl')
    decisions=json.loads((data/'audit/decisions.json').read_text(encoding='utf-8'))
    evidence={r['qa']['id']:r for r in json.loads((data/'audit/evidence.json').read_text(encoding='utf-8'))}
    page_map=json.loads((data/'audit/page_map.json').read_text(encoding='utf-8'))
    manual={r['id']:r for r in csv.DictReader((data/'qa/human_review_original.csv').open(encoding='utf-8-sig'))}
    fixes={d['id']:d for d in decisions};clean=[];audit=[];excluded=[]
    for original in qas:
        q=copy.deepcopy(original);d=fixes.get(q['id']);prior=manual.get(q['id'],{})
        if prior.get('verified','').upper()=='FALSE' and not d:
            excluded.append({'id':q['id'],'reason':'人工标为FALSE且未完成本轮修订'});continue
        q['original_question']=q['question'];q['original_answer']=q['answer']
        q['human_review']=prior;q['review_status']='ai_pdf_reviewed' if d else 'automated_only_not_fact_checked'
        q['verified']=prior.get('verified','')
        if d:
            q['question']=d.get('question',q['question']);q['answer']=d.get('answer',q['answer'])
        for i,s in enumerate(q['sources']):
            if d:
                physical,label=d['pages'][i];ps=[physical];labels=[label]
                s['text']='\n'.join(p['text'] for p in evidence[q['id']]['source_pages'] if p['source_index']==i and p['pdf_page']==physical)
                if not s['text']:raise ValueError(f"missing inspected page: {q['id']}")
            else:
                ps=sorted({int(s['pdf_page'])}|{int(p) for p in re.findall(r'【第(\d+)[页頁]】',s['text'])})
                labels=[page_map[s['source_file']][str(p)]['printed_page'] for p in ps]
            s['pdf_page']=ps[0];s['pdf_pages']=ps;s['printed_pages']=labels;s['printed_page_verified']=bool(d)
            s['citation']=citation(s)
        q['pdf_page']=q['sources'][0]['pdf_page'];q['printed_page']=q['sources'][0]['printed_pages'][0]
        q['source_text']='\n\n'.join(s['text'] for s in q['sources'])
        q['answer']=q['answer'].rstrip()+'\n来源：'+'；'.join(s['citation'] for s in q['sources'])+'。'
        clean.append(q)
        if d:
            audit.append({'id':q['id'],'qa_type':q['qa_type'],'year':q['year'],'report_type':q['report_type'],
              'reviewer':'AI对原始PDF图像及全文复核（不是新增人工签核）',
              'original_content_result':'需修订' if ('question' in d or 'answer' in d) else '原文支持',
              'issue':d['issue'],'original_question':original['question'],'original_answer':original['answer'],
              'corrected_question':q['question'],'corrected_answer':q['answer'],
              'source_file':' ; '.join(s['source_file'] for s in q['sources']),
              'printed_page':' ; '.join(str(s['printed_pages'][0]) if s['printed_pages'][0] is not None else '未印页码' for s in q['sources']),
              'pdf_page':' ; '.join(str(s['pdf_page']) for s in q['sources']),
              'review_note':d['note'],'prior_human_verified':prior.get('verified',''),
              'prior_human_note':prior.get('review_note','')})
    write_jsonl(data/'qa/qa_clean.jsonl',clean);write_csv(out/'audit_30.csv',audit)
    write_jsonl(out/'excluded.jsonl',excluded)
    # Union-find: same source page or same normalized question must stay in one split.
    parent={q['id']:q['id'] for q in clean}
    def find(x):
        while parent[x]!=x:parent[x]=parent[parent[x]];x=parent[x]
        return x
    def union(a,b):parent[find(a)]=find(b)
    seen={}
    for q in clean:
        keys=[('question',re.sub(r'\W+','',q['question']).lower())]
        keys += [('page',s['source_file'],p) for s in q['sources'] for p in s['pdf_pages']]
        for key in keys:
            if key in seen:union(q['id'],seen[key])
            else:seen[key]=q['id']
    groups=collections.defaultdict(list)
    for q in clean:groups[find(q['id'])].append(q)
    rng=random.Random(20261009);testgroups=set();testrefs=[]
    # One fact-checked reference per category. Hold out every connected source page.
    for cat in ['income','balance','cashflow','metric','compare','subjective']:
        candidates=[q for q in clean if q['qa_type']==cat and q['id'] in fixes and find(q['id']) not in testgroups]
        candidates.sort(key=lambda q:(len(groups[find(q['id'])]),q['id']))
        if candidates:
            q=candidates[0];testgroups.add(find(q['id']));testrefs.append(q)
    rest=[g for g in groups if g not in testgroups];rng.shuffle(rest)
    valgroups=set();valcount=0
    # Keep reviewed seen questions in training; validate on other groups.
    for g in rest:
        if valcount>=15:break
        if any(q['id'] in fixes for q in groups[g]):continue
        valgroups.add(g);valcount+=len(groups[g])
    splits={'train':[],'validation':[],'test':[]}
    for g,rows in groups.items():splits['test' if g in testgroups else 'validation' if g in valgroups else 'train']+=rows
    assert set(q['id'] for q in splits['train']).isdisjoint(q['id'] for q in splits['test'])
    for name,rows in splits.items():
        write_jsonl(data/f'prepared/{name}_records.jsonl',rows)
        write_jsonl(data/f'prepared/{name}.jsonl',[{'id':q['id'],'review_status':q['review_status'],'messages':[{'role':'system','content':SYSTEM},{'role':'user','content':q['question']},{'role':'assistant','content':q['answer']}]} for q in rows])
        write_jsonl(data/f'prepared/{name}_reviewed_only.jsonl',[{'id':q['id'],'messages':[{'role':'system','content':SYSTEM},{'role':'user','content':q['question']},{'role':'assistant','content':q['answer']}]} for q in rows if q['id'] in fixes])
    trainrefs=[]
    for cat in ['income','balance','cashflow','metric','compare','subjective']:
        candidates=[q for q in splits['train'] if q['qa_type']==cat and q['id'] in fixes]
        if candidates:trainrefs.append(sorted(candidates,key=lambda q:q['id'])[0])
    evalcases=[]
    paraphrases=json.loads((data/'audit/paraphrases.json').read_text(encoding='utf-8'))
    for membership,refs in [('seen',trainrefs),('unseen',testrefs)]:
        for q in refs:
            for wording,question in [('original',q['question']),('paraphrase',paraphrases[q['id']])]:
                evalcases.append({'case_id':f"{q['id']}_{wording}",'qa_id':q['id'],'membership':membership,'wording':wording,
                 'question':question,'reference_answer':q['answer'],'sources':q['sources'],'qa_type':q['qa_type'],
                 'year':q['year'],'report_type':q['report_type'],'reference_review':'ai_pdf_reviewed'})
    write_jsonl(data/'prepared/eval_cases.jsonl',evalcases)
    stat={'raw':len(qas),'clean':len(clean),'audit_count':len(audit),'categories':dict(collections.Counter(r['qa_type'] for r in audit)),
          'content_revision_count':sum(r['original_content_result']=='需修订' for r in audit),
          'cross_page_count':sum('跨页引用' in r['issue'] for r in audit),'original_answer_page_errors':sum('答案页码错误' in r['issue'] for r in audit),
          'split_counts':{k:len(v) for k,v in splits.items()},'split_grouping':'same-source-file-and-page connected components',
          'reviewed_training_count':sum(q['id'] in fixes for q in splits['train']),
          'unreviewed_training_count':sum(q['id'] not in fixes for q in splits['train']),
          'test_reference_count':len(testrefs),'seen_reference_count':len(trainrefs),'eval_cases':len(evalcases),
          'excluded':len(excluded),'model_experiment_status':'NOT_RUN','review_policy':'AI复核30条，保留13条用户既有人工记录；其余不标已验证。默认教学试验允许未抽检样本，另提供reviewed_only文件。'}
    (out/'preparation_summary.json').write_text(json.dumps(stat,ensure_ascii=False,indent=2),encoding='utf-8')
    lines=['# 太古地产财报QA：30条原始PDF抽检','',
      '复核日期：2026-10-09（Asia/Singapore）。复核者：AI；保留用户原有人工记录，不冒充人工签核。',
      '抽样：六类每类5条，共30条；包括用户已经填写的全部13条，其余17条按固定种子补足。因纳入已知问题，是分层风险抽检，不能把错误比例推广到全部300条。','',
      f"原文支持的内容 {30-stat['content_revision_count']} 条；需修订问法/口径/表述 {stat['content_revision_count']} 条。", 
      f"跨页起始页错定位 {stat['cross_page_count']} 条，原答案混用物理页/印刷页 {stat['original_answer_page_errors']} 条（分类可重叠）。",
      '未发现可确认的凭空捏造数值；发现公司本部/集团口径歧义、全报告否定缺乏支持、地域范围过宽。全部补充印刷页与物理页双定位。',
      '修订后30条保留。其余270条未逐条进行事实复核，不声称全部正确。无印刷页码的十年概要明确以章节及PDF第8页定位。','',
      '## 训练与评估边界',json.dumps(stat['split_counts'],ensure_ascii=False),
      '同一原始页面涉及的题全部落入同一数据集；测试原问法及改写均不进入训练。未训练过指本次LoRA训练未见，基础模型是否预训练见过财报不可证明。',
      '训练仅输入问题，不给来源文本；RAG对照只读取检索语料，不读取标准答案。自动分数只是诊断，语义正确性仍需按参考证据确认。','']
    for r in audit:
        lines += [f"## {r['id']} · {r['qa_type']}",f"原始内容结论：{r['original_content_result']}；问题标签：{r['issue']}",
          f"原问题：{r['original_question']}",f"原答案：{r['original_answer']}",f"复核依据：{r['review_note']}",
          f"修订问题：{r['corrected_question']}",f"修订答案：{r['corrected_answer']}",
          f"历史人工记录：{r['prior_human_verified'] or '未填'}；{r['prior_human_note']}",'']
    (out/'audit_30.md').write_text('\n\n'.join(lines),encoding='utf-8')
    print(json.dumps(stat,ensure_ascii=False,indent=2));return stat

if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('--root',type=Path,default=Path(__file__).resolve().parents[1]);a=ap.parse_args();prepare(a.root)
