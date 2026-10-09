"""Three routes on fixed four-cell questions; no reference answer enters inference."""
import argparse,collections,contextlib,csv,json,math,re,time
from pathlib import Path
from prepare_assignment import SYSTEM,read_jsonl,write_jsonl,write_csv

def normalize(text):
    try:
        from opencc import OpenCC
        return OpenCC('t2s').convert(text).lower()
    except ImportError:return text.lower()

def tokens(text):
    text=normalize(text);parts=re.findall(r'[a-z0-9]+|[\u4e00-\u9fff]+',text)
    return [t for part in parts for t in ([part] if not re.fullmatch(r'[\u4e00-\u9fff]+',part) else ([part] if len(part)==1 else [part[i:i+2] for i in range(len(part)-1)]))]

def parse_filters(question):
    years=set(re.findall(r'20\d{2}',question))
    digit=dict(zip('零〇一二三四五六七八九','00123456789'))
    for s in re.findall(r'二[零〇][一二][零〇一二三四五六七八九]',question):years.add(''.join(digit[c] for c in s))
    rt='中期报告' if any(w in question for w in ['半年','中期','上半年','前六个月','六月三十日止六个月','六月三十日止六個月']) else ('年度报告' if any(w in question for w in ['年报','年報','年度','全年','年底','年末','十二月三十一日','十年财务概要','十年財務概要']) else None)
    return sorted(years),rt

class HybridRetriever:
    def __init__(self,root,vector=True):
        import numpy as np
        self.np=np;self.root=Path(root);self.docs=read_jsonl(self.root/'data/chunks/all_chunks.jsonl')
        self.doc_by_id={d['chunk_id']:d for d in self.docs};self.counters=[collections.Counter(tokens(d['text'])) for d in self.docs]
        self.lengths=[sum(c.values()) for c in self.counters];self.avg=sum(self.lengths)/len(self.lengths)
        self.df=collections.Counter(t for c in self.counters for t in c);self.encoder=None;self.matrix=None;self.vector_ids=None
        if vector:
            from sentence_transformers import SentenceTransformer
            from huggingface_hub import snapshot_download
            model_id='BAAI/bge-large-zh-v1.5'
            path=model_id
            for cache_dir in [None,str(Path.home()/'.cache/huggingface')]:
                try:
                    path=snapshot_download(model_id,cache_dir=cache_dir,local_files_only=True)
                    break
                except (OSError,ValueError):continue
            self.encoder=SentenceTransformer(path,device='cpu')
            cache=np.load(self.root/'data/retrieval/vector_snapshot.npz',allow_pickle=False)
            self.matrix=cache['embeddings'].astype('float32');self.vector_ids=cache['ids'].tolist()
            assert len(self.vector_ids)==len(self.matrix)==len(self.docs)
            self.matrix/=np.linalg.norm(self.matrix,axis=1,keepdims=True).clip(min=1e-9)
            self.matrix_rows={k:i for i,k in enumerate(self.vector_ids)}
    def retrieve(self,question,k=6):
        years,rt=parse_filters(question)
        candidate=[i for i,d in enumerate(self.docs) if (not years or str(d['year']) in years) and (rt is None or d['report_type']==rt)]
        # Do not silently discard year constraints when no document matches.
        if not candidate:return []
        terms=set(tokens(question));N=len(self.docs)
        def score(i):
            c=self.counters[i];length=self.lengths[i];result=0
            for t in terms:
                f=c.get(t,0)
                if f:result+=math.log(1+(N-self.df[t]+.5)/(self.df[t]+.5))*f*2.5/(f+1.5*(.25+.75*length/self.avg))
            return result
        bm=sorted(candidate,key=lambda i:(-score(i),self.docs[i]['chunk_id']))[:50]
        ranks=collections.defaultdict(float)
        for rank,i in enumerate(bm,1):ranks[i]+=1/(60+rank)
        vec=[]
        if self.encoder is not None:
            v=self.encoder.encode(['为这个句子生成表示以用于检索相关文章：'+normalize(question)],normalize_embeddings=True,convert_to_numpy=True)[0]
            scored=[(float(self.matrix[self.matrix_rows[self.docs[i]['chunk_id']]]@v),i) for i in candidate]
            vec=[i for _,i in sorted(scored,reverse=True)[:50]]
            for rank,i in enumerate(vec,1):ranks[i]+=1/(60+rank)
        ordered=sorted(ranks,key=lambda i:(-ranks[i],self.docs[i]['chunk_id']))
        picked=[]
        if len(years)>1:
            quota=max(1,k//len(years))
            for year in years:picked += [i for i in ordered if str(self.docs[i]['year'])==year][:quota]
        picked += [i for i in ordered if i not in picked][:k-len(picked)]
        hits=[]
        for i in picked[:k]:
            d=self.docs[i]
            hits.append({'chunk_id':d['chunk_id'],'text':d['text'],'year':d['year'],'report_type':d['report_type'],
             'source_file':d['source_file'],'pdf_pages':d.get('pdf_pages',[d['pdf_page']]),'printed_pages':d.get('printed_pages',[]),
             'section':d.get('section',''),'rrf_score':round(ranks[i],6),'filters':{'years':years,'report_type':rt},
             'bm25_rank':bm.index(i)+1 if i in bm else None,'vector_rank':vec.index(i)+1 if i in vec else None})
        return hits

def retrieval_context(hits):
    return '\n\n'.join(f"[{i}] 《太古地产{h['year']}年{h['report_type']}》报告页码{h['printed_pages']}，PDF页码{h['pdf_pages']}，章节{h['section']}\n{h['text']}" for i,h in enumerate(hits,1))

def run(root,limit=0,retrieval_only=False):
    root=Path(root);out=root/'output';cases=read_jsonl(root/'data/prepared/eval_cases.jsonl')
    if limit:cases=cases[:limit]
    retriever=HybridRetriever(root);contexts={}
    retrieval=[]
    for c in cases:
        hits=retriever.retrieve(c['question']);contexts[c['case_id']]=hits
        gold={(s['source_file'],p) for s in c['sources'] for p in s['pdf_pages']}
        recalled={(h['source_file'],p) for h in hits for p in h['pdf_pages']}
        retrieval.append({'case_id':c['case_id'],'question':c['question'],'membership':c['membership'],'wording':c['wording'],
          'gold_pages_recall':len(gold&recalled)/len(gold),'hits':hits})
    write_jsonl(out/'retrieval_evaluation.jsonl',retrieval)
    if retrieval_only:print(f'检索完成 {len(cases)}题（未调用回答模型）。');return retrieval
    del retriever
    import torch
    from transformers import AutoTokenizer,AutoModelForImageTextToText,AutoModelForCausalLM,BitsAndBytesConfig
    from peft import PeftModel
    info=json.loads((out/'training_run.json').read_text(encoding='utf-8'))
    if not torch.cuda.is_available():raise RuntimeError('对照生成需要Colab GPU。')
    actual_train_ids={r['id'] for r in read_jsonl(root/('data/prepared/train_reviewed_only.jsonl' if info['reviewed_only'] else 'data/prepared/train.jsonl'))}
    actual_train_ids -= {r['id'] for r in info.get('rejected_oversize_samples',[])}
    for c in cases:
        if c['membership']=='seen' and c['qa_id'] not in actual_train_ids:raise RuntimeError(f"{c['qa_id']}被编码阶段排除，不能标成训练过。请重新构造seen题。")
        if c['membership']=='unseen' and c['qa_id'] in actual_train_ids:raise RuntimeError('测试题泄漏进实际训练。')
    model_id=info['model_id'];bf16=torch.cuda.is_bf16_supported();dtype=torch.bfloat16 if bf16 else torch.float16
    tokenizer=AutoTokenizer.from_pretrained(out/'lora_adapter')
    loader=AutoModelForImageTextToText if 'Qwen3.5' in model_id else AutoModelForCausalLM
    quant=BitsAndBytesConfig(load_in_4bit=True,bnb_4bit_quant_type='nf4',bnb_4bit_use_double_quant=True,bnb_4bit_compute_dtype=dtype)
    model=loader.from_pretrained(model_id,quantization_config=quant,device_map={'':0},dtype=dtype,attn_implementation='eager',revision=info.get('model_revision') or 'main')
    model=PeftModel.from_pretrained(model,out/'lora_adapter');model.eval()
    rows=[];answers=[]
    with torch.inference_mode():
        for route in ['base','lora','rag_base']:
            for c in cases:
                hits=contexts[c['case_id']] if route=='rag_base' else []
                system=SYSTEM
                user=c['question']
                if route=='rag_base':
                    system+='只依据下面检索到的材料回答；若未找到答案，明确拒绝臆测。引用证据编号及报告页码。'
                    user+='\n\n检索材料：\n'+retrieval_context(hits)
                prompt=tokenizer.apply_chat_template([{'role':'system','content':system},{'role':'user','content':user}],tokenize=False,add_generation_prompt=True,enable_thinking=False)
                inputs=tokenizer(prompt,return_tensors='pt',add_special_tokens=False).to('cuda')
                if inputs['input_ids'].shape[1]>16384:raise RuntimeError('RAG证据超过上下文预算，停止而不静默截断。')
                ctx=model.disable_adapter() if route!='lora' else contextlib.nullcontext()
                started=time.time()
                with ctx:
                    output=model.generate(**inputs,max_new_tokens=384,do_sample=False,use_cache=True,pad_token_id=tokenizer.pad_token_id or tokenizer.eos_token_id)
                answer=tokenizer.decode(output[0,inputs['input_ids'].shape[1]:],skip_special_tokens=True)
                # String overlap is a diagnostic, never called factual accuracy.
                ref=normalize(c['reference_answer'].split('\n来源：')[0]);pred=normalize(answer)
                rtoks=set(tokens(ref));ptoks=set(tokens(pred));overlap=len(rtoks&ptoks)/max(1,len(rtoks))
                expected_numbers=set(re.findall(r'(?<!\d)\d[\d,]*(?:\.\d+)?',ref))
                pred_numbers=set(re.findall(r'(?<!\d)\d[\d,]*(?:\.\d+)?',pred))
                number_recall=len(expected_numbers&pred_numbers)/len(expected_numbers) if expected_numbers else None
                rows.append({'case_id':c['case_id'],'qa_id':c['qa_id'],'membership':c['membership'],'wording':c['wording'],'route':route,
                  'question':c['question'],'reference_answer':c['reference_answer'],'model_answer':answer,
                  'reference_token_recall':round(overlap,4),'literal_number_recall':number_recall,
                  'seconds':round(time.time()-started,1),'fact_correct':'','citation_correct':'','tone_correct':'','review_note':''})
                answers.append({**rows[-1],'retrieved_chunks':hits,'model_id':model_id,'prompt_sha256':__import__('hashlib').sha256(prompt.encode()).hexdigest()})
                print(route,c['case_id'],'done',flush=True)
                write_jsonl(out/'model_answers.jsonl',answers)
    write_csv(out/'comparison_to_review.csv',rows)
    (out/'evaluation_run.json').write_text(json.dumps({'status':'GENERATION_COMPLETED_REVIEW_PENDING','cases':len(cases),'routes':['base','lora','rag_base'],
      'same_model_id':model_id,'greedy':True,'reference_answers_in_prompts':False,'human_scores_pending':True},ensure_ascii=False,indent=2),encoding='utf-8')
    return rows

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--root',type=Path,default=Path(__file__).resolve().parents[1]);p.add_argument('--limit',type=int,default=0);p.add_argument('--retrieval-only',action='store_true');a=p.parse_args();run(a.root,a.limit,a.retrieval_only)
