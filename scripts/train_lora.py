"""LoRA on Qwen3.5-2B. Requires a CUDA Colab runtime; no API key."""
import argparse,hashlib,importlib.metadata,json,math,random,time
from pathlib import Path
from prepare_assignment import read_jsonl

def run(root,model_id='Qwen/Qwen3.5-2B',epochs=2,reviewed_only=False,max_length=1024):
    import torch
    from transformers import AutoTokenizer,AutoModelForImageTextToText,AutoModelForCausalLM,BitsAndBytesConfig,Trainer,TrainingArguments,set_seed
    from peft import LoraConfig,get_peft_model,prepare_model_for_kbit_training
    from datasets import Dataset
    root=Path(root);set_seed(20261009)
    if not torch.cuda.is_available():raise RuntimeError('需要Colab GPU：运行时 -> 更改运行时类型 -> T4 GPU。本机CPU不启动训练。')
    suffix='_reviewed_only' if reviewed_only else ''
    rawtrain=read_jsonl(root/f'data/prepared/train{suffix}.jsonl')
    rawval=read_jsonl(root/f'data/prepared/validation{suffix}.jsonl')
    if len(rawtrain)<5:raise RuntimeError('训练样本少于5，请检查抽检及分割结果。')
    bf16=torch.cuda.is_bf16_supported();dtype=torch.bfloat16 if bf16 else torch.float16
    quant=BitsAndBytesConfig(load_in_4bit=True,bnb_4bit_quant_type='nf4',bnb_4bit_use_double_quant=True,bnb_4bit_compute_dtype=dtype)
    tokenizer=AutoTokenizer.from_pretrained(model_id)
    if tokenizer.pad_token_id is None:tokenizer.pad_token=tokenizer.eos_token
    tokenizer.padding_side='right'
    loader=AutoModelForImageTextToText if 'Qwen3.5' in model_id else AutoModelForCausalLM
    model=loader.from_pretrained(model_id,quantization_config=quant,device_map={'':0},dtype=dtype,attn_implementation='eager')
    model.config.use_cache=False
    if hasattr(model.config,'text_config'):model.config.text_config.use_cache=False
    model=prepare_model_for_kbit_training(model,use_gradient_checkpointing=True)
    endings={'q_proj','k_proj','v_proj','o_proj','gate_proj','up_proj','down_proj'}
    targets=[n for n,_ in model.named_modules() if n.split('.')[-1] in endings and 'visual' not in n and 'vision' not in n and 'mtp' not in n]
    if not targets:raise RuntimeError('未识别到语言模型LoRA模块，停止以防训练错误部分。')
    model=get_peft_model(model,LoraConfig(r=8,lora_alpha=16,lora_dropout=.05,target_modules=targets,bias='none',task_type='CAUSAL_LM'))
    model.print_trainable_parameters()
    rejected=[]
    def encode(records):
        rows=[]
        for r in records:
            messages=r['messages'];prefix=tokenizer.apply_chat_template(messages[:-1],tokenize=False,add_generation_prompt=True,enable_thinking=False)
            full=tokenizer.apply_chat_template(messages,tokenize=False,add_generation_prompt=False,enable_thinking=False)
            p=tokenizer(prefix,add_special_tokens=False)['input_ids'];f=tokenizer(full,add_special_tokens=False)['input_ids']
            if f[:len(p)]!=p:raise RuntimeError(f"{r['id']} chat template prefix mismatch，不能可靠遮罩提示词。")
            if len(f)>max_length:rejected.append({'id':r['id'],'reason':'超长而未静默截断','tokens':len(f)});continue
            rows.append({'input_ids':f,'attention_mask':[1]*len(f),'labels':[-100]*len(p)+f[len(p):]})
        return Dataset.from_list(rows)
    training=encode(rawtrain);validation=encode(rawval) if rawval else None
    if len(training)<5:raise RuntimeError('编码后训练样本不足。')
    if validation is not None and len(validation)==0:validation=None
    def collate(batch):
        n=max(len(x['input_ids']) for x in batch)
        return {'input_ids':torch.tensor([x['input_ids']+[tokenizer.pad_token_id]*(n-len(x['input_ids'])) for x in batch]),
                'attention_mask':torch.tensor([x['attention_mask']+[0]*(n-len(x['input_ids'])) for x in batch]),
                'labels':torch.tensor([x['labels']+[-100]*(n-len(x['labels'])) for x in batch])}
    destination=root/'output/lora_adapter';destination.mkdir(parents=True,exist_ok=True)
    args=TrainingArguments(output_dir=str(root/'output/checkpoints'),num_train_epochs=epochs,per_device_train_batch_size=1,
      per_device_eval_batch_size=1,gradient_accumulation_steps=8,learning_rate=1e-4,warmup_steps=max(1,math.ceil(math.ceil(len(training)/8)*epochs*.1)),weight_decay=.01,
      fp16=not bf16,bf16=bf16,gradient_checkpointing=True,gradient_checkpointing_kwargs={'use_reentrant':False},
      logging_steps=5,eval_strategy='epoch' if validation is not None else 'no',save_strategy='epoch',
      save_total_limit=2,report_to='none',optim='paged_adamw_8bit',remove_unused_columns=False,seed=20261009)
    started=time.time();trainer=Trainer(model=model,args=args,train_dataset=training,eval_dataset=validation,data_collator=collate)
    trainer.train();trainer.save_model(str(destination));tokenizer.save_pretrained(str(destination))
    packages={p:importlib.metadata.version(p) for p in ['torch','transformers','peft','datasets','accelerate','bitsandbytes']}
    runinfo={'status':'TRAINING_COMPLETED','model_id':model_id,'model_revision':getattr(model.config,'_commit_hash',None),
      'gpu':torch.cuda.get_device_name(0),'epochs':epochs,'max_length':max_length,'reviewed_only':reviewed_only,
      'train_samples_used':len(training),'validation_samples_used':len(validation) if validation is not None else 0,
      'unreviewed_training_records':sum(r.get('review_status')!='ai_pdf_reviewed' for r in rawtrain),
      'r':8,'alpha':16,'learning_rate':1e-4,'seconds':round(time.time()-started,1),'packages':packages,
      'training_data_sha256':hashlib.sha256((root/f'data/prepared/train{suffix}.jsonl').read_bytes()).hexdigest(),
      'rejected_oversize_samples':rejected,'loss_history':trainer.state.log_history,
      'test_used_for_training':False,'prompt_tokens_masked':True,'source_text_in_training_prompt':False}
    (root/'output/training_run.json').write_text(json.dumps(runinfo,ensure_ascii=False,indent=2),encoding='utf-8')
    print('训练完成：',destination);return runinfo

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--root',type=Path,default=Path(__file__).resolve().parents[1]);p.add_argument('--model',default='Qwen/Qwen3.5-2B');p.add_argument('--epochs',type=float,default=2);p.add_argument('--reviewed-only',action='store_true');a=p.parse_args();run(a.root,a.model,a.epochs,a.reviewed_only)
