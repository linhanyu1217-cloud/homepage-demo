# 太古地产财报董秘：第三周作业B

## 完成状态
原始材料27份，现有语料3,246块，问答300条。已做30条AI原始PDF抽检，并保留用户既有13条人工记录。LoRA训练及三路线对照尚未执行；不报告虚构准确率。

## 使用
1. 在Colab打开根目录的ipynb，选择T4 GPU。
2. 顺序运行单元格，上传随附代码zip。
3. 自动完成数据分割、LoRA训练、基础/微调/RAG模型生成和附件导出。
4. 逐题CSV中的事实评分可以填写TRUE/FALSE，再运行报告脚本更新PDF和PNG。

命令行：
```sh
python scripts/prepare_assignment.py
python scripts/train_lora.py
python scripts/evaluate_models.py
python scripts/make_report.py
```

`data/qa/qa_raw.jsonl`为原始QA；`human_review_original.csv`是人工历史记录；`output/audit_30.csv`、`audit_30.md`是AI复核与修订记录。`data/qa/qa_clean.jsonl`含原始与修订答案及双页码。

## 方法与限制
抽样按六类各5条，包含已知问题，不能推断整体错误率。同一来源页面形成连接组，train/validation/test为{'train': 276, 'validation': 15, 'test': 9}。四格“未见”指本次LoRA未训练，不意味着基础模型预训练未见。默认允许未逐条抽检的样本做教学训练，另附reviewed_only数据。来源页码优先印刷页，同时保留PDF物理页。没有印刷页码时明确使用章节与PDF页。

向量快照来自已有BAAI/bge-large-zh-v1.5向量库，未重新编码，最大512token有截断风险；BM25采用繁简归一的汉字二元词项，并用RRF合并。问题自身解析年报/中报及年份，参考答案不输入RAG。

自动词项/字面数字匹配仅作诊断，不等同语义事实正确率；错误、拒答和未评分必须如实记录。原有文本抽取没有完整还原复杂表格行列，原始PDF视觉核验覆盖本轮30条，而非全库。

## 数据来源与重新下载
[太古地产官方财务报告](https://ir.swireproperties.com/zh-hk/ir/reports.php)。`data/source_manifest.json`记录27份报告文件名、期别、原始SHA256、官网对应链接。`python scripts/download_reports.py`下载到忽略目录。官网文件可能与本地繁简转换或重新排版文件不同，必须重新抽取并核对页码，不能沿用本次标注。

## 造数据脚本
保留原来的`01_extract_pdf.py`、`02_clean_text.py`、`05_generate_qa.py`。如需重新生成QA，复制`.env.example`为`.env`填写自己的服务配置后运行`05_generate_qa.py`；训练和对照本身不需要API key。造数据脚本的抽检CSV重写逻辑已改为按id保留已有人工字段。本包不附原PDF、个人路径、密钥和原始Chroma数据库。

## 技术参考
[Qwen3.5-2B官方模型](https://huggingface.co/Qwen/Qwen3.5-2B)
[Transformers Qwen3.5](https://huggingface.co/docs/transformers/model_doc/qwen3_5)
[PEFT量化训练](https://huggingface.co/docs/peft/developer_guides/quantization)

已在Colab Tesla T4完成两轮训练（839.8秒）与72份模型回答。实际版本见requirements-colab.txt；原始训练日志model_revision为空，不能声称锁定模型revision。

## 本轮实际结果
逐条辅助复核：严格事实通过基础0/24、LoRA 0/24、RAG 8/24；RAG核心回答正确13/24、引用通过6/24。严格评分与核心指标分开。保留原始未评分CSV和全部模型回答，复核CSV记录逐条理由，未标为用户人工签核。个别改写新增子问题，四格对比不能全归因泛化。当前评分和失败不能被损失下降替代。
