"""Generate a one-page honest conclusion and comparison PNG from real saved runs."""
import argparse,csv,json,statistics
from pathlib import Path
from xml.sax.saxutils import escape

def make(root):
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.colors import HexColor,white
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.platypus import Paragraph,Table,TableStyle
    from reportlab.lib.styles import ParagraphStyle
    root=Path(root);out=root/'output';out.mkdir(exist_ok=True)
    prep=json.loads((out/'preparation_summary.json').read_text(encoding='utf-8'))
    scorepath=out/'comparison_to_review.csv'
    rows=list(csv.DictReader(scorepath.open(encoding='utf-8-sig'))) if scorepath.exists() else []
    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    style=ParagraphStyle('zh',fontName='STSong-Light',fontSize=10.5,leading=16,textColor=HexColor('#263548'))
    small=ParagraphStyle('small',parent=style,fontSize=8.5,leading=12)
    title=ParagraphStyle('title',parent=style,fontSize=21,leading=27)
    cellstyle=ParagraphStyle('cell',parent=style,fontSize=9,leading=13)
    def para(text,st=style):return Paragraph(escape(text).replace('\n','<br/>'),st)
    summary=[]
    for membership in ['seen','unseen']:
        for wording in ['original','paraphrase']:
            for route in ['base','lora','rag_base']:
                selected=[r for r in rows if r['membership']==membership and r['wording']==wording and r['route']==route]
                reviewed=[r for r in selected if r.get('fact_correct','').upper() in ['TRUE','FALSE']]
                summary.append({'membership':membership,'wording':wording,'route':route,'n':len(selected),
                  'token_recall':statistics.mean(float(r['reference_token_recall']) for r in selected) if selected else None,
                  'human_fact_accuracy':sum(r['fact_correct'].upper()=='TRUE' for r in reviewed)/len(reviewed) if reviewed else None,
                  'human_reviewed':len(reviewed)})
    generated=bool(rows);all_reviewed=generated and all(r.get('fact_correct','').upper() in ['TRUE','FALSE'] for r in rows)
    state='训练与72份模型回答已完成；逐条复核评分已填写' if all_reviewed else ('模型对照已生成；人工语义评分未完成' if generated else '抽检与数据准备完成；LoRA训练和模型对照尚未运行')
    filename='一页结论.pdf' if generated else '阶段结论_训练待运行.pdf'
    path=out/filename;c=canvas.Canvas(str(path),pagesize=A4);W,H=A4;c.setTitle('太古地产财报董秘问答：第三周作业B');c.setAuthor('')
    c.setFillColor(HexColor('#193754'));c.rect(0,H-112,W,112,fill=1,stroke=0)
    t=Paragraph('太古地产财报董秘问答',ParagraphStyle('head',parent=title,textColor=white));t.wrap(W-84,60);t.drawOn(c,42,H-66)
    c.setFont('STSong-Light',10);c.setFillColor(white);c.drawString(42,H-91,'第三周作业 B  |  2026-10-09  |  '+('实验结果' if generated else '阶段记录，尚非最终实验结论'))
    y=H-138
    def section(heading,text):
        nonlocal y
        c.setFont('STSong-Light',12);c.setFillColor(HexColor('#193754'));c.drawString(42,y,heading);y-=11
        p=para(text);_,h=p.wrap(W-84,1000);p.drawOn(c,42,y-h);y-=h+20
    section('完成状态',state+'。原始财报、既有抽检记录及全部模型回答保留。')
    section('材料与质量检查',f"27份太古地产年报与中期报告，2013—2026报告期；知识库3,246块，QA原始300条。分层抽检30条，六类各5条，含既有13条人工记录。{prep['content_revision_count']}条修订问法、口径或答案表述（含页码），{prep['cross_page_count']}条调整跨页证据定位。未发现可确认的凭空捏造数值；这不代表其余270条均正确。")
    section('可复现的对照设计',f"数据按同一报告页面的连接组划分：训练{prep['split_counts']['train']}、验证{prep['split_counts']['validation']}、测试{prep['split_counts']['test']}条。测试题及同页相关题不进入LoRA训练。四格为本次训练见过/未见过 × 原问法/改写；对照基础小模型、LoRA模型、向量+BM25检索后同一基础模型。")
    if generated:
        values=[]
        for route in ['base','lora','rag_base']:
            subset=[r for r in rows if r['route']==route]
            values.append((route,statistics.mean(float(r['reference_token_recall']) for r in subset)))
        review=json.loads((out/'AI_review_summary.json').read_text(encoding='utf-8')) if (out/'AI_review_summary.json').exists() else None
        if review:
            section('真实结果与错误分析','逐条严格事实通过：基础模型0/24、LoRA模型0/24、RAG模型8/24（33.3%）；RAG引用通过6/24。辅助指标“核心回答正确”RAG为13/24，不代表整段答案通过。LoRA反复生成1,066等错误金额及第10/11页；RAG仍出现单位放大、普通/基本利息倍数混淆和引用印刷页错误。')
            section('训练与实验限制','Tesla T4；Qwen3.5-2B；4-bit NF4，LoRA r=8、alpha=16、2轮，276训练/15验证样本，耗时约14分钟。每格6题，共12道独立题、24种问法、72份回答；个别改写增加子问题，不能把差异全归因泛化。252条训练数据未逐条核验，董秘口径未单独评分，模型revision未记录。评分为逐条辅助复核，待用户确认。')
        else:
            observed='；'.join(f'{r}参考词项召回{v:.1%}' for r,v in values)
            section('真实结果与限制',observed+'。词项重合不等同事实正确率；请核对逐题CSV。')
    else:
        section('已观察到的问题','抽检发现：有的chunk跨页后仍引用起始页；一题的5,604百万元属于公司本部，集团合并口径为8,121百万元。仅检查数字是否出现在原文中，无法排除主体或页码错误。自动化现已补充双页码并修订问题口径。')
    section('结论',('本轮财报金额、年份和出处问题更适合建库检索。LoRA学到引用式回答格式，却未可靠记住金额和页码；RAG虽优于本轮无检索模型，仍需表格行列还原、口径约束、短答案和页码校验。结果仅适用于本次样本和配置，不推论所有微调方法。' if generated else '当前证据支持先保证财报口径与出处准确。金额、年份及页码问题需要回到原文核对；董秘口径能否通过LoRA学会，必须等四格实验完成后判断，当前不报告训练效果。'))
    c.setFont('STSong-Light',8);c.setFillColor(HexColor('#68798a'));c.drawString(42,31,'依据：audit_30.csv、preparation_summary.json'+('、comparison_to_review.csv' if generated else '。训练未执行，无实验分数。'))
    if y<45:raise RuntimeError('PDF exceeds one-page layout budget')
    c.save()
    # Printable table is a real rendering of saved metrics, not an invented UI screenshot.
    gridpath=out/('四格对照表.pdf' if generated else '四格对照计划_待运行.pdf')
    gc=canvas.Canvas(str(gridpath),pagesize=(940,590));gc.setTitle('四格模型对照');gc.setAuthor('')
    gc.setFont('STSong-Light',21);gc.setFillColor(HexColor('#193754'));gc.drawString(36,547,'太古地产财报：四格对照'+('结果' if generated else '计划（尚未运行）'))
    gc.setFont('STSong-Light',11);gc.drawString(36,522,state)
    headings=['题组','基础模型','LoRA微调模型','向量 + BM25 + 基础模型']
    tabledata=[[para(h,cellstyle) for h in headings]]
    for membership,wording,label in [('seen','original','训练过 / 原问法'),('seen','paraphrase','训练过 / 换问法'),('unseen','original','未训练 / 原问法'),('unseen','paraphrase','未训练 / 换问法')]:
        cells=[label]
        for route in ['base','lora','rag_base']:
            s=next(x for x in summary if x['membership']==membership and x['wording']==wording and x['route']==route)
            if s['n']==0:cells.append('尚未运行\n无实验分数')
            else:cells.append(f"题数：{s['n']}\n参考词项召回：{s['token_recall']:.1%}\n复核事实通过："+(f"{s['human_fact_accuracy']:.1%}（已评{s['human_reviewed']}题）" if s['human_fact_accuracy'] is not None else '未填写'))
        tabledata.append([para(s,cellstyle) for s in cells])
    table=Table(tabledata,colWidths=[184,222,222,240],rowHeights=[42,76,76,76,76]);table.setStyle(TableStyle([
       ('BACKGROUND',(0,0),(-1,0),HexColor('#e2ebf3')),('BACKGROUND',(0,1),(0,-1),HexColor('#f3f6f9')),
       ('GRID',(0,0),(-1,-1),.5,HexColor('#b7c7d6')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(0,0),(-1,-1),14)]))
    table.wrapOn(gc,900,440);table.drawOn(gc,36,112)
    gc.setFont('STSong-Light',10);gc.setFillColor(HexColor('#68798a'));gc.drawString(36,78,'严格事实评分要求答对核心问题且无重大附加错误；每格6题，评分为逐条辅助复核，待用户确认。')
    gc.drawString(36,58,'未训练过仅指本次LoRA未见；基础模型预训练是否见过这些财报未知。');gc.save()
    import pymupdf
    for source in [path,gridpath]:
        doc=pymupdf.open(source);assert len(doc)==1
        doc[0].get_pixmap(matrix=pymupdf.Matrix(1.6,1.6)).save(source.with_suffix('.png'));doc.close()
    (out/'comparison_summary.json').write_text(json.dumps({'status':state,'generated':generated,'semantic_review_complete':all_reviewed,'cells':summary},ensure_ascii=False,indent=2),encoding='utf-8')
    print(path);print(gridpath.with_suffix('.png'))
    return path

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--root',type=Path,default=Path(__file__).resolve().parents[1]);a=p.parse_args();make(a.root)
