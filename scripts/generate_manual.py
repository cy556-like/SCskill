# -*- coding: utf-8 -*-
#!/usr/bin/env python3
"""
SCskill - 质量手册生成脚本（智能版）
=====================================
功能：基于用户填写的体系调研信息，对 IATF16949 质量手册模板进行
      章节级智能填充（公司名整体替换、人名按位置填入、质量方针/目标
      按章节替换、公司联络按字段填入），生成最终质量手册 docx。

用法：
    python generate_manual.py --survey-json '{"sv_company_name":"xxx",...}' \
                              --output-dir /path/to/output
"""
import os
import sys
import json
import re
import argparse
import tempfile
import subprocess
import shutil
from pathlib import Path
from datetime import datetime

# 确保依赖
def ensure_packages():
    import importlib
    for name in ['docx']:
        try:
            importlib.import_module(name)
        except ImportError:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install',
                                   'python-docx', '--quiet'])

ensure_packages()

from docx import Document
from docx.oxml.ns import qn
from copy import deepcopy

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
TEMPLATES_DIR = SKILL_ROOT / "templates"

# 模板文件（备用，优先使用知识库手册分类下的文件）
TEMPLATE_FILE = "IATF16949_quality_manual_template.docx"


# ===================================================================
# 调研数据加载
# ===================================================================

def load_survey_data(survey_json_str):
    """加载体系调研数据"""
    if not survey_json_str:
        return {}
    try:
        data = json.loads(survey_json_str)
        if isinstance(data, str):
            data = json.loads(data)
        return data
    except (json.JSONDecodeError, TypeError):
        return {}


def extract_company_short_name(full_name):
    """从公司全称提取简称（用于文件编号等）
    例：
        "诸暨正和金属有限公司"      -> "诸暨正和"
        "山东AAA机械制造有限公司"   -> "山东AAA"
        "深圳市XX科技股份有限公司"   -> "深圳市XX"
    """
    if not full_name:
        return 'QM'
    name = str(full_name).strip()
    # 去掉常见后缀（注意顺序：长后缀优先）
    for suffix in ['股份有限公司', '有限责任公司', '有限公司', '集团公司', '集团', '公司']:
        if name.endswith(suffix):
            name = name[:-len(suffix)]
            break
    name = name.strip()
    if not name:
        return 'QM'
    # 取前 4 个字符（汉字/字母均算 1 个）
    return name[:4] if len(name) >= 2 else name


def build_company_name_variants():
    """返回模板里所有可能出现的"原公司名"变体（长字符串优先匹配）。
    这些都是历史占位/真实公司名，需要整体替换为用户填写的公司名。"""
    return [
        '山东AAA机械制造有限公司',
        '山东纳赫汽车零部件有限公司',
        'AAA机械制造有限公司',
        'AAA机械制造',
        'AAA汽车零部件有限公司',
        'AAA有限公司',
        # 'AAA' 单独放在最后，避免误伤 'AAA-QM-2021' 这类编号
    ]


# ===================================================================
# 段落 / Run 操作工具
# ===================================================================

def is_toc_paragraph(p):
    """判断段落是否是目录条目（toc 1/2/3 样式）"""
    try:
        style_name = (p.style.name or '').lower() if p.style else ''
    except Exception:
        style_name = ''
    return 'toc' in style_name


def find_section_title(doc, prefix, keyword, exact_match=None):
    """在正文中查找章节标题段（跳过目录条目）。
    prefix: 标题前缀，如 '1.1' / '2.1' / '2.3'
    keyword: 标题中必须包含的关键字，如 '公司简介'
    exact_match: 可选，精确匹配整段文字（去掉空格后比较）
    返回段落索引，找不到返回 None。"""
    for i, p in enumerate(doc.paragraphs):
        if is_toc_paragraph(p):
            continue
        text = p.text.strip()
        if not text:
            continue
        if exact_match and text.replace(' ', '').replace('　', '') == exact_match:
            return i
        if text.startswith(prefix) and keyword in text:
            return i
    return None


def set_paragraph_text(p, new_text):
    """保留段落第一个 run 的格式，把整段文本设为 new_text，其余 run 清空。"""
    if new_text is None:
        new_text = ''
    if not p.runs:
        # 没有 run，直接添加
        p.add_run(str(new_text))
        return
    p.runs[0].text = str(new_text)
    for r in p.runs[1:]:
        r.text = ''


def clear_paragraph(p):
    """清空段落所有文本，保留段落本身。"""
    for r in list(p.runs):
        r.text = ''


def replace_text_in_paragraph(p, old, new):
    """在段落文本中做整体替换（跨 run 安全）。
    思路：先拼出整段文本 → 替换 → 重新写回第一个 run。"""
    if not p.runs or old not in p.text:
        return False
    full = p.text
    new_full = full.replace(old, new)
    if new_full == full:
        return False
    set_paragraph_text(p, new_full)
    return True


def replace_text_in_cell(cell, old, new):
    """在表格单元格内做整体替换"""
    changed = False
    for p in cell.paragraphs:
        if replace_text_in_paragraph(p, old, new):
            changed = True
    for t in cell.tables:
        for row in t.rows:
            for c in row.cells:
                if replace_text_in_cell(c, old, new):
                    changed = True
    return changed


# ===================================================================
# 1. 公司名整体替换（关键！避免 "诸暨正和金属有限公司机械制造有限公司"）
# ===================================================================

def replace_company_name_everywhere(doc, survey):
    """把模板里所有"原公司名变体"整体替换为用户填写的公司名。
    长字符串优先匹配，避免误伤。"""
    company_name = (survey.get('sv_company_name') or '').strip()
    if not company_name:
        return 0

    variants = build_company_name_variants()
    # 按长度降序，确保先匹配长的再匹配短的
    variants = sorted([v for v in variants if v], key=len, reverse=True)

    count = 0

    def replace_in_paragraph(p):
        nonlocal count
        for v in variants:
            if v in p.text:
                if replace_text_in_paragraph(p, v, company_name):
                    count += 1

    def replace_in_cell(cell):
        nonlocal count
        for p in cell.paragraphs:
            for v in variants:
                if v in p.text:
                    if replace_text_in_paragraph(p, v, company_name):
                        count += 1
        for t in cell.tables:
            for row in t.rows:
                for c in row.cells:
                    replace_in_cell(c)

    # 段落
    for p in doc.paragraphs:
        replace_in_paragraph(p)

    # 表格
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                replace_in_cell(c)

    # 页眉页脚
    for sec in doc.sections:
        for hf in [sec.header, sec.first_page_header, sec.even_page_header,
                   sec.footer, sec.first_page_footer, sec.even_page_footer]:
            if not hf:
                continue
            for p in hf.paragraphs:
                replace_in_paragraph(p)
            for t in hf.tables:
                for row in t.rows:
                    for c in row.cells:
                        replace_in_cell(c)

    return count


# ===================================================================
# 2. 文件编号替换 (AAA-QM-2021 / AAA-QM-2020 -> 简称-QM-2024)
# ===================================================================

def replace_file_numbers(doc, survey):
    """替换文件编号中的公司简称部分"""
    company_name = (survey.get('sv_company_name') or '').strip()
    short = extract_company_short_name(company_name)
    year = datetime.now().strftime("%Y")
    # 替换 AAA-QM-2021, AAA-QM-2020 等变体
    patterns = [
        (re.compile(r'AAA[-‐–—]QM[-‐–—]20\d\d'), f'{short}-QM-{year}'),
    ]

    count = 0

    def apply(p):
        nonlocal count
        for pat, repl in patterns:
            if pat.search(p.text):
                new_text = pat.sub(repl, p.text)
                if new_text != p.text:
                    set_paragraph_text(p, new_text)
                    count += 1
                    return

    def apply_cell(cell):
        for p in cell.paragraphs:
            apply(p)
        for t in cell.tables:
            for row in t.rows:
                for c in row.cells:
                    apply_cell(c)

    for p in doc.paragraphs:
        apply(p)
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                apply_cell(c)
    for sec in doc.sections:
        for hf in [sec.header, sec.first_page_header, sec.even_page_header,
                   sec.footer, sec.first_page_footer, sec.even_page_footer]:
            if not hf:
                continue
            for p in hf.paragraphs:
                apply(p)
            for t in hf.tables:
                for row in t.rows:
                    for c in row.cells:
                        apply_cell(c)
    return count


# ===================================================================
# 3. 颁布令：填入总经理人名 + 实施日期
# ===================================================================

def fill_publisher_letter(doc, survey):
    """颁布令区域：
       - 段落" 总经理：" 后填入总经理人名
       - "2022年1月1日起实施" 替换为当年日期
    """
    gm = (survey.get('sv_gm') or '').strip()
    today = datetime.now()
    today_str = f"{today.year}年{today.month}月{today.day}日"
    count = 0

    for p in doc.paragraphs:
        text = p.text
        # 颁布令的" 总经理："行（保持原有缩进格式）
        if '总经理：' in text and len(text.strip()) <= 10:
            # 把" 总经理：" 改为 " 总经理：{人名}    （签名）"
            if gm:
                new_text = f" 总经理：{gm}    （签名）"
            else:
                new_text = " 总经理：              （签名）"
            set_paragraph_text(p, new_text)
            count += 1
        # 实施日期
        m = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日起实施', text)
        if m:
            new_text = text.replace(m.group(0), f"{today_str}起实施")
            set_paragraph_text(p, new_text)
            count += 1
    return count


# ===================================================================
# 4. 公司简介（1.1）—— 智能生成，替换原 3 段
# ===================================================================

def build_company_intro_paragraphs(survey):
    """根据调研数据生成公司简介正文（1~3 段）。"""
    company_name = (survey.get('sv_company_name') or '本公司').strip()
    products = (survey.get('sv_products') or '').strip()
    location = (survey.get('sv_location') or '').strip()
    address = (survey.get('sv_address') or '').strip()
    area = (survey.get('sv_area') or '').strip()
    building_area = (survey.get('sv_building_area') or '').strip()
    staff_total = (survey.get('sv_staff_total') or '').strip()
    staff_mgmt = (survey.get('sv_staff_mgmt') or '').strip()
    staff_edu = (survey.get('sv_staff_edu') or '').strip()
    equipment = (survey.get('sv_equipment') or '').strip()
    customers = (survey.get('sv_customers') or '').strip()
    purpose = (survey.get('sv_purpose') or '').strip()

    # 第 1 段：基本信息
    para1_parts = []
    para1_parts.append(f"    {company_name}")
    if location:
        para1_parts.append(f"位于{location}")
    if address:
        para1_parts.append(f"，地址：{address}")
    if products:
        para1_parts.append(f"，是一家专注于{products}的研发、生产与销售的专业制造商")
    else:
        para1_parts.append("，是一家专业制造商")
    para1_parts.append("。")
    para1 = ''.join(para1_parts)

    # 第 2 段：规模信息
    para2_parts = []
    if area or building_area:
        size_bits = []
        if area:
            size_bits.append(f"占地面积{area}")
        if building_area:
            size_bits.append(f"建筑面积{building_area}")
        para2_parts.append(f"公司{ '，'.join(size_bits) }；")
    if staff_total:
        para2_parts.append(f"现有员工{staff_total}人")
        if staff_mgmt:
            para2_parts.append(f"，其中管理及技术人员{staff_mgmt}人")
        if staff_edu:
            para2_parts.append(f"，{staff_edu}")
        para2_parts.append("；")
    if equipment:
        para2_parts.append(f"主要生产设备：{equipment}；")
    para2 = '    ' + ''.join(para2_parts) if para2_parts else ''

    # 第 3 段：客户/宗旨
    para3_parts = []
    if customers:
        para3_parts.append(f"主要客户包括{customers}。")
    if purpose:
        para3_parts.append(f"公司经营理念：{purpose}。")
    para3 = '    ' + ''.join(para3_parts) if para3_parts else ''

    return [p for p in [para1, para2, para3] if p]


def fill_company_intro(doc, survey):
    """替换 1.1 公司简介小节的内容段落。
    定位策略：找到"1.1公司简介"标题段，向下找到第一段正文，
    把它替换为生成的新段落（最多 3 段）。"""
    new_paras = build_company_intro_paragraphs(survey)
    if not new_paras:
        return 0

    # 找到"1.1公司简介"标题（跳过目录条目）
    target_idx = find_section_title(doc, '1.1', '公司简介')
    if target_idx is None:
        return 0

    # 找到紧接着的正文段（跳过空段）
    body_indices = []
    for j in range(target_idx + 1, min(target_idx + 8, len(doc.paragraphs))):
        t = doc.paragraphs[j].text.strip()
        # 遇到下一个标题就停止
        if re.match(r'^\d+\.\d+', t) or t.startswith('1.2') or t.startswith('2.'):
            break
        if t:
            body_indices.append(j)

    if not body_indices:
        return 0

    # 把第一段写入 body_indices[0]
    set_paragraph_text(doc.paragraphs[body_indices[0]], new_paras[0])
    count = 1

    # 如果有更多新段落，需要追加；为简化，把后续内容拼到第二段（如果有）
    if len(new_paras) > 1 and len(body_indices) >= 2:
        set_paragraph_text(doc.paragraphs[body_indices[1]],
                           '\n'.join(new_paras[1:]))
        count += 1
        # 清空第 3 段（如果有）
        for j in body_indices[2:]:
            t = doc.paragraphs[j].text.strip()
            if '公司经营理念' in t or '研发实力' in t or '设计理念' in t:
                clear_paragraph(doc.paragraphs[j])
                count += 1
    else:
        # 只有一段正文，把所有新内容拼进去
        if len(new_paras) > 1:
            current = doc.paragraphs[body_indices[0]].text
            set_paragraph_text(doc.paragraphs[body_indices[0]],
                               current + '\n' + '\n'.join(new_paras[1:]))

    return count


# ===================================================================
# 5. 公司联络（1.2）—— 按字段填入
# ===================================================================

def fill_company_contact(doc, survey):
    """替换 1.2 公司联络区域的地址/邮编/电话/传真/网址"""
    address = (survey.get('sv_address') or '').strip()
    phone = (survey.get('sv_phone') or survey.get('sv_mobile') or '').strip()
    fax = (survey.get('sv_fax') or '').strip()
    count = 0

    # 找到"1.2公司联络"标题（跳过目录条目）
    contact_start = find_section_title(doc, '1.2', '公司联络')
    if contact_start is None:
        return 0

    # 在标题后 10 段范围内替换
    for j in range(contact_start + 1, min(contact_start + 12, len(doc.paragraphs))):
        p = doc.paragraphs[j]
        text = p.text
        # 遇到下一个章节标题，停止
        if re.match(r'^\d+\.\d+', text.strip()) or text.strip().startswith('2.'):
            break
        if not text.strip():
            continue

        # 地址
        if text.strip().startswith('地址'):
            new_text = f'地址：{address}' if address else '地址：'
            set_paragraph_text(p, new_text)
            count += 1
        # 邮编（调研表里没有邮编字段，保留原样或清空）
        elif text.strip().startswith('邮编'):
            # 不动，保留模板原邮编或清空（这里保留原样）
            pass
        # 电话
        elif text.strip().startswith('电话'):
            new_text = f'电话：{phone}' if phone else '电话：'
            set_paragraph_text(p, new_text)
            count += 1
        # 传真
        elif text.strip().startswith('传真'):
            new_text = f'传真：{fax}' if fax else '传真：'
            set_paragraph_text(p, new_text)
            count += 1
        # 网址（保留）
        elif text.strip().startswith('网址'):
            pass

    return count


# ===================================================================
# 6. 质量方针（2.1）—— 按章节替换
# ===================================================================

def fill_quality_policy(doc, survey):
    """替换 2.1 质量方针章节的正文段落。
    原模板：段落104="全员具备质量意识，提供高质量产品和服务；"
            段落105="持续改进，使质量水平处于中国第一。"
    用户填写的方针可能是一行或多行，按行填入。"""
    policy = (survey.get('sv_quality_policy') or '').strip()
    if not policy:
        return 0

    # 按换行分割用户方针
    lines = [l.strip() for l in re.split(r'[\n\r；;]+', policy) if l.strip()]
    if not lines:
        return 0

    # 找到"2.1质量方针"标题（跳过目录条目）
    target_idx = find_section_title(doc, '2.1', '质量方针')
    if target_idx is None:
        return 0

    # 收集标题后的正文段（跳过空段、停止于"2.2企业责任方针"或下一个 2.x 标题）
    body_indices = []
    for j in range(target_idx + 1, min(target_idx + 10, len(doc.paragraphs))):
        t = doc.paragraphs[j].text.strip()
        if not t:
            continue
        # 遇到下一个章节标题，停止
        if re.match(r'^2\.\d+', t) or '企业责任方针' in t or '质量目标' in t or \
           t.startswith('"质量方针"') or t.startswith('“质量方针”'):
            break
        body_indices.append(j)

    if not body_indices:
        return 0

    # 第一行写到第一段，其余行拼到第二段（保留原模板两段结构）
    set_paragraph_text(doc.paragraphs[body_indices[0]], '    ' + lines[0])
    count = 1
    if len(body_indices) >= 2:
        if len(lines) > 1:
            set_paragraph_text(doc.paragraphs[body_indices[1]], '    ' + '；'.join(lines[1:]))
        else:
            clear_paragraph(doc.paragraphs[body_indices[1]])
        count += 1

    # 清空第 3 段及之后（如果有）
    for j in body_indices[2:]:
        clear_paragraph(doc.paragraphs[j])
        count += 1

    return count


# ===================================================================
# 7. 质量目标（2.3）—— 按章节替换
# ===================================================================

def fill_quality_goal(doc, survey):
    """替换 2.3 质量目标章节的正文段落。
    原模板：116-119 段是模板的目标描述。"""
    goal = (survey.get('sv_quality_goal') or '').strip()
    if not goal:
        return 0

    # 找到"2.3质量目标"标题（跳过目录条目）
    target_idx = find_section_title(doc, '2.3', '质量目标')
    if target_idx is None:
        return 0

    # 收集标题后的正文段
    body_indices = []
    for j in range(target_idx + 1, min(target_idx + 12, len(doc.paragraphs))):
        t = doc.paragraphs[j].text.strip()
        if not t:
            continue
        # 遇到下一个章节标题，停止（如"3.公司组织图"）
        if re.match(r'^\d+\.', t) and not t.startswith('2.3'):
            break
        # "质量方针和质量目标/指标..." 这种总结性段落也停止
        if '质量方针和质量目标' in t and '是公司质量管理体系' in t:
            break
        body_indices.append(j)

    if not body_indices:
        return 0

    # 按行分割用户目标
    lines = [l.strip() for l in re.split(r'[\n\r]+', goal) if l.strip()]

    # 第一行写到第一段
    set_paragraph_text(doc.paragraphs[body_indices[0]], '    ' + lines[0])
    count = 1

    # 第二行写到第二段（如果有）
    if len(body_indices) >= 2:
        if len(lines) > 1:
            set_paragraph_text(doc.paragraphs[body_indices[1]], '    ' + '；'.join(lines[1:]))
        else:
            clear_paragraph(doc.paragraphs[body_indices[1]])
        count += 1

    # 清空后续段（保留总结性段落）
    for j in body_indices[2:]:
        t = doc.paragraphs[j].text.strip()
        if '质量方针和质量目标' in t:
            continue  # 保留总结段
        clear_paragraph(doc.paragraphs[j])
        count += 1

    return count


# ===================================================================
# 8. 公司宗旨（公司经营理念）—— 替换原"凝聚志同道合..."段
# ===================================================================

def fill_company_purpose(doc, survey):
    """替换公司经营理念段落。原模板段落94："公司经营理念：凝聚志同道合的人才..." """
    purpose = (survey.get('sv_purpose') or '').strip()
    if not purpose:
        return 0

    count = 0
    for p in doc.paragraphs:
        text = p.text.strip()
        if text.startswith('公司经营理念') and ('凝聚志同道合' in text or '制造安全可靠' in text):
            set_paragraph_text(p, f'公司经营理念：{purpose}')
            count += 1
            break
    return count


# ===================================================================
# 9. 表格 0：填入"制订/审核/批准"行的人名 + 实施日期
# ===================================================================

def fill_doc_control_table(doc, survey):
    """表格 0（文件控制表）：
       row[1] = ['实施日期', '制  订', '审  核', '批  准']
       row[2] = ['2022.1.1', '', '', '']  -> 填入日期 + ISO办主任/管代/总经理
    """
    iso_head = (survey.get('sv_iso_office_head') or '').strip()
    mgmt_rep = (survey.get('sv_mgmt_rep') or '').strip()
    gm = (survey.get('sv_gm') or '').strip()
    today = datetime.now()
    today_str = f"{today.year}.{today.month}.{today.day}"

    count = 0
    if not doc.tables:
        return 0

    t0 = doc.tables[0]
    if len(t0.rows) < 3 or len(t0.columns) < 4:
        return 0

    # 检查 row[1] 是否为 "实施日期|制订|审核|批准"
    header = [c.text.strip() for c in t0.rows[1].cells]
    if not any('制订' in h or '制  订' in h for h in header):
        return 0

    # row[2][0] = 实施日期
    set_paragraph_text(t0.rows[2].cells[0].paragraphs[0], today_str)
    count += 1
    # row[2][1] = 制订 -> ISO办主任
    if iso_head:
        set_paragraph_text(t0.rows[2].cells[1].paragraphs[0], iso_head)
        count += 1
    # row[2][2] = 审核 -> 管理者代表
    if mgmt_rep:
        set_paragraph_text(t0.rows[2].cells[2].paragraphs[0], mgmt_rep)
        count += 1
    # row[2][3] = 批准 -> 总经理
    if gm:
        set_paragraph_text(t0.rows[2].cells[3].paragraphs[0], gm)
        count += 1

    return count


# ===================================================================
# 主流程
# ===================================================================

def find_template():
    """查找模板文件：优先知识库手册分类，其次 SCskill 内置模板。
    支持 .docx（直接编辑）和 .doc（需转换）。"""
    project_root = SKILL_ROOT.parent
    manual_dir = project_root / "data" / "documents" / "agent_dfmea-risk-agent" / "手册"

    # 1. 知识库手册分类下的 .docx
    if manual_dir.exists():
        for f in sorted(os.listdir(str(manual_dir))):
            if f.lower().endswith('.docx') and not f.startswith('~$'):
                print(f"[INFO] 从知识库手册分类找到 .docx 模板: {f}")
                return manual_dir / f, False  # need_convert=False

    # 2. SCskill 内置 .docx
    builtin = TEMPLATES_DIR / TEMPLATE_FILE
    if builtin.exists():
        print(f"[INFO] 使用 SCskill 内置模板: {TEMPLATE_FILE}")
        return builtin, False

    # 3. 知识库手册分类下的 .doc（需转换）
    if manual_dir.exists():
        for f in sorted(os.listdir(str(manual_dir))):
            if f.lower().endswith('.doc') and not f.startswith('~$'):
                print(f"[INFO] 从知识库手册分类找到 .doc 模板: {f}，需转换")
                return manual_dir / f, True

    return None, False


def convert_doc_to_docx(doc_path):
    """用 LibreOffice 或 Word COM 把 .doc 转成 .docx，返回临时 .docx 路径。"""
    tmp_dir = tempfile.mkdtemp(prefix='doc2docx_')
    # 方法 1：LibreOffice
    soffice_paths = [
        'soffice',
        '/usr/bin/soffice',
        'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
        'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    ]
    for sp in soffice_paths:
        try:
            result = subprocess.run(
                [sp, '--headless', '--convert-to', 'docx',
                 str(doc_path), '--outdir', tmp_dir],
                capture_output=True, text=True, timeout=120
            )
            if result.returncode == 0:
                basename = os.path.splitext(os.path.basename(str(doc_path)))[0]
                converted = os.path.join(tmp_dir, basename + '.docx')
                if os.path.exists(converted):
                    print(f"[INFO] LibreOffice 转换成功: {converted}")
                    return converted
        except Exception as e:
            print(f"[WARN] LibreOffice ({sp}) 转换失败: {e}")

    # 方法 2：Windows Word COM
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        d = word.Documents.Open(str(doc_path))
        out = os.path.join(tmp_dir, 'converted.docx')
        d.SaveAs2(out, FileFormat=16)
        d.Close()
        word.Quit()
        pythoncom.CoUninitialize()
        print(f"[INFO] Word COM 转换成功: {out}")
        return out
    except Exception as e:
        print(f"[WARN] Word COM 转换失败: {e}")

    return None


def generate_manual(survey_data, output_dir):
    """主生成函数"""
    # 1. 查找模板
    template_path, need_convert = find_template()
    if template_path is None:
        return {
            "status": "error",
            "message": "未找到模板文件。请在内部知识库[手册]分类下上传 .docx/.doc 模板，"
                       "或确保 SCskill/templates/IATF16949_quality_manual_template.docx 存在。"
        }

    # 2. .doc 转 .docx
    actual_template = template_path
    if need_convert:
        print(f"[INFO] 正在将 .doc 模板转换为 .docx ...")
        converted = convert_doc_to_docx(template_path)
        if not converted:
            return {
                "status": "error",
                "message": "无法将 .doc 模板转换为 .docx。请安装 LibreOffice，或上传 .docx 模板。"
            }
        actual_template = Path(converted)

    print(f"[INFO] 加载模板: {actual_template}")
    doc = Document(str(actual_template))

    # 3. 执行智能填充（按顺序）
    stats = {}

    print(f"[INFO] [1/8] 整体替换公司名 ...")
    stats['公司名替换'] = replace_company_name_everywhere(doc, survey_data)

    print(f"[INFO] [2/8] 替换文件编号 ...")
    stats['文件编号替换'] = replace_file_numbers(doc, survey_data)

    print(f"[INFO] [3/8] 填充颁布令（总经理+日期）...")
    stats['颁布令填充'] = fill_publisher_letter(doc, survey_data)

    print(f"[INFO] [4/8] 填充公司简介 ...")
    stats['公司简介填充'] = fill_company_intro(doc, survey_data)

    print(f"[INFO] [5/8] 填充公司联络 ...")
    stats['公司联络填充'] = fill_company_contact(doc, survey_data)

    print(f"[INFO] [6/8] 填充质量方针 ...")
    stats['质量方针填充'] = fill_quality_policy(doc, survey_data)

    print(f"[INFO] [7/8] 填充质量目标 ...")
    stats['质量目标填充'] = fill_quality_goal(doc, survey_data)

    print(f"[INFO] [8/8] 填充公司经营理念 + 文件控制表人名 ...")
    stats['公司宗旨填充'] = fill_company_purpose(doc, survey_data)
    stats['文件控制表人名'] = fill_doc_control_table(doc, survey_data)

    total = sum(stats.values())
    print(f"[INFO] 智能填充完成，共修改 {total} 处：{stats}")

    # 4. 保存输出
    company_name = (survey_data.get('sv_company_name') or '企业').strip()
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', company_name)
    today = datetime.now().strftime("%Y%m%d")
    filename = f"质量管理手册_{safe_name}_{today}.docx"
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, filename)
    doc.save(output_path)
    print(f"[OK] 质量手册已生成: {output_path}")

    return {
        "status": "success",
        "filename": filename,
        "file_path": output_path,
        "company_name": company_name,
        "fill_stats": stats,
        "total_replacements": total,
    }


def parse_args():
    parser = argparse.ArgumentParser(description="质量手册智能生成器")
    parser.add_argument("--survey-json", required=False, default=None,
                        help="体系调研数据（JSON 字符串）")
    parser.add_argument("--output-dir", required=False, default=".",
                        help="输出目录")
    return parser.parse_args()


def main():
    if sys.platform == "win32":
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

    args = parse_args()
    survey = load_survey_data(args.survey_json) if args.survey_json else {}

    if not survey:
        print("[ERROR] 未提供体系调研数据")
        result = {"status": "error", "message": "未提供体系调研数据"}
    else:
        result = generate_manual(survey, args.output_dir)

    print(f"\n[RESULT_JSON]")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
