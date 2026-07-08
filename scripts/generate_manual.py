#!/usr/bin/env python3
"""
SCskill - 质量手册生成脚本
============================
功能：从模板文件生成质量手册，根据用户填写的体系调研信息替换内容
用法：python generate_manual.py --survey-json '{"company_name":"xxx",...}' --output-dir /path/to/output
"""
import os
import sys
import json
import re
import argparse
import tempfile
import subprocess
from pathlib import Path
from datetime import datetime

# 确保依赖
def ensure_packages():
    import importlib
    for name in ['docx']:
        try:
            importlib.import_module(name)
        except ImportError:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'python-docx', '--quiet'])

ensure_packages()

from docx import Document
from docx.shared import Pt

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_ROOT = SCRIPT_DIR.parent
TEMPLATES_DIR = SKILL_ROOT / "templates"

# 模板文件
TEMPLATE_FILE = "IATF16949_quality_manual_template.docx"


def load_survey_data(survey_json_str):
    """加载体系调研数据"""
    if not survey_json_str:
        return {}
    try:
        return json.loads(survey_json_str)
    except json.JSONDecodeError:
        return {}


def build_replacement_map(survey):
    """根据体系调研数据构建替换映射"""
    company_name = survey.get('sv_company_name', '')
    certs = survey.get('sv_certs', [])
    
    replacements = {}
    
    # 公司基本信息
    replacements['山东AAA机械制造有限公司'] = company_name or '企业名称'
    replacements['AAA'] = company_name or '企业名称'
    replacements['AAA成立于2018年5月4日'] = f'{company_name}成立于{survey.get("sv_area", "")}'
    
    # 公司简介
    company_intro = f'{company_name}成立于{survey.get("sv_area", "")}年'
    if survey.get('sv_products'):
        company_intro += f'，是一家专注于生产{survey["sv_products"]}的生产制造商。'
    if survey.get('sv_staff_total'):
        company_intro += f'公司现有正式员工{survey["sv_staff_total"]}人'
    if survey.get('sv_staff_mgmt'):
        company_intro += f'，管理和技术人员{survey["sv_staff_mgmt"]}人。'
    if survey.get('sv_equipment'):
        company_intro += f'{survey["sv_equipment"]}'
    
    replacements['    AAA成立于2018年5月4日，是一家专注于生产半挂车车轴、抗翻空悬、欧式空气悬架、超级盘刹制动器、智能空气悬挂控制系统（iCAS)、EBS制动系统及相关配件的生产制造商。'] = f'    {company_intro}'
    
    # 地址
    address = survey.get('sv_address', '')
    replacements['XX省XX市XX区XX路11号'] = address or '（请填写公司地址）'
    replacements['地址：XX省XX市XX区XX路11号'] = f'地址：{address}' if address else '地址：'
    
    # 电话
    phone = survey.get('sv_phone', '') or survey.get('sv_mobile', '')
    replacements['1XXXXXXXXXX'] = phone or '（请填写联系电话）'
    replacements['电话：1XXXXXXXXXX'] = f'电话：{phone}' if phone else '电话：'
    
    # 邮编
    if survey.get('sv_area'):
        replacements['276200'] = ''
    
    # 质量方针
    quality_policy = survey.get('sv_quality_policy', '')
    if quality_policy:
        replacements['全员具备质量意识，提供高质量产品和服务；\n持续改进，使质量水平处于中国第一。'] = quality_policy
    
    # 质量目标
    quality_goal = survey.get('sv_quality_goal', '')
    if quality_goal:
        # 找到质量目标段落替换
        pass  # 在段落级别处理
    
    # 公司宗旨
    purpose = survey.get('sv_purpose', '')
    if purpose:
        replacements['凝聚志同道合的人才，制造安全可靠的产品，提供客户满意'] = purpose
    
    # 管理者代表
    mgmt_rep = survey.get('sv_mgmt_rep', '')
    if mgmt_rep:
        replacements['管理者代表'] = f'管理者代表：{mgmt_rep}'
    
    # 总经理
    gm = survey.get('sv_gm', '')
    
    # 董事长
    chairman = survey.get('sv_chairman', '')
    
    # 认证标准
    cert_str = '、'.join(certs) if certs else 'IATF16949:2016&ISO9001:2015'
    
    return replacements, {
        'company_name': company_name,
        'address': address,
        'phone': phone,
        'quality_policy': quality_policy,
        'quality_goal': quality_goal,
        'purpose': purpose,
        'mgmt_rep': mgmt_rep,
        'gm': gm,
        'chairman': chairman,
        'certs': cert_str,
        'products': survey.get('sv_products', ''),
        'staff_total': survey.get('sv_staff_total', ''),
        'staff_mgmt': survey.get('sv_staff_mgmt', ''),
        'area': survey.get('sv_area', ''),
        'filler_name': survey.get('sv_filler_name', ''),
        'filler_phone': survey.get('sv_filler_phone', ''),
    }


def replace_in_runs(paragraph, replacements):
    """在段落 runs 中执行替换"""
    for run in paragraph.runs:
        if run.text:
            for old, new in replacements.items():
                if old in run.text:
                    run.text = run.text.replace(old, new)


def replace_in_cell(cell, replacements):
    """在表格单元格中执行替换"""
    for p in cell.paragraphs:
        replace_in_runs(p, replacements)
    for table in cell.tables:
        for row in table.rows:
            for c in row.cells:
                replace_in_cell(c, replacements)


def replace_in_header_footer(hf, replacements):
    """在页眉页脚中执行替换"""
    if not hf:
        return
    for p in hf.paragraphs:
        replace_in_runs(p, replacements)
    for table in hf.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_in_cell(cell, replacements)


def generate_manual(survey_data, output_dir):
    """生成质量手册"""
    
    # 1. 查找模板文件
    # 优先从知识库"手册"分类目录下查找模板（.docx 或 .doc）
    project_root = SKILL_ROOT.parent
    manual_dir = project_root / "data" / "documents" / "agent_dfmea-risk-agent" / "手册"
    template_path = None
    need_convert = False  # 是否需要从 .doc 转换
    
    if manual_dir.exists():
        # 先找 .docx
        for f in os.listdir(str(manual_dir)):
            if f.lower().endswith('.docx'):
                template_path = manual_dir / f
                print(f"[INFO] 从知识库"手册"分类找到模板: {f}")
                break
        # 再找 .doc
        if template_path is None:
            for f in os.listdir(str(manual_dir)):
                if f.lower().endswith('.doc'):
                    doc_file = manual_dir / f
                    print(f"[INFO] 从知识库"手册"分类找到 .doc 模板: {f}")
                    # 尝试用 olefile 直接读取
                    try:
                        import olefile
                        from langchain_core.documents import Document
                        import re as _re
                        
                        ole = olefile.OleFileIO(str(doc_file))
                        if ole.exists('WordDocument'):
                            stream = ole.openstream('WordDocument')
                            data = stream.read()
                            decoded = data.decode('utf-16-le', errors='ignore')
                            chunks = _re.findall(r'[一-鿿 -~　-〿＀-￯]+', decoded)
                            ole.close()
                            if chunks and len('\n'.join(chunks)) > 100:
                                # olefile 能读到内容，但 python-docx 不能直接编辑 .doc
                                # 需要转成 .docx 才能用 python-docx 编辑
                                need_convert = True
                                template_path = doc_file
                                print(f"[INFO] .doc 文件可读取，但需要转换为 .docx 进行编辑")
                    except Exception as e:
                        print(f"[WARN] olefile 读取 .doc 失败: {e}")
                    # 如果 olefile 失败，尝试用二进制方式
                    if template_path is None:
                        try:
                            with open(str(doc_file), 'rb') as rf:
                                raw = rf.read()
                            decoded = raw.decode('utf-16-le', errors='ignore')
                            chunks_list = re.findall(r'[一-鿿 -~　-〿＀-￯]+', decoded)
                            if chunks_list and len('\n'.join(chunks_list)) > 100:
                                need_convert = True
                                template_path = doc_file
                                print(f"[INFO] .doc 文件可读取（二进制方式），需要转换为 .docx")
                        except Exception:
                            pass
                    break
    
    # 如果知识库没有，查找 SCskill 自带的模板
    if template_path is None:
        builtin_path = TEMPLATES_DIR / TEMPLATE_FILE
        if builtin_path.exists():
            template_path = builtin_path
            print(f"[INFO] 使用内置模板: {TEMPLATE_FILE}")
    
    # 最后查找 .doc 文件（需要转换）
    if template_path is None and manual_dir.exists():
        for f in os.listdir(str(manual_dir)):
            if f.lower().endswith('.doc'):
                # 尝试用 olefile 读取
                try:
                    import olefile
                    from langchain_core.documents import Document
                    # 直接用之前的方式读取
                    print(f"[INFO] 从知识库找到 .doc 模板: {f}（尝试转换）")
                    # 用 LibreOffice 或 Word 转换
                    import subprocess, tempfile
                    with tempfile.TemporaryDirectory() as tmp_dir:
                        result = subprocess.run(
                            ['soffice', '--headless', '--convert-to', 'docx', str(manual_dir / f), '--outdir', tmp_dir],
                            capture_output=True, text=True, timeout=60
                        )
                        if result.returncode == 0:
                            basename = os.path.splitext(f)[0]
                            converted = os.path.join(tmp_dir, basename + '.docx')
                            if os.path.exists(converted):
                                template_path = Path(converted)
                                print(f"[INFO] .doc 转换成功")
                except Exception as e:
                    print(f"[WARN] .doc 模板转换失败: {e}")
    
    if template_path is None:
        return {
            "status": "error",
            "message": "模板文件未找到。请在"企业内部体系文件 → 手册"分类下上传一个 .docx 格式的质量手册模板文件。"
        }
    
    print(f"[INFO] 使用模板: {template_path}")
    
    # 如果是 .doc 文件，需要转成 .docx 才能用 python-docx 编辑
    actual_template = template_path
    if str(template_path).lower().endswith('.doc'):
        print(f"[INFO] 正在将 .doc 模板转换为 .docx 格式...")
        import tempfile
        import subprocess
        
        # 方法1：尝试用 pywin32 + Word COM
        converted = False
        try:
            import win32com.client
            import pythoncom
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc_obj = word.Documents.Open(str(template_path))
            tmp_docx = tempfile.mktemp(suffix='.docx')
            doc_obj.SaveAs2(tmp_docx, FileFormat=16)
            doc_obj.Close()
            word.Quit()
            pythoncom.CoUninitialize()
            actual_template = Path(tmp_docx)
            converted = True
            print(f"[INFO] Word COM 转换成功")
        except Exception as e:
            print(f"[WARN] Word COM 转换失败: {e}")
        
        # 方法2：尝试用 LibreOffice
        if not converted:
            try:
                with tempfile.TemporaryDirectory() as tmp_dir:
                    soffice_paths = ['soffice', 
                                   'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
                                   'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe']
                    for sp in soffice_paths:
                        try:
                            result = subprocess.run(
                                [sp, '--headless', '--convert-to', 'docx', str(template_path), '--outdir', tmp_dir],
                                capture_output=True, text=True, timeout=60
                            )
                            if result.returncode == 0:
                                basename = os.path.splitext(os.path.basename(str(template_path)))[0]
                                converted_path = os.path.join(tmp_dir, basename + '.docx')
                                if os.path.exists(converted_path):
                                    # 复制到临时文件（因为 tmp_dir 会被删除）
                                    tmp_docx = tempfile.mktemp(suffix='.docx')
                                    import shutil
                                    shutil.copy2(converted_path, tmp_docx)
                                    actual_template = Path(tmp_docx)
                                    converted = True
                                    print(f"[INFO] LibreOffice 转换成功")
                                    break
                        except Exception:
                            continue
            except Exception as e:
                print(f"[WARN] LibreOffice 转换失败: {e}")
        
        if not converted:
            return {
                "status": "error",
                "message": "无法将 .doc 模板转换为 .docx 格式。请安装 Microsoft Word 或 LibreOffice，或者上传 .docx 格式的模板文件。"
            }
    
    print(f"[INFO] 正在分析体系调研数据...")
    
    # 2. 构建替换映射
    replacements, info = build_replacement_map(survey_data)
    print(f"[INFO] 识别到 {len(replacements)} 处需要替换的内容")
    print(f"[INFO] 公司: {info.get('company_name','?')}, 产品: {info.get('products','?')}")
    print(f"[INFO] 正在加载模板文件...")
    
    # 3. 加载模板
    doc = Document(str(actual_template))
    print(f"[INFO] 模板加载完成，开始替换内容...")
    
    # 3. 加载模板
    doc = Document(str(template_path))
    
    # 4. 替换段落
    para_count = 0
    for p in doc.paragraphs:
        old_text = p.text
        replace_in_runs(p, replacements)
        if p.text != old_text:
            para_count += 1
    
    # 5. 替换表格
    table_count = 0
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                old_text = cell.text
                replace_in_cell(cell, replacements)
                if cell.text != old_text:
                    table_count += 1
    
    # 6. 替换页眉页脚
    hf_count = 0
    for section in doc.sections:
        for hf in [section.header, section.first_page_header, section.even_page_header,
                    section.footer, section.first_page_footer, section.even_page_footer]:
            if hf:
                for p in hf.paragraphs:
                    old_text = p.text
                    replace_in_runs(p, replacements)
                    if p.text != old_text:
                        hf_count += 1
                for table in hf.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            old_text = cell.text
                            replace_in_cell(cell, replacements)
                            if cell.text != old_text:
                                hf_count += 1
    
    print(f"[INFO] 替换完成: 段落{para_count}处, 表格{table_count}处, 页眉页脚{hf_count}处")
    print(f"[INFO] 正在生成文件...")
    
    # 7. 生成文件名
    company_name = info.get('company_name', '企业')
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', company_name)
    today = datetime.now().strftime("%Y%m%d")
    filename = f"质量管理手册_{safe_name}_{today}.docx"
    output_path = os.path.join(output_dir, filename)
    
    # 8. 保存
    os.makedirs(output_dir, exist_ok=True)
    doc.save(output_path)
    print(f"[OK] 质量手册已生成: {output_path}")
    
    return {
        "status": "success",
        "filename": filename,
        "file_path": output_path,
        "replacements": {
            "paragraphs": para_count,
            "tables": table_count,
            "headers_footers": hf_count
        },
        "company_name": company_name,
    }


def parse_args():
    parser = argparse.ArgumentParser(description="质量手册生成器")
    parser.add_argument("--survey-json", required=False, default=None,
                       help="体系调研数据（JSON 字符串）")
    parser.add_argument("--output-dir", required=False, default=".",
                       help="输出目录")
    return parser.parse_args()


def main():
    # Windows 终端 UTF-8
    if sys.platform == "win32":
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")
    
    args = parse_args()
    
    # 加载调研数据
    survey = {}
    if args.survey_json:
        survey = load_survey_data(args.survey_json)
    
    if not survey:
        print("[ERROR] 未提供体系调研数据")
        result = {"status": "error", "message": "未提供体系调研数据"}
    else:
        result = generate_manual(survey, args.output_dir)
    
    print(f"\n[RESULT_JSON]")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
