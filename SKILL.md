# SCskill - 体系文件生成技能

> **版本**：1.0.0
> **功能**：基于用户填写的体系调研信息，从模板生成质量手册等体系文件

## 工作原理

1. 用户在前端填写体系调研表
2. 点击"一键生成手册"按钮
3. 后端调用 `scripts/generate_manual.py`
4. 脚本从 `templates/` 读取模板（.docx）
5. 根据调研信息替换模板中的占位内容
6. 生成新的 .docx 文件供下载

## 使用方法

```bash
python scripts/generate_manual.py \
  --survey-json '{"sv_company_name":"xxx",...}' \
  --output-dir /path/to/output
```

## 文件结构

```
SCskill/
├── SKILL.md
├── scripts/
│   └── generate_manual.py
└── templates/
    └── IATF16949_quality_manual_template.docx
```
