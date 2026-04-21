---
name: excel-master-skill
description: >
  Generate formatted Excel (.xlsx) files from structured or unstructured data using the excel-master-cli (epm) tool.
  Use this skill whenever the user wants to: create test case documents (测试用例), requirements tracking matrices (需求追踪矩阵),
  project plans (项目计划表), or any structured Excel document. Triggers on mentions of: 测试用例, 需求文档, 项目计划, 任务列表,
  Excel generation, xlsx output, formatted spreadsheet, epm, test case table, requirement matrix, or when a user provides
  structured data (in conversation or files) that maps naturally to a tabular format. Also use when a user asks to convert
  JSON/YAML/Markdown/text data into a formatted Excel file, when analyzing UI screenshots to generate test cases,
  or when designing custom Excel templates for business documents.
---

# Excel Master Skill

Generate formatted Excel documents from structured data via the `epm` CLI tool.

## When This Skill Applies

You should use this skill when:
- The user wants to generate an Excel (.xlsx) file from structured data
- The user describes test cases, requirements, or project tasks and wants them in a formatted document
- The user has data in JSON/YAML/Markdown/text/docx format and wants it converted to a styled spreadsheet
- The user has UI screenshots or design images and wants to generate test cases from them
- The user says things like "帮我写测试用例", "整理需求到Excel", "生成项目计划表", "把这些数据放到Excel里", "根据截图生成测试用例", "分析图片生成Excel"

## Reference Files

Read these files on-demand when you need specifics — don't load them all upfront.

| When to read | File |
|-------------|------|
| Preparing test-case data | `schemas/test-case.yaml` |
| Preparing requirements data | `schemas/requirements.yaml` |
| Preparing project-plan data | `schemas/project-plan.yaml` |
| Unsure of input format | `references/test-case.json`, `references/test-case.yaml`, `references/test-case.md` |
| Something went wrong | `TROUBLESHOOTING.md` |

Each schema file defines the exact fields, validation values, defaults, and auto-behaviors for that template. Read the relevant schema before preparing input data.

## Step 0 — Pre-check

Before any operation, verify the tool is installed:

```bash
epm --version
```

If not installed:

```bash
cd /Users/lindaw/Documents/excel-master/excel-master-cli && pip install -e .
```

## Step 1 — Identify Template

Map the user's intent to a built-in template:

| User intent | Template |
|------------|----------|
| 测试用例 / test case / 用例 | `test-case` |
| 需求 / requirement / 需求追踪矩阵 | `requirements` |
| 项目计划 / project plan / 任务列表 / 排期 | `project-plan` |

If the user names a custom template, use it directly. If unsure, list available templates:

```bash
epm list-templates
```

To inspect a template's fields:

```bash
epm show-schema <template-name>
```

## Step 2 — Prepare Input Data

There are two data sources: **inline** (user provides data in conversation) or **file** (user has a data file).

### Source A — Inline data (user describes content in conversation)

When the user provides data in their message or across multiple messages:

1. Read the relevant schema file (e.g., `schemas/test-case.yaml`) to confirm required fields, validation values, and defaults
2. Extract structured records from the user's description — each distinct item becomes one record
3. Map user's language to schema field names (e.g., "模块" → `module`, "优先级" → `priority`)
4. Write the records to a JSON file — JSON is the most reliable format because it avoids parser ambiguity
5. **Omit `id` fields** — they are auto-generated (TC-001, R-001, PP-001)

Example — user says: "帮我写三个登录模块的测试用例，第一个是正常登录优先级高..."

Read `schemas/test-case.yaml`, then write:

```json
[
  {
    "module": "登录模块",
    "title": "正常登录",
    "priority": "高",
    "steps": "1. 打开登录页\n2. 输入正确的用户名\n3. 输入正确的密码\n4. 点击登录",
    "expected": "登录成功，跳转到首页"
  }
]
```

### Source B — File-based data

When the user provides a file path:

1. Check the file extension to determine format: `.json`, `.yaml`/`.yml`, `.md`, `.txt`, `.docx`
2. Read the file and verify it contains valid structured data
3. If format is unclear or the file doesn't parse, check `TROUBLESHOOTING.md`
4. Use the file directly as input — no need to rewrite it

### Source C — Image-based data (screenshots → test cases)

When the user has UI screenshots or design images in a folder and wants to generate test cases:

1. **Scan the folder** for image files (PNG/JPG/JPEG/BMP):
   ```bash
   ls -1 <folder>/*.png
   ```

2. **Analyze each image** using agent's built-in vision capability (Read tool or analyze_image MCP tool):
   - Read each screenshot and describe UI elements: fields, buttons, dialogs, lists, menus, tabs, validation messages
   - Filenames are often descriptive — use them to understand the feature context
   - Focus on: query conditions, form fields, action buttons, list columns, error messages, workflow states

3. **Generate JSON test cases** based on analysis:
   - Read the relevant schema (e.g., `schemas/test-case.yaml`) for field definitions
   - Each distinct feature/interaction from screenshots becomes one test case
   - Map observed UI behavior to: module, title, priority, precondition, steps, expected result
   - Write all records to a JSON file (e.g., `<folder>/testcases.json`)
   - **Omit `id` fields** — they are auto-generated
   - **Omit `status` fields** — default is `待执行`

4. **Generate Excel** from the JSON:
   ```bash
   epm create -t test-case -i <folder>/testcases.json -f image -o <folder>/测试用例.xlsx
   ```

Example workflow:
```
User: "根据 tests/项目计划/ 文件夹下的截图生成测试用例"

Agent:
1. ls tests/项目计划/*.png  →  finds 31 screenshots
2. Read & analyze each image → extract UI elements, features, behaviors
3. Write tests/项目计划/testcases.json with 32 test case records
4. epm create -t test-case -i tests/项目计划/testcases.json -f image -o tests/项目计划/项目计划测试用例.xlsx
5. Report: "Generated 32 test cases from 31 screenshots"
```

### Input file structure (JSON array)

```json
[
  {"field1": "value1", "field2": "value2"},
  {"field1": "value3", "field2": "value4"}
]
```

### Field name flexibility

For text/markdown inputs, fields accept aliases:
- `module` also matches `模块`, `module`, `功能模块`, `所属模块`
- `priority` also matches `优先级`, `priority`
- `owner` also matches `负责人`, `owner`, `责任人`

For JSON, use the exact `field` key from the schema.

### Multiline content

For fields marked `multiline: true` (e.g., `steps` in test-case):
- JSON: use literal `\n` (not `\\n`) inside string values
- YAML: use `|` block scalar
- Row height auto-adjusts based on line count

## Step 3 — Generate Excel

```bash
epm create --template <template-name> --input <input-file> --output <output.xlsx>
```

Options:
- `--format json|yaml|md|txt|docx` — explicit format (auto-detected from extension if omitted)
- `--clean` — ignore cache, start fresh (use if previous run was interrupted)
- Use `-` as input for stdin: `echo '[...]' | epm create --template test-case --input - --format json --output out.xlsx`

Pick a descriptive output filename based on content (e.g., `登录模块测试用例.xlsx`, `Q2项目计划.xlsx`).

### Output styling

The generated Excel has consistent styling:
- **Header row**: black bold font + blue background + border + center alignment
- **Data rows**: black normal font + no background + border + inherited alignment
- **Dropdown validation**: columns with `validate` options get Excel dropdown lists (e.g., 优先级: 高/中/低)
- **Auto row height**: multiline fields (like `steps`) auto-expand row height
- **Auto ID**: columns with `auto_generate` produce sequential IDs (TC-001, TC-002, ...)

## Step 4 — Report and Handle Issues

Tell the user:
- The output file path
- Number of records generated
- Any warnings from stderr (invalid values replaced, missing required fields)

If errors occurred:
1. Check `TROUBLESHOOTING.md` for the specific error message
2. Common fixes:
   - `command not found` → re-run install step
   - `Template not found` → check `epm list-templates` for correct name
   - `Invalid value` warnings → correct the field value to match the schema's `validate` list
   - `Missing required field` → add the missing field to the input data
   - Parse errors → verify JSON/YAML syntax (trailing commas, quotes, indentation)
3. Fix the input and re-run

## Custom Templates

Users can create templates for any structured document:

```bash
# Create template scaffold
epm template-init <name>
# This creates ~/.epm/templates/<name>/ with template.xlsx and schema.yaml
```

Edit `template.xlsx` for visual styles (header row fonts, borders, colors, alignment).
Edit `schema.yaml` for field definitions:

```yaml
name: my-template
display_name: 我的模板
columns:
  - field: id
    header: 编号
    auto_generate: true    # Auto sequential ID
  - field: title
    header: 标题
    required: true
  - field: priority
    header: 优先级
    default: 中
    validate: [高, 中, 低]  # Invalid → replaced with default
  - field: description
    header: 描述
    multiline: true         # Preserve \n, auto row height
  - field: category
    header: 分类
    extract:
      - keywords: [分类, category, 类型]  # Aliases for text parsers
```

Column properties: `field` (required), `header` (required), `width` (default 15), `required`, `default`, `validate`, `auto_generate`, `multiline`, `extract`.

## Key Rules

- **Never manually create .xlsx files** — always use `epm create`
- **Omit `id` fields** from input — auto-generated
- **Use JSON** for programmatic data preparation — most reliable
- **Check stderr** for warnings about invalid/missing fields
- **Read the schema** before preparing data — it has the authoritative field definitions
- **For image-based workflows** — use agent's built-in vision (Read tool or analyze_image), no external API key or model needed
- **Use `-f image`** when the input JSON was generated from image analysis
