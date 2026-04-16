---
name: lark-workflow-doc-to-slides
version: 1.0.0
description: "文档转幻灯片工作流：把飞书文档或 Wiki 内容先整理成可审阅的 slide outline，经用户确认后新建 Slides，或追加到已有 Slides。适用于使用 doc_url、doc_token、doc_name 作为来源，或把文档内容追加到 target_slides_url 指向的现有演示文稿。"
metadata:
  requires:
    bins: ["lark-cli", "python3"]
---

# Lark Workflow Doc To Slides

把文档变成 PPT 时，先把“文档内容”变成“可审阅的大纲”，再把已批准的大纲变成 Slides。

适用触发语义：

- “把这篇文档转成飞书幻灯片 / PPT”
- “根据 `doc_url` / `doc_token` / `doc_name` 生成 Slides”
- “把这个 Wiki 做成汇报幻灯片”
- “把这篇文档追加到现有 Slides”
- “用 `target_slides_url` 继续往当前 deck 里加几页”

**CRITICAL — 开始前必须先读取已安装的 `lark-shared`、`lark-doc`、`lark-slides`、`lark-drive` skills；本机默认路径是 `~/.codex/skills/lark-shared/SKILL.md`、`~/.codex/skills/lark-doc/SKILL.md`、`~/.codex/skills/lark-slides/SKILL.md`、`~/.codex/skills/lark-drive/SKILL.md`。**

## Required Inputs

必须且只能提供一个来源标识：

- `doc_url`
- `doc_token`
- `doc_name`

可选输入：

- `content_mode`: `faithful` | `report`
- `theme`: `briefing` | `document` | `spotlight` | `minimal` | `sunset`
- `cover_style`: `editorial` | `modern`
- `title`
- `max_slides`

append 模式额外必填：

- `target_slides_url`

默认值：

- 未提供 `target_slides_url` 时，`target_mode = new`
- 仅提供 `target_slides_url` 时，`target_mode = append`
- 未提供 `content_mode` 时，默认 `report`
- 未提供 `theme` 时，`report -> briefing`，`faithful -> document`
- 未提供 `cover_style` 时，按 `content_mode -> theme -> cover_style` 自动补默认值：`report -> briefing -> editorial`，`faithful -> document -> editorial`，`spotlight -> modern`
- 显式提供 `theme` 时，覆盖 `content_mode` 推导出的默认主题
- `max_slides` 是 outline 约束，不是底层脚本参数；若写入 `outline.json` 的 `presentation.max_slides`，`validate-outline` 会拒绝超出页数预算的结果

## Workflow Invariants

- 必须先生成 outline，再做 XML render 和 Slides 发布。
- 用户未明确确认 outline 前，禁止创建新 deck，也禁止向已有 deck 追加页面。
- `doc_name` 命中多个候选时，必须停下来让用户选；禁止自动猜测。
- 当 `doc_name` 命中多个候选并写出 `resolved-source.json` 后，可用 `choose-source --resolved-source ... --candidate-index <1-based index>` 固化用户选择，再继续 `fetch`。
- `/wiki/...` 不能当成可直接 fetch / publish 的目标；无论它来自 `doc_url`、`doc_name` 搜索命中，还是 `target_slides_url`，都必须先调用 `lark-cli wiki spaces get_node --as user` 解析到底层 `obj_type` / `obj_token`。
- `doc_token` 既可能是底层 `doc` / `docx` token，也可能是 wiki node token；如果它是 wiki token，同样必须先调用 `lark-cli wiki spaces get_node --as user` 解析。
- wiki source 只有在解析后 `obj_type` 为 `doc` / `docx` 时才允许继续 fetch；wiki target 只有在解析后 `obj_type` 为 `slides` 时才允许继续 append。
- 涉及 wiki source / wiki target 解析时，必须具备 `wiki:wiki:readonly`；`doc_name` 搜索仍然需要 `search:docs:read`。
- append 模式必须在 outline approval 前完成 target preflight（`resolve-target`）；无论 `target_slides_url` 是 `/slides/...`，还是指向 Slides 的 `/wiki/...` URL，都不能跳过这一步。
- `target_slides_url` 表示“直接追加到这个已有 deck”。
- append 模式只允许新增 slides；禁止重写、删除、重排已有页面。
- append 模式默认不生成通用“封面 / 目录”；只有用户明确要求新章节分隔页时，才允许使用 `role = "section"`、`layout = "title-only"` 且 `section_divider: true` 的 divider 页。
- append 模式下禁止再使用 `role = "cover"`，即使 `section_divider = true` 也不允许；章节页必须显式写成 `role = "section"`、`layout = "title-only"`、`section_divider = true`。
- 内置 `theme` 预设包括 `briefing`、`document`、`spotlight`、`minimal`、`sunset`；未知值必须视为非法 outline。
- 内置 `cover_style` 只有 `editorial`、`modern`；如果未显式提供，则按 `content_mode -> theme -> cover_style` 推导默认值：`report -> briefing -> editorial`，`faithful -> document -> editorial`，`spotlight/minimal/sunset -> modern`。
- `editorial` 的语义是克制、专业的封面；`modern` 的语义是更偏 presentation-style 的封面，可理解为 accent bar + title stack + subtitle panel 的组合，但仍必须在浅色主题下保持可读和成立，不能把它理解成仅适用于深色背景的变体。
- 当前封面副标题应按实际 render 顺序理解：优先取 `presentation.subtitle`，没有时再回退 cover slide 的 `objective`，最后才回退 `key_points[0]`。若同时填写这些字段，优先保持它们一致，避免 outline 语义分叉。

## Routing

- 无 `target_slides_url`：
  - 读取 [`references/workflow-new-slides.md`](references/workflow-new-slides.md)
  - 再读取 [`references/content-modes.md`](references/content-modes.md)
  - 再读取 [`references/slide-authoring-rules.md`](references/slide-authoring-rules.md)
- 仅有 `target_slides_url`：
  - 读取 [`references/workflow-append-slides.md`](references/workflow-append-slides.md)
  - 再读取 [`references/content-modes.md`](references/content-modes.md)
  - 再读取 [`references/slide-authoring-rules.md`](references/slide-authoring-rules.md)

## Execution Shape

执行脚本位于 `scripts/doc_to_slides.py`。工作流子命令：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py resolve-source ...
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py choose-source ...
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py resolve-target ...
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py fetch ...
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py validate-outline ...
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py render ...
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py publish ...
```

说明：

- 当前脚本已在 `resolve-target` / `publish` 子命令上支持 `--template-slides-url`

运行目录约定：

```text
.lark-workflow-doc-to-slides/runs/<timestamp>-<slug>/
```

常见产物：

- `resolved-source.json`
- `resolved-target.json`
- `source.json`
- `source.md`
- `outline.json`
- `slides.json`
- `render-summary.json`
- `publish-result.json`

`publish-result.json` 采用归一化的顶层字段：

- `target_mode`
- `xml_presentation_id`
- `url`
- `slide_ids`
- `slides_added`
- `run_dir`

发布保护：

- `publish` 只接受与当前 `outline.json` 匹配的 `slides.json`
- 如果 `render-summary.json` 缺失，或 `outline` / `slides` 指纹不一致，必须先重新 render

## Operator Rules

- 用 AI 负责理解原文和写 outline；用脚本负责校验、渲染、发布。
- `doc_token` 可以是底层 `doc` / `docx` token，也可以是 wiki token；如果是 wiki token，先解析到 `obj_type` / `obj_token`，再继续 workflow。
- 生成 outline 时，优先对齐 `templates/outline.json` 的字段形状；`presentation.theme` 可显式写入内置预设，不传时由 render 按 `content_mode` 自动补默认主题；`presentation.cover_style` 可显式写成 `editorial` 或 `modern`，不传时按 `content_mode -> theme -> cover_style` 补默认值（`report -> briefing -> editorial`，`faithful -> document -> editorial`，`spotlight/minimal/sunset -> modern`）；`cover_style` 只影响真实封面页的 `role = cover` + `layout = title-only` 渲染，不影响 append 模式章节分隔页；封面页优先使用 `role = cover` + `layout = title-only`，不要把封面写成普通 `title-body` 正文页；每页 slide 保留显式布尔字段 `section_divider`，默认 `false`，append 模式需要章节分隔页时应改成 `role = "section"`、`layout = "title-only"` 并把 `section_divider` 设为 `true`。
- 写封面语义时，`editorial` 默认对应 restrained professional cover；`modern` 对应 presentation-style cover，可按 accent bar + title stack + subtitle panel 理解，但不要把它写成只能依赖暗色背景成立的视觉说明。
- 写封面副标题时，优先填在 `presentation.subtitle`；只有在它为空时才会退回 cover slide 的 `objective`，最后才回退 `key_points[0]`。如果 outline 同时包含这些字段，先保持它们一致。
- append 模式在进入 `validate-outline` / `render` / `publish` 前，必须确认 `outline.json` 中 `presentation.target_mode = "append"`，不要沿用模板里的其他值。
- 发布阶段只允许使用现有 `lark-cli` 命令链；不存在 `slides +create-from-outline` 这种快捷命令。
- 如果权限或身份不对，按 `lark-shared` 规则先修正身份与授权，再继续工作流。
- 涉及 wiki source / wiki target 时，按当前已安装 refs / schema 查看真实返回字段，不要在 skill 文案里猜测自定义响应包裹层。
