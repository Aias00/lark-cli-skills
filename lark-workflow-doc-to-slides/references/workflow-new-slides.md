# Workflow: New Slides

适用场景：

- 用户要从 `doc_url`、`doc_token` 或 `doc_name` 生成一份新的飞书 Slides
- 当前没有 `target_slides_url`

只要用户提供了 `target_slides_url`，这就不再是 new 流程；必须切到 `workflow-append-slides.md`，并先完成 append preflight（`resolve-target`）。

## Inputs

必填：

- 一个且仅一个来源：`doc_url` / `doc_token` / `doc_name`

可选：

- `content_mode`: `faithful` 或 `report`
- `title`
- `max_slides`

默认：

- `target_mode = new`
- `content_mode = report`
- `max_slides` 是 outline 约束；若写入 `presentation.max_slides`，`validate-outline` 会拒绝超出页数预算的结果

## Required Scopes

- 基础流程：`docx:document:readonly`
- 使用 `doc_name` 搜索：`search:docs:read`
- 创建新 Slides：`slides:presentation:create`、`slides:presentation:write_only`
- 来源是 Wiki，或 `doc_name` 最终选中 Wiki 候选：`wiki:wiki:readonly`

## Required Reading

开始执行前：

1. 先读 `../SKILL.md`
2. 再读已安装的 `~/.codex/skills/lark-shared/SKILL.md`
3. 再读已安装的 `~/.codex/skills/lark-doc/SKILL.md`
4. 再读已安装的 `~/.codex/skills/lark-slides/SKILL.md`
5. 再读 `content-modes.md`
6. 再读 `slide-authoring-rules.md`

## Step 1: Resolve The Source

先把来源归一化成一个明确可抓取的文档目标。

- `doc_url`：
  - `/docx/`、`/doc/`：直接进入抓取
  - `/wiki/`：先解析 wiki node，再决定是否可抓取
- `doc_token`：
  - 底层 `doc` / `docx` token：直接进入抓取
  - wiki token：先解析 wiki node，再决定是否可抓取
- `doc_name`：先搜索，再按结果处理

`doc_name` 规则：

- `0` 个候选：停止，要求用户补充更精确的名称
- `1` 个明确候选：自动继续
- `>1` 个可信候选：停止，要求用户从候选列表中明确选择

补充约束：

- 搜索结果里的 wiki URL 只是候选命中，不等于已经拿到了可直接 `docs +fetch` 的底层对象
- 如果候选是 wiki，或 `doc_url` / `doc_token` 本身就是 wiki 引用，先执行 `lark-cli wiki spaces get_node --as user --params '{"token":"<wiki_token>"}'`
- wiki source 只有在返回 `node.obj_type = doc` 或 `docx` 时才允许继续 fetch，并使用返回的真实 `node.obj_token`
- 这一步依赖 `wiki:wiki:readonly`
- 不要在这里假设“search 命中的 wiki 一定能直接抓取”

先执行：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py resolve-source \
  --run-dir <run_dir> \
  --doc-url <url>
```

或：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py resolve-source \
  --run-dir <run_dir> \
  --doc-token <token>
```

或：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py resolve-source \
  --run-dir <run_dir> \
  --doc-name "<document name>"
```

期望产物：

- `resolved-source.json`

如果 `doc_name` 返回多个候选，需要在继续 `fetch` 前固化用户选择：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py choose-source \
  --resolved-source <run_dir>/resolved-source.json \
  --candidate-index <1-based index>
```

## Step 2: Fetch The Full Document

抓取全文，不是只抓第一页。

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py fetch \
  --run-dir <run_dir> \
  --resolved-source <run_dir>/resolved-source.json
```

执行要求：

- 使用 `lark-cli docs +fetch --as user --format json`
- 如果 Step 1 解析过 wiki，这里要用解析后的底层 `obj_token`，而不是 wiki token
- 如果返回 `has_more`，必须继续翻页直到全文抓完
- 把结构化结果写入 `source.json`
- 把聚合后的 Markdown 写入 `source.md`

## Step 3: Choose The Content Mode

按 [`content-modes.md`](content-modes.md) 选择模式：

- `faithful`：原文结构应保持可识别
- `report`：默认，按汇报逻辑重组

如果用户没有指定，默认 `report`。不要把默认值写得模糊。

## Step 4: Draft The Outline

基于 `source.md` 起草 `outline.json`，并遵守 [`slide-authoring-rules.md`](slide-authoring-rules.md)。

最少要包含：

- `presentation.title`
- `presentation.source`
- `presentation.content_mode`
- `slides[]`
- 每页的 `no`、`role`、`section_divider`、`title`、`layout`、`key_points`

`presentation.target_mode` 取值规则：

- 写成 `new`

建议输出一个给用户审阅的摘要：

```text
[PPT 标题] — [目标受众 / 汇报语境]
1. [封面]
2. [背景 / 目标]
3. [现状 / 问题]
4. [方案 / 进展]
5. [风险 / 决策]
6. [下一步]
```

## Step 5: Hard Approval Gate

这一步按推荐工作流应停下来。

在用户明确回复“确认 / 可以生成 / 继续发布 / approve outline”之前：

- 不要运行 `validate-outline`
- 不要运行 `render`
- 不要运行 `publish`
- 不要创建任何新的 Slides 资源
- 当前脚本本身不会强制执行这个 gate；如果需要人工确认，必须由调用方在这里显式停住

## Step 6: Validate And Render

用户确认 outline 后，再进入脚本阶段：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py validate-outline \
  --outline <run_dir>/outline.json

python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py render \
  --run-dir <run_dir> \
  --outline <run_dir>/outline.json
```

期望产物：

- `slides.json`
- `render-summary.json`

## Step 7: Publish

### 新建 deck

发布规则是固定的：

- 先用 `lark-cli slides +create --as user --title ...` 创建空 deck
- 再逐页调用 `lark-cli slides xml_presentation.slide create --as user`

说明：

- 这样做比 `slides +create --slides` 更啰嗦
- 但它能稳定保留“部分成功时已经创建的 presentation 和 slide_ids”，更符合本 workflow 的可恢复性要求

执行：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py publish \
  --run-dir <run_dir> \
  --outline <run_dir>/outline.json \
  --slides-json <run_dir>/slides.json
```

发布前提：

- `slides.json` 必须来自当前 `run_dir` 的最近一次 `render`
- `render-summary.json` 必须存在
- 如果 `outline.json` 和 `slides.json` 不再一致，必须先重新 render，不能直接 publish

冻结的 publish-result 契约：

- `target_mode = new`
- 新 deck 的 `xml_presentation_id`
- `url`
- `slide_ids`
- `slides_added`
- `run_dir`
- `publish-result.json`

## Stop Conditions

遇到以下情况必须停止，不要强行继续：

- 来源无法解析
- `doc_name` 有多个候选但用户尚未选择
- 抓取未完成或权限不足
- outline 校验失败
- XML render 失败
- 发布只完成了一部分页面

如果发布部分成功，也要保留 `run_dir` 和已返回的 `slide_ids`，不要假装整次发布完成。
