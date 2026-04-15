# Workflow: Append Slides

适用场景：

- 用户要把文档内容追加到现有演示文稿
- 输入里包含 `target_slides_url`

append 模式的核心约束：**只新增，不改旧内容。**

## Inputs

必填：

- 一个且仅一个来源：`doc_url` / `doc_token` / `doc_name`
- `target_slides_url`

可选：

- `content_mode`
- `title`
- `max_slides`

默认：

- `target_mode = append`
- `content_mode = report`
- `max_slides` 是 outline 约束；若写入 `presentation.max_slides`，`validate-outline` 会拒绝超出页数预算的结果

## Required Scopes

- 基础抓取：`docx:document:readonly`
- 使用 `doc_name` 搜索：`search:docs:read`
- 追加到 Slides：`slides:presentation:write_only`
- 来源或目标任一侧是 Wiki：`wiki:wiki:readonly`

## Required Reading

开始执行前：

1. 先读 `../SKILL.md`
2. 再读已安装的 `~/.codex/skills/lark-shared/SKILL.md`
3. 再读已安装的 `~/.codex/skills/lark-doc/SKILL.md`
4. 再读已安装的 `~/.codex/skills/lark-slides/SKILL.md`
5. 再读 `content-modes.md`
6. 再读 `slide-authoring-rules.md`

## Step 1: Resolve The Source

来源解析与 new 模式相同：

- `doc_url`：
  - `/docx/`、`/doc/`：直接继续
  - `/wiki/`：先解析 wiki node，再决定是否可抓取
- `doc_token`：
  - 底层 `doc` / `docx` token：直接继续
  - wiki token：先解析 wiki node，再决定是否可抓取
- `doc_name`：搜索并在歧义时停下来让用户选择

先运行 `resolve-source`，生成 `resolved-source.json`。

如果 `doc_name` 返回多个候选，需要在继续 `fetch` 前固化用户选择：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py choose-source \
  --resolved-source <run_dir>/resolved-source.json \
  --candidate-index <1-based index>
```

补充约束：

- 搜索命中的 wiki 只算候选，不算已解析完成的 fetch target
- 如果最终选中的是 wiki，或 `doc_url` / `doc_token` 本身就是 wiki 引用，先执行 `lark-cli wiki spaces get_node --as user --params '{"token":"<wiki_token>"}'`
- wiki source 只有在返回 `node.obj_type = doc` 或 `docx` 时才允许继续 fetch，并使用返回的真实 `node.obj_token`
- 这一步依赖 `wiki:wiki:readonly`

## Step 2: Resolve The Target Deck

`target_slides_url` 必须先解析成真实的 `xml_presentation_id`，再允许发布。

这一步属于 append 的 preflight，必须发生在 outline approval 前；不论目标是 `/slides/...` 还是指向 Slides 的 `/wiki/...` URL，都不能跳过。

先运行：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py resolve-target \
  --run-dir <run_dir> \
  --target-slides-url <target_slides_url>
```

期望产物：

- `resolved-target.json`

支持两种目标形式：

- `/slides/<xml_presentation_id>`：直接取 presentation id
- `/wiki/<wiki_token>`：先查询 wiki node，再按当前 refs / schema 确认它解析到 slides 对象，并使用返回的真实 `node.obj_token` 作为 presentation id

解析规则：

- 如果 wiki 目标解析后不是 `slides` 类型，立即停止
- wiki target 解析依赖 `wiki:wiki:readonly`
- 如果 URL 既不是 `/slides/` 也不是 `/wiki/`，立即停止
- 解析失败时，禁止开始追加

文档约束：

- 这里不要假设自定义响应 envelope
- 以已安装 `lark-slides` skill 和 `lark-cli schema` 的当前返回字段为准

## Step 3: Fetch The Full Source Document

用 `fetch` 子命令抓取全文，并写入：

- `source.json`
- `source.md`

分页必须抓完整，不能只取首页内容。

如果来源是 Wiki，这里必须抓取 Step 1 已解析出的底层 `obj_token`，不能直接用 wiki token。

## Step 4: Draft An Append-Safe Outline

append 模式的 outline 仍然先给用户审阅，但内容约束更严格。

必须在 `presentation` 中明确：

- `target_mode = "append"`
- `content_mode`
- 来源信息
- 不能沿用模板里的其他 `target_mode`

append 模式默认只写“新增段落”或“新增章节”需要的页面，不重做整份 deck 的封面和目录。

如果确实要插入章节分隔页，应在对应 slide 上显式写：

```json
{
  "role": "section",
  "section_divider": true,
  "layout": "title-only"
}
```

append 模式下不允许再使用 `role = "cover"`，即使把 `section_divider` 设为 `true` 也不行。

## Step 5: Avoid Duplicate Covers

默认禁止：

- 通用封面页
- 通用目录页
- 重复“汇报标题 / 项目名称 / 周报封面”页

只有当用户明确说“我就是要插入一个章节分隔页”时，才允许生成 section divider。

建议做法：

- 使用更像“章节页”的标题，而不是整份 deck 的封面文案
- 在对应 slide 上写 `section_divider: true`，明确这是有意为之

如果没有这个明确意图，append 模式第一张 slide 不应是 `role = cover`。

## Step 6: Hard Approval Gate

append 模式按推荐工作流也应先停在 outline 审阅。

在用户明确确认前：

- 不要 render
- 不要 publish
- 不要向现有 deck 添加任何 slide
- 当前脚本本身不会强制执行这个 gate；如果需要人工确认，必须由调用方在这里显式停住

## Step 7: Validate, Render, Publish

用户确认后执行。进入这一步前，先确认同一轮 append run 已经产出 `resolved-target.json`，并且 `outline.json` 里的 `presentation.target_mode = "append"`：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py validate-outline \
  --outline <run_dir>/outline.json

python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py render \
  --run-dir <run_dir> \
  --outline <run_dir>/outline.json

python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py publish \
  --run-dir <run_dir> \
  --outline <run_dir>/outline.json \
  --slides-json <run_dir>/slides.json \
  --target-slides-url <target_slides_url>
```

publish 阶段规则：

- 使用 append preflight 已确认过的目标 presentation id
- 逐页调用 `lark-cli slides xml_presentation.slide create --as user`
- 只记录新增页的 `slide_ids`

## Safe Publication Rules

- 不删除原有 slide
- 不重排原有 slide
- 不覆盖原有 slide 内容
- 不把 append 流程伪装成“重新生成整套 deck”
- 如果中途失败，保留已成功追加的页面元数据，并如实报告部分成功

## Expected Result

冻结的 publish-result 契约：

- `target_mode = append`
- `xml_presentation_id`
- `slide_ids`
- `slides_added`
- `url`
- `run_dir`

## Stop Conditions

以下情况必须停止：

- `target_slides_url` 无法解析
- wiki 目标不是 slides
- 用户没有确认 outline
- outline 校验失败
- render 失败
- 追加到一半报错

append 模式出错时，不要自动回滚；保留产物，报告成功追加到第几页。
