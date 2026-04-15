# lark-workflow-doc-to-slides

把飞书文档 / Wiki 内容先整理成可审阅的幻灯片大纲，再生成新的飞书 Slides，或追加到已有 Slides。

这不是“直接把文档生吞成 PPT”的黑盒工具。它的核心价值是：

- 先出 outline，再生成 Slides
- 支持 `doc_url`、`doc_token`、`doc_name` 三种来源
- 支持新建 deck、以及追加到已有 deck
- 追加模式只新增页面，不改旧内容

## 适用场景

- 把技术方案文档转成汇报 PPT
- 把周报文档变成飞书 Slides
- 根据文档名称生成一份新演示文稿
- 把文档内容追加到已有 Slides 作为新章节

## 前置条件

本机需要：

- `python3`
- `lark-cli`

并且已经完成：

```bash
lark-cli config init --new
lark-cli auth login --scope "docx:document:readonly search:docs:read slides:presentation:create slides:presentation:write_only wiki:wiki:readonly"
```

如果你要读取云空间文件、图片、表格等内容，还要确保对应 scope 已授权。

涉及 Wiki 来源或 Wiki 目标时，还要注意：

- `wiki:wiki:readonly` 是必需的；workflow 会先调用 `lark-cli wiki spaces get_node --as user` 解析 wiki 节点
- Wiki 不能直接 fetch / publish
- Wiki source 会先解析到真实的 `obj_type` / `obj_token`，再继续抓取
- Wiki target 会先解析到真实的 slides `obj_token`，再作为 `xml_presentation_id` 追加页面

## 支持的输入

必须提供一个来源，三选一：

- `doc_url`
- `doc_token`
- `doc_name`

可选提供：

- `content_mode`
- `theme`
- `cover_style`
- `title`
- `max_slides`

append 模式额外必填：

- `target_slides_url`

默认规则：

- 不传 `target_slides_url`：新建一份 Slides
- 只传 `target_slides_url`：进入 append 模式，并且必须先执行 target preflight（`resolve-target`）
- 不传 `content_mode`：默认 `report`
- 不传 `theme`：`report -> briefing`，`faithful -> document`
- 不传 `cover_style`：按 `content_mode -> theme -> cover_style` 自动补默认值，`report -> briefing -> editorial`，`faithful -> document -> editorial`，`spotlight -> modern`
- 显式传 `theme`：覆盖 `content_mode` 推导出的默认主题

补充说明：

- `doc_url` 可以是 `/docx/`、`/doc/` 或 `/wiki/` URL
- `doc_token` 可以是底层 `doc` / `docx` token，也可以是 wiki node token
- 如果 `doc_token` 实际上是 wiki token，也必须先走 `wiki spaces get_node` 解析，不能直接 fetch
- `max_slides` 是 outline contract 的上界：在 `validate-outline` 阶段按 `presentation.max_slides` 校验已批准的大纲页数，不能把它当成 publish 阶段再二次推导的目标页数
- `target_slides_url` 可以是 `/slides/` URL，也可以是指向 Slides 的 `/wiki/` URL
- `max_slides` 是 outline 约束，不是底层脚本参数；若写入 `outline.json` 的 `presentation.max_slides`，`validate-outline` 会拒绝超出页数预算的结果

## 内容模式

### `report`

默认模式。适合汇报、周报、项目更新。

典型结构：

- 背景 / 目标
- 问题 / 现状
- 方案 / 进展
- 风险 / 决策点
- 下一步

### `faithful`

适合技术设计、架构说明、教程类文档。

特点：

- 尽量保留原文章节顺序
- 不主动改写成管理汇报口径
- 保留限制、假设、 caveat

## 主题预设

`presentation.theme` 是可选字段，当前内置预设：

- `briefing`：浅色汇报风格；`report` 模式默认使用
- `document`：暖色文档风格；`faithful` 模式默认使用
- `spotlight`：高对比强调风格；适合需要更强视觉层级的 deck

如果不传 `theme`，render 阶段会按 `content_mode` 自动补默认值；如果传了 `theme`，就以显式值为准。

## 封面样式

`presentation.cover_style` 是可选字段，当前内置值：

- `editorial`：克制、专业的封面表达；适合汇报型和文档型 deck 的默认封面
- `modern`：更像 presentation opener 的封面表达；语义上应是 accent bar + title stack + subtitle panel 的组合，并且在浅色主题下也要成立，不应依赖暗色背景才有效

默认链路是 `content_mode -> theme -> cover_style`：

- `report -> briefing -> editorial`
- `faithful -> document -> editorial`
- `spotlight -> modern`

`cover_style` 只影响真实封面页的 `role = cover` + `layout = title-only` 渲染；append 模式下的章节分隔页仍然按 `role = "section"`、`layout = "title-only"`、`section_divider = true` 处理，不应把 divider 写成 `cover_style` 的变体。

当前封面副标题的取值顺序应与 render 实现保持一致：优先使用 `presentation.subtitle`，没有时再回退到 cover slide 的 `objective`，最后才回退到 `key_points[0]`。如果你在 outline 里同时填写这些字段，优先把它们保持一致，避免语义分叉。

## 推荐用法：直接对 AI 说

### 新建 Slides

```text
把这篇文档转成 PPT：
doc_url=https://www.feishu.cn/docx/xxxx
```

```text
把这个 Wiki 转成 PPT：
doc_url=https://www.feishu.cn/wiki/wikcnxxxx
```

```text
根据文档名称生成 Slides：
doc_name=项目周报
content_mode=report
```

### 追加到已有 Slides

```text
把这篇文档追加到现有 Slides：
doc_url=https://www.feishu.cn/docx/xxxx
target_slides_url=https://xxx.feishu.cn/slides/yyyy
```

```text
把这个 Wiki 章节追加到知识库里的 Slides：
doc_url=https://www.feishu.cn/wiki/wikcnSource
target_slides_url=https://xxx.feishu.cn/wiki/wikcnTarget
```

### 使用忠实提炼模式

```text
把这篇架构文档转成 Slides，尽量保留原结构：
doc_token=doxcnxxxx
content_mode=faithful
```

```text
把这个 wiki token 对应的文档转成 Slides：
doc_token=wikcnxxxx
content_mode=report
```

## 运行规则

### 1. 先解析来源

- `doc_url`：
  - `/docx/`、`/doc/`：直接使用
  - `/wiki/`：先查 wiki node，再解析到底层对象 token
- `doc_token`：
  - 底层 `doc` / `docx` token：直接使用
  - wiki token：先查 wiki node，再解析到底层对象 token
- `doc_name`：先搜索

`doc_name` 的行为是固定的：

- 0 个候选：停止，并要求你提供更准确的名称
- 1 个明确候选：自动继续
- 多个候选：停止，并要求你明确选择

不会自动猜。

如果你要在同一 `run_dir` 里继续处理多候选结果，可以在确认候选序号后执行：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py choose-source \
  --resolved-source <run_dir>/resolved-source.json \
  --candidate-index <1-based index>
```

如果来源是 Wiki，不论它来自 `doc_url`、`doc_token` 还是 `doc_name` 搜索命中，都会先执行：

```bash
lark-cli wiki spaces get_node --as user --params '{"token":"<wiki_token>"}'
```

然后再按返回结果继续：

- `node.obj_type` 是 `doc` / `docx`：使用 `node.obj_token` 继续 `fetch`
- 其他 `obj_type`：停止，提示当前 Wiki 不是可抓取的文档来源

### 2. 先出 outline，再生成 Slides

这是推荐工作流约定。

在你确认 outline 之前：

- 不会创建新的 Slides
- 不会追加到已有 Slides
- 不会写任何 slide XML 到目标 deck
- 当前脚本不会强制阻止 `validate-outline` / `render` / `publish`；如果你要保持人工审阅流程，需要由调用方在这一阶段显式停住

### 3. 追加模式的规则

在 append 模式里，不论 `target_slides_url` 是 `/slides/...` 还是 `/wiki/...`，都必须先完成目标解析，再进入 outline 审阅。

追加模式：

- 只新增页面
- 不删除旧页
- 不覆盖旧页
- 不重排旧页

默认不生成通用封面 / 目录。

如果你明确需要章节分隔页，append outline 应写成：

```json
{
  "role": "section",
  "section_divider": true,
  "layout": "title-only"
}
```

append 模式下不允许再使用 `role = "cover"`，即使把 `section_divider` 设为 `true` 也不行。

## 手工脚本用法

如果你不想通过自然语言触发，也可以直接调用脚本。

### Step 1: 解析来源

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py resolve-source \
  --doc-url "https://www.feishu.cn/docx/xxxx" \
  --run-dir /tmp/doc-to-slides-run
```

或：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py resolve-source \
  --doc-name "项目周报" \
  --run-dir /tmp/doc-to-slides-run
```

如果输入的是 Wiki URL，`resolve-source` 产物会先保留该 Wiki 引用；真正执行 `fetch` 前，workflow 会用 `wiki spaces get_node` 把它解析成底层文档 token。

如果输入的是 wiki token，同样要先完成这一步解析，再继续 `fetch`。

### Step 2: 抓取全文

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py fetch \
  --resolved-source /tmp/doc-to-slides-run/resolved-source.json \
  --run-dir /tmp/doc-to-slides-run
```

说明：

- 会自动处理分页
- `source.md` 应该是全文，不是第一页摘录

### Step 2A: 追加模式必须先解析目标

如果你要 append 到已有 Slides，必须先执行 target preflight。这个步骤不是可选项，不论 `target_slides_url` 是 `/slides/...` URL，还是指向 Slides 的 `/wiki/...` URL，都要先产出 `resolved-target.json`，再进入 outline 审阅：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py resolve-target \
  --target-slides-url "https://xxx.feishu.cn/wiki/wikcnTarget" \
  --run-dir /tmp/doc-to-slides-run
```

产物：

- `resolved-target.json`

说明：

- `/slides/...` 会直接提取 `xml_presentation_id`
- `/wiki/...` 会先解析 wiki node
- 只有当 `node.obj_type = slides` 时才会继续

### Step 3: 写 outline

把 outline 保存到：

```text
/tmp/doc-to-slides-run/outline.json
```

字段形状请参考：

- [templates/outline.json](./templates/outline.json)

仅演示 `presentation.theme` / `presentation.cover_style` 字段的片段：

```json
{
  "presentation": {
    "content_mode": "report",
    "theme": "briefing",
    "cover_style": "editorial"
  }
}
```

这里的示例默认值 `editorial` 表示“克制、专业的封面”；如果你显式改成 `modern`，语义应接近更有舞台感的 presentation-style opener，可理解为 accent bar + title stack + subtitle panel 的封面组合，但仍要能在浅色主题里正常工作。

模板里的 `presentation.subtitle` 与 cover slide 的 `objective` 仍然保持一致，便于在人工审阅 outline 时一眼看清封面副标题；但 render 阶段会优先读取 `presentation.subtitle`。

### Step 4: 校验 outline

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py validate-outline \
  --outline /tmp/doc-to-slides-run/outline.json
```

### Step 5: 渲染 slides.json

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py render \
  --outline /tmp/doc-to-slides-run/outline.json \
  --run-dir /tmp/doc-to-slides-run
```

### Step 6A: 新建 Slides（无模板）

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py publish \
  --outline /tmp/doc-to-slides-run/outline.json \
  --slides-json /tmp/doc-to-slides-run/slides.json \
  --run-dir /tmp/doc-to-slides-run
```

### Step 6B: 追加到已有 Slides

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py publish \
  --outline /tmp/doc-to-slides-run/outline.json \
  --slides-json /tmp/doc-to-slides-run/slides.json \
  --target-slides-url "https://xxx.feishu.cn/slides/yyyy" \
  --run-dir /tmp/doc-to-slides-run
```

append 模式要求在 outline 审阅前就已经完成 target preflight；`publish` 这里消费的是同一轮 append run 里已经写出的 `resolved-target.json`，不会再把原始 `target-slides-url` 当成新的解析输入：

- append publish 必须读到同一 `run_dir` 下的 `resolved-target.json`
- 实际追加使用的是 preflight 已确认过的 `xml_presentation_id`
- 如果你仍然传入 `--target-slides-url`，它只会被当成与 preflight 结果做一致性校验的断言

不要把目标解析拖到用户批准 outline 之后；append preflight 失败时应直接停止，不进入 `validate-outline` / `render` / `publish`。

## 运行产物

每次运行都会在 `run_dir` 下留下中间产物：

- `resolved-source.json`
- `resolved-target.json`（append 模式必有）
- `source.json`
- `source.md`
- `outline.json`
- `slides.json`
- `render-summary.json`
- `publish-result.json`

这些文件的作用：

- 可追溯
- 可恢复
- 可复查
- 出错时方便定位问题

## `publish-result.json`

发布成功后，结果格式固定为顶层字段：

- `target_mode`
- `xml_presentation_id`
- `url`
- `slide_ids`
- `slides_added`
- `run_dir`

## 重要保护机制

### 不允许发布旧的 `slides.json`

`publish` 前会检查：

- `render-summary.json` 是否存在
- 当前 `outline.json` 是否和 `slides.json` 一致

如果你改过 outline 但没重新 render，`publish` 会直接拒绝继续。

### 新建 deck 统一使用“先建空 deck，再逐页追加”

虽然更慢一点，但优点是：

- 失败时可以保留部分成功的结果
- `publish-result.json` 能准确记录已经成功创建了哪些 slide

## 常见问题

### 1. `doc_name` 搜到很多结果

这是预期行为。工具会停下来让你选，不会自动猜。

### 2. `fetch` 成功但内容不完整

这不应该发生。当前实现会分页抓到全文。

如果你发现内容仍然缺失，优先检查：

- 文档本身是否还有访问权限问题
- `source.json` 里的 `raw_pages`

如果来源是 Wiki，也先检查是否已授权 `wiki:wiki:readonly`，以及 Wiki 节点最终解析到的是 `doc` / `docx`。

### 3. `publish` 说 outline 和 slides 不一致

说明你改过 `outline.json`，但还没重新执行 `render`。

重新执行：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py render \
  --outline /tmp/doc-to-slides-run/outline.json \
  --run-dir /tmp/doc-to-slides-run
```

### 4. 追加模式生成了不该有的封面

追加模式默认不生成通用封面。

只有你明确需要章节页时，才应该在那一页设置：

```json
"section_divider": true
```

## 已验证的回归路径

当前仓库内已通过本地回归验证的路径：

- `validate-outline` 可校验模板 outline
- `render` 可从模板 outline 生成 `slides.json`
- `doc_token=wik...` 会先归一化为 wiki source，再解析到底层 `doc` / `docx`
- `resolve-target --target-slides-url=/wiki/...` 会先解析到底层 Slides，再写出 `resolved-target.json`

仍建议在真实飞书环境补做一次手工 smoke：

- `doc_url=/wiki/...` → `resolve-source` + `fetch`
- `target_slides_url=/wiki/...` → `resolve-target` + append publish

## 真实 Smoke Checklist

下面这轮 checklist 面向真实飞书环境，默认你已经有：

- 一个可抓取的 Wiki 文档来源
- 一个可追加的 Slides 目标，或一个指向 Slides 的 Wiki 目标
- 当前用户身份的 `lark-cli` 已配置完成

### 0. 前置授权检查

先确认当前机器是否真的具备 wiki 路径需要的授权：

```bash
lark-cli auth status
lark-cli auth check --scope "docx:document:readonly search:docs:read slides:presentation:create slides:presentation:write_only wiki:wiki:readonly"
```

如果 `wiki:wiki:readonly` 缺失，先补授权：

```bash
lark-cli auth login --scope "wiki:wiki:readonly"
```

### 1. 准备一次独立 run_dir

```bash
RUN_DIR=/tmp/doc-to-slides-smoke-$(date +%Y%m%d-%H%M%S)
mkdir -p "$RUN_DIR"
echo "$RUN_DIR"
```

### 2. 验证 Wiki Source 能解析到底层文档

把 `<SOURCE_WIKI_URL>` 换成真实 wiki 文档链接：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py resolve-source \
  --doc-url "<SOURCE_WIKI_URL>" \
  --run-dir "$RUN_DIR"

cat "$RUN_DIR/resolved-source.json"
```

检查点：

- 命令成功退出
- `resolved-source.json` 已生成
- 如果来源本身是 wiki，后续 `fetch` 不应再直接拿 wiki token 去抓

### 3. 验证 Wiki Source 抓取全文

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py fetch \
  --resolved-source "$RUN_DIR/resolved-source.json" \
  --run-dir "$RUN_DIR"

cat "$RUN_DIR/source.json"
sed -n '1,80p' "$RUN_DIR/source.md"
```

检查点：

- `source.json`、`source.md` 已生成
- `source.json.resolved_fetch_target` 是底层 `doc` / `docx` token，而不是 wiki token
- `source.md` 有正文，不是空文件

### 4. 验证 Wiki Target 预解析

把 `<TARGET_WIKI_URL>` 换成真实指向 Slides 的 wiki 链接：

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py resolve-target \
  --target-slides-url "<TARGET_WIKI_URL>" \
  --run-dir "$RUN_DIR"

cat "$RUN_DIR/resolved-target.json"
```

检查点：

- 命令成功退出
- `resolved-target.json` 已生成
- `xml_presentation_id` 存在，说明目标 wiki 已被解析到底层 Slides

### 5. 准备最小 append outline 并渲染

- 把 `presentation.target_mode` 改成 `append`
- 首张 slide 如果是章节页，应写成 `role = "section"` 且 `section_divider = true`

```bash
cp lark-workflow-doc-to-slides/templates/outline.json "$RUN_DIR/outline.json"

python3 - "$RUN_DIR/outline.json" <<'PY'
import json, sys
path = sys.argv[1]
with open(path, "r", encoding="utf-8") as f:
    data = json.load(f)
presentation = data.setdefault("presentation", {})
presentation["target_mode"] = "append"
slides = data.setdefault("slides", [])
if slides:
    slides[0]["role"] = "section"
    slides[0]["section_divider"] = True
    slides[0]["layout"] = "title-only"
    slides[0]["key_points"] = ["本次新增内容"]
with open(path, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)
    f.write("\n")
PY

grep '"target_mode": "append"' "$RUN_DIR/outline.json"
grep '"section_divider": true' "$RUN_DIR/outline.json"

python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py validate-outline \
  --outline "$RUN_DIR/outline.json"

python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py render \
  --outline "$RUN_DIR/outline.json" \
  --run-dir "$RUN_DIR"
```

检查点：

- `outline.json` 里的 `presentation.target_mode` 已明确是 `append`
- 首张 slide 已变成 append-safe 的章节页，而不是模板里的通用封面
- `validate-outline` 成功
- `slides.json`、`render-summary.json` 已生成

### 6. 执行真实 append publish

```bash
python3 lark-workflow-doc-to-slides/scripts/doc_to_slides.py publish \
  --outline "$RUN_DIR/outline.json" \
  --slides-json "$RUN_DIR/slides.json" \
  --target-slides-url "<TARGET_WIKI_URL>" \
  --run-dir "$RUN_DIR"

cat "$RUN_DIR/publish-result.json"
```

检查点：

- `publish-result.json` 已生成
- `target_mode = append`
- `xml_presentation_id` 存在
- `slide_ids` 非空
- `slides_added` 与实际新增页数一致

### 7. 人工验收

在飞书里打开目标 Slides，确认：

- 新增页确实被追加到了目标 deck
- 旧页没有被删除、覆盖或重排
- 新增页内容与 `slides.json` 对应
- 如果目标是 wiki 链接，最终落到的是该 wiki 背后的 Slides，而不是新建了别的 deck

### 8. 失败时优先排查

- `wiki:wiki:readonly` 是否已授权
- source wiki 最终是否解析为 `doc` / `docx`
- target wiki 最终是否解析为 `slides`
- `outline.json` 是否在 `render` 后又被改过
- `publish-result.json` 是否记录了部分成功的 slide ids

## 相关文件

- [SKILL.md](./SKILL.md)
- [workflow-new-slides.md](./references/workflow-new-slides.md)
- [workflow-append-slides.md](./references/workflow-append-slides.md)
- [content-modes.md](./references/content-modes.md)
- [slide-authoring-rules.md](./references/slide-authoring-rules.md)
- [outline.json](./templates/outline.json)
