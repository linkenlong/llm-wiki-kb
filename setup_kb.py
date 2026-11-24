#!/usr/bin/env python3
"""一键生成 LLM Wiki 知识库骨架。

用法：
    python setup_kb.py                # 在当前目录下创建 knowledge-base/
    python setup_kb.py my-kb          # 指定目录名
    python setup_kb.py /path/to/my-kb # 指定完整路径
    python setup_kb.py my-kb --no-git # 不初始化 git（检查点和回滚功能将不可用）
    python setup_kb.py my-kb --update # 升级已有知识库：用新版覆盖 AGENTS.md、kb_tools.py 和技能
                                      # （SCHEMA、index、log、templates、wiki 内容和 raw 资料都不会动）

特点：
- 只使用标准库；已存在的文件不会被覆盖，可以安全地重复运行。
- 如果本脚本旁边有 AGENTS.md、kb_tools.py 和 skills/ 目录，会一并复制到知识库里
  （skills/ 下的技能会安装到 .agents/skills/，Codex 会从那里自动发现）。
- 默认会把知识库初始化为 git 仓库并提交第一个版本（检查点和回滚依赖它）。
"""
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path


def _force_utf8():
    """输出被管道捕获时，Windows 等环境的默认编码（cp1252、cp936 等）无法输出中文，会直接抛异常。
    统一改成 UTF-8；个别无法编码的字符用 ? 代替，绝不因为编码崩溃。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError, OSError):
            pass


_force_utf8()

TODAY = date.today().isoformat()

# 空目录会放一个 .gitkeep，方便 git 记录。
DIRS = [
    "inbox",
    "raw/articles",
    "raw/papers",
    "raw/notes",
    "raw/assets",
    "wiki/sources",
    "wiki/concepts",
    "wiki/entities",
    "wiki/comparisons",
    "wiki/queries",
    "outputs",
    "templates",
]

# 从脚本所在目录复制到知识库根目录的配套文件
COPY_FILES = ["AGENTS.md", "kb_tools.py", "convert_to_md.py"]
# 从脚本所在目录复制到知识库的配套目录：(源目录, 目标目录)
COPY_DIRS = [("skills", ".agents/skills")]

FILES = {}

FILES[".gitignore"] = """\
# 由 kb_tools 添加的默认忽略项（不需要的项可以删掉，但请保留这一行，否则会再次自动补充）
.obsidian/
.DS_Store
Thumbs.db
__pycache__/
*.pyc
"""

FILES["inbox/README.md"] = """\
# 收件箱

把还没处理的新资料先丢到这里（文章、PDF、笔记、剪藏都可以），不用先分类。

处理流程：

1. 您对 agent 说：「处理 inbox」。
2. 默认和自动模式下，归类明确的文件可先归档；分类不清时先询问。谨慎模式下，归档前先确认。
3. 文件移入 `raw/` 并登记哈希后，列出建页和更新计划，按 AGENTS.md 的确认分级处理。默认模式的前三次摄取先确认写页计划。
4. 处理完成后，inbox 里不应再有该文件。

注意：这里的文件还没有被摄取，不属于知识库的正式内容。
"""

FILES["wiki/SCHEMA.md"] = """\
# Wiki Schema

> 本文件定义这个知识库的领域、约定和标签体系，是 agent 的行为约束。可以随时修改；agent 在摄取等需要用到它的任务开始时读取。

## 领域

【待填写】这个知识库覆盖什么？越具体越好，例如「大模型应用与 RAG 研究」「个人健康管理」「创业公司情报」。

## 写作语言

中文（专有名词保留原文）。

## 确认模式

当前：默认

可选：谨慎 / 默认 / 自动。含义和分级规则见 AGENTS.md 的「确认分级」。在对话中说「自动模式」只对本次会话有效；想长期生效，请直接修改上面的「当前」。

## 约定

- 文件名与页面名一致，不含空格，单词之间用连字符；中文名或英文名均可，但同一个知识库保持一种风格。
- [[wikilink]] 必须与文件名（不含 .md）完全一致。
- 每个页面以 YAML frontmatter 开头（格式见下）。
- 按实际关系添加 [[wikilink]]，出链数量仅作整理提示。来源页与已有衍生页互链；暂无相关页面时可只保留原文路径。
- 更新页面时同步更新 updated 日期。
- frontmatter 里的 title 和 summary，如果含有「: 」（冒号加空格）或「 #」（空格加井号），必须用双引号包起来，否则 YAML 会报错或把后半段当注释截断。
- wiki/index.md 由 `python kb_tools.py index` 根据各页面的 summary 自动生成，不要手动编辑；每次操作都要追加到 wiki/log.md。
- summary 是一句话说明（不超过 80 字），要具体到让人不打开页面就能判断是否相关；索引和搜索都依赖它。
- 溯源标记：综合了 3 个以上来源的页面，在依据某个具体来源的段落末尾加 ^[raw/articles/文件名.md]，方便追溯每个论断。单来源页面只需 sources 字段。

## 写作风格

页面是给人读的综合说明，不是资料的复述，也不是流水账。

- 按概念、决策、争议和下一步来组织内容，而不是按时间顺序或资料顺序罗列。
- 先写结论，再写依据；第一句话就让读者知道这页讲什么。
- 日期、状态、编号只在有意义时出现（比如说明两种矛盾说法的先后），不要为了显得完整而堆砌。
- 不要写成一行一条的清单。有逻辑关系的内容用短段落连接，只有并列的条目才用列表。
- 避免空话和填充语，例如「本文介绍了」「值得注意的是」。
- 保持中立、可核对：事实写明来源；推断和观点要标明是推断和观点。
- 页面要能在 30 秒内读完，细节放到单独的页面，用链接连接。

## Frontmatter

```yaml
---
title: 页面标题
type: source | concept | entity | comparison | query
summary: 一句话说明（不超过 80 字），会自动显示在 index 中
created: YYYY-MM-DD
updated: YYYY-MM-DD
tags: [必须来自下面的标签体系]
sources: [raw/articles/source-name.md]
# 以下为可选的质量信号：
confidence: high | medium | low     # 论断的支持程度
contested: true                      # 存在尚未解决的矛盾时设置
contradictions: [other-page]         # 与哪些页面冲突
---
```

confidence 和 contested 对观点性强或变化快的主题尤其推荐。lint 会把 contested 为 true 和 confidence 为 low 的页面列出来供复查，避免薄弱的论断悄悄变成「公认事实」。

## raw 资料的 frontmatter

由 `python kb_tools.py stamp` 自动写入，用于以后发现资料内容的变化：

```yaml
---
source_url: https://example.com/article
ingested: YYYY-MM-DD
sha256: （对正文计算的哈希）
---
```

PDF 等不能写 frontmatter 的文件，会在旁边生成一个同名的 .meta.md 记录文件。

## 标签体系

（尚未定义。首次摄取前，agent 与您一起确定 10 到 20 个顶层标签，然后在这里逐条列出，每个标签用反引号包起来。新增标签时，先加到这里，再使用。）

规则：页面上的每个标签都必须出现在本节。这样可以防止标签随意膨胀。

## 建页阈值

- 新建页面：某个实体或概念出现在 2 份以上来源中，或者是某一份来源的核心内容。
- 补充已有页面：来源提到了已经有页面的内容。
- 不要建页：一笔带过的内容、次要细节、领域之外的东西。
- 拆分页面：超过约 200 行时，拆成子主题并互相链接。
- 归档页面：内容被完全取代时，移到 wiki/_archive/，并从 index 中移除。

## 各类页面应包含的内容

- 来源摘要页：来源信息、核心观点、衍生页面、与已有知识的关系。
- 概念页：定义或解释、当前的认识、不同来源的说法与开放问题、相关概念。
- 实体页：是什么、关键事实和日期、与其他实体的关系、来源。
- 对比页：对比什么以及为什么、对比维度（优先用表格）、结论或综合、来源。
- 查询页：原始问题、结论、依据的页面、日期。

## 更新策略

新信息与已有内容冲突时：

1. 先看日期，较新的来源通常优先。
2. 如果确实矛盾，并列写出两种说法，各自注明日期和来源。
3. 在 frontmatter 中标记 contradictions，必要时设置 contested: true。
4. 在 lint 报告中提醒用户复查。

## 规模规则

- index.md 某个分区超过 50 条：按首字母或子领域拆分。
- index.md 总条目超过 200：创建 wiki/_meta/topic-map.md，按主题串联页面。
- log.md 超过 500 条：重命名为 log-YYYY.md，另起新文件。
"""

FILES["wiki/index.md"] = """\
# 知识库索引

> 本文件由 `python kb_tools.py index` 自动生成，请勿手动编辑；每条说明来自页面 frontmatter 的 summary 字段。
> 最后更新：{today} | 页面总数：0

## 来源

## 概念

## 实体

## 对比

## 查询
"""

FILES["wiki/log.md"] = """\
# 操作日志

> 按时间顺序记录所有操作，只追加，不修改。
> 格式：`## [YYYY-MM-DD] 操作类型 | 对象`，下面列出涉及的文件。
> 操作类型：ingest（摄取）、update、query、lint、create、archive、delete。
> 超过 500 条时轮转：重命名为 log-YYYY.md，另起新文件。

## [{today}] create | 知识库初始化
- 结构已创建：SCHEMA.md、index.md、log.md
- 下一步：填写 SCHEMA.md 的领域和标签体系
"""

FILES["templates/source.md"] = """\
---
title: {{资料标题}}
type: source
summary: {{一句话说明，不超过 80 字，会显示在 index 中}}
created: {{YYYY-MM-DD}}
updated: {{YYYY-MM-DD}}
tags: [{{标签，必须来自 SCHEMA.md 的标签体系}}]
sources: [{{raw/articles/xxx.md}}]
---

{{一两句话概括这份资料讲了什么。}}

## 核心观点
- {{观点 1}}
- {{观点 2}}

## 衍生页面
- 概念：[[{{概念页名}}]]
- 实体：[[{{实体页名}}]]

## 与已有知识的关系
{{补充、印证或与哪些已有页面矛盾。没有则写"无"。}}
"""

FILES["templates/concept.md"] = """\
---
title: {{概念名称}}
type: concept
summary: {{一句话说明，不超过 80 字，会显示在 index 中}}
created: {{YYYY-MM-DD}}
updated: {{YYYY-MM-DD}}
tags: [{{标签}}]
sources: [{{raw/articles/xxx.md}}]
confidence: {{high | medium | low}}
---

{{一句话定义。}}

## 要点
- {{要点 1}}
- {{要点 2}}

## 不同来源的说法
- {{来源 A}}（{{日期}}）：{{观点}}
- {{来源 B}}（{{日期}}）：{{观点；与 A 矛盾时请明确标注}}

## 开放问题
- {{尚无定论的问题}}

## 相关页面
- [[{{相关概念或实体}}]]
- 来源：[[{{来源页名}}]]
"""

FILES["templates/entity.md"] = """\
---
title: {{实体名称}}
type: entity
summary: {{一句话说明，不超过 80 字，会显示在 index 中}}
created: {{YYYY-MM-DD}}
updated: {{YYYY-MM-DD}}
tags: [{{标签}}]
sources: [{{raw/articles/xxx.md}}]
confidence: {{high | medium | low}}
---

{{一句话说明这个实体是什么（人物、组织、工具、项目、模型等）。}}

## 关键事实
- {{事实 1，注明日期}}
- {{事实 2}}

## 与其他实体的关系
- [[{{相关实体}}]]：{{关系}}

## 相关页面
- 来源：[[{{来源页名}}]]
"""

FILES["templates/comparison.md"] = """\
---
title: {{对比主题，如 A 与 B}}
type: comparison
summary: {{一句话说明，不超过 80 字，会显示在 index 中}}
created: {{YYYY-MM-DD}}
updated: {{YYYY-MM-DD}}
tags: [{{标签}}]
sources: [{{raw/articles/xxx.md}}]
confidence: {{high | medium | low}}
---

{{为什么要对比这些对象。}}

## 对比维度

| 维度 | [[{{对象 A}}]] | [[{{对象 B}}]] |
|---|---|---|
| {{维度 1}} | {{...}} | {{...}} |
| {{维度 2}} | {{...}} | {{...}} |

## 结论
{{综合判断，以及适用场景。}}
"""

FILES["templates/query.md"] = """\
---
title: {{问题的简短标题}}
type: query
summary: {{一句话说明，不超过 80 字，会显示在 index 中}}
created: {{YYYY-MM-DD}}
updated: {{YYYY-MM-DD}}
tags: [{{标签}}]
sources: [{{依据的来源文件，没有则留空}}]
confidence: {{high | medium | low}}
---

**问题：** {{原始问题}}

## 结论
{{回答。}}

## 依据
- [[{{依据的页面 A}}]]
- [[{{依据的页面 B}}]]
"""


def _git(target, *args, ident=False):
    cmd = ["git", "-C", str(target)]
    if ident:
        ok = all(
            subprocess.run(["git", "-C", str(target), "config", k],
                           capture_output=True, text=True).stdout.strip()
            for k in ("user.name", "user.email")
        )
        if not ok:
            cmd += ["-c", "user.name=kb-tools", "-c", "user.email=kb-tools@localhost"]
    return subprocess.run(cmd + list(args), capture_output=True, text=True,
                          encoding="utf-8", errors="replace")


def init_git(target: Path) -> str:
    """初始化 git 并提交第一个版本，返回给用户看的说明。"""
    if shutil.which("git") is None:
        return "未找到 git，已跳过 git 初始化（安装 git 后可运行 python kb_tools.py checkpoint --init）"
    if (target / ".git").exists():
        return "已经是 git 仓库，未重复初始化"
    r = _git(target, "rev-parse", "--show-toplevel")
    if r.returncode == 0:
        return (f"该目录位于另一个 git 仓库（{r.stdout.strip()}）之中，已跳过初始化；"
                "检查点功能要求知识库本身是独立的仓库")
    r = _git(target, "init")
    if r.returncode != 0:
        return f"git init 失败：{r.stderr.strip()}"
    _git(target, "add", "-A")
    r = _git(target, "commit", "-m", "init: 知识库骨架", ident=True)
    if r.returncode != 0:
        return f"已初始化 git，但首次提交失败：{r.stderr.strip() or r.stdout.strip()}"
    return "已初始化 git 仓库，并提交了第一个版本"


def main() -> None:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    no_git = "--no-git" in sys.argv
    update = "--update" in sys.argv
    target = Path(args[0] if args else "knowledge-base").resolve()
    created, skipped, copied, updated = [], [], [], []

    for d in DIRS:
        (target / d).mkdir(parents=True, exist_ok=True)

    for rel, content in FILES.items():
        path = target / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists():
            skipped.append(rel)
            continue
        path.write_text(content.replace("{today}", TODAY), encoding="utf-8")
        created.append(rel)

    here = Path(__file__).resolve().parent
    missing = []
    for name in COPY_FILES:
        src, dst = here / name, target / name
        if not src.exists():
            missing.append(name)
        elif dst.exists():
            if update and dst.read_bytes() != src.read_bytes():
                shutil.copy2(src, dst)
                updated.append(name)
            else:
                skipped.append(name)
        else:
            shutil.copy2(src, dst)
            copied.append(name)

    for src_name, dst_rel in COPY_DIRS:
        src = here / src_name
        if not src.is_dir():
            missing.append(f"{src_name}/")
            continue
        for f in sorted(src.rglob("*")):
            if not f.is_file():
                continue
            rel = f"{dst_rel}/{f.relative_to(src).as_posix()}"
            dst = target / rel
            if dst.exists():
                if update and dst.read_bytes() != f.read_bytes():
                    shutil.copy2(f, dst)
                    updated.append(rel)
                else:
                    skipped.append(rel)
                continue
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(f, dst)
            copied.append(rel)

    for d in DIRS:
        path = target / d
        if not any(path.iterdir()):
            (path / ".gitkeep").touch()

    git_note = "已按 --no-git 跳过 git 初始化" if no_git else init_git(target)

    print(f"知识库骨架已就绪：{target}\n")
    print(f"新建文件 {len(created)} 个：")
    for f in created:
        print(f"  + {f}")
    if copied:
        print(f"\n已复制配套文件：{', '.join(copied)}")
    if updated:
        print(f"\n已用新版覆盖 {len(updated)} 个配套文件：{', '.join(updated)}")
        print("（旧版本如果已提交到 git，可以用 git log 和 git checkout 找回）")
    if update and not updated:
        print("\n配套文件已经是最新版本，没有需要升级的。")
    if skipped:
        print(f"\n已存在，未覆盖 {len(skipped)} 个：{', '.join(skipped)}")
    if missing:
        print(f"\n注意：本脚本旁边没有 {', '.join(missing)}，请手动放到知识库里"
              "（AGENTS.md、kb_tools.py 放根目录，skills/ 里的技能放到 .agents/skills/）。")
    print(f"\ngit：{git_note}")
    print("\n知识库中的技能文件：")
    for skill in ("kb-ingest", "kb-query", "kb-lint", "kb-maintain"):
        installed = target / ".agents" / "skills" / skill / "SKILL.md"
        print(f"  {skill}：{'已安装' if installed.is_file() else '缺失'}")
    print("\n下一步：")
    print("  1. 在该目录启动 Codex")
    print("  2. 对它说：「帮我确定知识库的领域和标签体系」（会触发 kb-maintain 技能）")
    print("  3. 之后把资料放进 inbox/，说「处理 inbox」（会触发 kb-ingest 技能）")


if __name__ == "__main__":
    main()
