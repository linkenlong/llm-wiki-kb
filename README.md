# llm-wiki-kb

把手头的文档整理成可以提问、可以查回原文的本地 Markdown 知识库。你挑选资料、确定收录范围，Codex 根据项目规则归档原文、整理知识页，并在新资料涉及已有主题时更新相应页面。

例如，你先后收录两份项目会议纪要。每份纪要各有一张来源页，同一项目的决定可以汇总在一张主题页。之后问“这个决定最早出现在哪次会议？”，就可以从回答中的页面链接查回纪要原文。

## 搭建第一个知识库

准备 Python 3（批量转换模块需要 Python 3.10 或更高版本）。建库、检索和单份 Word 转换只用 Python 标准库。安装 Git 可启用检查点和撤销；用 Codex 打开生成的目录即可按项目规则维护知识页。

1. 下载并解压 [v3.2 发布包](https://github.com/linkenlong/llm-wiki-kb/releases/tag/v3.2)，或下载当前分支的源码以使用新版搜索。在包含 `setup_kb.py` 的目录打开终端。
2. 执行下列命令。示例把新库建在工具目录旁边，以便单独管理资料和 Git 历史。

   ```shell
   python setup_kb.py ../my-kb
   cd ../my-kb
   python kb_tools.py index
   python kb_tools.py lint
   ```

3. 在 Codex 中打开 **my-kb**，输入“帮我确定知识库的领域和标签体系”。选定的范围和标签记录在 `wiki/SCHEMA.md` 中。
4. 把第一份资料复制到 `inbox/`，输入“处理 inbox”。默认模式下，前三次摄取会先列出准备新建和更新的页面，等你确认后再写页。
5. 处理完成后，打开 `wiki/index.md` 浏览页面，或直接问 Codex：“这份资料有哪些关键决定？请列出依据的页面和原文路径。”

## 资料存在哪里

| 位置 | 用途 |
| --- | --- |
| `inbox/` | 待处理资料的收件箱 |
| `raw/` | 归档原文。来源页的 `sources` 字段指向这里 |
| `wiki/` | 来源、概念、实体、对比和查询页。索引在 `index.md`，操作记录在 `log.md` |
| `outputs/` | 一次性报告、批量转换的阅读稿和待复核截图 |
| `templates/`、`.agents/skills/` | 页面模板，以及摄取、查询、检查和维护四项技能 |

页面之间按内容关联添加双向链接，可以从一份原文和一张来源页开始。以后再收录相关资料时，补上有用的链接即可。

## 检索和检查

摄取和查询都可以直接浏览索引、跟随页面链接，或用编辑器全文搜索 `wiki/`。建页前仍须查重：原名查不到，就换用主题词或简称，打开候选页核对。`kb_tools.py search` 是可选工具，适合中文问句、结果排序和按标签筛选。

以下命令在生成的知识库根目录运行。

```shell
python kb_tools.py search 会议 决定
python kb_tools.py recent 5
python kb_tools.py index
python kb_tools.py lint
```

`search` 在页面的标题、摘要、标签和正文中检索。`recent` 查看最近操作。`index` 根据页面摘要重建目录。`lint` 可发现失效的页面链接、来源页引用不到的原件和索引不同步等问题。对内容含义或结论的疑问，仍需查回原文。

### 用日常问法搜索（v3.2 之后的源码版本）

```shell
python kb_tools.py search "急会诊多久要到场"
python kb_tools.py search 首页质控 --type concept --limit 5
```

搜索会优先显示含完整查询词的页面，再用“多久”“怎么”“哪些”等问句线索拆出片段，也会识别问句中完整出现的中文页名和已登记的中文别名。结果会标明命中位置和问句片段。已有页面可直接参与搜索。

如果读者常用简称，可以在对应知识页的 YAML 页头添加一行，例如在“病案首页质量控制”页中写 `aliases: [首页质控, 病案首页质控]`。别名只用于该页匹配。检索范围是 Wiki 正文和页面信息；未登记的同义表达可能漏检，可换关键词重试。

## 把 Word、PDF 和其他资料转成 Markdown

**单份 Word** 可以直接在本地转换。把示例路径换成已归档的实际文件名。

```shell
python kb_tools.py convert "raw/notes/会议纪要.docx"
```

阅读稿生成在原件旁边，并登记来源和哈希。Word 中的图片需要对照原件查看。

**批量转换** 适合一个归档目录中混有 PDF、Word、PowerPoint、Excel、HTML 或文本的情况。首次使用时安装可选依赖，再选择要处理的目录。

```shell
python -m pip install "markitdown[docx,pdf,pptx,xlsx]" pymupdf
python kb_tools.py convert-batch raw/notes
```

每批结果保存在 `outputs/转换/` 下的独立目录。先看 `report.md`，再核对标为“待复核”的阅读稿和截图。不同目录或不同扩展名的同名资料会分别保存。退出码 `0` 表示正常完成，`2` 表示有资料待复核，`1` 表示发生错误。

默认在本地处理。PDF 中缺少可提取文字或疑似乱码的页面会保存截图。若要使用 Anthropic 转写这些页面，需另装 `anthropic`，设置 `ANTHROPIC_API_KEY` 和 `FALLBACK_MODEL`，并在命令中显式加入 `--cloud`；该选项允许发送整页截图。旧版 `.doc`、`.ppt`、`.xls` 转换另需 LibreOffice。

## 更新已有知识库

从新版工具目录运行 `python setup_kb.py ../my-kb --update`。它会更新配套脚本、维护规则和技能文件，已有同名模板及 `wiki/SCHEMA.md` 会保留。如果自行改过工具或技能，建议先提交当前版本，便于核对升级差异。

项目分工与确认规则见 [AGENTS.md](AGENTS.md)。具体流程见 [资料摄取](skills/kb-ingest/SKILL.md)、[知识查询](skills/kb-query/SKILL.md)、[库内检查](skills/kb-lint/SKILL.md)和[知识库维护](skills/kb-maintain/SKILL.md)。
