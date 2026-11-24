---
name: kb-query
description: 查询库内知识、核对出处或按要求保存查询结果。
---

# kb-query

1. 使用 `python kb_tools.py search 关键词` 检索，必要时换同义词。读取相关页面，沿来源页的 `sources` 核对原文。
2. 根据找到的依据作答，列出页面和出处，指出资料缺口或相互矛盾的说法。
3. 普通查询直接答复。需要保存时按 AGENTS.md 确认分级，先建检查点，按 query 模板保存，再重建索引、追加日志、lint 并以 query 类型提交。
