"""使用 10 个目标主题和干扰页验证排名。python test_search.py [kb_tools.py]"""
import contextlib
import importlib.util
import io
from pathlib import Path
import re
import sys
import tempfile
import unittest

# 正式名称、页面维护者登记的别名、日常问法、可检索的正文。
CASES = [
    ("急会诊流程", "急会诊", "急会诊多久要到场", "记录到场时间和处置结果。"),
    ("非计划再次手术", "非计再手术", "非计再手术怎么统计？", "统计时核对手术记录。"),
    ("危急值报告", "危值上报", "请问危值上报由谁接收", "记录通知、接收和复核。"),
    ("病案首页质量控制", "首页质控", "首页质控要检查哪些内容", "复核诊断编码及填报内容。"),
    ("不良事件上报", "事件直报", "事件直报怎么操作", "按步骤提交事件记录。"),
    ("会议纪要归档", "纪要归档", "纪要归档放在哪里", "按会议日期保存。"),
    ("设备预防性维护", "设备保养", "设备保养多久做一次", "检查设备状态并填写记录。"),
    ("项目经费报销", "项目报账", "项目报账需要哪些凭证", "申请时提交发票和审批单。"),
    ("个人防护装备", "PPE", "PPE怎么选", "根据作业场景选择防护用品。"),
    ("持续集成", "CI", "怎么配置 CI？", "代码提交后运行构建。"),
]


def load_tool(path):
    spec = importlib.util.spec_from_file_location("kb_search_target", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


TOOL = load_tool(Path(__file__).with_name("kb_tools.py"))


class SearchTest(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.pages = self.root / "wiki" / "concepts"
        self.pages.mkdir(parents=True)
        for title, alias, _, body in CASES:
            self.add(title, body, alias)
        for term in ("时间", "设备", "经费", "质量", "上报", "归档", "操作", "接收"):
            self.add(term + "统计说明", "登记时间、交接时间、操作步骤和相关内容。")

    def add(self, title, body, alias=""):
        header = f"aliases: [{alias}]\n" if alias else ""
        (self.pages / (title + ".md")).write_text(
            f"---\ntitle: {title}\ntype: concept\ntags: [资料]\n{header}---\n{body}\n",
            encoding="utf-8")

    def query(self, *words, tag="", kind="", limit=5):
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            status = TOOL.cmd_search(self.root, words, kind, tag, limit)
        self.assertEqual(status, 0)
        return re.findall(r"^\d+\. .*?\[concept\]\s+(.+)$", out.getvalue(), re.M)

    def test_ranking(self):
        scores = []
        for column in range(3):
            positions = []
            for case in CASES:
                results = self.query(case[column])
                positions.append(results.index(case[0])+1 if case[0] in results else 0)
            scores.append(positions)
        print("正式名 / 简称 / 问句（目标页前五名次，0=未进入）", scores)
        self.assertEqual(scores[0], [1]*10)
        self.assertTrue(all(scores[1]) and all(scores[2]))

    def test_old_inputs(self):
        self.assertEqual(self.query("经费", tag="未登记"), [])
        self.assertEqual(self.query("设备", kind="entity"), [])
        self.assertEqual(len(self.query(tag="资料", limit=2)), 2)
        self.assertEqual(self.query("火星矿产"), [])
        self.assertIn("急会诊流程", self.query("急会诊", "无关词"))
        self.assertEqual(self.query("怎么"), [])

    def test_literal_before_expansion(self):
        self.add("真实问答", "上次会议的提问是急会诊多久要到场。")
        results = self.query("急会诊多久要到场")
        self.assertEqual(results[:2], ["真实问答", "急会诊流程"])
        self.assertNotIn("时间统计说明", results)
        self.assertEqual(self.query("星际运输需要多久"), [])

    def test_question_without_alias(self):
        self.add("复核办法", "交接班记录交班明细及签字。")
        self.assertIn("复核办法", self.query("交接班怎么记录"))
        self.assertEqual(self.query("ppe"), ["个人防护装备"])


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1].endswith(".py"):
        TOOL = load_tool(Path(sys.argv.pop(1)))
    unittest.main()
