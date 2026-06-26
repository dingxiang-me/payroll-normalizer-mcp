# -*- coding: utf-8 -*-
"""payroll-normalizer MCP server —— 让任意支持 MCP 的 AI 工具都能"自动整理工资表"。

暴露 4 个工具：
  - inspect_payroll(folder)        逐文件返回表头/样本/自动识别的字段映射（供模型判断非标表头）
  - normalize_payroll(folder, ...) 整理为社保测算标准模板 + 报告（支持 overrides 修正非标表）
  - generate_blank_template(...)   生成带下拉+说明的空白标准模板
  - standard_columns()             返回标准 10 列定义与口径，便于模型理解
"""
import json
from typing import Optional

from mcp.server.fastmcp import FastMCP

from . import core

mcp = FastMCP("payroll-normalizer")


@mcp.tool()
def standard_columns() -> str:
    """返回社保测算标准模板的 10 列定义、身份类型可选值与关键口径（应发≠实发等）。
    模型在做列映射前应先读这个。"""
    return json.dumps({
        "columns": core.COLS,
        "required": ["工号/唯一标识", "姓名", "所属主体", "所属年月", "应发工资(税前合计)"],
        "staff_types": core.STYPES,
        "rules": [
            "一行 = 某人某月某主体；同一人跨主体/跨月分多行，按自然人归并。",
            "应发工资 = 税前合计(基本+岗位+绩效+奖金+补贴…)，不含单位社保；个人社保/个人公积金计入应发。切勿用实发。",
            "若原表只有实发：应发 = 实发 + 个人社保 + 个人公积金 + 个税；或各收入项合计 − 考勤扣款。",
            "身份类型留空默认『全日制正式』；前3类(全日制正式/试用期/老板股东)参保计算，后5类剔除。",
            "所属年月格式 YYYY-MM。",
        ],
    }, ensure_ascii=False, indent=2)


@mcp.tool()
def inspect_payroll(folder: str) -> str:
    """检查文件夹内每个工资表(.xlsx/.xls/.csv)：返回表头、前3行样本、自动识别到的字段映射、应发口径与问题。
    当某文件无法自动映射时，模型据此判断每列含义，再在 normalize_payroll 的 overrides 里给出 column_map/entity/ym。

    folder: 存放工资表的文件夹绝对路径。
    """
    try:
        return json.dumps(core.inspect(folder), ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def normalize_payroll(folder: str, output_dir: Optional[str] = None, overrides_json: Optional[str] = None) -> str:
    """把文件夹内所有工资表整理成『社保测算标准模板』(xlsx) + 整理报告(md)，按自然人跨主体跨月归并。

    folder: 存放工资表的文件夹绝对路径。
    output_dir: 产出目录(默认与 folder 相同)。
    overrides_json: 可选 JSON 字符串，用于修正自动识别不了的文件，形如：
        {"某文件.xlsx": {"entity": "甲公司", "ym": "2025-03",
                         "column_map": {"name": "员工", "gross": "税前总额", "id": 0}}}
      column_map 的值可为"表头文字"或列序号(从0起)；字段名取标准字段：
      name/id/entity/ym/gross/total/unit_ss/net/stype/gjj_paid/gjj_base/gjj_ratio。

    返回 JSON：产出路径、记录数、应发口径分布、需人工接管的文件清单、报告 markdown。
    """
    try:
        overrides = json.loads(overrides_json) if overrides_json else None
        res = core.normalize(folder, output_dir, overrides)
        return json.dumps(res, ensure_ascii=False, indent=2)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


@mcp.tool()
def generate_blank_template(output_path: Optional[str] = None) -> str:
    """生成空白的『社保测算标准工资模板』(xlsx)：含顶部填表说明、身份类型/是否缴公积金下拉、示例行。
    output_path: 输出文件绝对路径(默认当前目录下 社保测算-标准工资模板.xlsx)。返回实际路径。
    """
    try:
        return json.dumps({"template_path": core.generate_blank_template(output_path)}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


def main():
    mcp.run()


if __name__ == "__main__":
    main()
