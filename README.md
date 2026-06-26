# payroll-normalizer-mcp

> 让**任意支持 MCP 的 AI 工具**（Claude Code / Claude Desktop / Cursor / Windsurf / Cline / Zed / 支持 MCP 的 ChatGPT 等）都能把**五花八门的工资表一键整理成「社保测算标准模板」**。

把企业各种格式的工资表（.xlsx/.xls/.csv，多主体多月份）按"自然人跨主体跨月"归并，自动识别列名、把实发换算回**应发(税前)**，输出 10 列标准模板 + 整理报告。遇到非标表头时，AI 客户端可先 `inspect_payroll` 看表头样本、判断列含义，再带 `overrides` 调 `normalize_payroll`。

## 工具（MCP tools）
| 工具 | 作用 |
|---|---|
| `standard_columns` | 返回标准 10 列定义、身份类型可选值、应发≠实发等口径（映射前先读） |
| `inspect_payroll(folder)` | 逐文件返回表头、前 3 行样本、自动识别的字段映射、应发口径与问题 |
| `normalize_payroll(folder, output_dir?, overrides_json?)` | 整理为标准模板 xlsx + 报告 md；`overrides` 修正非标表 |
| `generate_blank_template(output_path?)` | 生成带下拉+说明的空白标准模板 |

## 安装：在各家工具里加这个 MCP server

无需先发布到 PyPI——用 `uvx` 直接从 GitHub 运行（需本机有 [uv](https://docs.astral.sh/uv/)）。通用配置：

```json
{
  "mcpServers": {
    "payroll-normalizer": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/dingxiang-me/payroll-normalizer-mcp", "payroll-normalizer-mcp"]
    }
  }
}
```

放到对应位置即可：

- **Claude Code**（一条命令）：
  ```bash
  claude mcp add payroll-normalizer -- uvx --from git+https://github.com/dingxiang-me/payroll-normalizer-mcp payroll-normalizer-mcp
  ```
- **Claude Desktop**：`claude_desktop_config.json` → `mcpServers`（路径见 Settings › Developer）。
- **Cursor**：项目根 `.cursor/mcp.json`（或全局 `~/.cursor/mcp.json`）→ 同上 `mcpServers`。
- **Windsurf**：`~/.codeium/windsurf/mcp_config.json` → 同上。
- **Cline / Zed / 其他**：填到各自的 MCP 配置里，`command`/`args` 一致。

> 想更快启动可先发布到 PyPI，再把 `args` 换成 `["payroll-normalizer-mcp"]`。

## 用法
配置好后，直接对 AI 说：

> 「把 `/path/to/工资表文件夹` 里的工资表整理成社保测算标准模板」

AI 会自动调用 `inspect_payroll` →（必要时）判断非标列 → `normalize_payroll`，在该文件夹产出 `社保测算标准模板_整理结果.xlsx` 和 `整理报告.md`。

## 依赖
- [uv](https://docs.astral.sh/uv/)（提供 `uvx`）
- 运行时自动拉取 `mcp`、`openpyxl`；旧版 `.xls` 另需 `xlrd`（或先另存为 .xlsx）

## 配套
产出的标准模板可直接导入「社保公积金薪酬优化测算工具」做测算。本服务**只做数据整理，不做社保/个税计算**。

## 许可
MIT
