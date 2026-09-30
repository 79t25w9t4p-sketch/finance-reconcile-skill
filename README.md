# 财务明细核对 Skill

**首次确认核对方式，以后换两份文件即可重跑。** 面向重复的台账/系统明细核对，AI组织流程，本地脚本精确计算，用户在Excel里处理例外。

## 你会拿到什么

- 一份复核工作簿：汇总、待处理、已匹配、历史记录。
- 可筛选的差异、原始来源行、黄色处理状态与意见列。
- 重跑时沿用未变化条目的意见；内容变化则要求重新复核；本次未出现的条目保留在历史。
- 本地核对配置，保存字段映射、工作表和表头，省掉重复解释。

## 安装与使用

下载仓库，将目录放到 `~/.codex/skills/finance-reconcile/`。运行环境为Node.js 18+、Python 3.9+，Python环境需要 `python3 -m pip install -r requirements.txt`。已有依赖可直接复用。新开Codex会话：

> 使用 $finance-reconcile 核对这两份表，保存这次核对方式，给我可以填写处理意见的Excel。

以后说：

> 按上次的配置核对这两份新表，沿用这份上次结果里的处理意见。

技能会运行以下流程，用户不必手工拼命令。

## 独立运行

首次（在仓库目录，示例数据全部虚构）：

```sh
python3 scripts/workflow.py --left examples/左表.csv --right examples/右表.csv --left-keys '["客户","编号"]' --right-keys '["客户名称","单据号"]' --left-amount 金额 --right-amount 余额 --save-config 核对配置.json --out 首次结果
```

在 `首次结果/复核工作簿.xlsx` 的黄色列填写状态与意见。以后换新文件重跑：

```sh
python3 scripts/workflow.py --config 核对配置.json --left 新左表.csv --right 新右表.csv --previous 首次结果/复核工作簿.xlsx --out 本次结果
```

配置和结果不覆盖已有文件。非首行表头用 `--left-header-row 3`，多工作表用 `--left-sheet 明细`，右表有对应参数。用哪个Python装依赖，就用哪个Python运行。Node不在PATH时用 `--node` 指定路径。

## 意见复用规则

| 情况 | 重跑行为 |
|---|---|
| 同一编号/联合键，完整行内容与核对结果未变化 | 沿用处理状态和意见，即使源行顺序变了 |
| 金额、说明或匹配关系变化 | 保留旧意见，状态改为需重新复核 |
| 新编号 | 标记新增 |
| 上次有、本次没有 | 移到历史记录，不当作已解决 |
| 重复编号 / 空编号 | 重复键始终要求复核；空编号不继承意见 |
| 配置或表头变化 | 停止，要求核实格式和映射 |

请保留Excel表头及隐藏标识列，处理意见填文本。历史重现条目按新增处理，不自动恢复久远意见。复核意见保存在你指定的本地工作簿，AI默认只读汇总。

## 能力边界

当前 v0.5.0。支持UTF-8 CSV/TSV及XLSX、联合键、精确到分的金额核对；重复/无效金额保留为异常。分批回款、合并付款由人核实。XLSX公式需先重算并导出值，不支持旧xls。最多5000条明细、50列；金额字段保留两位小数。更多边界见 [SKILL.md](SKILL.md)。底层CSV模式仍可用 `node scripts/reconcile.cjs --help`。

工作簿金额以精确文本显示，避免Excel长数字截断；这是复核结果，不是公式计算模型。脚本不联网，AI对话本身是否联网由宿主决定。

## 验证与贡献

```sh
node --test tests/*.test.cjs
python3 tests/test_integration.py
python3 tests/test_workflow.py
```

测试覆盖配置复用、完整行变化、意见保留、历史记录、单元格公式防护和失败不留半成品。没有声称真实用户节省了多少时间。提交问题请附最小虚构样例，不要上传真实账单和个人信息。

主要参考 [finance-copilot-skills](https://github.com/cxtx/finance-copilot-skills) 与 [cpa-skills](https://github.com/adoptai/cpa-skills) 的工作流设计；早期参考 [csv-diff](https://github.com/simonw/csv-diff) 与 [csv-diff-browser](https://github.com/SaurabhBaranawal/csv-diff-browser)。实现由AI协助编写，没有复制上述仓库代码；比较见 [参考对比.md](参考对比.md)。MIT许可。
