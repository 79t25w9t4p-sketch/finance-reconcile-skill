# 财务明细核对 Skill

给 Codex 的可复用技能：提供两份 Excel/CSV，按编号核对金额，得到差异表、异常清单和可复核运行记录。无需网页、服务器、API Key 或云端上传。

## 使用

将此文件夹复制到你自己的 `~/.codex/skills/finance-reconcile/`，在新的 Codex 会话中调用：

> 使用 $finance-reconcile 核对这两份表：左表按编号和金额，右表按单据号和余额。把差异整理出来。

运行环境：Node.js 18+。CSV/TSV 无第三方依赖；XLSX 另需 Python 3.9+ 与 openpyxl，可由使用者自行运行 `python3 -m pip install -r requirements.txt`，或使用已有环境。

XLSX 表头在第3行时，加 `--left-header-row 3`；多工作表用 `--left-sheet 明细` 指定。右表有对应参数。输出编号定位到原工作表行号。

也可单独运行脚本（在本文件夹中）：

```sh
node scripts/reconcile.cjs --left examples/左表.csv --right examples/右表.csv --left-key 编号 --left-amount 金额 --right-key 单据号 --right-amount 余额 --out 样例结果
node --test tests/*.test.cjs
```

输出目录必须是新目录。模拟样例包含一分钱级差异、重复编号、单边记录、空编号和无效金额；样例全部虚构，不含个人账目。限制与精确处理口径见 SKILL.md。

## 和一句提示词有什么区别

SKILL.md 告诉 AI 何时调用、如何确认口径；脚本负责可重复的精确运算；样例和测试让别人能够检查结果。重复单据不会被静默覆盖，未知金额不会被当作零。

## 参考与贡献

主要参考现成技能 [finance-copilot-skills](https://github.com/cxtx/finance-copilot-skills) 的本地确定性执行、运行清单与汇总披露设计，以及 [cpa-skills](https://github.com/adoptai/cpa-skills) 的记录完整性与证据追查思路。详细比较见 [参考对比.md](参考对比.md)。

早期参考 [simonw/csv-diff](https://github.com/simonw/csv-diff) 的按键匹配思路；参考 [csv-diff-browser](https://github.com/SaurabhBaranawal/csv-diff-browser) 的双表差异工作流。没有复制其代码。此实现由 AI 协助编写，验证记录见 VALIDATION.md。欢迎提交可复现的匿名样例及预期结果。

当前为 v0.3.0 初版。财务差异需回查业务原始记录。本包仅做明细核对，不代替审计判断。

## 输出长什么样

| 编号 | 状态 | 示例说明 |
|---|---|---|
| AR003 | 金额差异 | 左表 1,250.10，右表 1,250.00，差额 0.10 |
| AR002 | 重复编号 | 左表出现两笔，全部保留，等待确认核对粒度 |
| AR004 / AR005 | 单边记录 | 需要检查遗漏或截止日期 |

完整模拟产物在 [examples/核对结果](examples/核对结果)。`运行清单.json` 可复算输入输出 SHA-256 并检查每条输入记录是否恰好出现一次。这个检查保证处理完整性，不保证业务内容真实。

## 贡献与隐私

提交问题时写明版本、文件格式、运行参数及预期结果，仅附最小模拟样例。请勿提交真实账单、个人信息或凭据。项目由 AI 协助编写，维护者负责复核；MIT 许可允许使用和修改。
