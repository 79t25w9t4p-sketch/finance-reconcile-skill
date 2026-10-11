# 客户明细核对 · finance-reconcile

**交两份有共同业务单号的明细，确认对应列和金额口径，得到内部复核表与待核实跟进清单。**

面向已经在用本地 Agent、反复核对固定格式客户明细的财务助理。脚本负责匹配、比金额、筛异常和整理结果；你核实业务事实。换表时可以复用规则和已填写的旧意见。

[![自动测试](https://github.com/79t25w9t4p-sketch/finance-reconcile-skill/actions/workflows/tests.yml/badge.svg)](https://github.com/79t25w9t4p-sketch/finance-reconcile-skill/actions/workflows/tests.yml)

当前 **v0.8.6 可在确认后自动制作核对副本**，处理合计、小计和重复表头，保留原文件与业务来源位置。需要能读写本地文件、执行 Python 的 Agent，以及 Python 3.9+、openpyxl。教程提供 Codex 路径；普通聊天窗口不能直接执行。演示均为虚构数据，尚无真实用户节省时间的测量。

## 先看是否适合你的两份表

- 双方有共同业务编号；同一客户范围、期间、币种、含税口径和明细粒度可以确认。
- 常见场景是客户采购/结算明细与我方销售/结算明细逐单比较。
- 一单多行、无共同编号、分批回款核销，需要先确定关系；本版不自动合并或模糊匹配。差额不直接等于错账。

## 已经有两份表：直接核对

把下面文字中的路径和业务范围换成自己的信息，交给有本地文件和脚本能力的 Agent。不知道的项目写“不知道”，让助手集中说明需要确认什么。

```text
请按这个入口获取并试用客户对账 Skill；已经安装则直接使用：
https://cdn.jsdelivr.net/gh/79t25w9t4p-sketch/finance-reconcile-skill@main/docs/Agent入口-v0.8.6.md

我方文件：「文件路径」；客户文件：「文件路径」。
本次比较「客户、期间、币种」下的「含税或未税金额」。
先检查两表，集中确认缺少或冲突的字段和业务口径；确认后给我复核表与跟进清单，并保存规则。直接处理这两份文件，不先跑模拟示例。
缺少必需Python库时，允许仅在本项目独立环境准备；缺Python先说明方案。
```

获取优先使用本 GitHub 仓库；失败时，入口列出了固定版本的备用文件源和校验要求。能否获取取决于 Agent 所在环境的网络。首次优先复用已有 Python；缺项时说明具体准备方式。试用默认在当前任务目录完成，不安装到用户技能目录、不寻找其他真实业务文件。

## 还没有文件：先看或运行示例

先看[跟进清单](examples/月末客户对账/跟进清单.xlsx)：它列出待查单号、两边金额、来源行和核实建议。这个结果适合你的工作，再让 Agent 试运行：

```text
请按这个入口运行客户对账的虚构示例，交给我复核表和跟进清单，先不安装到用户技能目录：
https://cdn.jsdelivr.net/gh/79t25w9t4p-sketch/finance-reconcile-skill@main/docs/Agent入口-v0.8.6.md
缺少必需Python库时，允许仅在本项目独立环境准备；缺Python先说明方案。
```

示例里我方8行、客户7行，共8个编号组；3组指定金额一致、5组待核实。SH-002差50.00；另外包含单边记录、重复编号和缺失金额。见[两份输入和结果说明](examples/月末客户对账/说明.md)。这些数据用于核验流程，不能证明实际省时。

## 拿到结果后做什么

| 文件 | 先看哪里 | 接下来做什么 |
|---|---|---|
| 复核工作簿.xlsx | 汇总、本次重点 | 点击处理入口，在原记录黄色列填写状态和内部意见 |
| 跟进清单.xlsx | 待核实事项 | 按来源行回查原表，检查后用于业务核实；反馈需整理回内部表 |

默认交付这两份 Excel，助手在对话里简述待核实数量、分类和下一步。文本摘要、CSV、JSON及哈希记录留在结果目录，排查时再看。

结果是生成时快照；手动改状态不会实时刷新汇总和跟进表，重跑后更新。重复、空编号、无效金额先回查原表。本次重点是只读视图；保留原表头及隐藏标识列，便于复用意见。跟进表不含内部意见，但仍有业务编号和金额，分享前检查接收范围。

## 以后再用

确认适合后，可对 Agent 说“把刚才试用成功的完整 Skill 安装到当前宿主的用户技能目录，以后继续使用，不重复跑示例”。Codex 对应 `~/.agents/skills/finance-reconcile/`；已有同名目录先说明，不覆盖其他技能。[首次使用与安装细节](docs/新手第一次使用.md)

下次交新明细和填写过的上次复核表：

> 使用 $finance-reconcile，按上次确认的规则核对这两份新明细，沿用这份上次复核表的意见，结果存到新目录。

未变化的已解决记录沿用；变化项重新复核；旧的未解决事项仍要跟进。本次未出现不等于结案。配置只保存列的对应方式；换客户或编号重新使用时，不能套用旧意见。[换表复用的虚构示例](examples/换表复用/说明.md)

## 当前边界

支持 UTF-8 CSV/TSV 和 XLSX，最多5000条明细、50列。旧xls需另存；公式需在原软件重算并导出值。金额精确到分，重复、空编号和无效金额均保留；不自动填零、合并、跨行求和或解释差异原因。

业务文件在指定本地目录计算，不提交公开仓库。AI宿主如何处理对话数据由其决定；默认只向模型提供摘要，具体明细按问题查看。[技能规则](SKILL.md) · [验证记录](VALIDATION.md)

<details>
<summary>脚本与环境排查</summary>

```sh
python3 scripts/doctor.py
python3 scripts/suggest.py --left 我方.xlsx --right 客户.xlsx
python3 examples/month_end_demo.py 新的演示目录
python3 -m unittest discover -s tests -p 'test_*.py'
```

脚本参数和复用方式见 [SKILL.md](SKILL.md)。首次检查XLSX各页前20行：唯一可见编号与金额表头候选才建议使用，多候选或隐藏页请确认；自定义列名或更晚的表头可显式指定。疑似合计/小计及重复表头先确认，再由固定脚本制作仅含所选明细值的新副本，并记录原文件哈希、排除理由和位置；原文件不改。底层不自动删除输入行。纯空尾列可忽略，有数据的无名列仍报错。两份结果表列出用于核对的文件、工作表及表头行，结合来源行回查。缺Python库时，优先在独立环境安装requirements.txt，安装和执行使用同一个解释器；缺Python时先说明，不直接修改系统环境。旧Node入口改为 `python scripts/reconcile.py`。

</details>

## 反馈

试用后最有用的反馈是：**原来怎么做、哪里卡住、拿到结果后还得手动补什么、下次愿不愿意继续用**。[反馈说明](docs/试用反馈.md) · [提交反馈](https://github.com/79t25w9t4p-sketch/finance-reconcile-skill/issues/new?template=first-use.md)

先验证现有流程是否值得用，再按实际阻碍改进。[当前改良原则](docs/定位与改良方案.md) · [分享草稿](docs/分享文案.md) · [版本记录](CHANGELOG.md)

由AI协助实现，参考 [finance-copilot-skills](https://github.com/cxtx/finance-copilot-skills)、[cpa-skills](https://github.com/adoptai/cpa-skills)、[csv-diff](https://github.com/simonw/csv-diff) 的工作流思路，没有复制其代码。MIT许可。
