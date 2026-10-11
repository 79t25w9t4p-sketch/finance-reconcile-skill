# 客户对账 Skill：Agent 接入入口

本文件供用户要求“获取、试用并安装这个 Skill”时使用。代码主仓库：

https://github.com/79t25w9t4p-sketch/finance-reconcile-skill

用户只需发一条请求；Agent 完成下面的获取、检查和试用步骤。普通聊天界面若不能联网、读写工作目录或执行 Python/Node，应直接说明缺少哪项能力，不能把读过本文件说成已安装或运行。

## 1. 获取固定版本的完整运行文件

当前接入版本：v0.7.0，代码提交 `cc26afaa88378f4af3fa50421ff1e1370c45e488`。

优先从 GitHub 获取这个提交的完整仓库到当前任务的新目录。已有目录不覆盖，不读取真实业务文件。保留目录结构，不只下载 `SKILL.md`。

GitHub 获取失败时，允许使用下面列出的 jsDelivr 备用文件源；不自行搜索其他代理站。先读取运行文件清单：

https://cdn.jsdelivr.net/gh/79t25w9t4p-sketch/finance-reconcile-skill@main/agent-manifest.json

清单字段 `revision` 必须等于上面的代码提交；否则先说明入口与清单版本不一致，不拼接新旧版本。清单中的每个 `path` 对应：

```text
https://cdn.jsdelivr.net/gh/79t25w9t4p-sketch/finance-reconcile-skill@cc26afaa88378f4af3fa50421ff1e1370c45e488/{path}
```

路径中的中文和空格按 UTF-8 URL 编码，保留目录分隔符 `/`。下载清单 `files` 内的全部文件，按原目录保存。当前清单包含41个运行和配套说明文件，保留全部脚本与技能说明；`omitted_prebuilt_examples` 列出的6个预生成工作簿不需下载。2026-10-11测试时，这些xlsx的备用请求返回403；两种示例脚本能在本地生成对应工作簿。不要把备用路径称为完整GitHub仓库克隆。

下载完成后，对每个文件核对字节数 `size` 与 `git_blob_sha1`。Git blob 的校验方法是 `SHA1(b"blob " + str(len(data)).encode("ascii") + b"\0" + data)`，不是直接对文件做普通 SHA1。校验失败的文件不能用于运行；报告路径、预期值和实际值，不执行错误页或缺失文件。不接受包含 `..`、绝对路径或反斜杠的清单路径。

这些校验用于确认下载内容与清单一致，不代表脚本安全审计。读取技能说明和实际要运行的脚本后再执行。本入口规定首次接入方式；获取到的历史版本 README 不替代本入口的首次获取步骤。

## 2. 检查现有环境，运行虚构示例

- 需要 Python 3.9+、Node 18+、`openpyxl>=3.1,<4`。先运行 `scripts/doctor.py`；优先复用当前宿主已有运行时。
- 用户要求试用并安装本技能时，允许仅在当前项目独立虚拟环境安装 `requirements.txt` 中的 Python 库。缺 Python 或 Node 时先说明缺项、安装位置和方案，不直接修改系统环境。
- 按 `SKILL.md` 运行 `examples/month_end_demo.py`，结果放在新的工作目录。保留真实运行输出；失败时说明具体原因。
- 把生成的 `复核工作簿.xlsx`、`跟进清单.xlsx` 交给用户，带用户查看 SH-002 的 50.00 金额差异。数据均为虚构；预期共8个编号组，3组金额一致、5组待核实。不能只复述预期值来声称运行成功。

## 3. 试用成功后安装，以后直接调用

用户只要求试用时，完成示例后停止安装步骤。要求“获取、试用并安装”时，示例成功后再安装：

- Codex：用户技能目录 `~/.agents/skills/finance-reconcile/`，里面直接包含 `SKILL.md`、`scripts/` 和 `examples/`，Windows 使用对应用户目录。
- 其他 Agent：先确认其支持的本地 Skill 格式和目录，不能套用 Codex 路径；当前没有逐一验证其他宿主。
- 保留技能所需完整文件；不复制 `.git`、临时虚拟环境、缓存或真实业务文件。已有同名技能先说明，不覆盖其他技能。

报告安装位置、文件获取来源、实际代码提交、运行结果及缺项。安装完成后让用户按宿主的技能加载说明操作。Codex 以后可说：

```text
使用 $finance-reconcile 核对这两份客户明细。先建议共同业务编号和金额列，让我确认客户、期间、币种及金额口径，再输出内部复核表和跟进清单。
```

已安装后不重复下载。确认业务字段和口径是必要输入，不能为了“一步调用”擅自猜测。

## 获取方式的边界

jsDelivr 是 GitHub 文件的 CDN 访问路径，代码仍来自上述公开仓库；它不复制 GitHub 的登录、评论和 Star，也不是远程执行对账的 API。源文件有缓存，入口使用 `main`，运行文件固定到明确提交，避免同次获取混用版本。所有来源都取决于实际执行环境的网络，不承诺国内各运营商均可达。

维护者更改运行代码时，应同时更新本入口中的代码提交和 `agent-manifest.json`，核验全部运行文件并跑通示例后再分享。仅调整推广文案不需要切换运行代码版本。

参考：[jsDelivr 官方 GitHub 文件说明](https://github.com/jsdelivr/jsdelivr#github)、[Codex 官方本地技能目录说明](https://learn.chatgpt.com/docs/build-skills#where-codex-loads-local-skills)。
