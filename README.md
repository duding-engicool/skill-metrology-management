# 计量管理（metrology-management）

面向计量管理员、实验室主管的计量器具台账与校准周期管理技能。自动计算下次校准日期、三级超期预警，生成 **MD 管理台账（默认）**，可按需导出 **Word / Excel** 版本。

## 适用角色
- 计量管理员
- 实验室主管 / 质检负责人

## 核心能力
- 维护计量器具台账（新增 / 更新 / 查询）。
- 依据"上次校准日期 + 校准周期"自动计算下次校准日期。
- 三级预警：红色超期 / 橙色 7 天内 / 黄色 30 天内 / 正常。
- 生成送检计划与合规状态报告。
- 提示 MSA（测量系统分析）关联待办。

## 使用流程
1. 选择动作：查询到期 / 单台查询 / 合规检查 / 送检计划 / 新增器具。
2. 提供器具信息（名称、编号、上次校准日、周期；缺失按类别默认周期）。
3. 技能计算下次校准日期并判定预警状态。
4. 调用 `scripts/build_report.py` 渲染 **MD 台账**（默认）。
5. 生成后询问用户是否需要 **Word / Excel**，确认后导出。

## 脚本说明（scripts/build_report.py）
- `python scripts/build_report.py` —— 使用内置 5 台样本，覆盖四态，默认产出 MD。
- `python scripts/build_report.py --format xlsx` —— 导出 Excel（状态列带颜色）。
- `python scripts/build_report.py --format docx`  —— 导出 Word。
- `python scripts/build_report.py --format all`   —— MD + DOCX + XLSX 全导出。
- `python scripts/build_report.py --data-file instruments.json --out-dir ./out` —— 自定义台账（默认输出到当前工作目录）。
- `python scripts/build_report.py --as-of 2026-07-13` —— 指定基准日期（默认今天）。
- 输出：`metrology_report.md`（默认），可选 `metrology_report.docx` / `metrology_report.xlsx`（主色 #C8102E）。

## 能力边界
- 不替代计量机构出具证书；不判定器具合格与否（以证书为准）。
- 强检设备超期直接提示"判定不合格"。
- 周期为通用建议，以企业制度与实际证书为准（标注 待企业补充）。

## 联动技能
- 表单生成（量具校准有效期核查）
- 不合格处理（超期量具数据追溯）
- 内审不符合项（条款 7.1.5）
- 体系文件类技能

## TRACE
五维各 ≥8，总分 45/50（详见 SKILL.md ⑩）。
