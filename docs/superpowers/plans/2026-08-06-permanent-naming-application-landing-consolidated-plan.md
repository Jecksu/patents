# 永久命名申请文件落稿汇总强化实现计划

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推荐）或 superpowers:executing-plans 逐任务实现此计划。步骤使用复选框（`- [ ]`）语法来跟踪进度。任务之间存在强顺序依赖（任务 4 起的正文内容由任务 1 至 3 的权利要求确定，任务 8 的产物由任务 6、7 的脚本确定），不得乱序执行。

**目标：** 以专利通过率最大化为目的，把已完成的永久命名强化设计（1240 行）和锚点现有技术核查（172 行）的全部结论，实际落到会被提交和审查的四份产物上——申请正文 Markdown、申请文件 DOCX 生成器、技术交底书生成器、说明书附图——使独立权利要求恰为七步、八项绑定闭环逐步承载、六类已知手段全部降级、指标与数值实施例可复算、算法特征与 CAD 数据处理绑定，并消除现存的旧架构表述、被禁公式和生成器结构缺陷。

**架构：** 申请正文 Markdown 是 DOCX 中摘要、权利要求书和说明书的唯一内容来源（`build_material()` 用 `section_between()` 抽取三段），因此必须先改正文；附图标签和技术交底书正文各自硬编码在两份脚本中，不继承 Markdown 改写，必须并行改脚本；最后统一生成 PNG、DOCX、PDF 并逐页复核。全程以强化设计稿第 15、16、17 节为落稿规范，以锚点现有技术补充稿第 5 节为权利要求地位约束。

**技术栈：** 中文 Markdown、Python 3.12（`python-docx`、`Pillow`）、Word COM 或等效工具做 DOCX→PDF、PyMuPDF 或等效工具做 PDF 分页栅格化、`rg` 文本核查、PowerShell 数值复算。

---

## 输入依据与本计划定位

本计划汇总以下三份已完成材料，不重复其结论，只负责把结论落到正式产物：

1. `docs/superpowers/plans/2026-08-06-permanent-naming-metrics-and-model-diff-strengthening.md`（218 行，7 任务 32 步全部完成）——上一轮的执行记录与验证方法论；其任务 7 的 `rg` 扫描、数值复算、现有技术边界核对和 Git 范围检查是本轮可复用的验证模式。
2. `docs/superpowers/specs/2026-08-06-permanent-naming-metrics-and-model-diff-strengthening-design.md`（1240 行，17 节）——本轮唯一技术内容来源。第 4 至 11 节为定义与流程，第 12 节为可复算数值实施例，第 13 节为异常场景，第 14 节为证据记录与技术效果限定，第 15 节为落稿映射，第 16 节为 14 项交付标准，第 17 节为 13 项最终确认清单。
3. `docs/submission/2026-08-06-permanent-naming-anchor-prior-art-supplement.md`（172 行）——权利要求地位约束来源。第 5.1 节六项不得作独权核心，第 5.2 节八项绑定闭环，第 5.3 节六条技术效果因果链，第 6 节剩余风险与复检要求。

**本轮必须关闭的缺口**（强化设计稿第 1 节第 11 行原文）：“本文档是申请文件和技术交底书的前置设计审查稿。本轮不直接修改正式申请正文、权利要求、技术交底书生成脚本或附图。”

因此三份输入的强化成果目前**一行也没有进入会被提交的文件**。会被提交的四份产物仍是旧架构或含被禁内容：

| 产物 | 当前状态 | 具体问题 |
|---|---|---|
| `docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md` | 完全旧架构 | 摘要仍写“基于永久命名标识对两个CAD模型的实体集合执行集合运算，以确定新增实体、删除实体和共有实体；对集合运算结果和共有实体执行命名一致性审计”，即被禁的“先集合差、后审计”顺序；权利要求 1 为旧 11 步；权利要求 5 直接由单侧键存在导出新增／删除／共有；无关系基数、`R_seed`、三类假设、联合选择、`FROZEN`、双侧分区、`COMPLETE_FAIL`、`CLOSURE_PASS`、类型化事件；附图说明声明 5 幅且主题与生成器不一致；有益效果与复杂度分析含“降低计算复杂度”“显著降低计算量”等无支持性能断言 |
| `scripts/generate_technical_disclosure_docs.py` 的 `PERMANENT` 字典 | 含全部六类被禁内容 | `C_A=|K_A|/|U_A|`、`C_pair=2|K_A∩K_B|/(|K_A|+|K_B|)`、`Q_cons=Σ(q_k·1[q_k≥τ_cons])/|K_A∩K_B|`、`C_map`、`D=Σw_i d_i`、`d_num(a,b)=min(|a-b|/max(|a|,|b|,ε),1)`、“若命名键仅存在于第二模型，则先标记为新增候选”、无冻结集合、数值例用“面、边和特征层级”且给出 `C_pair≈93.88%`／`Q_cons≈96.74%`／`τ_pair=90%`（与设计稿第 12 节的 `l=face`、`|K_11|=90`、`87/90`、`84.78/90`、`C_cross≈3.33%` 全部冲突） |
| `scripts/generate_patent_submission_docs.py` 的 `permanent_naming` 条目与 `make_metrics_diff_flow()` | 旧标签＋被禁符号 | `abstract_figure`／`system_modules`／`decision_flow` 全部是“执行命名集合运算”“命名集合比较模块”“异常闭环／指纹筛选”等旧主线；图 4 渲染出设计稿第 15.8 节明令不得使用的 `Q_cons` 与隐含 `C_map` 的综合冲突表述，以及 `综合差异 D = Σ w_i d_i`、“异常候选：一对一／一对多／多对一”并列而无联合选择与冻结 |
| `docs/submission/assets/permanent_naming-fig*.png`、DOCX、PDF | 与旧正文一致 | 仅有 fig1／fig2／fig3，fig4 尚未生成；`rendered/` 下各目录均为旧版本 |

另有两项结构性缺陷必须在本轮修掉，否则任务 8 无法可靠执行：

- `generate_technical_disclosure_docs.py` 中 `TEMPLATE` 与 `PERMANENT["output"]` 是同一路径（均为 `docs/submission/技术交底书/永久命名-CAD模型差异对比-专利技术交底书-权利要求增强版.docx`）。脚本读取自身产物作为模板并覆盖写回；轻量化交底书也以永久命名文件为模板生成。虽然 `clear_cell()` 使行内容不累积、单次运行结果尚可用，但模板与产物不可分离、每次运行都改写被 Git 跟踪的模板文件，属结构性缺陷。
- 本机 `python-docx` 在 3.13 与 3.12 均缺失，`Pillow` 仅 3.12 有，且无 `soffice`／`libreoffice`／`pdftoppm`。两份脚本当前**不可运行**，必须先补齐环境（任务 0）。

**授权风险到任务的对应关系：**

| 授权风险 | 对策 | 落在 |
|---|---|---|
| 独权被 `US10816957B2`＋`US11288411B2` 组合评价为显而易见 | 独权恰为七步并逐步承载八项闭环；六类已知手段全部降为从属或实施方式 | 任务 1、2 |
| 算法特征被认定为智力活动规则或不具技术性 | 七步每一步均绑定具体 CAD 数据对象或数据读取动作 | 任务 1 步骤 3 |
| 权利要求得不到说明书支持、公开不充分 | 第 4 至 11 节完整定义与第 12 节可复算实施例进入具体实施方式 | 任务 4、5 |
| 清楚性缺陷（符号未定义、口径混淆、空分母） | 先定义后使用扫描；一致性通过率与平均得分分离；空分母记不可计算 | 任务 5、9 |
| 形式缺陷通知（摘要超字数、附图与正文不符、引用关系错误） | 摘要 300 字上限核查；附图说明与生成器 4 幅严格对齐；从属项引用关系核查 | 任务 6、9 |
| 无支持的性能断言被要求删除并削弱有益效果 | 有益效果严格按第 14 节限定表述；复杂度分析改为算法上界 | 任务 4 步骤 4 |

---

## 文件结构

- 修改：`docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md` — 正式申请正文，DOCX 中摘要／权利要求书／说明书的唯一来源，本轮全文重写（任务 1 至 6）。
- 修改：`scripts/generate_technical_disclosure_docs.py` — 重写 `PERMANENT` 字典，分离模板路径并加自覆盖断言（任务 7）。
- 修改：`scripts/generate_patent_submission_docs.py` — 重写 `MATERIALS` 中 `permanent_naming` 条目的图形标签，并重写 `make_metrics_diff_flow()`（任务 6）。
- 修改：`docs/superpowers/specs/2026-08-06-permanent-naming-metrics-and-model-diff-strengthening-design.md` — 仅第 17 节逐项打勾并附一行核对结果，不改动第 1 至 16 节任何技术定义、公式或数值。
- 新增：`docs/submission/技术交底书/_模板/专利技术交底书-结构模板.docx` — 从当前被跟踪副本复制出的独立结构模板。
- 生成：`docs/submission/assets/permanent_naming-fig1..fig4-*.png`、`docs/submission/永久命名-CAD模型差异对比-发明专利申请文件.docx`、`docs/submission/技术交底书/永久命名-CAD模型差异对比-专利技术交底书-权利要求增强版.docx`、`docs/submission/rendered/permanent-naming-consolidated/`（PDF 与分页 PNG）。
- 参考：`docs/superpowers/plans/2026-08-06-permanent-naming-metrics-and-model-diff-strengthening.md`、`docs/submission/2026-08-06-permanent-naming-anchor-prior-art-supplement.md`、`docs/submission/指定8件现有专利-独立权利要求级核查报告.md`。
- 不修改：`docs/patent/2026-07-07-cad-model-diff-lightweight-summary-cross-kernel-revised-patent-draft.md`、`generate_technical_disclosure_docs.py` 的 `LIGHTWEIGHT` 字典、`generate_patent_submission_docs.py` 的 `MATERIALS[0]`（`lightweight`）条目、`docs/superpowers/plans/2026-08-06-cross-kernel-lightweight-summary-strengthening.md`、`docs/superpowers/specs/2026-08-06-cross-kernel-lightweight-summary-strengthening-design.md` — 跨 CAD 内核轻量摘要案归属 `codex/lightweight-summary-strengthening` 分支，本轮只保持设计稿第 3.4 节边界，不改其任何内容。

---

## 全局硬约束

每个任务的每一步都必须同时满足以下约束，任何一步违反即视为该步未完成。

- **H1 七步唯一主线。** 方法独立权利要求恰为设计稿第 15.2 节七步，顺序与前后依赖不变。摘要、系统／装置权利要求、介质权利要求、说明书技术方案、具体实施方式、技术交底书和图 4 必须复用同一条依赖链，不得另造旁路功能。
- **H2 闭包前禁止确定性增删。** 任何产物中都不得出现“单侧命名键即判为新增／删除”“先集合差、后审计”“先标记为新增候选”。新增／删除必须由同代次 `COMPLETE_FAIL_A/B` 端点级失败凭证支持，并在 `CLOSURE_PASS` 后物化。
- **H3 被禁符号清零。** 全部产物（含附图标签）不得出现 `Q_cons`、`C_map`、`C_pair`、`C_intra`、`d_num(a,b)=min(...)`、`D=Σw_i d_i`、`τ_pair`。一致性用 `R_cons`（通过率，分母 `K_11`）与 `S_cons`（平均得分，分母 `K_score`）分开表述，并另报 `C_hard`、`C_lowinfo`。
- **H4 六类已知手段降级。** 锚点传播、邻接签名传播、普通候选排序／最大权匹配、版本指纹快速跳过、摘要压缩、局部 B-rep 验证只出现在从属权利要求或具体实施方式，且每项都写明其受七步主线约束的条件（设计稿第 15.4 节）。普通一对一不得先行占用可能属于 `R_1m/R_m1` 的端点。
- **H5 效果表述限定。** 有益效果只能表述为：降低错误身份传播风险、给出候选规模上界、避免把证据不足端点提前归入 `ADD/DEL`、保持结论可复算可追溯（设计稿第 14 节末段）。不得出现“显著”“大幅”“提高效率”“降低计算复杂度”等无对照试验支持的性能断言，不得给出任何真实性能百分比。
- **H6 风险陈述不入申请正文。** 检索范围限定、组合风险、“不构成授权保证”“官方数据库复检”等表述只存在于设计稿、锚点核查稿和技术交底书的内部备注中，不写入申请正文的摘要、权利要求书或说明书。
- **H7 反引号与 Markdown 表格禁用。** 申请正文 Markdown 中一律不使用反引号包裹符号，也不使用 Markdown 表格。原因：`add_md_content()` 无反引号与表格处理，会把 `` ` `` 和 `|` 原样写入 DOCX。符号写作纯文本（如 `R_11` 写成 R_11，`FROZEN_A` 写成 FROZEN_A），表格内容改写为编号段落。
- **H8 不生成永久命名。** 保留负面边界：本方案读取比较开始前已经存在且可读取的永久命名关系，比较规范键仅用于本次比较且不回写 CAD、CAD 内核或 PDM，本方案不生成、修复或更新永久命名规则。

---

### 任务 0：环境与备份前置

**文件：**
- 新增：`tmp/backup-before-consolidated-landing/`（备份目录，不纳入提交）

- [ ] **步骤 1：备份两份用户已有未提交脚本改动**

工作区中 `scripts/generate_patent_submission_docs.py` 与 `scripts/generate_technical_disclosure_docs.py` 是用户已有的 ` M` 未暂存改动。上一轮计划的“不修改”约束仅适用于该轮，本轮明确将两文件列为修改目标；但用户改动不得被静默丢弃或整体回退。

运行：

```bash
mkdir -p tmp/backup-before-consolidated-landing && git diff -- scripts/ > tmp/backup-before-consolidated-landing/scripts-user-uncommitted.diff && cp scripts/generate_patent_submission_docs.py scripts/generate_technical_disclosure_docs.py tmp/backup-before-consolidated-landing/
```

预期：备份目录含一份 diff 和两份脚本副本。全程禁止对这两个文件执行 `git checkout --`、`git restore` 或 `git stash`。

- [ ] **步骤 2：确认用户改动中可保留与必须替换的部分**

可保留：`generate_patent_submission_docs.py` 中新增第 4 幅附图的机制（`extra_figure` 键、`figure_titles` 第 4 项、`build_material()` 中的 `extra_fig` 分支、`make_metrics_diff_flow()` 的函数签名与双列渲染骨架）；`generate_technical_disclosure_docs.py` 中把 `TEMPLATE` 从外部绝对路径改为仓库内路径这一方向性修正。

关于该 `TEMPLATE` 改动的事实认定（已核实）：原路径为企业微信聊天缓存中的空白模板 `C:/Users/xiaodi.su/Documents/WXWorkLocalPro_data/Profiles/32A711E5667E8FD210E3A23C66FA1C18/Cache/chat/file/202607/专利技术交底书模板_V2_副本.docx`；该文件在本机不存在（当前账户为 `sxd`，非 `xiaodi.su`），全盘搜索也未找到同名模板。用户因此把 `TEMPLATE` 改指向仓库内已被跟踪的生成产物 `docs/submission/技术交底书/永久命名-CAD模型差异对比-专利技术交底书-权利要求增强版.docx`，属于在缺失原模板情况下的可用性恢复，方向正确；但由此产生 `TEMPLATE` 与 `PERMANENT["output"]` 同一路径的自读自写，需在任务 7 步骤 1 分离。

必须替换：`make_metrics_diff_flow()` 的全部标签文本、`permanent_naming` 的三份标签列表、`PERMANENT` 字典的 `background`／`solution`／`advantages`／`key_points`／`figures`、`TEMPLATE` 与产物同路径的设置。

在计划执行记录中写明该判定，再进入后续任务。

- [ ] **步骤 3：补齐运行环境**

本机 Python 3.13 为默认解释器但缺 `python-docx` 与 `Pillow`；Python 3.12 有 `Pillow` 缺 `python-docx`；无 `soffice`、`libreoffice`、`pdftoppm`；Word 位于 `C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE`。

运行：

```bash
py -3.12 -m pip install python-docx PyMuPDF
```

随后运行：

```bash
py -3.12 -c "import docx, PIL, fitz; print('docx', docx.__version__ if hasattr(docx,'__version__') else 'ok'); print('PIL ok'); print('fitz', fitz.__doc__.strip().splitlines()[0])"
```

预期：三个模块均导入成功。此后任务 8 的全部脚本调用统一使用 `py -3.12`。若 `pip install` 因网络受限失败，停止并向用户报告，不得跳过任务 8 而声称已交付。

---

### 任务 1：重写摘要与方法独立权利要求

**文件：**
- 修改：`docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md` 第 3 至 20 行（摘要与权利要求 1）

- [ ] **步骤 1：按七项主链重写摘要**

依设计稿第 15.1 节的七项主链改写，依次覆盖：读取既有永久命名索引、实体—比较规范键多重关联关系与既有键对应关系；分离硬身份约束、软身份证据与允许变化属性并计算质量状态、执行全局与逐键门控；形成仍待身份／指纹验证的临时种子关系并生成稀疏的一对一、一对多、多对一关系假设；同一局部分量联合双侧端点互斥选择、不能唯一消歧者转入冻结；双侧分区不变量检查与失败回滚扩读；仅在闭包通过后分类；输出统一类型化差异事件与证据、代次和闭包快照。

摘要不展开公式，不把普通匹配技巧写成发明支点，不出现“集合运算”“先集合差”“集合补集”。

- [ ] **步骤 2：控制摘要字数不超过 300 字**

《专利法实施细则》要求摘要文字部分不超过 300 个字。七项主链须压缩到该上限内，优先保留“既有关系基数—质量门控—三类关系联合互斥选择—未决冻结—双侧分区闭包—闭包后分类—类型化事件输出”这条骨架，删去可放入说明书的修饰。

运行：

```bash
py -3.12 -c "import re,pathlib; t=pathlib.Path('docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md').read_text(encoding='utf-8'); m=re.search(r'## 摘要\n(.*?)\n## 权利要求书', t, re.S); s=re.sub(r'\s+','',m.group(1)); print(len(s))"
```

预期：输出的字符数不超过 300。

- [ ] **步骤 3：把权利要求 1 重写为恰好七步并绑定 CAD 数据处理**

逐字对照设计稿第 15.2 节的七步撰写，保持“一种CAD模型差异对比方法，其特征在于，包括：”的前序＋特征部分形式，七步分别承载：第 1 步关系基数保真；第 2 步身份／变化分离与质量状态门控；第 3、4 步三类关系联合选择及未决冻结；第 5 步双侧分区检查与失败恢复；第 6 步闭包后分类；第 7 步可审计类型化输出。不得多于或少于七步，不得把任一项拆成与其余步骤无关的普通匹配结果。

同时满足算法特征的技术性要求：七步中每一步都必须至少出现一个具体 CAD 数据对象或数据读取动作，例如永久命名实体索引、B-rep 面／边／顶点、拓扑邻接关系、特征生成关系、装配约束、局部 B-rep 读取范围、PDM 映射表。禁止出现任何一步只描述纯数学运算或纯规则判断。

运行：

```bash
py -3.12 -c "
import re,pathlib
t=pathlib.Path('docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md').read_text(encoding='utf-8')
c1=re.search(r'\n1\. 一种CAD模型差异对比方法(.*?)\n2\. ', t, re.S).group(1)
steps=re.findall(r'^\s*S?(\d)[\.、）\)]', c1, re.M)
print('步数标记:', steps)
kw=['永久命名实体索引','B-rep','拓扑邻接','特征','装配','局部','读取']
for i,seg in enumerate(re.split(r'(?m)^\s*(?=[1-7][\.、）\)])', c1)[1:], 1):
    print(i, [k for k in kw if k in seg] or '缺CAD数据对象')
"
```

预期：步数标记恰为 1 至 7；每一步都至少命中一个 CAD 数据对象关键词，无“缺CAD数据对象”。

- [ ] **步骤 4：核对八项绑定闭环逐步落位**

按锚点核查稿第 5.2 节八项逐项在权利要求 1 中定位：关系基数保真、身份／变化分离、状态门控、三类关系联合互斥选择、未决冻结、双侧分区检查、失败回滚或局部扩读、闭包后判增删。八项必须都在权利要求 1 正文中，不得任何一项只出现在说明书。在执行记录中写明每项对应第几步。

- [ ] **步骤 5：删除旧摘要与旧权利要求 1 的全部残留**

运行：

```bash
rg -n "集合运算|集合补集|先集合差|命名集合" docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md
```

预期：在摘要与权利要求 1 范围内零命中。全文其他位置若仍有命中，须在任务 2 至 5 中一并消除。

---

### 任务 2：从属权利要求十簇重排与已知手段降级

**文件：**
- 修改：`docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md` 权利要求 2 至 29（旧编号）

- [ ] **步骤 1：按十个簇建立从属项骨架**

依设计稿第 15.3 节，按从靠近独权核心到实现与计量的顺序建立十簇，每簇 2 至 4 项：关系基数与三重唯一簇；质量公式与空分母簇；三态门控簇；证据与变化分离簇；三类假设与稀疏预算簇；联合互斥与冻结不动点簇；失败回流与代次簇；端点展平与双侧分区簇；统一事件与证据缓存簇；候选压缩与局部读取计量簇。

独权不得硬塞公式、求解器、摘要种类或具体 B-rep 特征；公式、空分母规则、预算参数、求解器选择全部下放到对应簇。

- [ ] **步骤 2：把六类已知手段写成受约束的从属项或实施方式**

逐项按设计稿第 15.4 节的限定条件撰写，不得并列成创造性支点：

锚点传播——已验证关系只作可选候选来源或验证上下文，仍受三类关系联合选择、冻结和闭包约束。邻接签名传播——只生成或补强候选证据，不绕过局部门控与双侧分区检查。普通候选排序／最大权匹配——最大权、稳定选择、整数优化或受约束搜索只是联合目标的可选求解器，不得先做普通一对一匹配并占用可能属于 `R_1m/R_m1` 的端点。版本指纹快速跳过——复用条件限于模型版本、直接依赖摘要和当前代次键均有效，且只在一对一身份／指纹验证已通过、关系闭包成立、内容指纹可验证覆盖当前层级全部变化维度及当前函数版本时用于免去重复读取；不得以指纹相同直接判定保持不变。摘要压缩——属证据表示或存储实施方式，保持与跨 CAD 内核轻量摘要案的边界，不替代既有永久命名关系入口。局部 B-rep 验证——只作身份／关系证据取得方式，受预定读取预算、失败回流、代次失效和冻结收敛约束。

- [ ] **步骤 3：处置旧从属项**

必须删除：旧权利要求 5（“所述集合运算包括”，直接由单侧键存在导出新增／删除／共有，违反 H2）。

必须改写后归簇：旧 12（共有实体摘要或版本指纹比较）→ 版本指纹从属项，加上步骤 2 的复用条件；旧 13（候选缩减＋局部验证）→ 候选压缩与局部读取计量簇；旧 17 至 20（审计规则、审计评分、双阈值分流、异常队列记录）→ 三态门控簇与失败回流簇，其中“单侧缺失命名实体但另一模型存在候选实体”一项必须改写为形成待审计残余端点，不得表述为确定增删；旧 21、22（摘要辅助匹配与替代候选生成）→ 三类假设与稀疏预算簇，须补入双向索引查询与预算溢出冻结；旧 24、25（跨系统映射）→ 质量公式与空分母簇的映射覆盖、边冲突、端点冲突口径分离；旧 26（差异证据记录）与旧 28、29（证据记录与增量缓存）→ 统一事件与证据缓存簇，须补入非空代次键、闭包快照与同批原子发布；旧 27（差异结果字段）→ 统一事件簇的事件字段限定。

可基本保留（仅按新术语调整措辞）：旧 2（实体种类）、旧 3（实体属性）、旧 7 至 11（各类摘要与版本指纹构成）、旧 14（实体依赖关系）、旧 15、16（语义聚合与影响传播）。

必须保留并加强负面边界：旧 6 与旧 23 合并为一项从属项，明确本方法不生成、修复或更新 CAD 系统、CAD 内核或产品数据管理系统中的永久命名规则，比较规范键仅由已有名称标准化得到、仅用于本次比较且不回写。该负面限定必须在说明书具体实施方式中有对应支持段落（任务 5 步骤 5）。

- [ ] **步骤 4：连续重编号并修正全部引用关系**

从权利要求 2 起连续编号至方法从属项结束，逐项检查引用：只能引用编号更小的权利要求；多项引用只能采用“根据权利要求 X 或 Y 所述的方法”的择一形式；多项引用的从属权利要求不得作为另一个多项引用从属权利要求的引用基础。

运行：

```bash
py -3.12 -c "
import re,pathlib
t=pathlib.Path('docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md').read_text(encoding='utf-8')
cl=re.search(r'## 权利要求书\n(.*?)\n## 说明书', t, re.S).group(1)
items=re.findall(r'(?m)^(\d+)\. (.{0,60})', cl)
nums=[int(n) for n,_ in items]
print('编号连续:', nums==list(range(1,len(nums)+1)), '总项数:', len(nums))
for n,head in items:
    for r in re.findall(r'权利要求(\d+)', head):
        if int(r)>=int(n): print('前向或自引用:', n, r)
"
```

预期：编号连续为真；无前向或自引用输出。

- [ ] **步骤 5：核对从属项未复活被禁架构**

运行：

```bash
rg -n "集合运算|集合补集|先标记为新增|Q_cons|C_map|C_pair|C_intra" docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md
```

预期：权利要求书范围内零命中。

---

### 任务 3：系统、装置与介质权利要求同步

**文件：**
- 修改：`docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md` 旧权利要求 30 至 35

- [ ] **步骤 1：按七步重划系统权利要求模块**

依设计稿第 15.5 节，系统独立权利要求的模块只能是七步功能的对应划分：既有命名关系读取模块、质量状态与门控模块、种子验证与假设生成模块、联合选择与冻结模块、分区检查与失败回流模块、闭包后差异分类模块、类型化事件输出模块。存储器保存既有永久命名关系、候选配置、证据和代次，处理器被配置为执行所述七步。

删除旧模块中的“命名集合比较模块”“指纹比较模块”“命名异常闭环处理模块”等隐含旧主线的划分，并确保不引入“自动生成永久名称”或“闭包前判定新增／删除”等正文不存在的功能。

- [ ] **步骤 2：同步系统从属项**

旧 31（映射模块）、旧 32（异常识别与辅助匹配）、旧 33（增量缓存模块）改写为新模块划分下的从属项，其中异常识别项须明确异常端点进入三类关系假设与联合选择，不直接产出最终分类。

- [ ] **步骤 3：更新电子设备与计算机可读存储介质权利要求**

两项独立权利要求引用的方法权利要求编号范围必须改为任务 2 重编号后的实际方法项区间（旧文为“权利要求1至29任一项”，重编号后须实测更新）。介质权利要求须写明其程序被处理器执行时使系统完成同一七步及同代次发布约束，不另造独立于 `R_seed/R_11`、`FROZEN`、`CLOSURE_PASS` 或 `CAD_DIFF_EVENT_V1` 的处理主线。

- [ ] **步骤 4：核对单一性与类别一致**

发明名称、摘要、方法独权、系统独权、设备与介质权利要求必须指向同一发明构思与同一条七步依赖链。在执行记录中写明四类权利要求的编号与各自映射的七步位置。

---

### 任务 4：重写说明书正文

**文件：**
- 修改：`docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md` 技术领域、背景技术、发明内容（要解决的技术问题、技术方案、有益效果）

- [ ] **步骤 1：改写背景技术**

保留技术领域。背景技术改为陈述客观技术状况与问题：同一命名域或同一 PDM 链路中的 CAD 实体通常具有比较开始前已经存在的永久命名；实际编辑中会出现命名重复、命名漂移、命名失效、命名分裂、命名合并、低信息命名和局部数据读取失败；已有做法多关注永久命名的生成与维持，或在两侧元素已经直接对应的前提下比较并列出差异，或以高置信匹配辅助低置信匹配并在匹配完成后把剩余实体判为新增或删除。

背景技术可依《专利法实施细则》引证已知文献，但**不得**写入检索范围限定、组合创造性风险评价或“不构成授权保证”等内部风险表述（H6）。这些只留在设计稿与锚点核查稿中。

- [ ] **步骤 2：改写要解决的技术问题**

删除旧第 1 项“计算复杂度高”式表述。技术问题统一为：在永久命名部分可靠且存在重复、漂移、分裂、合并、低信息或读取失败的情况下，如何保持名称—实体映射基数、分离身份判定与变化判定、冻结未决端点，并在一对一、一对多、多对一关系互斥闭包成立后才输出差异，以控制错误身份传播与增删误报。

- [ ] **步骤 3：改写技术方案概述**

按七步顺序概述，逐步点明所承载的闭环项，并说明七步把八项绑定闭环组合在一条 CAD 数据处理链上。此段是权利要求 1 的支持基础，用语必须与权利要求 1 一致，不得引入权利要求中不存在的步骤或功能。

- [ ] **步骤 4：改写有益效果与复杂度分析**

有益效果只保留设计稿第 14 节末段允许的四类表述，并按锚点核查稿第 5.3 节的六条因果链逐条给出“技术特征→作用机理→技术效果”的因果说明：映射基数保真防止普通名称去重丢失重复名称信息；身份与变化两阶段判断防止合法尺寸变化被误判为名称漂移或复用；组关系与一对一关系联合选择防止分裂／合并端点被普通一对一匹配提前占用；未决冻结防止读取失败或候选竞争端点被错误计为新增／删除；分区不变量驱动关系回滚或局部扩读，使其成为控制算法而非结果展示公式；完成闭包后再分类使双侧实体结果可复算且不重复计数。

复杂度分析改写为算法上界与计量口径：给出稀疏候选规模上界 `|C| <= min(|U_A^abn|*|U_B^abn|, |U_A^abn|*B_edge,A→B^l + |U_B^abn|*B_edge,B→A^l)`（写入正文时按 H7 去掉反引号），以及候选边比例、局部验证实体比例、局部 B-rep 字节比例三个指标各自的分子分母与空分母规则。删除全部“显著降低计算量”“降低计算复杂度”“提高效率”表述。

运行：

```bash
rg -n "显著|大幅|明显提高|提高效率|降低计算复杂度|降低计算量" docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md
```

预期：零命中。

---

### 任务 5：重写具体实施方式与可复算数值实施例

**文件：**
- 修改：`docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md` 具体实施方式

- [ ] **步骤 1：写入第 4 至 11 节完整算法定义**

按设计稿第 15.6 节要求，把第 4 至 11 节的定义与流程完整写入具体实施方式：比较层级与可审计实体；比较规范键与实体—键关联关系及既有键对应关系；实体权重；覆盖率（`C_cov,X`、`C_unique,X`、种子端点覆盖率）；一致性（`K_11` 三重唯一、原始字段的硬约束／软证据／允许变化属性／组聚合不变量四类互斥用途、`K_score`／`K_hard`／`K_lowinfo` 划分、`q_κ`、`R_cons`、`S_cons`、`C_hard`、`C_lowinfo`）；冲突率（`C_dup,X`、`C_cross`、`C_mapcov,X`、`C_mapedgeconf,X`、`C_mapconf,X`、`C_conf`）；门控（元数据硬门控、全局质量门控三态路由、逐键可信／歧义／冲突／低信息、`pass_map_local(κ)` 与 `pass_map_dir(X)`、空分母与不可计算规则）；S100 至 S1000 全流程；变化计算（容差归一化数值差异、集合与多重集差异、类别／方向／关系差异、变化证据覆盖门控 `J_change^l`／`coverage_change`／`tau_change^l`／`tau_same^l`、可复算名称变化状态）；最终集合（类型化关系记录与 `end_A/end_B`、严格共有、可追踪范围、`BLOCKED_A/B` 与 `COMPLETE_FAIL_A/B`、`DEL`、`ADD`、双侧分区不变量与机器断言及失败动作、统一类型化事件字段）。

写入时按 H7 去掉全部反引号，符号写作纯文本。

- [ ] **步骤 2：写入第 12 节可复算数值实施例**

数值必须与设计稿第 12 节逐项一致，层级固定为面层级，不混入边、特征、体或装配：`|U_A^face|=100`、`|U_B^face|=105`、`|N_A|=96`、`|N_B|=100`、`|G_A|=94`、`|G_B|=97`、`|K_11|=90`；87 个键对 `q_κ=0.96`、3 个键对 `q_κ=0.42`，`K_score=K_11`、`K_hard=K_lowinfo=∅`；A 侧 2 个、B 侧 3 个已命名面受模型内重复键影响。

指标：`C_cov,A=96/100=96%`、`C_cov,B=100/105≈95.24%`、`C_unique,A=94/100=94%`、`C_unique,B=97/105≈92.38%`、`R_cons=87/90≈96.67%`、`S_cons=(87×0.96+3×0.42)/90=84.78/90=94.20%`、`C_hard=0/90=0%`、`C_lowinfo=0/90=0%`、`C_dup,A=2/96≈2.08%`、`C_dup,B=3/100=3%`、`C_cross=3/90≈3.33%`；映射类指标记为不适用而非 0。

事前固定阈值：`tau_cons=0.90`、`tau_conf=0.50`、`tau_cov,A=tau_cov,B=90%`、`tau_unique,A=tau_unique,B=85%`、`tau_rate=95%`、`tau_score=90%`、`tau_hard=0%`、`tau_lowinfo=0%`、`tau_dup,A=tau_dup,B=5%`、`tau_cross=5%`。

候选压缩：`|U_A^abn|=13`、`|U_B^abn|=18`、`E_cross^base=13×18=234`、`|C|=48`、`candidate_edge_ratio=48/234≈20.5128%`、`1-48/234≈79.4872%`，并写明这是算法构造实施例而非产品实测，全交叉只作对照分母。

最终关系与分类：94 个一对一、1 个一对多分裂（A 侧 1、B 侧 2）、1 个多对一合并（A 侧 2、B 侧 1），双侧各 97 个端点被有效关系覆盖，`FROZEN_A=FROZEN_B=∅`、`BLOCKED_A=BLOCKED_B=∅`；3 个 A 侧残余端点与 8 个 B 侧残余端点各持同代次 `COMPLETE_FAIL`，形成 `|DEL|=3`、`|ADD|=8`；`CLOSURE_PASS` 后在 94 个一对一关系上得到 84 个保持不变、8 个修改、2 个纯重命名，其中 1 个修改带重命名副标签。复核式：A 侧 94+1+2+3=100，B 侧 94+2+1+8=105。

局部读取指标：`local_validation_entity_ratio` 与 `local_brep_byte_ratio` 在本例记为不可计算，明确不据此声称任何实测提升。

- [ ] **步骤 3：写入第 12.7 节冻结、回滚与重新发布微型实施例**

面层级，`U_A^face={a_1,a_2}`、`U_B^face={b_1,b_2,b_3}`；`ONE_TO_ONE(a_1,b_1)` 已物化；`a_2` 对 `b_2`、`b_3` 的候选分数与绝对裕量均无法区分，`BLOCKED` 非空时不签发 `CLOSURE_PASS` 且任何把 `a_2` 记为删除或 `b_2/b_3` 记为新增的内部暂存分类均回滚。局部读取失败后冻结闭包为 `FROZEN_A={a_2}`、`FROZEN_B={b_2,b_3}`，A 侧 1+1=2、B 侧 1+2=3；复核通过后形成 `closure_snapshot_id=cs_1` 与冻结型复核事件并同批原子发布。扩读成功后唯一确认 `ONE_TO_ONE(a_2,b_2)`、为 `b_3` 签发同代次 `COMPLETE_FAIL_B`，得 `ADD={b_3}`、`DEL=∅`，A 侧 2、B 侧 2+1=3，形成 `cs_2`；旧冻结事件不得与 `cs_2` 混用。扩读失败或预算耗尽则以实际原因保持冻结并形成替代 `cs_1` 的新快照，不得转报增删。

- [ ] **步骤 4：写入第 13 节异常场景（编号段落形式）**

设计稿第 13 节为 Markdown 表格，共 12 个场景。按 H7 改写为 12 个编号段落，每段包含“场景—指标或检测信号—处理方式—最终输出”四要素，**不使用 Markdown 表格**。

- [ ] **步骤 5：写入区分性说明与六类可选实施方式的正反实施例**

必须显式写入：普通模型修改不等于命名冲突；名称变化不等于身份变化，名称变化只在已验证一对一关系上计算并与身份变化、允许变化属性分离；指标分母为空时记为不可计算，不补为 0 或 1；维度缺失与候选无法消歧时的处理。

并按设计稿第 15.6 节最后一项，为六类可选实施方式各给出一组正反实施例：正例说明其在七步主线约束下如何补强候选或证据；反例说明脱离约束后会产生的错误（如普通一对一先行占用分裂端点、指纹相同直接判定不变、单侧名称直接判增删）。

同时保留旧文四个实施例的可用部分并改造为符合新主线：同一命名域两版本对比、影响范围传播、命名异常闭环、跨系统映射；其中命名异常闭环实施例必须以联合选择与冻结改写，不得保留“集合运算结果作为后续审计基础输入”的顺序。

- [ ] **步骤 6：反引号与表格清零**

运行：

```bash
rg -n '`|^\|' docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md
```

预期：零命中。

---

### 任务 6：附图说明与附图重建

**文件：**
- 修改：`docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md` 附图说明
- 修改：`scripts/generate_patent_submission_docs.py` 中 `MATERIALS[1]`（`permanent_naming`）与 `make_metrics_diff_flow()`

- [ ] **步骤 1：把附图说明统一为 4 幅**

旧文声明 5 幅（方法流程、索引结构、集合＋指纹流程、语义聚合、系统结构），与生成器实际产出不一致。改为恰好 4 幅，与 `figure_titles` 逐字对应：

图1 基于永久命名的CAD模型差异对比方法流程图（同时作摘要附图）；图2 基于永久命名的CAD模型差异对比系统结构图；图3 命名域元数据硬门控与三态路由流程图；图4 关系闭包与差异分类两支汇合流程图。

旧附图说明（已核实为 5 幅）到新 4 幅的对应关系必须逐条处置：旧图1（方法流程）→ 新图1，内容按七步重画；旧图2（永久命名实体索引结构）→ 删除，其内容并入具体实施方式中实体—比较规范键关联关系的文字定义；旧图3（命名集合比较和指纹比较流程）→ 删除，被新图3（元数据硬门控与三态路由）和新图4（关系闭包两支汇合）取代；旧图4（实体级差异聚合为设计语义级差异）→ 删除，其语义聚合内容并入新图4 干线末端的闭包后分类与类型化事件节点；旧图5（系统结构）→ 新图2。

删除具体实施方式中对旧图2、旧图3、旧图4主题及全部“图5”的引用，并把原引用改指到新编号。附图中出现的每个符号都必须在说明书中已定义。

- [ ] **步骤 2：重写 `permanent_naming` 的三份标签列表**

`abstract_figure`（图1，7 项，纵向流程，每项不超过 14 个汉字以免超出框高）：读取既有命名索引与关联关系；计算质量状态并三态门控；种子关系验证与三类假设；联合互斥选择与未决冻结；双侧分区检查与失败回流；闭包通过后差异分类；输出类型化事件与证据。

`system_modules`（图2，7 项，三列网格，每项不超过 11 个汉字）：既有命名关系读取模块；质量状态与门控模块；种子验证与假设生成模块；联合选择与冻结模块；分区检查与失败回流模块；闭包后差异分类模块；类型化事件输出模块。

`decision_flow`（图3，恰好 6 项，对应 `make_decision_flow()` 固定的 top／left／mid／right／bottom_left／bottom_right 六框拓扑，每项不超过 10 个汉字）：读取既有命名关系；元数据门控失败停用直配；全局未通过仅逐键可信；全局通过进入种子路径；逐键四状态门控；异常端点转候选假设。

`figure_titles` 改为步骤 1 的四个标题。不得改动 `MATERIALS[0]`（`lightweight`）的任何字段。

- [ ] **步骤 3：重写 `make_metrics_diff_flow()` 为两支汇合流程**

依设计稿第 15.8 节，改为左支、右支各 5 个节点并汇合到 6 个节点的干线。保留现有 `draw_column()` 与 `arrow()` 骨架，只改标签、坐标与汇合连线。

左支：既有永久命名索引与实体—比较规范键多重关系；硬约束／软证据／允许变化属性分离与质量状态；元数据硬门控与全局三态路由、逐键四状态；临时种子关系 R_seed；身份／指纹局部验证。

右支：稀疏异常候选边生成与双向索引查询预算；一对一／一对多／多对一三类关系假设；组聚合不变量与局部 B-rep 证据；候选权重与归一化目标 Phi_G；绝对裕量 delta_margin 判定。

干线：同一局部分量联合端点互斥选择得 R_11／R_1m／R_m1；冻结最小不动点 FROZEN_A／FROZEN_B；end_A／end_B 双侧完备不交分区检查；失败回流：撤销、局部扩读、重建假设、重新选择；CLOSURE_PASS 与 closure_snapshot_id；闭包后分类并物化 CAD_DIFF_EVENT_V1。

建议几何：画布 `1900×2320`，标题区 `y=30..110`；左支 `x=90`、右支 `x=1030`、`box_w=760`、`box_h=150`、`y_values=[170,350,530,710,890]`，`wrap_label(..., 18)`；干线 `x=90`、`box_w=1700`、`y=[1130,1310,1490,1670,1850,2030]`，`wrap_label(..., 40)`。两支最后一框底部中心各画一条箭头汇入干线首框顶部的左三分点与右三分点。

绝对禁止出现 `Q_cons`、`C_map`、`C_pair`、`综合差异 D = Σ w_i d_i`、`τ_same`、`命名集合运算`，以及任何把一对一／一对多／多对一并列为“异常候选”而无联合选择与冻结的表述。

- [ ] **步骤 4：核对脚本改动范围与被禁符号**

运行：

```bash
rg -n "Q_cons|C_map|C_pair|C_intra|命名集合|集合运算|Σ w_i" scripts/generate_patent_submission_docs.py
git diff --stat -- scripts/generate_patent_submission_docs.py
```

预期：第一条零命中；第二条显示改动集中在 `MATERIALS[1]` 与 `make_metrics_diff_flow()`，`lightweight` 条目与通用绘图／DOCX 函数未被改写。

---

### 任务 7：技术交底书 `PERMANENT` 重写与模板分离

**文件：**
- 修改：`scripts/generate_technical_disclosure_docs.py`
- 新增：`docs/submission/技术交底书/_模板/专利技术交底书-结构模板.docx`

- [ ] **步骤 1：分离模板路径并加自覆盖断言**

运行：

```bash
mkdir -p "docs/submission/技术交底书/_模板" && cp "docs/submission/技术交底书/永久命名-CAD模型差异对比-专利技术交底书-权利要求增强版.docx" "docs/submission/技术交底书/_模板/专利技术交底书-结构模板.docx"
```

把 `TEMPLATE` 指向新模板路径；`PERMANENT["output"]`（`永久命名-CAD模型差异对比-专利技术交底书-权利要求增强版.docx`）与 `LIGHTWEIGHT["output"]`（`跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-专利技术交底书.docx`）保持原文件名不变，以免打断下游引用。在 `main()` 中于循环前加断言：若 `TEMPLATE` 出现在待生成产物路径集合中则抛出异常，禁止脚本再次读取自身产物作为模板。

说明：企业微信缓存中的原空白模板在本机不存在（见任务 0 步骤 2），因此本结构模板只能从已被跟踪的生成产物复制而来，其中残留的旧内容由 `build_disclosure()` 调用的 `clear_cell()` 在每次生成时清空，不会累积；表格结构、字体和页眉页脚正是需要保留的部分。若日后取得原空白模板，应直接替换该结构模板文件，无需再改脚本。

- [ ] **步骤 2：重写 `PERMANENT["background"]`**

三段，与任务 4 步骤 1 的背景技术同源：既有永久命名的存在与作用；命名重复、漂移、失效、分裂、合并、低信息与读取失败带来的误判风险；已有做法的关注点与本方案关注点的差别。不写入组合创造性风险评价（H6）。

- [ ] **步骤 3：重写 `PERMANENT["solution"]`，逐项消除六类被禁内容**

按设计稿第 15.7 节六项逐一替换，每项都必须能在改写后的文本中指出替代物：

用有效键数量替代有效命名实体数量计算覆盖率 → 改为按有效命名实体计算 `C_cov,X=|N_X|/|U_X^l|`、`C_unique,X=|G_X|/|U_X^l|`。将平均质量与通过率混为一式 → 拆为 `R_cons`（分母全部 `K_11`）与 `S_cons`（分母 `K_score`），并另报 `C_hard`、`C_lowinfo`。把命名键交叠率作硬门槛 → 删除 `C_pair` 与 `τ_pair`，改为 `K_11` 三重唯一规则加元数据硬门控、全局质量门控三态路由、逐键四状态局部门控。异常匹配完成前直接确定单侧新增／删除 → 改为形成待审计残余端点，须取得同代次 `COMPLETE_FAIL_A/B` 并在 `CLOSURE_PASS` 后物化。未定义冻结端点仍用集合补集 → 定义 `FROZEN_A/FROZEN_B` 与 `end_A/end_B`，改用双侧完备不交分区，删除补集算法。把名称拼写、别名或允许变化属性变化倒推为身份变化 → 名称变化仅在已验证一对一关系上以别名取商后的语义令牌集合与全部最大匹配计算，输出可复算的名称变化状态，与身份判定分离。

同时替换：`d_num` 改为设计稿第 10.1 节的容差归一化形式（`T(x,y)=eps_abs+eps_rel*max(|x|,|y|)`，`d_num=clip((|x-y|-T)/max(T,delta),0,1)`）；`D=Σw_i d_i` 改为第 10.5 节的变化证据覆盖门控体系（`J_change^l`、`G_required^l`、`coverage_change`、`tau_change^l`、`change_gate`、`D`、`tau_same^l`）；冲突率改为 `C_dup,X`、`C_cross`、`C_mapcov,X`、`C_mapedgeconf,X`、`C_mapconf,X`、`C_conf` 并保留各自空分母规则。方案叙述顺序改为七步。

- [ ] **步骤 4：替换数值实施例**

删除“面、边和特征层级”“共有命名键92个”“89个通过”“`C_pair` 约为93.88%”“`Q_cons` 约为96.74%”“`τ_pair=90%`”等全部旧数值，改用任务 5 步骤 2 的第 12 节数值，并补入第 12.7 节冻结与重新发布微型例的双侧计数。

- [ ] **步骤 5：重写 `advantages` 与 `key_points`**

`advantages` 严格按 H5 的四类限定表述，并附锚点核查稿第 5.3 节六条因果链的简版；删除“将全量几何实体匹配转化为……提高效率”“可跳过完整几何验证，提高效率”等性能断言。

`key_points` 改为八项绑定闭环＋负面边界：既有名称—实体关联关系基数保真；硬身份约束／软身份证据／允许变化属性／组聚合不变量四类用途分离；元数据硬门控与全局三态、逐键四状态门控；一对一／一对多／多对一假设在同一局部分量联合互斥选择；未决端点冻结最小不动点；`end_A/end_B` 双侧完备不交分区检查；失败回流与代次失效；闭包成立后才分类并物化统一类型化事件；以及不生成或更新永久命名规则、比较规范键不回写的负面边界。

- [ ] **步骤 6：同步 `PERMANENT["figures"]` 标题**

四项图题改为与任务 6 步骤 1 完全一致的文字。不得改动 `LIGHTWEIGHT` 字典的任何字段。

- [ ] **步骤 7：核对被禁内容清零**

运行：

```bash
rg -n "Q_cons|C_map|C_pair|C_intra|τ_pair|93\.88|96\.74|先标记为新增候选|集合运算|面、边和特征层级" scripts/generate_technical_disclosure_docs.py
```

预期：零命中。

---

### 任务 8：生成、渲染与逐页复核

**文件：**
- 生成：`docs/submission/assets/permanent_naming-fig1..fig4-*.png`、两份 DOCX、`docs/submission/rendered/permanent-naming-consolidated/`

- [ ] **步骤 1：只生成永久命名材料的附图与申请文件 DOCX**

`main()` 遍历 `MATERIALS` 中两个材料，会连带重绘三份被 Git 跟踪的 `lightweight-fig*.png` 并重写轻量化申请文件 DOCX。为把改动范围限定在本轮归属，不调用 `main()`，只构建 `MATERIALS[1]`：

```bash
py -3.12 -c "import sys; sys.path.insert(0,'scripts'); import generate_patent_submission_docs as g; g.build_material(g.MATERIALS[1]); print(g.MATERIALS[1]['output'])"
```

预期：打印 `docs/submission/永久命名-CAD模型差异对比-发明专利申请文件.docx`；`docs/submission/assets/` 下出现 `permanent_naming-fig4-metrics-diff.png`，fig1／fig2／fig3 被更新。

随后运行：

```bash
git status --short docs/submission/assets/
```

预期：仅 `permanent_naming-fig1/2/3` 显示为已修改、`permanent_naming-fig4` 为新增；三份 `lightweight-fig*` 完全无改动。若出现 `lightweight-fig*` 改动，说明误调用了 `main()`，须查明后用 `git checkout -- docs/submission/assets/lightweight-fig1-method.png docs/submission/assets/lightweight-fig2-system.png docs/submission/assets/lightweight-fig3-decision.png` 还原（这三份是生成产物而非用户源文件，还原安全；两份脚本仍严禁还原）。

- [ ] **步骤 2：逐幅目视检查 4 幅附图**

逐个读取 4 个 PNG 并确认：文字无截断、无越出框线、框与框不重叠；箭头方向与流程一致；图4 两支各 5 框、干线 6 框且两条汇入箭头到位；无 `Q_cons`／`C_map`／`C_pair`；符号与说明书一致；仅黑白线条，无灰度或阴影。发现问题回到任务 6 步骤 3 调整坐标或缩短标签后重新生成。

- [ ] **步骤 3：只生成永久命名技术交底书 DOCX**

同理，`LIGHTWEIGHT["output"]`（`跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-专利技术交底书.docx`）已被 Git 跟踪，`main()` 会连带重写它。只构建 `PERMANENT`：

```bash
py -3.12 -c "import sys; sys.path.insert(0,'scripts'); import generate_technical_disclosure_docs as g; print(g.build_disclosure(g.PERMANENT))"
```

预期：打印 `docs/submission/技术交底书/永久命名-CAD模型差异对比-专利技术交底书-权利要求增强版.docx`；`docs/submission/技术交底书/_模板/专利技术交底书-结构模板.docx` 的修改时间不变（模板未被覆盖）；`docs/submission/技术交底书/跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-专利技术交底书.docx` 与 `永久命名-CAD模型差异对比-专利技术交底书.docx` 均无改动。

另需单独验证任务 7 步骤 1 的自覆盖断言生效：

```bash
py -3.12 -c "import sys; sys.path.insert(0,'scripts'); import generate_technical_disclosure_docs as g; g.TEMPLATE=g.PERMANENT['output']; g.main()" 2>&1 | tail -3
```

预期：抛出异常并指出模板不得与产物同路径。

- [ ] **步骤 4：DOCX 转 PDF 并分页栅格化**

本机无 LibreOffice，使用 Word COM：

```bash
powershell -NoProfile -Command "\$w=New-Object -ComObject Word.Application; \$w.Visible=\$false; \$src=(Resolve-Path 'docs/submission/永久命名-CAD模型差异对比-发明专利申请文件.docx').Path; \$dst=(Join-Path (Resolve-Path 'docs/submission/rendered').Path 'permanent-naming-consolidated/permanent-naming-consolidated.pdf'); New-Item -ItemType Directory -Force -Path (Split-Path \$dst) | Out-Null; \$d=\$w.Documents.Open(\$src); \$d.SaveAs([ref]\$dst,[ref]17); \$d.Close(); \$w.Quit()"
```

再运行：

```bash
py -3.12 -c "
import fitz, pathlib
p=pathlib.Path('docs/submission/rendered/permanent-naming-consolidated')
doc=fitz.open(p/'permanent-naming-consolidated.pdf')
print('页数', doc.page_count)
for i,pg in enumerate(doc, 1):
    pg.get_pixmap(dpi=140).save(p/f'page-{i:02d}.png')
"
```

预期：PDF 生成成功，分页 PNG 数量等于页数。对技术交底书 DOCX 重复同一流程，输出到 `docs/submission/rendered/disclosure-permanent-naming-consolidated/`。

- [ ] **步骤 5：逐页复核（设计稿第 16 节第 14 项）**

逐页读取全部 PNG，检查：公式与下标是否完整显示、编号段落是否错行、分页处是否切断权利要求项或实施例、附图是否清晰且未被压缩变形、中文字体是否正常（无方框或乱码）、中文标点与全角符号是否正确、页码是否连续、无残留反引号或竖线字符。逐页记录结论；任一页不合格则回到对应任务修正后重新生成，不得以“已生成”代替“已复核”。

---

### 任务 9：一致性、可复算与形式要求验证

**文件：**
- 验证：`docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md`、两份脚本
- 修改：`docs/superpowers/specs/2026-08-06-permanent-naming-metrics-and-model-diff-strengthening-design.md` 第 17 节

- [ ] **步骤 1：旧架构与被禁符号全局扫描**

运行：

```bash
rg -n "集合运算|命名集合|集合补集|先集合差|先标记为新增|Q_cons|C_map|C_pair|C_intra|τ_pair|R_name|REVIEW_A|REVIEW_B" docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md scripts/generate_patent_submission_docs.py scripts/generate_technical_disclosure_docs.py
```

预期：退出码 1，零命中。

- [ ] **步骤 2：占位符与模糊表达扫描**

运行：

```bash
rg -n "TODO|待定|后续补充|适当处理|类似上述|参见上文" docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md
```

预期：零命中。

- [ ] **步骤 3：新符号先定义后使用扫描**

运行：

```bash
rg -n "R_seed|R_11|R_1m|R_m1|FROZEN_A|FROZEN_B|BLOCKED_A|COMPLETE_FAIL|CLOSURE_PASS|CAD_DIFF_EVENT_V1|end_A|K_11|K_score|K_lowinfo|C_mapcov|closure_snapshot_id" docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md
```

预期：每个符号的首次出现位置都在其定义处或定义之后；权利要求中使用的每个符号在说明书具体实施方式中均有定义段落。逐符号记录首现行号与定义行号。

- [ ] **步骤 4：数值实施例复算**

运行：

```bash
py -3.12 -c "
print('C_cov,A', 96/100); print('C_cov,B', round(100/105*100,2)); print('C_unique,A', 94/100); print('C_unique,B', round(97/105*100,2))
print('R_cons', round(87/90*100,2)); print('S_cons', (87*0.96+3*0.42), (87*0.96+3*0.42)/90)
print('C_dup,A', round(2/96*100,2)); print('C_dup,B', 3/100); print('C_cross', round(3/90*100,2))
print('E_cross', 13*18); print('ratio', round(48/234*100,4), 'reduce', round((1-48/234)*100,4))
print('A侧', 94+1+2+3, 'B侧', 94+2+1+8); print('分类', 84+8+2)
print('微型例 A', 1+1, 'B', 1+2)
"
```

预期：`96%`、`95.24`、`94%`、`92.38`、`96.67`、`84.78`／`0.942`、`2.08`、`3%`、`3.33`、`234`、`20.5128`／`79.4872`、`100`／`105`、`94`、`2`／`3`，与申请正文和技术交底书中写入的数值逐项一致。

- [ ] **步骤 5：形式要求核查**

摘要不超过 300 字（任务 1 步骤 2 的命令）；权利要求编号连续且无前向引用（任务 2 步骤 4 的命令）；设备与介质权利要求引用的方法项区间与实际编号一致；附图说明声明的图数等于生成的 4 幅且图题逐字一致；说明书中出现的附图引用不含图5及旧主题。

运行：

```bash
rg -n "图[1-9]" docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md | rg -v "图1|图2|图3|图4"
```

预期：零命中。

- [ ] **步骤 6：现有技术边界与效果限定核查**

运行：

```bash
rg -n "US10816957B2|US20200218835A1|US11288411B2|不构成授权|官方数据库|检索范围" docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md
```

预期：申请正文中**不出现**“不构成授权”“官方数据库”“检索范围”等内部风险表述（H6）；专利号仅可出现在背景技术的引证位置。同时确认这些风险表述在设计稿第 3.3、14、17 节与锚点核查稿第 6 节中仍然完整保留。

- [ ] **步骤 7：逐项核对设计稿第 16 节 14 项交付标准**

逐项判定通过或不通过，并在计划执行记录中写明依据（引用具体行号或验证命令输出）。任一项不通过则回到对应任务修正。

- [ ] **步骤 8：勾选设计稿第 17 节 13 项最终确认清单**

第 17 节的 13 个复选框当前全部未勾选。本轮已完成正式改写，逐项核对后把 `- [ ]` 改为 `- [x]`，并在每项末尾追加一行核对结果（形如“结果：通过。……”），指明落位的文件与位置。仅改动第 17 节，不得改动第 1 至 16 节的任何技术定义、公式或数值。

---

### 任务 10：交付与 Git 纪律

**文件：**
- 验证：仓库全局

- [ ] **步骤 1：确认改动范围**

运行：

```bash
git status --short
```

预期改动限于：`docs/patent/2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md`、`scripts/generate_patent_submission_docs.py`、`scripts/generate_technical_disclosure_docs.py`、`docs/superpowers/specs/2026-08-06-permanent-naming-metrics-and-model-diff-strengthening-design.md`（仅第 17 节）、本计划文件、`docs/submission/assets/permanent_naming-fig*.png`、新增模板与 `rendered/` 新目录、未跟踪的 DOCX 产物。

必须未被改动：跨 CAD 内核轻量摘要案的正文、计划、设计稿，`MATERIALS[0]` 与 `LIGHTWEIGHT` 字典，三份 `lightweight-fig*.png`。

- [ ] **步骤 2：确认用户原有脚本改动未被丢弃**

运行：

```bash
git diff -- scripts/ > tmp/backup-before-consolidated-landing/scripts-after.diff && diff <(rg -c "" tmp/backup-before-consolidated-landing/scripts-user-uncommitted.diff) <(rg -c "" tmp/backup-before-consolidated-landing/scripts-after.diff) || true
```

预期：新 diff 是在用户改动基础上继续演进的结果，即用户新增的第 4 幅附图机制与模板路径修正仍在（`extra_figure` 键、`figure_titles` 第 4 项、`make_metrics_diff_flow` 函数、仓库内 `TEMPLATE` 相对路径均存在），只有被禁内容被替换。在执行记录中写明该核对结论。

- [ ] **步骤 3：不自动提交**

工作区含用户已有未提交改动与多份未跟踪审查稿。除用户另行明确要求外，不执行 `git add`、`git commit` 或 `git push`。

若用户要求提交，仅提交本分支归属的永久命名落稿材料（申请正文、两份脚本、设计稿第 17 节勾选、本计划、`permanent_naming-fig*.png`、新增结构模板、`rendered/` 下本轮新目录），并排除：`.planning/` 研究记录、`tmp/`（含本轮备份目录与 PDF 复核暂存）、跨内核轻量摘要计划与设计稿（归属 `codex/lightweight-summary-strengthening` 分支）、未跟踪的 DOCX 产物（除用户明确要求纳入）。

- [ ] **步骤 4：交付说明**

向用户报告时如实列出：已落位的四份产物与各自变更要点；设计稿第 16 节 14 项与第 17 节 13 项的逐项结论；逐页复核发现并修正的问题；仍需专利代理人完成的事项（按锚点核查稿第 6 节：最终独立权利要求在官方数据库的逐件复检、负面限定支持方式确认、算法特征与技术效果联系的确认、CNIPA《专利审查指南》最新文本的逐条补核）。不得表述为已获得授权保证或已完成穷尽检索。
