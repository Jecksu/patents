# 跨 CAD 内核轻量摘要专利强化实现计划

> **面向 AI 代理的工作者：** 必需子技能：使用 superpowers:subagent-driven-development（推荐）或 superpowers:executing-plans 逐任务实现此计划。步骤使用复选框（`- [ ]`）语法来跟踪进度。

**目标：** 以减少跨 CAD 内核场景中的完整 BREP 读取和实体两两比较为核心，形成包含候选关系消歧、局部 BREP 验证、拓扑谱系和最终差异闭包的强化申请材料。

**架构：** 以新的强化 Markdown 草案作为申请正文唯一内容源，保留原稿；申请文件生成脚本从强化草案生成独立命名的 DOCX 和附图，技术交底脚本同步维护同一技术闭环。通过文本、结构、公式复算和逐页渲染四类校验确保正文、交底书和附图一致。

**技术栈：** Markdown、Python 3、python-docx、Pillow、LibreOffice、Poppler/PDF 渲染。

---

## 文件结构

- 创建 `docs/patent/2026-08-06-cad-model-diff-lightweight-summary-strengthened-patent-draft.md`：强化后的摘要、权利要求书和说明书正文。
- 修改 `scripts/generate_patent_submission_docs.py`：指向强化草案，使用独立资源键和输出文件名，生成候选闭环附图。
- 修改 `scripts/generate_technical_disclosure_docs.py`：同步强化技术方案、差异公式、场景矩阵要点和附图。
- 创建 `docs/submission/跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-发明专利申请文件-强化版.docx`：最终申请文件。
- 创建 `docs/submission/技术交底书/跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-专利技术交底书-强化版.docx`：最终技术交底书。
- 创建 `docs/submission/assets/lightweight_strengthened-*.png`：强化版方法、系统、候选闭环和场景处理附图。
- 使用 `docs/submission/rendered/lightweight-strengthened-word/` 和 `docs/submission/rendered/disclosure-lightweight-strengthened-word/`：内部渲染验证目录。

### 任务 1：建立强化专利正文

**文件：**
- 创建：`docs/patent/2026-08-06-cad-model-diff-lightweight-summary-strengthened-patent-draft.md`
- 参考：`docs/superpowers/specs/2026-08-06-cross-kernel-lightweight-summary-strengthening-design.md`

- [x] **步骤 1：编写摘要和独立权利要求主链**

摘要和权利要求1必须依次包含：无可比较永久命名入口、初始阶段不执行完整 BREP 逐项比较、至少两类轻量摘要、变换假设、至少两个摘要闸门、候选/排除关系、互斥消歧、候选邻域局部 BREP、`R_11/R_1m/R_m1`、关系闭包后新增/删除/共有、语义聚合和双侧定位。

- [x] **步骤 2：编写从属方法权利要求**

逐项覆盖摘要来源、全局/局部/语义/关系摘要、对称变换多假设、闸门输入输出、差异公式、候选裕量、互斥求解、局部 BREP 范围、面/边验证、拓扑分裂、拓扑合并、集合闭包、装配发生、语义聚合、证据记录和压缩率。

- [x] **步骤 3：编写系统、设备和存储介质权利要求**

系统模块必须与方法步骤一一对应，系统独立权利要求包含候选关系消歧、局部 BREP 按需读取和差异闭包模块。

- [x] **步骤 4：编写说明书技术方案和公式**

写入 `d_num`、`d_set`、候选相似度 `S(a,b)`、候选裕量、验证差异 `D(a,b)`、`COMMON`、`DEL`、`ADD`，并定义缺失分量、空集合和硬冲突处理。

- [x] **步骤 5：编写场景实施例和可复算数量实施例**

至少覆盖刚体位姿、单位变化、未知仿射、对称歧义、跨内核拓扑分裂、拓扑合并、局部特征新增/删除、装配重复实例、摘要缺失和局部 BREP 读取失败。数量实施例使用 100/105 个实体、10,500 个理论实体对、126 个闸门后候选、94 个一对一关系、1 个分裂、1 个合并、3 个删除和8个新增。

- [x] **步骤 6：运行文本结构验证**

运行：

```powershell
rg -n "候选关系|排除关系|互斥|局部BREP|R_11|拓扑分裂|拓扑合并|COMMON|DEL|ADD|设计语义|定位" docs/patent/2026-08-06-cad-model-diff-lightweight-summary-strengthened-patent-draft.md
```

预期：每个必要术语至少在权利要求和说明书中各出现一次；没有占位标记或未定义公式。

### 任务 2：同步申请文件生成和附图

**文件：**
- 修改：`scripts/generate_patent_submission_docs.py`
- 创建：`tests/test_generate_patent_submission_docs.py`

- [x] **步骤 0：按TDD建立生成器行为测试并验证红灯**

使用标准库 `unittest` 验证强化材料采用独立键、强化 Markdown 和强化 DOCX 文件名，声明5张强化附图，并能生成图4候选闭环与图5场景处理；先运行测试并确认因行为尚未实现而失败。

- [x] **步骤 1：将轻量摘要材料源切换到强化草案和独立输出名**

将轻量材料键设置为 `lightweight_strengthened`，源文件设置为强化 Markdown，输出设置为强化版 DOCX；保留永久命名材料及其现有未提交改动。

- [x] **步骤 2：更新图1和图2标签**

方法图包含“摘要闸门候选/排除”“互斥消歧”“候选邻域局部 BREP”“关系闭包”；系统图包含候选关系消歧模块、按需 BREP 模块和差异闭包模块。

- [x] **步骤 3：新增候选关系与差异闭环图**

新增独立绘图函数，图中展示 `C -> 互斥消歧 -> 局部验证 -> R_11/R_1m/R_m1 -> COMMON/DEL/ADD -> 语义聚合`，输出 `lightweight_strengthened-fig4-candidate-closure.png`。

- [x] **步骤 4：新增场景处理图**

新增场景分流图，展示刚体/单位归一化、全局形变、对称多假设、分裂/合并和数据不足待复核，输出 `lightweight_strengthened-fig5-scenarios.png`。

- [x] **步骤 5：运行生成脚本**

运行：

```powershell
& 'C:\Users\sxd\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' scripts/generate_patent_submission_docs.py
```

预期：强化版申请 DOCX 和5张强化版附图存在，永久命名 DOCX仍可正常生成。

### 任务 3：同步技术交底书

**文件：**
- 修改：`scripts/generate_technical_disclosure_docs.py`
- 创建：`tests/test_generate_technical_disclosure_docs.py`

- [ ] **步骤 0：按TDD建立强化交底书测试并验证红灯**

使用标准库 `unittest` 锁定强化版独立输出名、五张强化附图、技术闭环、核心公式、数量实施例和永久命名材料不变；先运行测试确认因强化交底行为尚未实现而失败。

- [ ] **步骤 1：更新轻量摘要技术问题和方案段落**

将核心问题、顺序闸门、稀疏候选关系、候选消歧、按需局部 BREP、拓扑谱系和差异闭包写入 `LIGHTWEIGHT`。

- [ ] **步骤 2：加入公式、数量实施例和场景降级**

交底书必须给出 `S(a,b)` 与 `D(a,b)` 的区别、`COMMON/DEL/ADD` 的计算顺序以及 100/105 实体复算过程。

- [ ] **步骤 3：更新关键创新点和附图引用**

关键创新点不再依赖“凸包”或“变换”单点，而强调完整处理闭环；引用5张强化版附图。

- [ ] **步骤 4：运行交底书生成脚本**

运行：

```powershell
& 'C:\Users\sxd\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' scripts/generate_technical_disclosure_docs.py
```

预期：生成强化版轻量摘要技术交底书，同时不覆盖原轻量摘要交底书。

### 任务 4：结构和内容验证

**文件：**
- 验证：强化 Markdown、强化申请 DOCX、强化技术交底 DOCX、强化附图

- [ ] **步骤 1：提取 DOCX 文字并检查必要特征**

使用绑定 Python 和 `python-docx` 提取两份 DOCX，检查“候选关系”“排除关系”“互斥”“局部BREP”“拓扑分裂”“拓扑合并”“新增实体”“删除实体”“设计语义”和数量实施例均存在。

- [ ] **步骤 2：复算数量实施例**

验证：`100*105=10500`、`126/10500=1.2%`、关系端点 `94+1+2=97` 和 `94+2+1=97`、删除 `100-97=3`、新增 `105-97=8`。

- [ ] **步骤 3：检查权利要求引用和术语一致性**

检查从属权利要求引用的权利要求号存在；摘要、独权、系统权利要求和说明书对步骤顺序使用相同术语。

- [ ] **步骤 4：检查 Git 差异边界**

运行：

```powershell
git diff --stat
git diff -- docs/patent/2026-08-06-cad-model-diff-lightweight-summary-strengthened-patent-draft.md scripts/generate_patent_submission_docs.py scripts/generate_technical_disclosure_docs.py
```

预期：永久命名脚本中用户已有指标改动仍保留，没有无关文件被覆盖。

### 任务 5：DOCX 渲染和逐页视觉验证

**文件：**
- 验证：`docs/submission/跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-发明专利申请文件-强化版.docx`
- 验证：`docs/submission/技术交底书/跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-专利技术交底书-强化版.docx`

- [ ] **步骤 1：渲染强化申请文件**

运行 `documents` 技能的 `render_docx.py`，输出到 `docs/submission/rendered/lightweight-strengthened-word/` 并生成 PDF。

- [ ] **步骤 2：逐页检查强化申请文件**

逐页确认标题、权利要求分页、公式字符、5张附图、页眉页脚、图注和中文字体无裁切、重叠或乱码。

- [ ] **步骤 3：渲染并逐页检查强化技术交底书**

输出到 `docs/submission/rendered/disclosure-lightweight-strengthened-word/`，检查表格单元格、公式、长段落、附图和分页。

- [ ] **步骤 4：修复后重新渲染**

任何内容或版式修改后重新生成 DOCX 并重新检查全部页面，直到最新渲染无缺陷。

### 任务 6：完成审计

**文件：**
- 复核：全部强化交付物和当前工作区状态

- [ ] **步骤 1：逐项核对用户目标**

确认核心技术问题、详细差异计算流程、场景处理、减少 BREP 读取、减少两两比较、实体定位和语义差异均有直接正文证据。

- [ ] **步骤 2：进行专利性边界自检**

确认独权不退化为“可视化粗比较后进行 BREP 比较”，不以全模型图匹配、全量点云距离或凸包单点作为核心。

- [ ] **步骤 3：更新持久计划并记录最终验证结果**

将任务阶段标记完成，记录生成文件、页数、验证命令和剩余需代理人核实的法律问题。
