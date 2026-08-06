from __future__ import annotations

import warnings
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "docs" / "submission" / "技术交底书"
ASSET_DIR = ROOT / "docs" / "submission" / "assets"
TEMPLATE = OUT_DIR / "永久命名-CAD模型差异对比-专利技术交底书-权利要求增强版.docx"


def set_paragraph_text(paragraph, text: str, bold: bool = False, size: int | None = None) -> None:
    paragraph.clear()
    run = paragraph.add_run(text)
    run.bold = bold
    if size:
        run.font.size = Pt(size)


def clear_cell(cell) -> None:
    cell._tc.clear_content()
    cell.add_paragraph()


def add_heading(cell, text: str) -> None:
    p = cell.paragraphs[-1] if len(cell.paragraphs) == 1 and not cell.paragraphs[-1].text else cell.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(12)


def add_para(cell, text: str = "") -> None:
    p = cell.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    if text:
        run = p.add_run(text)
        run.font.size = Pt(10.5)


def add_bullets(cell, items: list[str]) -> None:
    for item in items:
        p = cell.add_paragraph(style=None)
        p.paragraph_format.left_indent = Pt(14)
        p.paragraph_format.first_line_indent = Pt(-10)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(f"• {item}")
        run.font.size = Pt(10.5)


def fill_row(cell, title: str, paragraphs: list[str] | None = None, bullets: list[str] | None = None) -> None:
    clear_cell(cell)
    add_heading(cell, title)
    for text in paragraphs or []:
        add_para(cell, text)
    if bullets:
        add_bullets(cell, bullets)


def add_figures(cell, figures: list[tuple[Path, str]]) -> None:
    for image_path, caption in figures:
        if not image_path.exists():
            continue
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(image_path), width=Inches(5.5))
        cap = cell.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cap.add_run(caption)
        run.font.size = Pt(9)


def fill_metadata(doc: Document, title: str) -> None:
    replacements = {
        2: f"专利申请名称：{title}",
        3: "此专利是否已经公开：否",
        4: "申请人及部门：待补充",
        5: "发明人或撰写人：待补充",
        6: "发明人或撰写人的电话：待补充",
        7: "类型：发明专利",
        8: "内部编号：待补充",
    }
    for idx, text in replacements.items():
        if idx < len(doc.paragraphs):
            set_paragraph_text(doc.paragraphs[idx], text, size=10.5)

    if doc.tables:
        table = doc.tables[0]
        if len(table.rows) >= 3:
            for row_idx in [1, 2]:
                for col_idx in range(1, len(table.rows[row_idx].cells)):
                    table.rows[row_idx].cells[col_idx].text = "待补充"


def validate_figure_paths(material: dict) -> None:
    missing = [image_path for image_path, _ in material["figures"] if not image_path.exists()]
    if not missing:
        return
    message = "附图文件不存在：\n" + "\n".join(str(path) for path in missing)
    if material.get("require_all_figures", False):
        raise FileNotFoundError(message)
    warnings.warn(message, RuntimeWarning, stacklevel=2)


def validate_template_structure(doc: Document) -> None:
    if len(doc.tables) < 2:
        raise ValueError("技术交底书模板应包含至少2个表格")
    if len(doc.tables[0].rows) < 3:
        raise ValueError("技术交底书模板的发明人信息表至少需要3行")
    if len(doc.tables[1].rows) < 6:
        raise ValueError("技术交底书模板的技术内容表至少需要6行")


LIGHTWEIGHT = {
    "short_title": "跨CAD内核模型差异快速对比方法",
    "full_title": "一种面向跨CAD内核的基于轻量摘要闸门和受限局部BREP验证的CAD模型差异对比方法及系统",
    "output": OUT_DIR / "跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-专利技术交底书-强化版.docx",
    "require_all_figures": True,
    "figures": [
        (ASSET_DIR / "lightweight_strengthened-fig1-method.png", "图1 方法闭环流程示意图"),
        (ASSET_DIR / "lightweight_strengthened-fig2-system.png", "图2 系统模块与比较会话数据流示意图"),
        (ASSET_DIR / "lightweight_strengthened-fig3-decision.png", "图3 多变换假设与摘要闸门示意图"),
        (ASSET_DIR / "lightweight_strengthened-fig4-candidate-closure.png", "图4 候选关系消歧、局部验证和差异闭包示意图"),
        (ASSET_DIR / "lightweight_strengthened-fig5-scenarios.png", "图5 场景处理与降级示意图"),
    ],
    "background": [
        "CAD模型版本对比、供应商模型验收和跨格式转换验证需要识别两个三维模型之间的实体级差异和设计语义级差异。在跨CAD内核、跨文件格式或经过不同转换链路的场景中，两个模型通常缺少可比较永久命名，已有名称也可能因来源和命名域不同而不可信，因而不能直接作为实体对应关系。",
        "完整BREP读取会引入几何、拓扑和参数数据的解析成本；若进一步对两个完整模型的实体笛卡尔积进行两两比较，候选数量为两个实体域基数的乘积。跨内核拓扑重构、位姿变化、单位变化、镜像或其他仿射变换，以及对称重复实体，还可能造成同一设计对象的表达差异、候选歧义和误判。仅输出模型整体相同或不同，也不能定位具体实体及其设计语义。",
        "现有技术文献CN119312543A公开的流程包括：可视化不一致时判定不同，可视化一致时进入完整/分层BREP比较。",
        "本方案由摘要闸门形成实体候选和排除关系，构造稀疏候选图，并通过互斥消歧、按判定需要读取候选邻域BREP、建立跨内核拓扑谱系以及保留待复核端点的差异闭包，输出实体级和设计语义级可定位结果。未知仿射或业务不允许的镜像作为全局差异记录，全局差异不默认终止局部定位。",
    ],
    "solution": [
        "步骤一，建立比较会话和实体域。当两个模型不存在可比较永久命名，或者名称来源、命名域或稳定性不可信时，系统按装配、零件、体、壳、面、边或特征等层级l分别构造模型A和模型B的实体域U_A^l、U_B^l，并将不同维度或不同实体类型放入独立比较池。系统可保存模型内瞬时定位符用于回到原模型选择实体，但该定位符不称为永久命名，也不作为跨模型身份相等的依据。",
        "步骤二，提取轻量摘要。初始阶段不执行两个完整模型BREP实体逐项比较，而从文件元数据、轻量网格、实体属性、拓扑统计和语义识别结果中提取全局摘要G_X、局部实体摘要L_X(e)、语义摘要SEM_X(e)和关系摘要H_X(e)中的至少两类。全局摘要可包括包围盒、质心、主轴、尺度和类型计数；局部摘要可包括面积、长度、体积、法向或曲率统计；语义摘要可包括孔、倒角、圆角、阵列或装配角色；关系摘要可包括邻接类型、多重度和局部度数。",
        "步骤三，生成并验证坐标变换假设。系统基于全局摘要和锚点摘要生成刚体、相似变换、仿射、镜像以及针对对称结构的对称多变换假设，并以摘要残差、语义锚点和局部关系一致性验证各假设。单位换算属于业务允许范围时执行尺度归一化；检测到未知仿射或业务不允许镜像时记录全局差异，但不默认终止局部定位，可在多个变换分支中继续产生候选和证据。",
        "步骤四，执行至少两个顺序摘要闸门。第一闸门可按实体层级、维度、类型、尺度区间和空间索引产生粗候选及硬排除；后续闸门可按局部摘要、语义摘要和关系摘要逐步收缩候选。每一闸门至少输出候选、排除、约束或证据中的一种，后闸门继承前闸门结果；闸门链最终形成实体级候选关系和排除关系，并以保留候选为边构造C_gate稀疏候选图。被硬冲突排除的实体对不进入后续相似度求解。",
        "摘要距离采用可复核口径。对于数值维度x、y，d_num(x,y)=clip((|x-y|-eps_abs)/(eps_abs+eps_rel*max(|x|,|y|)),0,1)，分母为0时采用该维度最小正尺度。对于邻接类型、语义标签或拓扑角色等带重复计数的数据，采用多重集而非普通集合：I_ms=Σ_t min(c_X(t),c_Y(t))、U_ms=Σ_t max(c_X(t),c_Y(t))、d_ms=1-I_ms/max(U_ms,1)，不得以普通集合Jaccard替代多重集。",
        "步骤五，计算候选相似度并消歧。定义J_cand(a,b)={i | 实体a和b在候选摘要维度i上的值均存在且可比较}。候选相似度为S(a,b)=Σ alpha_i(1-d_i)/Σ alpha_i，各求和仅遍历J_cand(a,b)，权重满足alpha_i≥0且Σ_{i∈J_cand(a,b)}alpha_i>0。对每个端点记录最佳候选、次佳候选及最佳/次佳裕量，结合双向排名、端点互斥、局部关系一致性执行局部最大权匹配或等价求解。候选相似度S只用于候选排序和消歧，不能替代验证后差异D。缺失维度从分子和分母剔除并重新归一化，全缺失进入低信息/REVIEW，类型、业务约束或确定拓扑角色的硬冲突直接排除。",
        "步骤六，按判定需要分级读取局部BREP。读取对象仅包括已选中候选、存在歧义的候选、阈值附近候选、摘要缺失候选、周期或等价表达候选、非流形候选以及分裂/合并候选。读取范围限于候选自身、至多h阶邻域、所属语义结构或候选组，并受到实体数预算、字节数预算、查询次数预算和邻域阶数预算的联合约束。预算超限、数据源读取失败或证据仍不充分时，系统保护相应REVIEW端点，或者按策略执行受控回退，不无条件读取两个模型的完整BREP。",
        "步骤七，执行面、边和体壳局部验证。面局部验证比较支承曲面类型和参数、裁剪区域覆盖、面积、方向、边界环及邻面关系；边局部验证比较支承曲线、端点、长度、方向、周期性和相邻面；体壳局部验证比较闭合性、连通分量、欧拉或壳层统计、体积和包含关系。容差按单位和几何尺度归一化，解析曲面与NURBS等价表达可通过映射到共同参数域或三维采样证据验证。",
        "步骤八，建立跨版本拓扑谱系。高置信一对一关系记为R_11⊆U_A^l×U_B^l；一对多关系记为R_1m⊆U_A^l×P_+(U_B^l)；多对一关系记为R_m1⊆P_+(U_A^l)×U_B^l，其中P_+表示非空幂集，R_1m和R_m1的组基数至少为2，且R_11、R_1m、R_m1三类关系两侧端点分别互斥，任一实体端点不能同时被两个有效主关系占用。",
        "修剪面分裂的拓扑关系为旧修剪面因新增拓扑分界对应多个新修剪面；分裂时，新模型相对于旧模型新增内部分界。存在共同参数化时，以参数域内各新修剪域的并集覆盖原修剪域；不存在共同参数化时，将各修剪域映射到三维后验证裁剪曲面区域覆盖。区域覆盖和度量守恒为必要条件，并结合支承几何、新增内部分界及外部邻接证据确认R_1m。修剪面合并的拓扑关系为多个旧修剪面对应一个新修剪面；合并时，旧模型中的内部分界在新模型中删除。存在共同参数化时，以参数域内各旧修剪域的并集覆盖新修剪域；不存在共同参数化时，将各旧修剪域映射到三维后验证其裁剪曲面区域覆盖新修剪面，并结合度量守恒、支承几何、已删除内部分界和外部邻接证据确认R_m1。重复键不是分裂/合并的充分证据。周期缝边或等价参数化若不满足组谱系条件，则作为R_11表示差异。",
        "步骤九，执行关系闭包和新增删除分类。令R=R_11∪R_1m∪R_m1、COMMON_11=R_11、TRACK_A=dom(R)、TRACK_B=ran(R)。待复核端点集合满足REVIEW_A⊆U_A^l、REVIEW_B⊆U_B^l。系统先占用全部有效关系端点，再保护证据不足的REVIEW_A和REVIEW_B端点，最终计算DEL=U_A^l-(dom(R)∪REVIEW_A)及ADD=U_B^l-(ran(R)∪REVIEW_B)。R_11不等于未修改，它只表示实体谱系为一对一；REVIEW保护闭包可防止证据不足的端点被强制归入新增或删除。",
        "步骤十，区分一对一实体的保持不变与修改。定义J_verify(a,b)={i | 实体a和b在局部验证维度i上的结果均存在且可比较}。验证后差异为D(a,b)=Σ w_i d_i/Σ w_i，各求和仅遍历J_verify(a,b)，权重满足w_i≥0且Σ_{i∈J_verify(a,b)}w_i>0；D与候选相似度S(a,b)明确分开，可组合几何、拓扑、参数、语义和关系差异。缺失维度同样从分子和分母剔除并重新归一化，全缺失时保持低信息/REVIEW。对于R_11中的实体对，以D和阈值tau_same区分保持不变与修改；保持不变、修改、分裂、合并、新增、删除和REVIEW等主分类端点互斥。",
        "步骤十一，执行语义聚合与双侧定位。系统沿生成、参数引用、拓扑邻接和装配约束关系，把实体差异聚合为孔、倒角、圆角、阵列、槽或装配约束等设计语义变化，并保存两个原模型中的双侧定位信息。对于装配重复发生，使用发生路径、父级上下文、变换和邻接关系区分同源实例；对于顺序重排，不仅依赖实例序号，而通过多重集、上下文和互斥匹配恢复对应。",
        "步骤十二，生成证据记录和可审计指标。每个结论记录摘要版本、变换假设、闸门排除原因、候选评分、裕量、双向排名、局部BREP读取范围、预算消耗、拓扑谱系、D值、语义聚合和双侧定位。候选保留率为rho_pair=|C_gate|/(|U_A^l|·|U_B^l|)，1-rho_pair表示相对于实体笛卡尔积的候选压缩比例。局部BREP读取率为rho_brep=B_local/(B_full,A+B_full,B)，其中B_local为本次比较实际读取的局部BREP数据量，B_full,A和B_full,B分别为模型A和模型B完整BREP数据量的基准值，三者采用同一计量单位。B_full可由元数据、离线标定、历史基准或采样估计获得，不要求本次完整读取；任一指标分母为0时指标不适用。",
        "一个可复算数量实施例中，示例参数仅用于复算，不构成对实际性能限定。模型A有100个实体，模型B有105个实体，理论实体笛卡尔积为10,500对；局部索引产生310对，闸门后C_gate保留126对，故rho_pair=126/10,500=1.2%，即rho_pair=1.2%。消歧和局部验证得到94个R_11、1个分裂关系消耗模型A侧1个、模型B侧2个端点、1个合并关系消耗模型A侧2个、模型B侧1个端点，因此两侧均为97个已跟踪端点。REVIEW为空时，DEL=100-97=3、ADD=105-97=8，即DEL=3、ADD=8；94个R_11中86个保持不变、8个修改。若两个完整模型BREP基准量分别为80MiB和88MiB，本次比较实际读取B_local=8.4MiB，则rho_brep=8.4/(80+88)=5%，即rho_brep=5%；该数值示例用于说明指标计算口径。",
        "场景处理与降级之一：刚体场景在刚体假设通过后归一化并继续闸门；单位变化在单位合法时按比例归一化，单位不明时保留多个尺度候选；未知仿射记录全局差异并继续可验证局部定位；对称/主轴不唯一时保留多个变换分支直至局部关系消歧；镜像场景在业务允许时归一化，在禁止时记录全局镜像差异且不自动放弃局部对应。",
        "场景处理与降级之二：修剪面分裂按R_1m验证新模型新增内部分界、区域覆盖、度量守恒和邻接证据；修剪面合并/几何修复按R_m1验证旧模型内部分界在新模型中删除、区域覆盖、度量守恒和邻接证据，证据不足则进入REVIEW；周期曲面缝边和解析曲面与NURBS等价优先按R_11验证表达差异；孔/倒角/圆角在实体关系稳定后聚合为语义变化；装配重复发生与重排使用发生路径、上下文、多重集和端点互斥消歧。",
        "场景处理与降级之三：摘要缺失进入低信息候选并按需读取局部BREP；BREP读取失败/预算超限时保护REVIEW端点或执行受控回退；非流形/无效拓扑单独标记数据质量证据并限制邻域扩张；整体不相关时输出低全局相关性，不强制建立实体对应；低置信REVIEW保留候选、失败原因、已用预算和原模型定位，等待后续数据或人工复核。",
    ],
    "advantages": [
        "通过实体候选和排除关系构造稀疏图，实现候选对压缩，并以rho_pair记录候选保留率、以1-rho_pair记录候选压缩比例。",
        "由判定需要和预算共同驱动按需局部读取BREP，以rho_brep记录实际局部读取量与完整数据量基准的比例。",
        "通过R_11、R_1m和R_m1及区域覆盖、度量守恒和邻接证据表达跨内核拓扑谱系，区分修改、分裂、合并与等价表达。",
        "摘要缺失、预算超限或证据冲突时保护REVIEW端点，做到低信息不强制定性，减少将未知错误归入新增或删除的风险。",
        "输出实体级和语义级双层定位，并保存两个原模型中的双侧定位、闸门、候选、预算和验证证据，形成可审计指标。",
    ],
    "key_points": [
        "在无可比较永久命名的实体域上，由顺序摘要闸门输出实体候选关系和排除关系，并构造C_gate稀疏候选图。",
        "以最佳/次佳双向裕量、双向排名、端点互斥和局部关系一致性完成候选消歧，而不把摘要相似直接认定为最终相同。",
        "采用判定需要驱动、受实体数、字节数、查询次数和邻域阶数预算约束的局部BREP验证，并对失败或超限执行REVIEW保护。",
        "以R_11、R_1m和R_m1表达一对一、分裂和合并谱系，以区域覆盖、度量守恒、支承几何和邻接证据约束组关系。",
        "在关系端点占用后执行REVIEW保护闭包，再计算新增和删除，并以D和tau_same区分R_11中的保持不变与修改。",
        "沿实体依赖和装配上下文执行语义聚合，保留两个原模型中的双侧定位以及rho_pair和rho_brep等证据指标。",
    ],
}


PERMANENT = {
    "short_title": "基于永久命名的CAD模型差异对比方法",
    "full_title": "一种基于CAD永久命名的模型差异对比方法及系统",
    "output": OUT_DIR / "永久命名-CAD模型差异对比-专利技术交底书-权利要求增强版.docx",
    "figures": [
        (ASSET_DIR / "permanent_naming-fig1-method.png", "图1 基于永久命名的CAD模型差异对比方法流程图"),
        (ASSET_DIR / "permanent_naming-fig2-system.png", "图2 基于永久命名的CAD模型差异对比系统结构图"),
        (ASSET_DIR / "permanent_naming-fig3-decision.png", "图3 命名映射与命名异常闭环处理流程图"),
        (ASSET_DIR / "permanent_naming-fig4-metrics-diff.png", "图4 命名域指标计算与模型具体差异计算流程图"),
    ],
    "background": [
        "同一CAD系统或同一产品数据管理链路中，CAD实体通常具有永久命名标识。直接利用永久命名可以减少全量几何匹配，但永久命名本身并不能说明同名实体是否发生几何、拓扑、参数或语义变化。",
        "实际编辑过程中可能出现命名复用、命名冲突、命名失效、命名漂移、命名分裂或命名合并。若差异对比系统直接信任永久命名，可能将真实修改误判为未变化，或者将命名异常误判为新增/删除。",
        "现有永久命名相关技术更多关注如何生成或维持永久命名，而本方案关注在差异对比流程中如何读取既有命名索引、审计其可信度，并在异常区域通过摘要辅助匹配和局部验证避免差异误判。",
    ],
    "solution": [
        "本方案首先读取由CAD系统、CAD内核或产品数据管理系统预先提供的永久命名实体索引，并判断两个索引是否属于同一命名域，或者是否存在已建立的统一命名域映射关系。",
        "命名域可信度判断：读取命名域标识、CAD系统标识、CAD内核标识、模型版本标识、映射表版本和映射来源，计算永久命名标识覆盖率、永久命名标识唯一性、永久命名冲突率和映射覆盖率等指标。若可信度满足阈值，则进入基于永久命名标识的差异对比流程；若不满足，则拒绝进入该流程，或仅对满足映射置信度的局部实体集合进入流程。",
        "命名域指标的计算口径如下：先按预先选择的实体层级（例如体、面、边、特征或参数）形成可审计实体集合U，并对永久命名索引中的原始标识执行格式校验、命名域拼接和映射版本归一化，形成规范命名键集合K。模型A或模型B的命名覆盖率分别为C_A=|K_A|/|U_A|、C_B=|K_B|/|U_B|；两模型共有命名覆盖率为C_pair=2|K_A∩K_B|/(|K_A|+|K_B|)。其中，分母只统计纳入本次对比范围且实体类型、几何有效性和索引可读性满足条件的实体，不把被明确排除的辅助显示对象计入分母。",
        "永久命名一致性用于判断同一规范命名键在两个模型中是否仍指向同一设计实体。对共有命名键k建立实体对(e_A,k,e_B,k)，计算实体身份一致性I_k、上下文一致性H_k和关系一致性R_k，并以预设权重得到q_k=w_I I_k+w_H H_k+w_R R_k，其中w_I+w_H+w_R=1。模型对比的一致性率为Q_cons=Σ(q_k·1[q_k≥τ_cons])/|K_A∩K_B|；采用业务权重时，将分母和分子中的1替换为实体权重。允许参数、尺寸和几何数值发生变化，但实体类型、语义类别、父级特征和拓扑角色发生不允许的跳转时，q_k降低并进入冲突审计，而不直接认定为普通修改。",
        "永久命名冲突率至少分为三类计算：同一模型内一个命名键对应多个实体的内部冲突率C_intra=Σ_k max(0,m_k-1)/N_valid，其中m_k为命名键k在单模型中的实体数，N_valid为有效命名实体数；两模型中同一命名键对应的实体身份或关系不一致的跨版本冲突率C_cross=|K_conflict|/|K_A∩K_B|；存在映射表时，多个源键映射到多个互不一致目标键的映射冲突率C_map=|M_conflict|/|M_valid|。综合冲突率可按业务权重计算C_conf=w_1C_intra+w_2C_cross+w_3C_map，缺少某类数据时对已计算类别重新归一化。普通的尺寸或参数变化只有在身份一致性通过且属于允许变化集合时才计为模型修改，不计入命名冲突率。",
        "命名域门控可采用C_A≥τ_cov、C_B≥τ_cov、C_pair≥τ_pair、Q_cons≥τ_cons且C_conf≤τ_conf的联合条件；跨CAD系统时还要求映射覆盖率C_mapcov=|K_mapped|/|K_valid|≥τ_map。若整体条件不满足，则输出命名域不可用及其证据；若仅部分实体满足条件，则按实体级q_k和映射置信度生成可用子集，对可用子集执行命名集合运算，其余实体转入摘要辅助匹配和局部验证。",
        "在满足命名域条件后，系统基于永久命名标识执行集合运算，识别新增实体、删除实体和共有实体。集合运算结果作为后续命名一致性审计、版本指纹比较和命名异常闭环处理的基础输入。",
        "命名一致性审计：输入包括永久命名标识、实体类型、父子层级、所属特征、拓扑邻接关系、装配约束关系、版本指纹、实体摘要和映射置信度。审计规则包括：同名实体的摘要或版本指纹差异超过阈值；同名实体的父子层级、所属特征、邻接关系或语义结构发生不满足业务规则的跳转；单侧缺失命名实体但另一模型存在摘要相似且局部关系可信的候选实体；一个命名实体对应多个候选实体或多个命名实体对应一个候选实体组。审计输出为命名正常实体集合、命名异常实体集合、异常类型和置信度。",
        "命名审计评分和异常队列：系统可根据实体摘要一致性、版本指纹一致性、父子层级一致性、拓扑邻接一致性、语义标签一致性、映射置信度和数据来源可信度计算命名审计评分。高于第一阈值的实体进入命名正常集合；介于第一阈值和第二阈值之间的实体进入命名异常集合并执行摘要辅助匹配；低于第二阈值的实体输出为命名异常待复核。命名异常队列记录永久命名标识、异常类型、触发规则、异常前后摘要、替代候选列表、审计评分和局部验证状态。",
        "命名异常闭环：命名异常实体集合中的异常类型至少包括命名复用或冲突、命名失效或疑似重命名、命名漂移、命名分裂和命名合并。系统不直接根据永久命名得出最终差异结论，而是将异常实体转入替代候选流程；根据后续验证结果，将异常实体重分类为新增、删除、修改、重命名、分裂、合并或命名异常待复核。",
        "摘要辅助匹配：对命名异常实体生成替代候选集合。可使用几何摘要、拓扑摘要、参数摘要、语义摘要、关系摘要和外包络摘要中的至少两类摘要；其中摘要用于候选生成和一致性校验，不作为永久命名生成规则，也不作为最终相似判定或最终差异判定的单一依据。候选集合可按实体类型、所属特征、空间邻域、拓扑邻接关系、语义标签和参数摘要检索，并计算候选匹配评分，按评分排序后保留满足条件的替代候选。",
        "局部验证：验证对象包括候选变化实体、命名异常待验证实体以及摘要辅助匹配得到的替代候选集合。验证内容包括局部几何验证、局部拓扑验证、参数验证、语义验证和装配约束验证。对于命名正常且版本指纹一致的共有实体，系统跳过完整几何验证；对于命名异常实体，仅在替代候选范围内执行局部验证。",
        "模型具体差异计算：对集合运算得到的共有实体对及异常候选对，分别计算几何差异d_g、拓扑差异d_t、参数差异d_p、语义差异d_s和依赖关系差异d_r，并归一化到[0,1]。数值属性a、b的差异可采用d_num(a,b)=min(|a-b|/max(|a|,|b|,ε),1)；实体集合或邻接集合的差异可采用d_set(X,Y)=1-|X∩Y|/max(|X∪Y|,1)；类型、语义标签或关系类型不一致时产生离散差异项。综合差异分数为D=w_gd_g+w_td_t+w_pd_p+w_sd_s+w_rd_r，其中各权重非负且总和为1。",
        "差异分类流程：若命名键仅存在于第二模型，则先标记为新增候选；若仅存在于第一模型，则标记为删除候选；若命名键在两模型中共有且q_k≥τ_cons，则比较版本指纹和D，D≤τ_same时判定为保持不变，D>τ_same时判定为修改。对命名异常实体，构造以实体摘要和关系摘要为边权的候选关系图：一对一高置信度关系判定为重命名或修改，一对多判定为分裂，多对一判定为合并，未达到τ_verify的关系保留为命名异常待复核。最后沿特征生成、参数引用、拓扑邻接和装配约束关系聚合为设计语义差异，并保留各项分数和验证证据。",
        "一个数值实施例：在选定的面、边和特征层级中，模型A有100个可审计实体、96个有效命名键，模型B有105个可审计实体、100个有效命名键，共有命名键92个，其中89个通过身份、上下文和关系一致性审计，3个进入跨版本冲突集合，且不存在同一模型内重复命名。则C_A=96%、C_B约为95.24%、C_pair约为93.88%、Q_cons约为96.74%、C_intra=0、C_cross约为3.26%。当阈值分别取τ_cov=90%、τ_pair=90%、τ_cons=95%、τ_conf=5%时，该命名域允许进入命名集合运算；3个冲突键进入异常闭环，未被冲突影响的共有实体按版本指纹和D值继续分类。",
        "差异证据记录和增量缓存：系统生成差异证据记录，记录集合运算结果、版本指纹比较结果、命名一致性审计结果、摘要辅助匹配结果、局部验证结果、重分类结果和语义聚合结果。系统还可缓存永久命名实体索引、实体摘要、版本指纹、命名一致性审计结果和命名异常处理结果，以支持增量版本对比。",
        "语义聚合：系统沿实体依赖关系执行影响传播，将实体级新增、删除、修改、重命名、分裂、合并和命名异常待复核结果聚合为孔变化、倒角变化、圆角变化、阵列变化、槽变化、草图变化、参数变化、装配约束变化等设计语义级差异，并输出关联实体、关联特征、关联参数、装配约束、制造特征、检测标注及影响范围。",
    ],
    "advantages": [
        "将全量几何实体匹配转化为命名集合运算、命名一致性审计、指纹筛选和候选局部验证。",
        "不保护永久命名生成机制，避免与已有永久命名生成/维护方案正面重叠。",
        "能够识别命名复用、命名冲突、命名失效、命名漂移、命名分裂和命名合并，降低错误命名导致的差异误判。",
        "版本指纹相同且命名审计正常的共有实体可跳过完整几何验证，提高同一命名域模型版本对比效率。",
        "通过设计语义聚合和影响传播，输出更适合工程审查的差异结果。",
    ],
    "key_points": [
        "永久命名实体索引作为既有输入被读取，不生成或更新CAD系统的永久命名规则。",
        "只有同一命名域或已有统一命名域映射关系、且命名域可信度满足预设条件时，才进入命名集合运算流程。",
        "命名域可信度可由永久命名标识覆盖率、唯一性、冲突率、映射覆盖率和映射置信度等指标确定。",
        "覆盖率的分母为本次选定实体层级中可审计实体总数；一致性率以两模型共有规范命名键为分母，按实体身份、上下文和关系一致性评分计算；冲突率分别统计单模型重复命名、跨版本身份冲突和映射多对多冲突，并可按业务权重合成为综合冲突率。",
        "命名域门控采用覆盖率、一致性率和冲突率的联合阈值；未通过整体门控时，只对通过实体级评分和映射置信度的子集执行命名比较，其余实体回退到摘要辅助匹配。",
        "命名一致性审计以永久命名标识、版本指纹、实体摘要、局部拓扑关系、语义结构和依赖关系为输入，输出命名正常实体集合、命名异常实体集合、异常类型和置信度。",
        "命名审计评分用于区分命名正常、命名异常待验证和命名异常待复核实体，异常队列记录触发规则、替代候选、评分和验证状态。",
        "命名异常闭环覆盖命名复用/冲突、命名失效/疑似重命名、命名漂移、命名分裂和命名合并，并将异常实体重分类为新增、删除、修改、重命名、分裂、合并或待复核。",
        "摘要辅助匹配使用几何摘要、拓扑摘要、参数摘要、语义摘要、关系摘要和外包络摘要生成替代候选集合，并可按候选匹配评分排序。",
        "局部验证仅作用于候选变化实体、命名异常待验证实体和替代候选集合；命名正常且版本指纹一致的共有实体跳过完整几何验证。",
        "差异证据记录保存集合运算、审计、匹配、验证、重分类和语义聚合过程，便于审查、追溯和人工复核。",
        "增量缓存保存永久命名索引、实体摘要、版本指纹、审计结果和异常处理结果，以支持后续版本快速对比。",
        "语义聚合将实体级差异、命名异常处理结果和影响传播结果输出为工程可理解的设计语义级差异。",
    ],
}


def build_disclosure(material: dict) -> Path:
    validate_figure_paths(material)
    doc = Document(TEMPLATE)
    validate_template_structure(doc)
    for rel_id, relationship in list(doc.part.rels.items()):
        if relationship.reltype == RT.IMAGE:
            doc.part.drop_rel(rel_id)
    fill_metadata(doc, material["short_title"])

    detail_table = doc.tables[1]
    fill_row(detail_table.rows[0].cells[0], "专利申请名称", [material["full_title"]])
    fill_row(detail_table.rows[1].cells[0], "同类技术的发展状况、缺陷和不足", material["background"])
    fill_row(detail_table.rows[2].cells[0], "专利技术方案", material["solution"])
    fill_row(
        detail_table.rows[3].cells[0],
        "附图说明",
        [
            "本交底书建议配套以下附图，附图用于说明方法流程、系统模块和关键判断流程。"
        ],
    )
    add_figures(detail_table.rows[3].cells[0], material["figures"])
    fill_row(detail_table.rows[4].cells[0], "本专利技术的优点和有益效果", bullets=material["advantages"])
    fill_row(detail_table.rows[5].cells[0], "需要保护的关键创新点", bullets=material["key_points"])

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    doc.save(material["output"])
    return material["output"]


def main() -> None:
    if not TEMPLATE.exists():
        raise FileNotFoundError(TEMPLATE)
    for material in [LIGHTWEIGHT, PERMANENT]:
        print(build_disclosure(material))


if __name__ == "__main__":
    main()
