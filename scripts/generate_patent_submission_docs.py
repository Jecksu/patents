from __future__ import annotations

import re
import textwrap
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
PATENT_DIR = ROOT / "docs" / "patent"
OUT_DIR = ROOT / "docs" / "submission"
ASSET_DIR = OUT_DIR / "assets"

CANDIDATE_CLOSURE_LABELS = (
    "摘要闸门后的稀疏候选边集合 C_gate",
    "候选裕量 + 双向排名 + 端点互斥\n竞争子图消歧",
    "预算受限局部BREP验证\n实体数/字节数/查询次数/邻域阶数",
    "形成 R_11/R_1m/R_m1\n一对一 / 分裂 / 合并关系",
    "有效关系端点占用 + REVIEW保护\n信息不足端点不进入确定性ADD/DEL",
    "差异闭包：COMMON_11 → 保持不变 / 修改\n未占用且非REVIEW端点 → DEL / ADD",
    "设计语义聚合 + 第一/第二CAD模型双侧定位",
)

SCENARIO_ROWS = (
    ("刚体/单位变化", "验证刚体或合法单位因子 → 归一化继续实体比较"),
    ("未知仿射/镜像", "记录全局差异继续局部定位 → 保留可验证关系"),
    ("对称多假设", "候选图与关系一致性 → 消歧或REVIEW"),
    ("修剪面分裂/合并", "读取候选组局部BREP → 组关系验证并形成R_1m/R_m1"),
    ("周期缝边/解析-NURBS", "循环对齐及几何不变量 → 等价表达验证"),
    ("数据缺失/读取失败/无效拓扑", "保存已有证据 → REVIEW或受控回退"),
    ("整体不相关", "输出全局证据+分区排除；不以一次粗比较替代确定性实体证据"),
)


MATERIALS = [
    {
        "key": "lightweight_strengthened",
        "source": PATENT_DIR / "2026-08-06-cad-model-diff-lightweight-summary-strengthened-patent-draft.md",
        "title": "一种面向跨CAD内核的基于轻量摘要闸门和受限局部BREP验证的CAD模型差异对比方法及系统",
        "output": OUT_DIR / "跨CAD内核轻量化摘要-CAD模型差异快速筛选与对比-发明专利申请文件-强化版.docx",
        "layout": "strengthened",
        "figure_kinds": ("method", "system", "decision", "candidate_closure", "scenarios"),
        "abstract_figure": [
            "审计入口：无可比较永久命名",
            "提取至少两类轻量摘要：全局/局部/关系/语义",
            "生成并验证多变换假设：刚体/单位/镜像/仿射/对称",
            "至少两个摘要闸门形成候选关系与排除关系",
            "稀疏候选图：裕量/双向排名/端点互斥消歧",
            "预算受限局部BREP按需验证",
            "形成R_11/R_1m/R_m1关系",
            "有效关系占用与REVIEW保护后的关系闭包",
            "语义聚合并输出双侧定位",
        ],
        "system_modules": [
            "审计入口模块",
            "轻量摘要提取模块",
            "变换假设模块",
            "摘要闸门模块",
            "候选关系消歧模块",
            "局部BREP按需读取模块",
            "关系/谱系与差异闭包模块",
            "语义聚合与双侧定位模块",
        ],
        "decision_flow": [
            "多变换假设生成",
            "刚体/单位/镜像",
            "仿射/对称多假设",
            "验证并保留多假设",
            "至少两个摘要闸门形成候选关系和排除关系",
            "全局差异不默认终止，继续局部定位",
        ],
        "figure_titles": [
            "图1 方法闭环流程示意图",
            "图2 系统模块与比较会话数据流示意图",
            "图3 多变换假设与摘要闸门示意图",
            "图4 候选关系消歧、局部验证和差异闭包示意图",
            "图5 场景处理与降级示意图",
        ],
    },
    {
        "key": "permanent_naming",
        "source": PATENT_DIR / "2026-07-07-cad-model-diff-permanent-naming-core-patent-draft.md",
        "title": "一种基于CAD永久命名的模型差异对比方法及系统",
        "output": OUT_DIR / "永久命名-CAD模型差异对比-发明专利申请文件.docx",
        "layout": "standard",
        "figure_kinds": ("method", "system", "decision", "metrics_diff"),
        "abstract_figure": [
            "获取第一/第二CAD模型",
            "读取已存在命名索引",
            "判断命名域/映射",
            "执行命名集合运算",
            "命名一致性审计",
            "异常闭环/指纹筛选",
            "局部验证与依赖传播",
            "输出语义差异结果",
        ],
        "system_modules": [
            "模型获取模块",
            "永久命名实体索引模块",
            "命名域判断模块",
            "命名集合比较模块",
            "命名一致性审计模块",
            "指纹比较模块",
            "命名异常闭环处理模块",
            "局部验证模块",
            "语义差异聚合模块",
            "永久命名映射模块",
            "差异输出模块",
        ],
        "decision_flow": [
            "读取已有永久命名",
            "同一命名域",
            "已有跨系统映射",
            "命名异常",
            "集合比较",
            "异常闭环重分类",
        ],
        "figure_titles": [
            "图1 基于永久命名的CAD模型差异对比方法流程图",
            "图2 基于永久命名的CAD模型差异对比系统结构图",
            "图3 命名映射与命名异常处理流程图",
            "图4 命名域指标计算与模型具体差异计算流程图",
        ],
    },
]


def font_path() -> str | None:
    for candidate in [
        Path("C:/Windows/Fonts/simsun.ttc"),
        Path("C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf"),
    ]:
        if candidate.exists():
            return str(candidate)
    return None


def get_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    path = font_path()
    if path:
        return ImageFont.truetype(path, size=size, index=1 if bold and path.endswith(".ttc") else 0)
    return ImageFont.load_default()


def wrap_label(label: str, width: int = 12) -> str:
    parts = []
    line = ""
    for ch in label:
        line += ch
        if len(line) >= width:
            parts.append(line)
            line = ""
    if line:
        parts.append(line)
    return "\n".join(parts)


def draw_centered_text(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], text: str, font) -> None:
    lines = text.splitlines()
    heights = []
    widths = []
    for line in lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        widths.append(bbox[2] - bbox[0])
        heights.append(bbox[3] - bbox[1])
    total_h = sum(heights) + 5 * (len(lines) - 1)
    y = box[1] + ((box[3] - box[1]) - total_h) / 2
    for line, w, h in zip(lines, widths, heights):
        x = box[0] + ((box[2] - box[0]) - w) / 2
        draw.text((x, y), line, fill="black", font=font)
        y += h + 5


def arrow(draw: ImageDraw.ImageDraw, start: tuple[int, int], end: tuple[int, int]) -> None:
    draw.line([start, end], fill="black", width=3)
    x1, y1 = start
    x2, y2 = end
    if y2 > y1:
        points = [(x2, y2), (x2 - 8, y2 - 14), (x2 + 8, y2 - 14)]
    elif y2 < y1:
        points = [(x2, y2), (x2 - 8, y2 + 14), (x2 + 8, y2 + 14)]
    elif x2 > x1:
        points = [(x2, y2), (x2 - 14, y2 - 8), (x2 - 14, y2 + 8)]
    else:
        points = [(x2, y2), (x2 + 14, y2 - 8), (x2 + 14, y2 + 8)]
    draw.polygon(points, fill="black")


def make_vertical_flow(path: Path, title: str, steps: list[str], strengthened: bool = False) -> None:
    w = 1500
    if strengthened:
        box_w, box_h = 1100, 155
        gap = 45
        h = max(1900, 260 + len(steps) * (box_h + gap))
        box_font = get_font(29)
        label_width = 20
        start_y = 170
    else:
        box_w, box_h = 900, 130
        gap = 52
        h = 1900
        box_font = get_font(34)
        label_width = 13
        start_y = 180
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    title_font = get_font(44, bold=True)
    draw_centered_text(draw, (80, 40, w - 80, 120), title, title_font)
    x = (w - box_w) // 2
    y = start_y
    boxes = []
    for step in steps:
        box = (x, y, x + box_w, y + box_h)
        boxes.append(box)
        draw.rounded_rectangle(box, radius=8, outline="black", width=3, fill="white")
        draw_centered_text(draw, box, wrap_label(step, label_width), box_font)
        y += box_h + gap
    for a, b in zip(boxes, boxes[1:]):
        arrow(draw, ((a[0] + a[2]) // 2, a[3]), ((b[0] + b[2]) // 2, b[1]))
    img.save(path)


def make_module_grid(path: Path, title: str, modules: list[str], strengthened: bool = False) -> None:
    w, h = 1700, 1300
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    title_font = get_font(44, bold=True)
    box_font = get_font(30)
    draw_centered_text(draw, (80, 40, w - 80, 120), title, title_font)
    cols = 3
    box_w, box_h = 460, 140
    gap_x, gap_y = 80, 85
    start_x = (w - cols * box_w - (cols - 1) * gap_x) // 2
    start_y = 210
    boxes = []
    for idx, module in enumerate(modules):
        row, col = divmod(idx, cols)
        display_col = col if not strengthened or row % 2 == 0 else cols - 1 - col
        x = start_x + display_col * (box_w + gap_x)
        y = start_y + row * (box_h + gap_y)
        box = (x, y, x + box_w, y + box_h)
        boxes.append(box)
        draw.rounded_rectangle(box, radius=8, outline="black", width=3, fill="white")
        draw_centered_text(draw, box, wrap_label(module, 10), box_font)
    if strengthened:
        for current, following in zip(boxes, boxes[1:]):
            current_y = (current[1] + current[3]) // 2
            following_y = (following[1] + following[3]) // 2
            if current_y == following_y:
                if following[0] > current[0]:
                    arrow(draw, (current[2], current_y), (following[0], following_y))
                else:
                    arrow(draw, (current[0], current_y), (following[2], following_y))
            else:
                arrow(
                    draw,
                    ((current[0] + current[2]) // 2, current[3]),
                    ((following[0] + following[2]) // 2, following[1]),
                )
    img.save(path)


def make_decision_flow(path: Path, title: str, labels: list[str], strengthened: bool = False) -> None:
    w, h = (1900, 1350) if strengthened else (1700, 1200)
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    title_font = get_font(44, bold=True)
    box_font = get_font(27 if strengthened else 30)
    draw_centered_text(draw, (80, 40, w - 80, 120), title, title_font)
    if strengthened:
        boxes = {
            "top": (450, 170, 1450, 350),
            "left": (100, 500, 600, 680),
            "mid": (700, 500, 1200, 680),
            "right": (1300, 500, 1800, 680),
            "bottom_left": (260, 880, 960, 1070),
            "bottom_right": (1040, 880, 1740, 1070),
        }
        label_width = 15
    else:
        boxes = {
            "top": (620, 180, 1080, 320),
            "left": (160, 470, 550, 610),
            "mid": (655, 470, 1045, 610),
            "right": (1150, 470, 1540, 610),
            "bottom_left": (390, 800, 780, 940),
            "bottom_right": (930, 800, 1320, 940),
        }
        label_width = 10
    for box, label in zip(boxes.values(), labels):
        draw.rounded_rectangle(box, radius=8, outline="black", width=3, fill="white")
        draw_centered_text(draw, box, wrap_label(label, label_width), box_font)
    top_c = ((boxes["top"][0] + boxes["top"][2]) // 2, boxes["top"][3])
    for name in ["left", "mid", "right"]:
        target = ((boxes[name][0] + boxes[name][2]) // 2, boxes[name][1])
        arrow(draw, top_c, target)
    for name in ["left", "mid"]:
        arrow(draw, ((boxes[name][0] + boxes[name][2]) // 2, boxes[name][3]), ((boxes["bottom_left"][0] + boxes["bottom_left"][2]) // 2, boxes["bottom_left"][1]))
    arrow(draw, ((boxes["right"][0] + boxes["right"][2]) // 2, boxes["right"][3]), ((boxes["bottom_right"][0] + boxes["bottom_right"][2]) // 2, boxes["bottom_right"][1]))
    img.save(path)


def make_metrics_diff_flow(path: Path, title: str) -> None:
    """Render the auditable namespace gate and entity-difference classification flow."""
    w, h = 1900, 1750
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    title_font = get_font(42, bold=True)
    box_font = get_font(27)
    small_font = get_font(23)
    draw_centered_text(draw, (70, 30, w - 70, 100), title, title_font)

    left_x, right_x = 90, 1030
    box_w, box_h = 760, 150
    y_values = [160, 390, 620, 850, 1080, 1310]
    left_labels = [
        "选定实体层级并形成可审计集合 U",
        "规范化命名键集合 K，计算覆盖率",
        "计算一致性率 Q_cons 与冲突率 C_conf",
        "联合阈值门控：覆盖率 / 一致性 / 冲突率",
        "通过：命名集合运算；未通过：局部子集或摘要回退",
        "输出命名域可用性及指标证据",
    ]
    right_labels = [
        "共有键建立实体对，单侧键形成新增 / 删除候选",
        "计算 d_g、d_t、d_p、d_s、d_r 并归一化",
        "综合差异 D = Σ w_i d_i",
        "D ≤ τ_same：保持不变；D > τ_same：修改",
        "异常候选：一对一 / 一对多 / 多对一",
        "重命名、修改、分裂、合并或待复核",
    ]

    def draw_column(x: int, labels: list[str]) -> list[tuple[int, int, int, int]]:
        boxes = []
        for y, label in zip(y_values, labels):
            box = (x, y, x + box_w, y + box_h)
            boxes.append(box)
            draw.rounded_rectangle(box, radius=8, outline="black", width=3, fill="white")
            draw_centered_text(draw, box, wrap_label(label, 18), box_font)
        for a, b in zip(boxes, boxes[1:]):
            arrow(draw, ((a[0] + a[2]) // 2, a[3]), ((b[0] + b[2]) // 2, b[1]))
        return boxes

    left_boxes = draw_column(left_x, left_labels)
    right_boxes = draw_column(right_x, right_labels)
    arrow(
        draw,
        (left_boxes[4][2], (left_boxes[4][1] + left_boxes[4][3]) // 2),
        (right_boxes[0][0], (right_boxes[0][1] + right_boxes[0][3]) // 2),
    )
    draw.text((740, 525), "命名键 / 映射可用", fill="black", font=small_font)
    draw.text((790, 1135), "指标证据", fill="black", font=small_font)
    img.save(path)


def make_candidate_closure_flow(path: Path, title: str) -> None:
    """Render candidate disambiguation, local validation, and relation closure."""
    w, h = 1900, 1950
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    title_font = get_font(42, bold=True)
    box_font = get_font(28)
    draw_centered_text(draw, (70, 30, w - 70, 105), title, title_font)

    box_w, box_h = 1440, 175
    x = (w - box_w) // 2
    y = 145
    gap = 80
    boxes = []
    for label in CANDIDATE_CLOSURE_LABELS:
        box = (x, y, x + box_w, y + box_h)
        boxes.append(box)
        draw.rounded_rectangle(box, radius=10, outline="black", width=3, fill="white")
        draw_centered_text(draw, box, wrap_label(label, 25), box_font)
        y += box_h + gap
    for current, following in zip(boxes, boxes[1:]):
        arrow(
            draw,
            ((current[0] + current[2]) // 2, current[3]),
            ((following[0] + following[2]) // 2, following[1]),
        )
    img.save(path)


def make_scenario_flow(path: Path, title: str) -> None:
    """Render scenario-specific processing and controlled degradation paths."""
    w, h = 2100, 1950
    img = Image.new("RGB", (w, h), "white")
    draw = ImageDraw.Draw(img)
    title_font = get_font(42, bold=True)
    header_font = get_font(29, bold=True)
    box_font = get_font(25)
    draw_centered_text(draw, (70, 25, w - 70, 95), title, title_font)
    draw_centered_text(draw, (80, 105, 790, 155), "输入场景", header_font)
    draw_centered_text(draw, (1070, 105, 2020, 155), "处理路径与输出/降级", header_font)

    left_x, left_w = 80, 710
    right_x, right_w = 1070, 950
    box_h, gap = 170, 70
    y = 180
    for scenario, action in SCENARIO_ROWS:
        left_box = (left_x, y, left_x + left_w, y + box_h)
        right_box = (right_x, y, right_x + right_w, y + box_h)
        draw.rounded_rectangle(left_box, radius=10, outline="black", width=3, fill="white")
        draw.rounded_rectangle(right_box, radius=10, outline="black", width=3, fill="white")
        draw_centered_text(draw, left_box, wrap_label(scenario, 18), box_font)
        draw_centered_text(draw, right_box, wrap_label(action, 25), box_font)
        arrow(draw, (left_box[2], (left_box[1] + left_box[3]) // 2), (right_box[0], (right_box[1] + right_box[3]) // 2))
        y += box_h + gap
    img.save(path)


def render_method(path: Path, title: str, material: dict) -> None:
    make_vertical_flow(
        path,
        title,
        material["abstract_figure"],
        strengthened=material["layout"] == "strengthened",
    )


def render_system(path: Path, title: str, material: dict) -> None:
    make_module_grid(
        path,
        title,
        material["system_modules"],
        strengthened=material["layout"] == "strengthened",
    )


def render_decision(path: Path, title: str, material: dict) -> None:
    make_decision_flow(
        path,
        title,
        material["decision_flow"],
        strengthened=material["layout"] == "strengthened",
    )


def render_metrics_diff(path: Path, title: str, material: dict) -> None:
    make_metrics_diff_flow(path, title)


def render_candidate_closure(path: Path, title: str, material: dict) -> None:
    make_candidate_closure_flow(path, title)


def render_scenarios(path: Path, title: str, material: dict) -> None:
    make_scenario_flow(path, title)


FIGURE_REGISTRY = {
    "method": {"number": 1, "suffix": "fig1-method.png", "renderer": render_method},
    "system": {"number": 2, "suffix": "fig2-system.png", "renderer": render_system},
    "decision": {"number": 3, "suffix": "fig3-decision.png", "renderer": render_decision},
    "metrics_diff": {"number": 4, "suffix": "fig4-metrics-diff.png", "renderer": render_metrics_diff},
    "candidate_closure": {
        "number": 4,
        "suffix": "fig4-candidate-closure.png",
        "renderer": render_candidate_closure,
    },
    "scenarios": {"number": 5, "suffix": "fig5-scenarios.png", "renderer": render_scenarios},
}

VALID_LAYOUTS = {"standard", "strengthened"}


def validate_figure_plan(material: dict) -> tuple[tuple[dict, str], ...]:
    layout = material.get("layout")
    if layout not in VALID_LAYOUTS:
        raise ValueError(f"未知 layout 布局: {layout!r}")

    kinds = material.get("figure_kinds")
    if not isinstance(kinds, (list, tuple)) or not kinds:
        raise ValueError("figure_kinds 必须是非空有序列表或元组")
    unregistered = [kind for kind in kinds if not isinstance(kind, str) or kind not in FIGURE_REGISTRY]
    if unregistered:
        raise ValueError(f"存在未注册的 figure kind: {unregistered!r}")
    duplicate_kinds = sorted({kind for kind in kinds if kinds.count(kind) > 1})
    if duplicate_kinds:
        raise ValueError(f"存在重复 figure kind: {duplicate_kinds!r}")

    registrations = [FIGURE_REGISTRY[kind] for kind in kinds]
    actual_numbers = [registration["number"] for registration in registrations]
    expected_numbers = list(range(1, len(kinds) + 1))
    if actual_numbers != expected_numbers:
        raise ValueError(
            f"figure kind 图号顺序无效: 实际 {actual_numbers!r}，预期 {expected_numbers!r}"
        )

    titles = material.get("figure_titles")
    if not isinstance(titles, (list, tuple)) or len(titles) != len(kinds):
        title_count = len(titles) if isinstance(titles, (list, tuple)) else 0
        raise ValueError(
            f"figure_titles 数量 {title_count} 与 figure_kinds 数量 {len(kinds)} 不一致"
        )
    for registration, title in zip(registrations, titles, strict=True):
        number = registration["number"]
        if not isinstance(title, str) or not re.match(rf"^图{number}(?!\d)", title):
            raise ValueError(f"图{number} 标题必须以对应图号开头: {title!r}")

    return tuple(zip(registrations, titles, strict=True))


def section_between(text: str, start: str, end: str | None = None) -> str:
    start_pat = re.escape(start)
    if end:
        end_pat = re.escape(end)
        m = re.search(start_pat + r"\n(.*?)\n" + end_pat, text, re.S)
    else:
        m = re.search(start_pat + r"\n(.*)", text, re.S)
    return m.group(1).strip() if m else ""


def clean_md_line(line: str) -> str:
    line = line.rstrip()
    line = re.sub(r"^\s*[-*]\s+", "", line)
    line = re.sub(r"^\s*#+\s*", "", line)
    return line


def set_run_font(run, name: str = "SimSun", size: int | None = None, bold: bool | None = None) -> None:
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run._element.rPr.rFonts.set(qn("w:ascii"), "Calibri")
    run._element.rPr.rFonts.set(qn("w:hAnsi"), "Calibri")
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold


def set_paragraph_spacing(paragraph, before=0, after=6, line=1.1) -> None:
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line


def configure_doc(doc: Document, title: str) -> None:
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Calibri"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")
    normal.font.size = Pt(11)

    for style_name, size, color in [
        ("Heading 1", 16, "2E74B5"),
        ("Heading 2", 13, "2E74B5"),
        ("Heading 3", 12, "1F4D78"),
    ]:
        style = styles[style_name]
        style.font.name = "Calibri"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "SimSun")
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(12)
        style.paragraph_format.space_after = Pt(6)
        style.paragraph_format.line_spacing = 1.1

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_run = footer.add_run("发明专利申请文件")
    set_run_font(footer_run, size=9)


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def add_title_page(doc: Document, title: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(18)
    r = p.add_run("发明专利申请文件")
    set_run_font(r, size=22, bold=True)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(30)
    r = p.add_run(title)
    set_run_font(r, size=18, bold=True)

    for label in ["申请人", "发明人", "代理机构", "申请日期"]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.left_indent = Cm(3)
        p.paragraph_format.space_after = Pt(12)
        r = p.add_run(f"{label}：____________________________")
        set_run_font(r, size=12)

    doc.add_page_break()


def add_directory(doc: Document) -> None:
    doc.add_heading("申请文件目录", level=1)
    entries = ["一、说明书摘要", "二、摘要附图", "三、权利要求书", "四、说明书", "五、说明书附图"]
    for entry in entries:
        p = doc.add_paragraph(entry)
        set_paragraph_spacing(p, after=6)
        for run in p.runs:
            set_run_font(run, size=11)
    doc.add_page_break()


def add_md_content(doc: Document, md: str, top_level: int = 1) -> None:
    for raw_line in md.splitlines():
        line = raw_line.rstrip()
        if not line:
            continue
        if line.startswith("#### "):
            p = doc.add_heading(clean_md_line(line), level=min(top_level + 2, 3))
        elif line.startswith("### "):
            p = doc.add_heading(clean_md_line(line), level=min(top_level + 1, 3))
        elif line.startswith("## "):
            p = doc.add_heading(clean_md_line(line), level=top_level)
        else:
            p = doc.add_paragraph()
            text = clean_md_line(line)
            run = p.add_run(text)
            set_run_font(run, size=11)
            set_paragraph_spacing(p, after=6)
            if re.match(r"^\s+\S", raw_line):
                p.paragraph_format.left_indent = Cm(0.74)
            if re.match(r"^\d+\.\s", text):
                p.paragraph_format.first_line_indent = Cm(-0.35)
                p.paragraph_format.left_indent = Cm(0.35)
        for run in p.runs:
            if p.style.name.startswith("Heading"):
                set_run_font(run, size=int(p.style.font.size.pt), bold=True)


def add_figure(doc: Document, image_path: Path, caption: str, width_in: float = 5.9) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(str(image_path), width=Inches(width_in))
    cap = doc.add_paragraph()
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cap.add_run(caption)
    set_run_font(run, size=10)
    set_paragraph_spacing(cap, after=12)


def build_material(material: dict) -> None:
    figure_plan = validate_figure_plan(material)
    source_text = material["source"].read_text(encoding="utf-8")
    abstract = section_between(source_text, "## 摘要", "## 权利要求书")
    claims = section_between(source_text, "## 权利要求书", "## 说明书")
    specification = section_between(source_text, "## 说明书")

    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    figures = []
    for registration, title in figure_plan:
        figure = ASSET_DIR / f"{material['key']}-{registration['suffix']}"
        registration["renderer"](figure, title, material)
        figures.append(figure)

    doc = Document()
    configure_doc(doc, material["title"])
    add_title_page(doc, material["title"])
    add_directory(doc)

    doc.add_heading("一、说明书摘要", level=1)
    add_md_content(doc, abstract, top_level=2)
    doc.add_page_break()

    doc.add_heading("二、摘要附图", level=1)
    add_figure(doc, figures[0], "摘要附图")
    doc.add_page_break()

    doc.add_heading("三、权利要求书", level=1)
    add_md_content(doc, claims, top_level=2)
    doc.add_page_break()

    doc.add_heading("四、说明书", level=1)
    add_md_content(doc, specification, top_level=2)
    doc.add_page_break()

    doc.add_heading("五、说明书附图", level=1)
    for index, figure in enumerate(figures):
        add_figure(doc, figure, material["figure_titles"][index])

    material["output"].parent.mkdir(parents=True, exist_ok=True)
    doc.save(material["output"])


def main() -> None:
    for material in MATERIALS:
        build_material(material)
        print(material["output"])


if __name__ == "__main__":
    main()
