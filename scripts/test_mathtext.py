"""``mathtext`` 的上下标切分回归用例。

用法::

    py -3.12 scripts/test_mathtext.py

退出码非零表示有用例未通过。用例分四组：应排为上下标的记号、希腊字母还原、
必须保持平排的状态名与字段名，以及逗号归属（词根下标跨逗号续读，实参与符号
分隔符中的逗号不得被吞入下标）。
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from mathtext import split_math

def fmt(text):
    out = ""
    for chunk, s in split_math(text):
        out += chunk if s is None else ("{" + ("_" if s == "sub" else "^") + chunk + "}")
    return out

CASES = [
    # 应当排为上下标
    ("一致性通过率R_cons以K_11为分母", "一致性通过率R{_cons}以K{_11}为分母"),
    ("|U_A^abn|与|U_B^abn|", "|U{_A}{^abn}|与|U{_B}{^abn}|"),
    ("B_edge,A→B^l", "B{_edge,A→B}{^l}"),
    ("命名覆盖率C_cov,X等于N_X的权重和", "命名覆盖率C{_cov,X}等于N{_X}的权重和"),
    ("维度子集J_κ^+", "维度子集J{_κ}{^+}"),
    ("候选配置Pi_cand^l", "候选配置Π{_cand}{^l}"),
    ("分量目标Phi_G(z)", "分量目标Φ{_G}(z)"),
    ("端点展平算子end_A和end_B", "端点展平算子end{_A}和end{_B}"),
    ("闭包快照cs_1", "闭包快照cs{_1}"),
    ("关系(a_1,b_1)；a_2在b_2与b_3之间", "关系(a{_1},b{_1})；a{_2}在b{_2}与b{_3}之间"),
    ("|C| <= min(", "|C| ≤ min("),
    # 希腊字母还原
    ("q_κ不小于tau_cons", "q{_κ}不小于τ{_cons}"),
    ("绝对裕量阈值delta_margin^l", "绝对裕量阈值δ{_margin}{^l}"),
    ("除以T与正下限delta的较大者", "除以T与正下限δ的较大者"),
    ("可复算置信元组rho(r)", "可复算置信元组ρ(r)"),
    ("合成容差eps_abs与eps_rel", "合成容差ε{_abs}与ε{_rel}"),
    ("预设正下限epsilon与假设质量", "预设正下限ε与假设质量"),
    ("严格正权重lambda_i、alpha_i、beta_j、gamma_i", "严格正权重λ{_i}、α{_i}、β{_j}、γ{_i}"),
    # 必须保持平排的标识符
    ("闭包通过标志CLOSURE_PASS失效", "闭包通过标志CLOSURE_PASS失效"),
    ("持有同代次COMPLETE_FAIL_A", "持有同代次COMPLETE_FAIL_A"),
    ("候选边比candidate_edge_ratio以", "候选边比candidate_edge_ratio以"),
    ("物化为CAD_DIFF_EVENT_V1事件", "物化为CAD_DIFF_EVENT_V1事件"),
    ("闭包快照标识closure_snapshot_id", "闭包快照标识closure_snapshot_id"),
    ("变化门控change_gate才判为通过", "变化门控change_gate才判为通过"),
    ("覆盖度coverage_change等于", "覆盖度coverage_change等于"),
    ("局部B-rep字节比local_brep_byte_ratio", "局部B-rep字节比local_brep_byte_ratio"),
    ("冻结集合FROZEN_A与FROZEN_B", "冻结集合FROZEN_A与FROZEN_B"),
    ("状态seed_status与hyp_status", "状态seed_status与hyp_status"),
    ("局部映射校验pass_map_local(κ)", "局部映射校验pass_map_local(κ)"),
    ("方向门控pass_map_dir(X)", "方向门控pass_map_dir(X)"),
    ("初选全败PRELIMINARY_ALL_FAIL", "初选全败PRELIMINARY_ALL_FAIL"),
    ("候选并列CANDIDATE_TIE", "候选并列CANDIDATE_TIE"),
    # 逗号下标只在词根下标后续读；实参与符号分隔符里的逗号不得被吞入下标。
    # 回归来源：``MANY_TO_ONE(A_h,b)`` 曾被排成 A{_h,b}，把实参 b 吞成下标。
    ("MANY_TO_ONE(A_h,b)", "MANY_TO_ONE(A{_h},b)"),
    ("ONE_TO_MANY(a,B_h)", "ONE_TO_MANY(a,B{_h})"),
    ("R_11,R_1m,R_m1", "R{_11},R{_1m},R{_m1}"),
    ("sum(lambda_i*d_i,i∈J)", "sum(λ{_i}*d{_i},i∈J)"),
    ("软证据组g_I,g_H", "软证据组g{_I},g{_H}"),
    ("键对(k_A,k_B)", "键对(k{_A},k{_B})"),
    ("w_κ=(w_a+w_b)/2", "w{_κ}=(w{_a}+w{_b})/2"),
    ("COMPLETE_FAIL_B(b_3,gkey(b_3))", "COMPLETE_FAIL_B(b{_3},gkey(b{_3}))"),
    # 词根下标仍须跨逗号续读
    ("有效映射边集合M_edge,valid", "有效映射边集合M{_edge,valid}"),
    ("冲突边集合M_edge,conf", "冲突边集合M{_edge,conf}"),
    ("映射覆盖率C_mapcov,X", "映射覆盖率C{_mapcov,X}"),
    ("端点集合V_local,A与V_local,B", "端点集合V{_local,A}与V{_local,B}"),
    ("双向映射C_mapcov,A→B", "双向映射C{_mapcov,A→B}"),
    ("未匹配令牌U_name,A", "未匹配令牌U{_name,A}"),
    # 名称对应上限与组规模上限、最大匹配集合三者字形须各自独立
    ("N_namemax与m_max^l与M_max", "N{_namemax}与m{_max}{^l}与M{_max}"),
]
bad = [(t, fmt(t), e) for t, e in CASES if fmt(t) != e]
print(f"{len(CASES) - len(bad)}/{len(CASES)} 用例通过")
for t, got, exp in bad:
    print(f"  输入 {t!r}\n  实得 {got!r}\n  期望 {exp!r}")
sys.exit(1 if bad else 0)
