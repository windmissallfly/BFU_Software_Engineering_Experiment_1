# -*- coding: utf-8 -*-
"""Generate D1-6 需求验证报告."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import _data as D

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def _set_cn_font(run, font="宋体", size_pt=10.5, bold=False):
    run.font.name = font
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), font)
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)


def H(doc, text, level=1):
    p = doc.add_heading(level=level)
    run = p.add_run(text)
    if level == 1:
        _set_cn_font(run, "黑体", 16, bold=True)
    elif level == 2:
        _set_cn_font(run, "黑体", 14, bold=True)
    else:
        _set_cn_font(run, "黑体", 12, bold=True)
    return p


def P(doc, text, indent_first=True, size=10.5):
    p = doc.add_paragraph()
    if indent_first:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    _set_cn_font(run, "宋体", size)
    return p


def BULLET(doc, text, size=10.5):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    _set_cn_font(run, "宋体", size)
    return p


def NUM(doc, text, size=10.5):
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    _set_cn_font(run, "宋体", size)
    return p


def make_table(doc, header, rows, col_widths_cm=None):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(header))
    tbl.style = "Light Grid Accent 1"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        cell = tbl.cell(0, i)
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        _set_cn_font(run, "黑体", 10, bold=True)
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            _set_cn_font(run, "宋体", 9)
    if col_widths_cm:
        for i, w in enumerate(col_widths_cm):
            for cell in tbl.columns[i].cells:
                cell.width = Cm(w)
    return tbl


def build():
    doc = Document()
    doc.styles["Normal"].font.name = "宋体"
    doc.styles["Normal"].font.size = Pt(10.5)

    # ===== 封面 =====
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(D.DOC_META["course"])
    _set_cn_font(r, "黑体", 14, bold=True)

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("需求验证报告")
    _set_cn_font(r, "黑体", 18, bold=True)

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run("Requirements Validation Report — D1-6")
    _set_cn_font(r, "Times New Roman", 12)

    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(f"—— 案例：{D.DOC_META['case']} ——")
    _set_cn_font(r, "宋体", 12)

    doc.add_paragraph()
    make_table(doc, ["项目", "内容"], [
        ["产出编号", "D1-6"],
        ["产出名称", "需求验证报告（IEEE 830 / ISO 29148 检查清单法）"],
        ["所属实验", D.DOC_META["exp"]],
        ["案例项目", D.DOC_META["case"]],
        ["版本", D.DOC_META["version"]],
        ["作者团队", D.DOC_META["team"]],
        ["验证范围", "D1-3 SRS v1.0 全部章节 + D1-2 全部 40 条需求"],
        ["验证方法", "检查清单（Checklist Inspection）+ 原型评审 + AI 反向质询"],
        ["结论", "通过 — 推荐冻结 requirements-v1.0 基线"],
        ["提交日期", D.DOC_META["submit"]],
    ], col_widths_cm=[4, 12])

    # 修订记录
    P(doc, "")
    H(doc, "修订记录", level=3)
    make_table(doc, ["版本", "日期", "作者", "变更摘要"], [
        ["v0.1", "2026-05-15", "陈思", "初稿，建立 16 项检查清单"],
        ["v0.5", "2026-05-15", "张明", "完成 D1-3 v0.5 版本逐条检查"],
        ["v0.9", "2026-05-16", "全员", "针对 SRS v0.9 复审，补充 D-007/D-008 缺陷"],
        ["v1.0", "2026-05-16", "全员", "8 条缺陷全部闭环，输出基线建议"],
    ], col_widths_cm=[2, 3, 3, 8])

    doc.add_page_break()

    # ===========================================================
    # 1 验证范围
    # ===========================================================
    H(doc, "1  验证范围", level=1)
    P(doc, "本次验证以 D1-3《软件需求规格说明书 SRS》v1.0 为主对象，并以 D1-2《需求清单与 MoSCoW 分级表》、"
           "D1-5《术语表与需求追溯矩阵》、D1-4 UML 模型集为交叉证据，目的是判断 v1.0 需求基线是否"
           "可被冻结并交付给设计/编码阶段。")

    H(doc, "1.1  覆盖的文档与条目", level=2)
    make_table(doc, ["对象", "范围", "条目数 / 章节数"], [
        ["D1-3 SRS", "§1 引言、§2 总体描述、§3 具体需求、§4 其他需求、§5 附录", "5 章 / 共约 24 节"],
        ["D1-2 需求清单", "FR-001 ~ FR-030 + NFR-001 ~ NFR-010", "40 条需求条目"],
        ["D1-4 UML 模型", "1 主用例图 + 2 子用例图 + 2 活动图 + 1 类图 + 2 顺序图 + 1 状态图", "9 张图 + 4 份说明文档"],
        ["D1-5 RTM", "Sheet Glossary 25 条术语 / Sheet RTM_FR / Sheet RTM_NFR / Sheet 覆盖度统计", "4 个工作表"],
        ["D1-7 访谈记录", "受访者 A/B/C 全部 36 题访谈对话原文", "3 轮 × 12 题"],
    ], col_widths_cm=[3, 9, 4])

    H(doc, "1.2  验证目标", level=2)
    NUM(doc, "确认每条需求都符合「正确、完整、一致、无二义、可验证、可追溯」六大质量属性；")
    NUM(doc, "确认 SRS 与 D1-2 需求清单、D1-4 UML 模型、D1-7 访谈纪要在术语与编号上完全一致；")
    NUM(doc, "标注每条需求的「可测试性来源」（即如何证明它能被测试）；")
    NUM(doc, "形成可关闭的「缺陷列表 + 处置追踪」，决定基线是否可冻结。")

    H(doc, "1.3  验证团队与时间", level=2)
    make_table(doc, ["阶段", "时间", "参与角色", "工时"], [
        ["检查清单制定", "2026-05-14 上午", "组长 + 全体", "0.5h × 4"],
        ["逐条 SRS 走查", "2026-05-15 上午", "陈思（主审）+ 张明（记录）", "2h × 2"],
        ["原型评审", "2026-05-15 下午", "全体 + AI 模拟干系人", "1h × 4"],
        ["AI 反向质询", "2026-05-15 晚", "李华（执行）", "1.5h"],
        ["缺陷处置 / 复审", "2026-05-16 上午", "全员", "2h × 4"],
        ["基线签字", "2026-05-16 下午", "组长 + 教师 / 课代表", "0.5h"],
    ], col_widths_cm=[4, 4, 5, 3])

    doc.add_page_break()

    # ===========================================================
    # 2 验证方法
    # ===========================================================
    H(doc, "2  验证方法", level=1)
    P(doc, "依据 IEEE Std 830-1998 §4 与 ISO/IEC/IEEE 29148:2018 §5.2.5 的指导，本次验证综合运用三种方法："
           "（1）检查清单法 — 主导方法；（2）原型评审 — 由小组与 AI 模拟干系人共审；（3）AI 反向质询 — 利用 LLM 对 SRS 找漏洞。")

    H(doc, "2.1  检查清单（Checklist Inspection）", level=2)
    P(doc, "本节列出对 SRS 进行检查所采用的 16 条清单；每条清单标注「通过 / 部分通过 / 不通过」与"
           "「证据/备注」；不通过或部分通过的项目在 §3 缺陷列表中详述并跟踪闭环。")
    rows = [[i + 1, c[0], c[1], c[2], c[3]] for i, c in enumerate(D.CHECKLIST)]
    make_table(doc, ["#", "维度", "检查项", "结论", "证据 / 备注"],
               rows, col_widths_cm=[1, 2.4, 5.5, 1.6, 5.5])

    P(doc, "")
    P(doc, "通过率统计：通过 13 / 16 ≈ 81.3%，部分通过 3 / 16 ≈ 18.7%，不通过 0 / 16 = 0%。"
           "所有「部分通过」项均已转化为 §3 缺陷条目并在 §4 中闭环。", indent_first=False)

    H(doc, "2.2  原型评审（Prototype Review）", level=2)
    P(doc, "工具与流程：")
    BULLET(doc, "原型工具：墨刀（顾客端 / 店员端） + Figma（总部后台）；")
    BULLET(doc, "评审会议：2026-05-15 14:00-18:00，全体组员 + AI 模拟干系人（顾客 A、店员 B、运营 C）共审；")
    BULLET(doc, "对照件：D1-3 SRS v0.9 + D1-4 用例图 + D1-2 验收准则。")
    P(doc, "评审主要发现：")
    BULLET(doc, "顾客端：FR-005 「安静桌位」在原型中无视觉提示 → 触发 D-001 二义性；")
    BULLET(doc, "店员端：高峰期换桌操作走查需 4 次点击，违反 NFR-008 → 重新调整原型为 3 次点击；")
    BULLET(doc, "总部端：FR-021 实时大屏在原型中刷新策略未注明 → 触发 D-006；")
    BULLET(doc, "管家端：FR-017 强制休息提醒原型未给出可调节时长入口 → 触发 D-003。")

    H(doc, "2.3  AI 反向质询（AI Adversarial Review）", level=2)
    P(doc, "目的：让 AI（Claude Sonnet 4.6）扮演「挑剔的评审专家」对 SRS 找漏洞，弥补人工走查的盲区。")
    P(doc, "Prompt 全文（已存档至 D1-7 附录 A.4）：")
    P(doc, "你现在是一名拥有 15 年经验的 IEEE 830 评审专家，请基于附件中的 SRS v0.9，"
           "围绕「正确性、完整性、一致性、无二义性、可验证性、可追溯性」六个维度找漏洞；"
           "对每个发现，请给出：(1) 涉及的需求 ID；(2) 漏洞类别（二义性 / 冲突 / 缺失 / 不可验证）；"
           "(3) 严重度（H/M/L）；(4) 建议处置方式。请尽可能挑剔，不要客气。", indent_first=True)
    P(doc, "AI 反向质询主要产出：")
    BULLET(doc, "新增缺陷 D-007（FR-011 全额退款 vs FR-009 第三方支付到账时差冲突）— 团队复核确认成立；")
    BULLET(doc, "新增缺陷 D-008（FR-026 评价管理未声明情感分析准确率指标）— 接受并量化；")
    BULLET(doc, "AI 还提出 5 条「值得关注」但当前不列为缺陷的提示，记录在 §5 结论的「下阶段输入」。")

    P(doc, "AI 反向质询的局限性：")
    BULLET(doc, "AI 缺乏门店一线经验，对硬件兼容性、网络稳定性等情境洞察弱（已通过店员 B 原型评审补足）；")
    BULLET(doc, "AI 倾向给出「形式合规」建议，对真实业务取舍判断有限（人工评审仍是最终裁决）。")

    doc.add_page_break()

    # ===========================================================
    # 3 缺陷列表
    # ===========================================================
    H(doc, "3  缺陷列表", level=1)
    P(doc, "本次验证累计发现 8 条缺陷，按缺陷类型分为「二义性 3 条 / 冲突 2 条 / 缺失 3 条」；按严重度"
           "分为「高 H 1 条 / 中 M 5 条 / 低 L 2 条」。所有缺陷已在 §4 闭环。")

    rows = [[d["id"], d["type"], d["req"], d["desc"], d["severity"], d["status"]]
            for d in D.DEFECTS]
    make_table(doc, ["缺陷ID", "类型", "所在需求 ID", "描述", "严重度", "状态"],
               rows, col_widths_cm=[1.5, 1.6, 2.6, 7.5, 1.4, 1.4])

    H(doc, "3.1  二义性问题", level=2)
    P(doc, "二义性指需求中存在模糊形容词或未量化的关键参数，导致开发与测试可能采用不同假设。")

    for d in D.DEFECTS:
        if d["type"] != "二义性":
            continue
        H(doc, f"{d['id']}  {d['req']}", level=3)
        P(doc, f"问题描述：{d['desc']}", indent_first=False)
        P(doc, f"分析：在 SRS v0.9 中此处描述未给出可量化的接受范围，测试团队反馈无法编写可执行用例。",
          indent_first=False)
        P(doc, f"处置建议：{d['fix']}", indent_first=False)
        P(doc, f"复审结论：通过（{d['closed']} 已合并入 SRS v1.0）", indent_first=False)

    H(doc, "3.2  冲突问题", level=2)
    P(doc, "冲突指两条或多条需求在并发场景或资源边界下相互制约，导致难以同时满足。")

    for d in D.DEFECTS:
        if d["type"] != "冲突":
            continue
        H(doc, f"{d['id']}  {d['req']}", level=3)
        P(doc, f"问题描述：{d['desc']}", indent_first=False)
        P(doc, f"分析：从 SRS §3.2/§3.3 交叉读取可见 — "
               f"该冲突在系统架构层而非简单文字层。需协调权衡或分时降级。", indent_first=False)
        P(doc, f"处置建议：{d['fix']}", indent_first=False)
        P(doc, f"复审结论：通过（{d['closed']}），相应描述已写入 SRS §3.2 与 §3.3 备注。",
          indent_first=False)

    H(doc, "3.3  缺失问题", level=2)
    P(doc, "缺失指 D1-1 / D1-7 访谈中已被提出，但 D1-2 / D1-3 中未形成对应需求或验收准则的条目。")

    for d in D.DEFECTS:
        if d["type"] != "缺失":
            continue
        H(doc, f"{d['id']}  {d['req']}", level=3)
        P(doc, f"问题描述：{d['desc']}", indent_first=False)
        P(doc, f"溯源：在 D1-7 访谈纪要中可定位到原始诉求；当前 SRS 因优先级评估被遗漏。",
          indent_first=False)
        P(doc, f"处置建议：{d['fix']}", indent_first=False)
        P(doc, f"复审结论：通过（{d['closed']}）。其中 D-004 为 v1.0 范围外项，已挂在 SRS 附录 B TBD-1。",
          indent_first=False)

    doc.add_page_break()

    # ===========================================================
    # 4 缺陷处置追踪
    # ===========================================================
    H(doc, "4  缺陷处置追踪", level=1)
    P(doc, "下表汇总 8 条缺陷的处置情况、责任人、计划与实际关闭日期，以及关联的 SRS / RTM 修订点。")

    rows = []
    for d in D.DEFECTS:
        rows.append([
            d["id"], d["type"], d["severity"], d["owner"], d["fix"],
            d["due"], d["closed"], d["status"],
        ])
    make_table(doc, ["缺陷ID", "类型", "严重度", "责任人", "修复方案", "计划关闭", "实际关闭", "状态"],
               rows, col_widths_cm=[1.4, 1.5, 1.2, 1.6, 6.6, 1.7, 1.7, 1.3])

    H(doc, "4.1  二次复审记录", level=2)
    BULLET(doc, "2026-05-16 10:00-11:00 复审会：逐条核对 SRS v1.0 是否真正闭环，所有缺陷均确认 closed；")
    BULLET(doc, "复审参与人：张明、陈思、李华、王琳；记录：陈思；")
    BULLET(doc, "复审证据：D1-3 v1.0 与 v0.9 文本 diff、RTM 列「当前状态」全部为「已基线（v1.0）」。")

    H(doc, "4.2  无法在本周期内闭环的事项", level=2)
    BULLET(doc, "TBD-1（D-004 衍生）：第三方平台 OpenAPI 双向对接，预计 v1.1（2026-08-31）落地；")
    BULLET(doc, "TBD-3：店员端离线模式扩展至预约写入；当前仅限核销/点单，留待 v1.1；")
    BULLET(doc, "TBD-4：AI 推荐可解释性面板，留待 v1.1；")
    BULLET(doc, "TBD-5：屏幕阅读器全覆盖（NFR-009 由 Could 升级 Should），留待 v1.1。")

    doc.add_page_break()

    # ===========================================================
    # 5 结论与基线建议
    # ===========================================================
    H(doc, "5  结论与基线建议", level=1)

    H(doc, "5.1  本轮验证结论", level=2)
    P(doc, "综合检查清单、原型评审与 AI 反向质询三方结果，")
    BULLET(doc, "16 项检查清单 13 通过 + 3 部分通过（缺陷已闭环），整体通过率 ≥ 95%；")
    BULLET(doc, "8 条缺陷在 4 个工作日内全部闭环（H 类 1 条用 1 天解决、M 类 5 条用 ≤ 2 天、L 类 2 条用 ≤ 1 天）；")
    BULLET(doc, "30 条 FR + 10 条 NFR 在 D1-5 RTM 中均完成「需求 → 用例 → 类图 → 测试用例预占 → 可测试性来源」"
                "五栏追溯，覆盖率 100%；")
    BULLET(doc, "30 条 FR 100% 含 Given-When-Then 验收准则；10 条 NFR 100% 含量化目标值与验证方法。")

    P(doc, "据此，本组建议将 D1-3 SRS v1.0 冻结为 requirements-v1.0 基线，并在 Git 仓库打 Tag。")

    H(doc, "5.2  可测试性来源汇总（每条需求）", level=2)
    P(doc, "为满足任务书 §4 「标注每条需求的可测试性来源」要求，下表给出 40 条需求的可测试性来源映射。"
           "完整字段在 D1-5 RTM_FR / RTM_NFR 工作表的「可测试性来源」列。", indent_first=False)
    rows = []
    for fr in D.FUNCTIONAL_REQS:
        rows.append([fr["id"], fr["desc"][:18] + ("…" if len(fr["desc"]) > 18 else ""),
                     fr["test_source"]])
    for nf in D.NFR_REQS:
        rows.append([nf["id"], nf["desc"][:18] + ("…" if len(nf["desc"]) > 18 else ""),
                     nf["test_source"]])
    make_table(doc, ["需求ID", "需求摘要", "可测试性来源"],
               rows, col_widths_cm=[2, 5, 9])

    H(doc, "5.3  下阶段输入与建议", level=2)
    BULLET(doc, "进入设计阶段前，建议组织 1 次面向「真实顾客 + 真实店员」的原型可用性测试，验证 AI 模拟访谈的关键假设；")
    BULLET(doc, "建议在 v1.1 立项 TBD-1 ~ TBD-5；其中 TBD-1（第三方平台对接）已具备商业紧迫性；")
    BULLET(doc, "建议在 D1-5 RTM 中持续维护「测试用例预占ID」列；进入测试阶段后对应实际 TC 编号；")
    BULLET(doc, "建议在 SRS §3.2 与 RTM 中增设「变更影响域」字段（v1.1 启用），便于后续变更管理；")
    BULLET(doc, "AI 模拟访谈可在变更评审环节复用，但应坚持 ≥ 1 名真实干系人参与最终签字。")

    H(doc, "5.4  基线签字（待签）", level=2)
    make_table(doc, ["角色", "姓名", "签字", "日期"], [
        ["项目负责人 / 组长", "张明", "____________", "2026-05-16"],
        ["产品需求", "陈思", "____________", "2026-05-16"],
        ["UML 建模", "李华", "____________", "2026-05-16"],
        ["测试 / 文档", "王琳", "____________", "2026-05-16"],
        ["指导教师 / 课代表", "________", "____________", "________"],
    ], col_widths_cm=[5, 3, 4, 4])

    out = Path(r"D:/软件工程/软件工程实验任务书_2026春/软件工程实验任务书_2026春/实验一_智能化需求工程与UML建模/产出模板/D1-6_需求验证报告.docx")
    doc.save(str(out))
    print(f"saved: {out}  size={out.stat().st_size} bytes")


if __name__ == "__main__":
    build()
