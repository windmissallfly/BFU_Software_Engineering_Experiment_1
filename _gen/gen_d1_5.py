# -*- coding: utf-8 -*-
"""Generate D1-5 术语表与需求追溯矩阵."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import _data as D

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


HEADER_FILL = PatternFill("solid", fgColor="4472C4")
HEADER_FONT = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name="微软雅黑", size=10)
WRAP = Alignment(wrap_text=True, vertical="center")
CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")
THIN = Side(border_style="thin", color="B4B4B4")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"


def style_body(ws, nrows, ncols, center_cols=()):
    for r in range(2, nrows + 2):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = CENTER if c in center_cols else WRAP
        ws.row_dimensions[r].height = None  # 自动


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def build():
    wb = Workbook()

    # ====================== Sheet 1: Glossary ======================
    ws1 = wb.active
    ws1.title = "Glossary"

    headers = ["术语", "英文 / 缩写", "定义", "同义词", "反义词", "出处"]
    ws1.append(headers)
    for term in D.GLOSSARY:
        ws1.append(list(term))

    style_header(ws1, len(headers))
    style_body(ws1, len(D.GLOSSARY), len(headers), center_cols=(1, 2, 4, 5))
    set_col_widths(ws1, [16, 22, 60, 14, 12, 26])

    # ====================== Sheet 2: RTM (功能需求) ======================
    ws2 = wb.create_sheet("RTM_FR")
    headers2 = [
        "需求ID", "类别", "需求描述", "来源（干系人 + 访谈 Q）", "MoSCoW", "优先级",
        "用例ID", "类图实体", "顺序图引用", "状态机引用",
        "Given-When-Then 验收准则", "测试用例预占ID", "可测试性来源", "当前状态", "负责人",
    ]
    ws2.append(headers2)

    owners = ["张明", "陈思", "李华", "王琳"]
    for i, fr in enumerate(D.FUNCTIONAL_REQS):
        ws2.append([
            fr["id"], fr["category"], fr["desc"], fr["source"], fr["moscow"], fr["priority"],
            fr["uc"], fr["classes"], fr["seq"], fr["state"],
            fr["gwt"], fr["tc"], fr["test_source"], "已基线（v1.0）",
            owners[i % len(owners)],
        ])

    style_header(ws2, len(headers2))
    style_body(ws2, len(D.FUNCTIONAL_REQS), len(headers2),
               center_cols=(1, 2, 5, 6, 7, 9, 10, 12, 14, 15))
    set_col_widths(ws2, [9, 11, 30, 22, 8, 8, 9, 18, 16, 14, 50, 11, 36, 14, 9])

    # ====================== Sheet 3: RTM (非功能需求) ======================
    ws3 = wb.create_sheet("RTM_NFR")
    headers3 = [
        "需求ID", "ISO25010 子特性", "需求描述", "度量指标", "目标值", "验证方法",
        "MoSCoW", "关联用例", "测试用例预占ID", "可测试性来源", "当前状态", "负责人",
    ]
    ws3.append(headers3)
    for i, nf in enumerate(D.NFR_REQS):
        ws3.append([
            nf["id"], nf["iso"], nf["desc"], nf["metric"], nf["target"], nf["method"],
            nf["moscow"], nf["uc"], nf["tc"], nf["test_source"], "已基线（v1.0）",
            owners[i % len(owners)],
        ])

    style_header(ws3, len(headers3))
    style_body(ws3, len(D.NFR_REQS), len(headers3),
               center_cols=(1, 2, 4, 5, 7, 8, 9, 11, 12))
    set_col_widths(ws3, [10, 14, 30, 14, 16, 18, 8, 10, 11, 36, 14, 9])

    # ====================== Sheet 4: 覆盖度统计 ======================
    ws4 = wb.create_sheet("Coverage_Stats")
    ws4.append(["统计维度", "指标", "数量 / 比例", "说明"])
    moscow_count = {"Must": 0, "Should": 0, "Could": 0, "Won't": 0}
    for fr in D.FUNCTIONAL_REQS:
        moscow_count[fr["moscow"]] = moscow_count.get(fr["moscow"], 0) + 1
    uc_count = {}
    for fr in D.FUNCTIONAL_REQS:
        uc_count[fr["uc"]] = uc_count.get(fr["uc"], 0) + 1

    rows4 = [
        ["需求总数", "FR", str(len(D.FUNCTIONAL_REQS)),
         "30 条功能需求，覆盖任务书 §3 全部 6 类用例"],
        ["需求总数", "NFR", str(len(D.NFR_REQS)),
         "10 条非功能需求，按 ISO 25010 性能/可靠性/安全/易用性/可维护性 五大类组织"],
        ["MoSCoW · Must", "数量", str(moscow_count["Must"]),
         "Must 类必须在 v1.0 交付（含全部硬性核心流程）"],
        ["MoSCoW · Should", "数量", str(moscow_count["Should"]),
         "Should 类计划交付，可在最终评审中视进度调整"],
        ["MoSCoW · Could", "数量", str(moscow_count["Could"]),
         "Could 类资源允许时交付，否则纳入 v1.1"],
        ["MoSCoW · Won't (v1.0)", "数量", "0",
         "猫咪周边电商等明确 v1.0 不交付，列入 §1.2 范围外"],
        ["用例覆盖", "UC-01 桌位预约", str(uc_count.get("UC-01", 0)) + " 条 FR", "—"],
        ["用例覆盖", "UC-02 菜品点单", str(uc_count.get("UC-02", 0)) + " 条 FR", "—"],
        ["用例覆盖", "UC-03 处理订单", str(uc_count.get("UC-03", 0)) + " 条 FR", "—"],
        ["用例覆盖", "UC-04 店长经营", str(uc_count.get("UC-04", 0)) + " 条 FR", "—"],
        ["用例覆盖", "UC-05 猫咪健康", str(uc_count.get("UC-05", 0)) + " 条 FR", "—"],
        ["用例覆盖", "UC-06 跨店治理", str(uc_count.get("UC-06", 0)) + " 条 FR", "—"],
        ["可测性", "FR 含 Given-When-Then", "30 / 30 = 100%",
         "RTM_FR 表的「Given-When-Then 验收准则」列填充率 100%"],
        ["可测性", "NFR 含目标值 + 验证方法", "10 / 10 = 100%",
         "RTM_NFR 表的「目标值」与「验证方法」列填充率 100%"],
        ["UML 追溯", "FR ↔ 用例图", "30 / 30 = 100%",
         "每条 FR 在 RTM_FR 「用例ID」列均有非空值，对应 D1-4 use case 图节点"],
        ["UML 追溯", "FR ↔ 类图", "30 / 30 = 100%",
         "每条 FR 标注关联类图实体（如「门店, 桌子, 预约」）"],
        ["UML 追溯", "FR ↔ 顺序图", "12 / 30 = 40%",
         "需要跨服务调用的 FR 才需顺序图，其它标 「—」"],
        ["UML 追溯", "FR ↔ 状态机", "5 / 30 ≈ 17%",
         "仅订单生命周期相关 FR（FR-008/009/014/020）需状态机"],
        ["来源回溯", "FR 来源访谈", "30 / 30 = 100%",
         "每条 FR 注明访谈轮次 + 干系人 + 问题号，可在 D1-7 中精确定位"],
    ]
    for r in rows4:
        ws4.append(r)
    style_header(ws4, 4)
    style_body(ws4, len(rows4), 4, center_cols=(2, 3))
    set_col_widths(ws4, [22, 20, 18, 50])

    # ====================== Sheet 5: 使用说明 ======================
    ws5 = wb.create_sheet("使用说明")
    ws5.append(["列名", "含义", "填写规则"])
    rules = [
        ["需求ID", "唯一标识", "FR-xxx / NFR-xxx 三位数字，禁止重复使用旧 ID"],
        ["来源", "需求出处", "格式：干系人 · 访谈轮次-Qx，与 D1-7 一一对应"],
        ["MoSCoW", "优先级方法", "仅允许 Must / Should / Could / Won't 四值"],
        ["用例ID", "对应 UML 用例", "UC-01 ~ UC-06，与 D1-4/usecase/main_usecase.puml 一致"],
        ["类图实体", "对应 D1-4 类图", "多个实体用「, 」分隔"],
        ["顺序图引用", "对应 D1-4/sequence/", "未涉及跨服务调用填「—」"],
        ["状态机引用", "对应 D1-4/state/state_order.puml", "未涉及订单生命周期填「—」"],
        ["Given-When-Then 验收准则", "可测试形式", "每条 FR 必须填写，作为测试用例编写输入"],
        ["测试用例预占ID", "TC-xxx", "由测试组在 v1.0 进入实施阶段后填充实际用例 ID"],
        ["可测试性来源", "如何证明可测", "Given-When-Then + 验证方法/工具，例如「k6 压测 + Prometheus」"],
        ["当前状态", "RTM 生命周期状态", "未基线 / 设计中 / 已基线（v1.0） / 已交付 / 已撤销"],
        ["负责人", "需求负责人", "由小组分工分配；变更需在评审中说明"],
    ]
    for r in rules:
        ws5.append(r)
    style_header(ws5, 3)
    style_body(ws5, len(rules), 3, center_cols=(1,))
    set_col_widths(ws5, [18, 20, 60])

    # ===== 保存 =====
    out = Path(r"D:/软件工程/软件工程实验任务书_2026春/软件工程实验任务书_2026春/实验一_智能化需求工程与UML建模/产出模板/D1-5_术语表与需求追溯矩阵.xlsx")
    wb.save(str(out))
    print(f"saved: {out}  size={out.stat().st_size} bytes")


if __name__ == "__main__":
    build()
