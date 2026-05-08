# -*- coding: utf-8 -*-
"""Generate D1-3 软件需求规格说明书 SRS（IEEE 830）."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import _data as D

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ---------------------------- helpers ----------------------------
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
    """Add a styled table. header: list of str; rows: list of list of str."""
    tbl = doc.add_table(rows=1 + len(rows), cols=len(header))
    tbl.style = "Light Grid Accent 1"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    # header
    for i, h in enumerate(header):
        cell = tbl.cell(0, i)
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        _set_cn_font(run, "黑体", 10, bold=True)
    # rows
    for r, row in enumerate(rows, start=1):
        for c, val in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            _set_cn_font(run, "宋体", 9)
    # col widths
    if col_widths_cm:
        for i, w in enumerate(col_widths_cm):
            for cell in tbl.columns[i].cells:
                cell.width = Cm(w)
    return tbl


# ---------------------------- main ----------------------------
def build():
    doc = Document()

    # set default font
    style = doc.styles["Normal"]
    style.font.name = "宋体"
    style.font.size = Pt(10.5)

    # ==== 封面信息 ====
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run(D.DOC_META["course"])
    _set_cn_font(r, "黑体", 14, bold=True)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = subtitle.add_run("软件需求规格说明书 SRS（IEEE 830）")
    _set_cn_font(r, "黑体", 18, bold=True)

    sub2 = doc.add_paragraph()
    sub2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub2.add_run("Software Requirements Specification — D1-3")
    _set_cn_font(r, "Times New Roman", 12, bold=False)

    sub3 = doc.add_paragraph()
    sub3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub3.add_run(f"—— 案例：{D.DOC_META['case']} ——")
    _set_cn_font(r, "宋体", 12)

    doc.add_paragraph()

    # 文档信息表
    meta = make_table(doc, ["项目", "内容"], [
        ["产出编号", "D1-3"],
        ["产出名称", "软件需求规格说明书 SRS（IEEE 830）"],
        ["所属实验", D.DOC_META["exp"]],
        ["案例项目", D.DOC_META["case"]],
        ["版本", D.DOC_META["version"]],
        ["作者团队", D.DOC_META["team"]],
        ["提交日期", D.DOC_META["submit"]],
    ], col_widths_cm=[4, 12])

    # 修订记录
    P(doc, "")
    H(doc, "修订记录", level=3)
    make_table(doc, ["版本", "日期", "作者", "变更摘要"], [
        ["v0.1", "2026-05-13", "张明", "初稿，依据 D1-1/D1-2 形成 IEEE 830 章节骨架"],
        ["v0.5", "2026-05-15", "陈思", "补全 6 个用例规约 + 3.3 性能与 3.5 软件系统属性"],
        ["v0.9", "2026-05-16", "张明", "依据 D1-6 验证报告处置 D-001~D-008 缺陷"],
        ["v1.0", "2026-05-16", "全员", "评审通过，作为 requirements-v1.0 基线候选"],
    ], col_widths_cm=[2, 3, 3, 8])

    doc.add_page_break()

    # ===========================================================
    # 1 引言
    # ===========================================================
    H(doc, "1  引言", level=1)

    H(doc, "1.1  目的", level=2)
    P(doc, "本《软件需求规格说明书》（以下简称 SRS）面向 NekoCafé 智慧餐饮预约平台 v1.0，"
           "旨在以 IEEE Std 830-1998 / ISO/IEC/IEEE 29148-2018 规定的形式描述系统的功能与非功能需求，"
           "作为开发、测试、运维、产品、运营五方共同的需求基线。")
    P(doc, "本 SRS 的预期读者包括：")
    BULLET(doc, "项目经理 / 产品经理：核对范围、优先级与里程碑；")
    BULLET(doc, "前后端 / AI / 数据 开发工程师：用作设计与编码输入；")
    BULLET(doc, "测试工程师：基于 §3 的 Given-When-Then 验收准则编写测试用例；")
    BULLET(doc, "运维 / SRE 工程师：核对 §3.3 性能、§3.5 可靠性目标；")
    BULLET(doc, "总部运营、店长、店员、猫咪管家、顾客代表：对照本基线在评审中签字确认。")

    H(doc, "1.2  范围", level=2)
    P(doc, "范围内（In Scope）：")
    BULLET(doc, "顾客端：小程序与 H5，覆盖门店浏览、桌位预约、AI 推荐、点单支付、跨店会员、分享打卡；")
    BULLET(doc, "店员端 / 猫咪管家端：iPad / Android POS / 手机端，覆盖桌位看板、订单分配、扫码核销、健康打卡、异常告警；")
    BULLET(doc, "店长后台：排班、退款审批、评价回复；")
    BULLET(doc, "总部运营平台：跨店数据看板、AI 转化报表、营销发券、菜品下发、权限治理；")
    BULLET(doc, "覆盖 30 条功能需求（FR-001 ~ FR-030）与 10 条非功能需求（NFR-001 ~ NFR-010）。")
    P(doc, "范围外（Out of Scope，标记为 v1.0 不交付）：")
    BULLET(doc, "美团 / 大众点评 / 抖音 第三方平台 OpenAPI 双向对接（D-004 已立项 v1.1）；")
    BULLET(doc, "猫咪周边小电商模块（来自访谈 III-Q12，MoSCoW=Won't）；")
    BULLET(doc, "猫咪健康 SaaS 对外输出（远期商业化路径）。")

    H(doc, "1.3  定义、缩写、缩略语", level=2)
    P(doc, "本节列出 SRS 内出现的核心术语；完整术语表参见 D1-5《术语表与需求追溯矩阵》Sheet1。")
    rows = [[t[0], t[1], t[2]] for t in D.GLOSSARY[:14]]
    make_table(doc, ["术语", "英文 / 缩写", "定义"], rows, col_widths_cm=[3, 3.5, 9.5])

    H(doc, "1.4  参考文献", level=2)
    refs = [
        "[1] IEEE Std 830-1998. Recommended Practice for Software Requirements Specifications.",
        "[2] ISO/IEC/IEEE 29148:2018. Systems and software engineering — Life cycle processes — Requirements engineering.",
        "[3] ISO/IEC 25010:2011. Systems and software Quality Requirements and Evaluation (SQuaRE).",
        "[4] Sommerville I. Software Engineering (10th Edition). Pearson, 2015.",
        "[5] 张海藩, 吕云翔. 《软件工程（第 4 版）》. 人民邮电出版社, 2021.",
        "[6] D1-1《需求获取计划与访谈剧本》v1.0. 项目内部文档.",
        "[7] D1-2《需求清单与 MoSCoW 分级表》v1.0. 项目内部文档.",
        "[8] D1-4《UML 模型集》v1.0（含用例图、活动图、类图、顺序图、状态图）. 项目内部文档.",
        "[9] D1-7《AI 使用记录与原始访谈对话》v1.0. 项目内部文档.",
        "[10] PlantUML 官方文档：https://plantuml.com/zh/",
        "[11] Conventional Commits 规范：https://www.conventionalcommits.org/zh-hans/",
    ]
    for ref in refs:
        BULLET(doc, ref)

    H(doc, "1.5  概览", level=2)
    P(doc, "本 SRS 后续章节组织方式：")
    NUM(doc, "第 2 章 总体描述：从产品愿景、功能概要、用户类、运行环境、约束与假设五个维度刻画系统的全景；")
    NUM(doc, "第 3 章 具体需求：按 IEEE 830 §3 推荐结构展开外部接口、功能、性能、设计约束与软件系统属性；功能需求 §3.2 按 6 个用例（UC-01 ~ UC-06）组织；")
    NUM(doc, "第 4 章 其他需求：补充数据库、国际化、合规等横切要求；")
    NUM(doc, "第 5 章 附录：放置分析模型缩略图与待解决问题（TBD）清单。")
    P(doc, "本 SRS 与 D1-2、D1-4、D1-5 互为输入输出：每条需求均可经 D1-5 RTM 反向追溯至访谈来源、UML 模型与测试用例。")

    doc.add_page_break()

    # ===========================================================
    # 2 总体描述
    # ===========================================================
    H(doc, "2  总体描述", level=1)

    H(doc, "2.1  产品愿景", level=2)
    P(doc, "「让每一次撸猫都被温柔以待」—— NekoCafé 智慧餐饮预约平台旨在打通全国 12 家门店的"
           "预约、点单、会员、排班与猫咪健康全链路，使顾客能够基于 AI 推荐确定性地享受指定猫咪的陪伴，"
           "店员从纸质流程中解放出来专注服务质量，店长 / 总部依靠实时数据看板做出科学经营决策。"
           "项目 v1.0 期望在两年内将集团翻台率由 1.8 次/天提升至 2.3 次/天、会员复购率由 34% 提升至 50% 以上。")

    H(doc, "2.2  产品功能概要", level=2)
    feats = [
        ("F1 门店浏览与桌位预约", "顾客可按城市、距离、标签筛选门店并预约带特定猫咪伴侣的桌位（FR-001~FR-003、FR-005、FR-006、FR-010、FR-011）。"),
        ("F2 AI 智能推荐与点单支付", "支持基于偏好的菜品推荐、扫码点单、聚合支付、跨店会员积分与海报分享（FR-004、FR-007~FR-009、FR-012、FR-029）。"),
        ("F3 店员桌位调度与订单处理", "实时桌位看板、订单分配、扫码核销、换桌操作（FR-013、FR-014、FR-018、FR-019）。"),
        ("F4 店长经营管理", "排班管理、客诉退款审批、评价回复（FR-015、FR-020、FR-026）。"),
        ("F5 猫咪健康档案与异常预警", "猫咪健康打卡、连续工作时长强制休息、黑名单管理（FR-016、FR-017、FR-030）。"),
        ("F6 总部跨店数据治理与营销", "实时大屏、AI 转化报表、跨店发券、权限治理、菜品下发、热力图分析（FR-021~FR-025、FR-027、FR-028）。"),
    ]
    rows = [[k, v] for k, v in feats]
    make_table(doc, ["能力域", "范围与覆盖需求"], rows, col_widths_cm=[5, 11])

    H(doc, "2.3  用户类与特征", level=2)
    users = [
        ["顾客", "高（每周 1 次以上）", "中（年轻消费群体，移动端熟练）",
         "预约确定性、猫咪健康可见、隐私可控", "FR-001~FR-012、FR-029"],
        ["店员", "极高（每日 8 小时）", "中（POS 操作熟练，对纸质流程疲倦）",
         "高峰期效率、离线可用、新手友好", "FR-013、FR-014、FR-018、FR-019"],
        ["猫咪管家", "高（每日 2 次打卡）", "中（手机端为主）",
         "猫咪档案数字化、异常一键告警", "FR-016、FR-017、FR-030"],
        ["店长", "中（每日 1~2 次）", "中（PC 后台为主）",
         "排班、退款、评价管理", "FR-015、FR-020、FR-026"],
        ["总部运营总监", "中（每日 1 次大屏）", "高（关注 KPI 与战略）",
         "数据看板、报表、权限、营销", "FR-021~FR-028"],
        ["顾客代表 - VIP", "中", "高（高 ARPU 群体）",
         "专享折扣、优先预约", "FR-025、子用例 VIPDiscount"],
        ["监管 / 审计员（间接）", "低（合规审计期间）", "高（合规要求）",
         "审计日志、动物保护合规", "NFR-006、NFR-007、§4.3"],
    ]
    make_table(doc, ["用户类", "使用频次", "技术水平", "关注点", "对应需求"],
               users, col_widths_cm=[3, 2.5, 3, 4, 3.5])

    H(doc, "2.4  运行环境", level=2)
    env_rows = [
        ["顾客端", "微信小程序（基础库 ≥ 2.27）/ H5（Chrome 90+ / Safari 14+）"],
        ["店员端", "iPad（iPadOS 15+）/ Android POS（Android 10+，1280x800 以上）"],
        ["猫咪管家端", "Android 10+ / iOS 14+ 手机"],
        ["店长 / 总部后台", "PC 浏览器（Chrome 90+ / Edge 90+），分辨率 ≥ 1920×1080"],
        ["服务端", "阿里云 ACK 容器集群（8c16g × ≥ 6 节点）+ MySQL 8.0 + Redis 7 + Elasticsearch 8 + Kafka 3"],
        ["AI 服务", "PaaS 推理服务，TPS ≥ 200，GPU/CPU 弹性扩缩容"],
        ["第三方依赖", "微信支付 / 支付宝 / 阿里云短信 / 高德地图 / 七牛云存储"],
        ["网络环境", "门店有线 + 4G/5G 双链路，店员端支持离线缓存（NFR-005 关联）"],
    ]
    make_table(doc, ["端 / 组件", "运行环境"], env_rows, col_widths_cm=[5, 11])

    H(doc, "2.5  设计与实现约束", level=2)
    BULLET(doc, "架构约束：采用微服务 + 容器化部署，按业务限界上下文（顾客域、订单域、猫咪域、运营域）解耦（NFR-010）；")
    BULLET(doc, "技术栈约束：后端 Java 17 / Spring Boot 3 / MyBatis-Plus；前端 Vue 3 + TypeScript / 小程序原生；"
                "数据库 MySQL 8.0（OLTP）+ Elasticsearch（搜索/报表）+ Redis（缓存/计数）；")
    BULLET(doc, "安全约束：所有外网通信强制 HTTPS（TLS 1.2+，NFR-007）；用户敏感数据 AES-256 加密落库（NFR-006）；")
    BULLET(doc, "合规约束：遵守《个人信息保护法》《网络安全法》《数据安全法》以及餐饮经营许可、动物保护相关条例；")
    BULLET(doc, "性能约束：核心接口 P95 ≤ 350ms，QPS ≥ 5000，首屏 ≤ 2.0s（NFR-001~003）；")
    BULLET(doc, "可用性约束：核心链路月度可用率 ≥ 99.9%，关键节点 RTO ≤ 1 分钟（NFR-004、NFR-005）；")
    BULLET(doc, "易用性约束：店员高峰期关键操作 ≤ 3 次点击（NFR-008），UI 满足 WCAG 2.1 AA（NFR-009）；")
    BULLET(doc, "审计约束：所有写操作（菜单变更、退款、权限变更）需记录审计日志，保留 ≥ 2 年。")

    H(doc, "2.6  假设与依赖", level=2)
    BULLET(doc, "假设 A1：微信支付 / 支付宝 沙箱与生产可用率均 ≥ 99.95%；")
    BULLET(doc, "假设 A2：高德地图 SDK 在国内门店所在城市覆盖 POI 与路网；")
    BULLET(doc, "假设 A3：阿里云短信通道 5 秒内可送达 ≥ 99% 短信；")
    BULLET(doc, "假设 A4：各门店至少配备 1 台 iPad + 1 台 POS + 稳定 Wi-Fi（带宽 ≥ 50 Mbps）；")
    BULLET(doc, "依赖 D1：AI 推荐服务（猫咪性格画像、口味画像）由集团 AI 中台输出 RESTful 接口；")
    BULLET(doc, "依赖 D2：会员中心 / 营销中心 / 支付中心 由集团已有中台提供，本平台仅作集成；")
    BULLET(doc, "依赖 D3：门店硬件（扫码枪、打印机）需符合 ESC/POS 协议；")
    BULLET(doc, "依赖 D4：监管报送接口由集团合规部提供模板，开发期内冻结。")

    doc.add_page_break()

    # ===========================================================
    # 3 具体需求
    # ===========================================================
    H(doc, "3  具体需求", level=1)
    P(doc, "本章按 IEEE 830 §3 推荐结构组织：3.1 外部接口、3.2 功能需求（按 6 个用例编号）、"
           "3.3 性能、3.4 设计约束、3.5 软件系统属性。每条功能需求附 Given-When-Then 验收准则与测试用例预占 ID（TC-xxx），"
           "便于在 D1-5 RTM 中前向追溯到测试，反向追溯到访谈与 UML 元素。")

    # ---- 3.1 外部接口 ----
    H(doc, "3.1  外部接口需求", level=2)

    H(doc, "3.1.1  用户接口（UI）", level=3)
    BULLET(doc, "顾客小程序 / H5：列表 + 详情 + 表单的标准移动布局；首屏 ≤ 2 秒（NFR-003）；支持深色模式与字体放大；")
    BULLET(doc, "店员 iPad / POS：实时桌位地图、订单看板、扫码核销；提供「新手模式 / 简洁模式」（覆盖访谈 II-Q10），关键操作 ≤ 3 次点击（NFR-008）；")
    BULLET(doc, "猫咪管家手机端：单手可达的打卡卡片，支持拍照与异常红色高亮；")
    BULLET(doc, "店长 / 总部 Web 后台：1920×1080 优先，左侧菜单 + 右侧工作区；多角色权限按 RBAC 渲染。")

    H(doc, "3.1.2  硬件接口", level=3)
    make_table(doc, ["接口", "硬件", "说明"], [
        ["扫码枪 / 摄像头", "门店扫码枪（USB-HID）/ iPad 摄像头", "用于预约码核销 FR-018，低光环境亦可识别"],
        ["小票打印机", "ESC/POS 兼容打印机", "FR-014 后厨/吧台小票"],
        ["收银 POS 机", "Sunmi/Verifone Android POS", "聚合支付硬件回调（FR-009）"],
        ["体温/计数传感（可选）", "门店摄像头 AI 客流计数", "辅助 FR-021 实时客流统计"],
    ], col_widths_cm=[3, 4, 9])

    H(doc, "3.1.3  软件接口", level=3)
    make_table(doc, ["类型", "对接方", "协议 / 数据格式", "用途", "异常码示例"], [
        ["支付", "微信支付 / 支付宝", "HTTPS REST + JSON", "FR-009 聚合支付", "PAY_TIMEOUT / PAY_SIG_FAIL"],
        ["短信", "阿里云短信", "HTTPS REST + 模板号", "FR-003 预约确认 / 客诉通知", "SMS_RATE_LIMIT"],
        ["地图", "高德地图 Web SDK", "JS SDK + REST", "门店列表距离 / 路径（FR-001）", "AMAP_QUOTA"],
        ["对象存储", "七牛云", "HTTPS + 七牛 SDK", "猫咪照片、海报存储（FR-006/029）", "OSS_403"],
        ["AI 推荐", "集团 AI 中台", "HTTPS REST + JSON", "FR-004 / FR-005 推荐", "AI_TIMEOUT / AI_NO_DATA"],
        ["会员中心", "集团会员中台", "HTTPS REST + JSON", "FR-012 / FR-025 积分等级", "MEMBER_404"],
        ["合规审计", "集团合规中台", "HTTPS + 异步 MQ", "审计日志归档", "AUDIT_RETRY"],
    ], col_widths_cm=[2.5, 3, 3, 4.5, 3])

    H(doc, "3.1.4  通信接口", level=3)
    BULLET(doc, "对外：HTTPS + TLS 1.2+（NFR-007），网关层 OAuth 2.0 + JWT 鉴权；")
    BULLET(doc, "实时推送：店员端 ↔ 服务端 采用 WebSocket（FR-013 桌位看板 1 秒刷新）；")
    BULLET(doc, "门店内：店员端与小票打印机通过局域网 ESC/POS / 蓝牙；")
    BULLET(doc, "异步：核心域之间通过 Kafka，主题命名 `nekocafe.<domain>.<event>`；")
    BULLET(doc, "时序：所有时间采用 UTC + 时区元数据传递，前端按本地时区渲染。")

    # ---- 3.2 功能需求（按 UC）----
    H(doc, "3.2  功能需求（按用例编号组织）", level=2)
    P(doc, "本节按 6 个核心用例（UC-01 ~ UC-06）组织功能需求。每个用例包括：参与者、目标、前置条件、"
           "主流程、备选流程、异常路径、后置条件、覆盖的 FR 列表。各 FR 的 Given-When-Then 验收准则汇总在表 3.2-X。")

    for uc in D.USE_CASES:
        H(doc, f"3.2.{int(uc['id'].split('-')[1])}  {uc['id']}  {uc['name']}", level=3)
        info = [
            ["主参与者", uc["actor_primary"]],
            ["次参与者", uc["actor_secondary"]],
            ["用例目标", uc["goal"]],
            ["前置条件", uc["precondition"]],
            ["后置条件", uc["postcondition"]],
            ["覆盖功能需求", " / ".join(uc["frs"])],
        ]
        make_table(doc, ["项", "内容"], info, col_widths_cm=[3, 13])

        P(doc, "主流程：", indent_first=False)
        for s in uc["main_flow"]:
            BULLET(doc, s)
        P(doc, "备选流程：", indent_first=False)
        for s in uc["alt_flow"]:
            BULLET(doc, s)
        P(doc, "异常路径：", indent_first=False)
        for s in uc["exception"]:
            BULLET(doc, s)

        # 该用例的 FR 验收表
        rows = []
        for fr in D.FUNCTIONAL_REQS:
            if fr["id"] in uc["frs"]:
                rows.append([fr["id"], fr["desc"], fr["moscow"], fr["gwt"], fr["tc"]])
        make_table(doc, ["FR-ID", "需求描述", "MoSCoW", "Given-When-Then 验收准则", "TC 预占"],
                   rows, col_widths_cm=[1.7, 3.6, 1.5, 7.4, 1.8])

    # ---- 3.3 性能 ----
    H(doc, "3.3  性能需求", level=2)
    P(doc, "本节给出系统在响应时间、吞吐量、并发能力、容量等方面的硬性约束。所有指标在 D1-5 RTM 中"
           "标注「可测试性来源」，并在 D1-6 验证报告中通过压测脚本 / Lighthouse / Prometheus 监控核验。")
    perf = [r for r in D.NFR_REQS if r["iso"] == "性能效率" or r["iso"] == "可靠性"]
    rows = [[r["id"], r["iso"], r["desc"], r["metric"], r["target"], r["method"]] for r in perf]
    make_table(doc, ["NFR-ID", "ISO 25010 子特性", "需求描述", "指标", "目标值", "验证方法"],
               rows, col_widths_cm=[1.6, 2.4, 5.0, 2.0, 2.5, 2.5])
    P(doc, "补充：")
    BULLET(doc, "高峰期定义：每周五 18:00-22:00、周末 11:00-22:00、法定节假日 11:00-22:00（消除访谈中「高峰期」二义性，对应 D-001 处置）；")
    BULLET(doc, "实时定义：FR-021 数据看板「实时」 = 数据延迟 ≤ 5 秒（消除 D-006 二义性）；")
    BULLET(doc, "容量规划：单门店峰值并发 200，总部大屏汇聚 ≤ 2 秒生成；6 个月数据热存，2 年冷存，超过 2 年归档至 OSS。")

    # ---- 3.4 设计约束 ----
    H(doc, "3.4  设计约束", level=2)
    BULLET(doc, "编程语言：后端 Java 17（Spring Boot 3.2）；前端 Vue 3.4 + TypeScript 5；移动端 Taro 3 + 小程序原生；")
    BULLET(doc, "数据库：MySQL 8.0（双 AZ 主从）+ Elasticsearch 8（报表）+ Redis 7（缓存/计数）；")
    BULLET(doc, "消息中间件：Kafka 3.x，跨域事件驱动；")
    BULLET(doc, "通信协议：对外 HTTPS / WSS；服务间 gRPC + JSON Schema 校验；")
    BULLET(doc, "代码与提交规范：Conventional Commits、CI 强制 SonarQube 圈复杂度 ≤ 15；")
    BULLET(doc, "合规：所有用户数据存储于境内可用区，跨境调用一律走脱敏中转。")

    # ---- 3.5 软件系统属性 ----
    H(doc, "3.5  软件系统属性", level=2)

    H(doc, "3.5.1  可靠性 (Reliability)", level=3)
    BULLET(doc, "NFR-004 月度可用率 ≥ 99.9%；")
    BULLET(doc, "NFR-005 关键节点 RTO ≤ 1 分钟；店员端离线缓存 30 分钟（D-005 处置补丁）；")
    BULLET(doc, "依据 ISO 25010：包含 Maturity / Fault Tolerance / Recoverability；混沌工程演练每季度一次。")

    H(doc, "3.5.2  安全性 (Security)", level=3)
    BULLET(doc, "NFR-006 顾客敏感字段 AES-256；密钥托管 KMS，季度轮换；")
    BULLET(doc, "NFR-007 全端 HTTPS（TLS 1.2+），HSTS、SameSite=Strict；")
    BULLET(doc, "FR-024 RBAC 强制隔离，所有写操作记录审计日志（含操作者 / IP / 修改前后值）；")
    BULLET(doc, "OWASP Top 10 防护：上线前 SAST + DAST 扫描必须 0 高危。")

    H(doc, "3.5.3  易用性 (Usability)", level=3)
    BULLET(doc, "NFR-008 店员高峰期关键操作 ≤ 3 次点击；")
    BULLET(doc, "NFR-009 颜色对比度满足 WCAG 2.1 AA；支持系统级字体放大；")
    BULLET(doc, "店员新手模式：连续登录 14 天内默认开启操作引导。")

    H(doc, "3.5.4  可维护性 (Maintainability)", level=3)
    BULLET(doc, "NFR-010 微服务 + 容器化；服务间零共享数据库；")
    BULLET(doc, "代码覆盖率：核心域单元测试 ≥ 70%；E2E 用例覆盖 100% Must 级 FR；")
    BULLET(doc, "可观测性：每个服务输出 RED 指标 + Trace + 结构化日志。")

    H(doc, "3.5.5  可移植性 (Portability) / 兼容性 (Compatibility)", level=3)
    BULLET(doc, "顾客小程序：兼容微信基础库 2.27 及以上；")
    BULLET(doc, "店员端：iPadOS 15+ / Android 10+；")
    BULLET(doc, "服务端：容器化部署，可在阿里云 ACK / 自建 K8s 之间迁移；")
    BULLET(doc, "数据：导入导出统一使用 JSON Schema + 协议 Buffer。")

    doc.add_page_break()

    # ===========================================================
    # 4 其他需求
    # ===========================================================
    H(doc, "4  其他需求", level=1)

    H(doc, "4.1  数据库需求", level=2)
    BULLET(doc, "数据保留：交易订单、支付凭证保留 ≥ 7 年（财税合规）；猫咪健康记录保留 ≥ 5 年；用户行为日志热存 6 个月、冷存 2 年；")
    BULLET(doc, "备份与归档：每日全备 + 每 15 分钟增量；快照保留 30 天，OSS 归档保留 7 年；")
    BULLET(doc, "灾备：跨可用区主从 + 异地 1 份冷备，RPO ≤ 5 分钟、RTO ≤ 1 分钟（与 NFR-005 一致）；")
    BULLET(doc, "数据脱敏：测试环境一律使用脱敏后数据，禁止生产数据流出；")
    BULLET(doc, "数据字典：参见 D1-4 类图（领域类图）的 12 个核心实体作为权威字段来源。")

    H(doc, "4.2  国际化需求", level=2)
    BULLET(doc, "v1.0 仅交付 zh-CN 中文版，所有文案预留 i18n key 以支持 v1.x 扩展（zh-HK / en-US）；")
    BULLET(doc, "时区：当前默认 Asia/Shanghai；预留按门店时区运营的扩展点；")
    BULLET(doc, "货币：v1.0 仅 CNY；积分价值 1 分 = 0.01 CNY 等价兑换；")
    BULLET(doc, "字符集：全链路 UTF-8；emoji / 颜文字（顾客评价）正常存取。")

    H(doc, "4.3  法律法规与合规", level=2)
    rows = [
        ["《中华人民共和国个人信息保护法》(2021)", "FR-006/007/012/023/024、NFR-006",
         "用户敏感字段加密存储；本人授权后才进行 AI 偏好画像；权限分级"],
        ["《网络安全法》(2017)", "NFR-006/007、§3.1.4",
         "全链路 HTTPS、关键日志保留 ≥ 6 个月、网络安全等级保护 III 级"],
        ["《数据安全法》(2021)", "§4.1、NFR-006",
         "数据分类分级管理；跨境/跨域数据流动需脱敏并审批"],
        ["《餐饮服务食品安全操作规范》", "FR-016、FR-027",
         "菜品备案、留样记录追溯；批次/过敏原标注"],
        ["《动物防疫法》《北京市动物防疫管理办法》", "FR-016、FR-017",
         "猫咪疫苗、强制休息、隔离档案"],
        ["《广告法》", "FR-022/023/028",
         "营销文案审核；推广转化率统计不得诱导未成年人消费"],
    ]
    make_table(doc, ["法规 / 标准", "对应需求", "落实点"], rows, col_widths_cm=[5, 4, 7])

    doc.add_page_break()

    # ===========================================================
    # 5 附录
    # ===========================================================
    H(doc, "5  附录", level=1)

    H(doc, "附录 A  分析模型", level=2)
    P(doc, "本附录列出 D1-4 UML 模型集中的关键图例索引（导出 PNG 位于 D1-4/exports/ 目录）：")
    rows = [
        ["主用例图", "usecase/main_usecase.puml", "exports/NekoCafe_Main_UseCase.png", "UC-01 ~ UC-06 全景"],
        ["子用例图（顾客）", "usecase/sub_usecase_customer.puml", "exports/NekoCafe_Sub_UseCase_Customer.png",
         "include / extend / 泛化（顾客 ↔ VIPCustomer）"],
        ["子用例图（管家）", "usecase/sub_usecase_keeper.puml", "exports/NekoCafe_Sub_UseCase_Keeper.png",
         "猫咪健康打卡 → 异常告警 extend"],
        ["活动图 - 桌位预约", "activity/reservation.puml", "exports/NekoCafe_Activity_Reservation.png",
         "覆盖 UC-01 主流程与 5a 取消分支"],
        ["活动图 - 猫咪健康打卡", "activity/healthCheck.puml", "exports/NekoCafe_Activity_HealthCheck.png",
         "覆盖 UC-05 主流程与 1a 异常分支"],
        ["类图（领域类图）", "class/class.puml", "exports/NekoCafe_ClassDiagram.png",
         "12 个核心实体：用户/顾客/员工/管理者/猫咪管家/门店/桌子/猫咪/猫咪健康记录/预约/订单/订单元素/菜单元素"],
        ["顺序图 - AI 点单", "sequence/sequence_ai_order.puml", "exports/NekoCafe_Sequence_AIOrder.png",
         "覆盖 FR-004/007/008/009 跨服务调用"],
        ["顺序图 - 排班", "sequence/sequence_staff_schedule.puml", "exports/NekoCafe_Sequence_StaffSchedule.png",
         "覆盖 FR-015 排班 / FR-017 异常告警"],
        ["状态图 - 订单状态机", "state/state_order.puml", "exports/NekoCafe_State_Order.png",
         "覆盖订单 7 个状态与 12 条迁移"],
    ]
    make_table(doc, ["图类型", "源文件", "导出图", "覆盖需求要点"], rows, col_widths_cm=[3, 4.5, 4.5, 4])

    H(doc, "附录 B  待解决问题清单（TBD）", level=2)
    rows = [
        ["TBD-1", "美团 / 大众点评 / 抖音 OpenAPI 双向对接", "v1.1 引入 BFF 适配层", "陈思", "2026-08-31",
         "源访谈 III-Q11 / D1-6 缺陷 D-004"],
        ["TBD-2", "猫咪周边小电商模块（v1.0 Won't）", "纳入 v2.0 商业化路线图", "刘建华", "2027-Q1",
         "源访谈 III-Q12"],
        ["TBD-3", "门店离线模式扩展至预约写入（当前仅核销/点单）", "需评估冲突合并策略", "张明", "2026-09-30",
         "NFR-005 后续优化"],
        ["TBD-4", "AI 模型可解释性（推荐结果原因）", "v1.1 增加「为什么推荐」面板", "李华", "2026-10-31",
         "源访谈 I-Q6"],
        ["TBD-5", "无障碍：屏幕阅读器全覆盖", "NFR-009 当前为 Could，v1.1 升至 Should", "王琳", "2026-11-30",
         "ISO 25010 易用性"],
    ]
    make_table(doc, ["TBD-ID", "问题描述", "解决方向", "责任人", "截止时间", "来源"],
               rows, col_widths_cm=[1.7, 4.5, 4, 1.8, 2.2, 2])

    # ====== 落盘 ======
    out = Path(r"D:/软件工程/软件工程实验任务书_2026春/软件工程实验任务书_2026春/实验一_智能化需求工程与UML建模/产出模板/D1-3_软件需求规格说明书_SRS.docx")
    doc.save(str(out))
    print(f"saved: {out}  size={out.stat().st_size} bytes")


if __name__ == "__main__":
    build()
