"""
CrossLingo ⚖️ — 涉外商务合同双语风险智能扫描器（旗舰全量版）
================================================================
四大核心功能：
  1. 文件上传解析 — 支持 Word(.docx) / PDF / TXT 直接上传
  2. 多合同模板 — 7 套合同类型（采购/外包/NDA/劳动/数据跨境/合资/授权书）一键切换
  3. 一键复制建议 — 每条修改建议独立复制，直接粘贴到修订版合同

技术栈：Python + Streamlit + DeepSeek API
作者：Chen | 版本：v5.0 Flagship
"""

# ========== 第1步：导入工具箱 ==========
import streamlit as st
from openai import OpenAI
import json
from datetime import datetime
import os
import io
import re
from fpdf import FPDF
from dotenv import load_dotenv

# 文件解析库
from docx import Document       # 解析 Word .docx 文件
from PyPDF2 import PdfReader    # 解析 PDF 文件

load_dotenv()

# ---- 初始化语言设置 ----
if "lang" not in st.session_state:
    st.session_state.lang = "zh"

# ---- 翻译字典 ----
L = {
    # 侧边栏
    "ai_engine":        {"zh": "🔑 接入 AI 引擎",           "en": "🔑 AI Engine"},
    "key_configured":   {"zh": "✅ API Key 已配置",         "en": "✅ API Key Configured"},
    "change_key":       {"zh": "🔄 更换 Key",               "en": "🔄 Change Key"},
    "enter_new_key":    {"zh": "输入新的 API Key",          "en": "Enter new API Key"},
    "temp_key_hint":    {"zh": "仅本次生效，不会修改 .env 文件", "en": "Temporary, won't modify .env"},
    "key_tip":          {"zh": "💡 提示：创建 .env 文件可永久保存 Key", "en": "💡 Tip: Create .env file to save your key permanently"},
    "key_placeholder":  {"zh": "在 platform.deepseek.com → API Keys 获取", "en": "Get it at platform.deepseek.com → API Keys"},
    "scan_dims":        {"zh": "📋 扫描维度",               "en": "📋 Scan Dimensions"},
    "dim_mismatch":     {"zh": "中英文不对应",               "en": "Translation Mismatch"},
    "dim_ambiguous":    {"zh": "模糊条款",                   "en": "Ambiguous Clauses"},
    "dim_missing":      {"zh": "缺失条款",                   "en": "Missing Clauses"},
    "dim_compliance":   {"zh": "跨文化合规风险",             "en": "Cross-Border Compliance"},
    "dim_mismatch_desc":  {"zh": "语义偏差 · 范围不一致 · 数字时间差异", "en": "Semantic gaps · Scope mismatch · Numeric discrepancies"},
    "dim_ambiguous_desc": {"zh": "vague terms · 无量化标准 · 单向解释权", "en": "Vague terms · No metrics · Unilateral interpretation"},
    "dim_missing_desc":   {"zh": "保密期限 · 管辖法院 · 知识产权归属", "en": "Confidentiality period · Jurisdiction · IP ownership"},
    "dim_compliance_desc": {"zh": "争议解决冲突 · 违约金差异 · 数据跨境", "en": "Dispute resolution · Penalty differences · Data transfer"},
    # 主区域
    "contract_input":   {"zh": "📄 合同内容输入",           "en": "📄 Contract Input"},
    "select_template":  {"zh": "📋 选择合同模板（快速填充示例）", "en": "📋 Select Template"},
    "upload_file":      {"zh": "📎 或上传合同文件（支持 .txt / .docx / .pdf）", "en": "📎 Or Upload File (.txt / .docx / .pdf)"},
    "upload_hint":      {"zh": "上传后自动解析文本内容并填入下方输入框", "en": "Auto-parses text into the editor below"},
    "parse_success":    {"zh": "已成功解析文件",             "en": "Successfully parsed file"},
    "parse_chars":      {"zh": "个字符",                     "en": "chars"},
    "parse_error":      {"zh": "不支持的文件格式",           "en": "Unsupported file format"},
    "parse_only":       {"zh": "仅支持 .txt / .docx / .pdf", "en": "Only .txt / .docx / .pdf supported"},
    "contract_label":   {"zh": "合同内容（中文 + 英文）— 可自由编辑", "en": "Contract Content (CN + EN) — Editable"},
    "contract_placeholder": {"zh": "请将中英双语合同全文粘贴到此处...", "en": "Paste bilingual contract text here..."},
    "scan_btn":         {"zh": "🔍 开始智能扫描",           "en": "🔍 Start Scan"},
    "no_key_warn":      {"zh": "⚠️ 请先在左侧边栏输入 DeepSeek API Key，激活扫描引擎", "en": "⚠️ Please enter your DeepSeek API Key in the sidebar"},
    # 加载动画
    "scanning_title":   {"zh": "AI 正在分析合同...",        "en": "AI is analyzing your contract..."},
    "scanning_sub":     {"zh": "DeepSeek 大模型逐条比对中英文条款，预计需要 10-30 秒", "en": "DeepSeek is comparing CN/EN clauses, ~10-30 seconds"},
    "scan_step1":       {"zh": "识别中英对应关系",           "en": "Identifying CN/EN pairs"},
    "scan_step2":       {"zh": "四维风险逐条扫描",           "en": "Four-dimension scanning"},
    "scan_step3":       {"zh": "生成结构化报告",             "en": "Generating report"},
    # 结果区
    "risk_dashboard":   {"zh": "📊 风险仪表盘",             "en": "📊 Risk Dashboard"},
    "ai_verdict":       {"zh": "AI 综合评级",               "en": "AI Verdict"},
    "stat_total":       {"zh": "🔍 发现总数",               "en": "🔍 Total Issues"},
    "stat_high":        {"zh": "🔴 高危风险",               "en": "🔴 High Risk"},
    "stat_medium":      {"zh": "🟡 中危风险",               "en": "🟡 Medium Risk"},
    "stat_low":         {"zh": "🟢 低危风险",               "en": "🟢 Low Risk"},
    "dim_brief":        {"zh": "📋 四维简报",               "en": "📋 Dimension Brief"},
    "detail_title":     {"zh": "🔍 逐条详情",               "en": "🔍 Detailed Findings"},
    "filter_label":     {"zh": "按严重等级筛选",             "en": "Filter by Severity"},
    "filter_all":       {"zh": "📋 全部",                   "en": "📋 All"},
    "filter_high":      {"zh": "🔴 仅看高危",               "en": "🔴 High Only"},
    "filter_medium":    {"zh": "🟡 仅看中危",               "en": "🟡 Medium Only"},
    "filter_low":       {"zh": "🟢 仅看低危",               "en": "🟢 Low Only"},
    "dim_good":         {"zh": "该维度未发现明显风险，条款质量良好！", "en": "No significant risks found in this dimension."},
    "problem_analysis": {"zh": "❌ 问题分析",                "en": "❌ Problem"},
    "clause_compare":   {"zh": "📝 条款对比",                "en": "📝 Clause Compare"},
    "cn_original":      {"zh": "🇨🇳 中文",                  "en": "🇨🇳 Chinese"},
    "en_original":      {"zh": "🇬🇧 English",               "en": "🇬🇧 English"},
    "cn_delete":        {"zh": "🔴 中文原文（删除）",        "en": "🔴 CN Original (Delete)"},
    "cn_revised":       {"zh": "🟢 中文修订版（新增）",      "en": "🟢 CN Revised (Add)"},
    "en_delete":        {"zh": "🔴 英文原文（删除）",        "en": "🔴 EN Original (Delete)"},
    "en_revised":       {"zh": "🟢 英文修订版（新增）",      "en": "🟢 EN Revised (Add)"},
    "suggestion_label": {"zh": "💡 修改建议",                "en": "💡 Suggestion"},
    "revised_compare":  {"zh": "📝 修订前后对比",            "en": "📝 Before & After"},
    # 下载
    "download_section": {"zh": "📥 导出风险报告",            "en": "📥 Export Report"},
    "download_txt":     {"zh": "📄 下载报告（TXT格式）",     "en": "📄 Download TXT"},
    "download_pdf":     {"zh": "📕 下载报告（PDF格式）",     "en": "📕 Download PDF"},
    # 空状态
    "hero_title":       {"zh": "AI 智能扫描引擎 · 已就绪",  "en": "AI Scan Engine · Ready"},
    "hero_desc":        {"zh": "基于 DeepSeek 大语言模型，四大维度逐条扫描中英双语合同，10-30 秒生成专业风险审核报告",
                          "en": "Powered by DeepSeek LLM, 4-dimension bilingual contract scanning, professional risk report in 10-30s"},
    "core_capabilities": {"zh": "五大核心能力",              "en": "Five Core Capabilities"},
    "feat1_title":      {"zh": "中英逐条比对",               "en": "Bilingual Comparison"},
    "feat1_desc":       {"zh": "自动识别中英文条款对应关系，<br>逐句比对语义差异、数字偏差、<br>范围不一致等关键问题",
                          "en": "Auto-identifies CN/EN clause pairs,<br>compares semantics, numbers,<br>and scope discrepancies"},
    "feat2_title":      {"zh": "模糊表述检测",               "en": "Vague Term Detection"},
    "feat2_desc":       {"zh": "智能识别 reasonable、appropriate<br>等 vague terms，评估在实际争议中<br>可能导致的法律不确定性",
                          "en": "Detects vague terms like reasonable,<br>appropriate, etc., assessing legal<br>uncertainty in real disputes"},
    "feat3_title":      {"zh": "缺失条款发现",               "en": "Missing Clause Discovery"},
    "feat3_desc":       {"zh": "检测保密期限、知识产权归属、<br>管辖法律等关键条款是否在中英文<br>版本中完整覆盖",
                          "en": "Detects missing key clauses —<br>confidentiality period, IP ownership,<br>governing law, etc."},
    "feat4_title":      {"zh": "跨境合规审查",               "en": "Cross-Border Compliance"},
    "feat4_desc":       {"zh": "识别争议解决方式冲突、违约金<br>制度差异、数据跨境传输等中西方<br>商业法律冲突风险",
                          "en": "Identifies dispute resolution conflicts,<br>penalty doctrine differences, cross-border<br>data transfer compliance risks"},
    "feat5_title":      {"zh": "文件直接上传",               "en": "File Upload"},
    "feat5_desc":       {"zh": "支持上传 Word (.docx) / PDF / TXT<br>合同文件，自动解析文本，<br>无需手动复制粘贴",
                          "en": "Upload Word (.docx) / PDF / TXT<br>contract files; auto-parsed,<br>no manual copy-paste needed"},
    "empty_info":       {"zh": "👆 选择合同模板或上传文件后，点击 **「🔍 开始智能扫描」** 即可体验完整分析流程。",
                          "en": "👆 Select a template or upload a file, then click **「🔍 Start Scan」** to begin."},
    # 国际化补漏
    "header_tagline":   {"zh": "涉外商务合同 · 双语风险智能扫描引擎", "en": "Bilingual Contract Risk Scanner · AI-Powered"},
    "tab_mismatch":     {"zh": "🔤 中英文不对应",              "en": "🔤 Translation Mismatch"},
    "tab_ambiguous":    {"zh": "🌫️ 模糊条款",                  "en": "🌫️ Ambiguous Clauses"},
    "tab_missing":      {"zh": "🕳️ 缺失条款",                  "en": "🕳️ Missing Clauses"},
    "tab_compliance":   {"zh": "🌐 跨文化合规风险",            "en": "🌐 Cross-Border Compliance"},
    "gauge_subtitle":   {"zh": "风险指数",                     "en": "Risk Index"},
    "dim_unit":         {"zh": "个问题",                       "en": " issues"},
    "label_risk_level": {"zh": "风险等级",                     "en": "Risk Level"},
    "label_issue_count": {"zh": "问题数",                      "en": "Issues"},
    "severity_high":    {"zh": "高风险",                       "en": "High"},
    "severity_medium":  {"zh": "中等风险",                     "en": "Medium"},
    "severity_low":     {"zh": "低风险",                       "en": "Low"},
    "severity_high_short": {"zh": "高",                        "en": "H"},
    "severity_medium_short": {"zh": "中",                      "en": "M"},
    "severity_low_short":  {"zh": "低",                        "en": "L"},
    "pdf_report_title": {"zh": "CrossLingo 合同风险扫描报告",   "en": "CrossLingo Contract Risk Scan Report"},
    "pdf_stats_title":  {"zh": "统计概览",                     "en": "Statistics Overview"},
    "pdf_risk_score":   {"zh": "综合风险指数",                 "en": "Risk Index"},
    "pdf_dim_good":     {"zh": "该维度未发现明显风险",          "en": "No significant risks found"},
    # 页脚
    "footer_text":      {"zh": "涉外商务合同双语风险扫描器", "en": "Bilingual Contract Risk Scanner"},
    "missing_clause_placeholder": {"zh": "（条款缺失）", "en": "(Missing)"},
}

def t(key):
    """获取当前语言的文本"""
    item = L.get(key, {})
    return item.get(st.session_state.lang, item.get("zh", key))

# ========== 第2步：网页基本设置 ==========
st.set_page_config(
    page_title="CrossLingo - 涉外合同风险扫描器",
    page_icon="⚖️",
    layout="wide"
)

# ========== 第3步：CSS 主题 ==========
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

    :root {
        --navy-deep: #080F1F; --navy: #0F1F3D; --navy-mid: #1A3A6B; --navy-light: #2B5EA7;
        --gold: #C9A06C; --gold-bright: #D4AF37; --gold-pale: #F5E6D3; --gold-wash: #FDF8F2;
        --red: #C0392B; --amber: #E67E22; --green: #27AE60;
        --gray-50: #F8F9FB; --gray-100: #EDF0F4; --gray-300: #B0BAC5; --gray-500: #6B7B8D; --gray-700: #3A4A5C;
        --shadow-sm: 0 1px 3px rgba(8,15,31,0.04); --shadow-md: 0 4px 16px rgba(8,15,31,0.07);
        --shadow-lg: 0 8px 32px rgba(8,15,31,0.10); --shadow-xl: 0 12px 48px rgba(8,15,31,0.14);
        --radius-sm: 8px; --radius-md: 12px; --radius-lg: 16px; --radius-xl: 20px;
    }

    * { font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'PingFang SC', 'Microsoft YaHei', sans-serif; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header[data-testid="stHeader"] {background: transparent;}
    .stApp { background: linear-gradient(180deg, #F8F9FB 0%, #FFFFFF 100%); }

    /* 缩减 Streamlit 默认容器上下边距 */
    .block-container { padding-top: 1.8rem !important; padding-bottom: 0.7rem !important; }

    /* ===== 品牌Header ===== */
    .brand-header {
        background: linear-gradient(135deg, #080F1F 0%, #0F1F3D 30%, #1A3A6B 60%, #15335A 100%);
        background-size: 200% 200%;
        animation: headerShimmer 12s ease-in-out infinite;
        padding: 16px 22px; border-radius: var(--radius-xl); margin-bottom: 14px;
        display: flex; align-items: center; justify-content: space-between;
        box-shadow: var(--shadow-xl); position: relative; overflow: hidden;
    }
    .brand-header::before {
        content: ''; position: absolute; top: -50%; right: -10%;
        width: 300px; height: 300px;
        background: radial-gradient(circle, rgba(201,160,108,0.08) 0%, transparent 70%);
        pointer-events: none;
    }
    @keyframes headerShimmer {
        0%, 100% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
    }
    .brand-header .logo-area { display: flex; align-items: center; gap: 18px; z-index: 1; }
    .brand-header .logo-icon {
        font-size: 44px; background: rgba(255,255,255,0.08); backdrop-filter: blur(8px);
        width: 68px; height: 68px; border-radius: var(--radius-lg);
        display: flex; align-items: center; justify-content: center;
        border: 1px solid rgba(255,255,255,0.1);
    }
    .brand-header h1 { color: #FFFFFF; font-size: 30px; font-weight: 800; margin: 0; letter-spacing: -0.3px; }
    .brand-header .tagline { color: rgba(255,255,255,0.65); font-size: 14px; margin-top: 3px; font-weight: 400; }
    .brand-header .badge {
        background: rgba(201,160,108,0.15); border: 1px solid rgba(201,160,108,0.35);
        color: #D4AF37; padding: 10px 20px; border-radius: 24px;
        font-size: 12px; font-weight: 700; letter-spacing: 1.5px; text-transform: uppercase; z-index: 1;
    }

    /* ===== 功能卡片 ===== */
    .feature-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin: 8px 0 24px 0; }
    .feature-card {
        background: #FFFFFF; border-radius: var(--radius-lg); padding: 28px 24px;
        text-align: center; border: 1px solid var(--gray-100); box-shadow: var(--shadow-sm);
        transition: all 0.25s ease; cursor: default;
    }
    .feature-card:hover { transform: translateY(-4px); box-shadow: var(--shadow-lg); border-color: #C9A06C40; }
    .feature-card .feature-title { font-size: 15px; font-weight: 700; color: var(--navy); margin: 0 0 6px 0; }
    .feature-card .feature-desc { font-size: 12px; color: var(--gray-500); line-height: 1.5; margin: 0; }

    /* ===== 分割线 ===== */
    .section-divider { display: flex; align-items: center; gap: 10px; margin: 14px 0 10px 0; }
    .section-divider .divider-line { flex: 1; height: 1px; background: linear-gradient(90deg, transparent, #C9A06C40, transparent); }
    .section-divider .divider-diamond { color: var(--gold); font-size: 10px; }

    /* ===== 统计卡片 ===== */
    .stat-card {
        background: #FFFFFF; border-radius: var(--radius-lg); padding: 20px 24px;
        box-shadow: var(--shadow-sm); border: 1px solid var(--gray-100);
        text-align: center; transition: all 0.2s ease; position: relative; overflow: hidden;
    }
    .stat-card::before { content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; }
    .stat-card:hover { transform: translateY(-2px); box-shadow: var(--shadow-lg); }
    .stat-card .stat-number { font-size: 40px; font-weight: 900; margin: 0; line-height: 1; letter-spacing: -1px; }
    .stat-card .stat-label { font-size: 12px; color: var(--gray-500); margin-top: 6px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }
    .stat-card.stat-total::before { background: linear-gradient(90deg, var(--navy-mid), var(--navy-light)); }
    .stat-card.stat-high::before { background: linear-gradient(90deg, #E74C3C, var(--red)); }
    .stat-card.stat-medium::before { background: linear-gradient(90deg, #F39C12, var(--amber)); }
    .stat-card.stat-low::before { background: linear-gradient(90deg, #2ECC71, var(--green)); }
    .stat-total .stat-number { color: var(--navy); }
    .stat-high .stat-number { color: var(--red); }
    .stat-medium .stat-number { color: var(--amber); }
    .stat-low .stat-number { color: var(--green); }

    /* ===== 维度简报 ===== */
    .dimension-strip { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin: 20px 0; }
    .dimension-mini {
        background: #FFFFFF; border-radius: var(--radius-md); padding: 16px 18px;
        border: 1px solid var(--gray-100); box-shadow: var(--shadow-sm); transition: all 0.2s ease;
    }
    .dimension-mini:hover { box-shadow: var(--shadow-md); border-color: #C9A06C30; }
    .dimension-mini .dim-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; }
    .dimension-mini .dim-icon { font-size: 22px; }
    .dimension-mini .dim-count { font-size: 24px; font-weight: 800; color: var(--navy); }
    .dimension-mini .dim-name { font-size: 12px; font-weight: 600; color: var(--gray-700); }
    .dimension-mini .dim-bar-bg { height: 4px; background: var(--gray-100); border-radius: 2px; margin-top: 8px; overflow: hidden; }
    .dimension-mini .dim-bar-fill { height: 100%; border-radius: 2px; transition: width 1s ease; }

    /* ===== 徽章 ===== */
    .severity-badge { display: inline-block; padding: 4px 12px; border-radius: 14px; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
    .severity-high { background: #FDECEA; color: #B03A2E; border: 1px solid #F5C6CB; }
    .severity-medium { background: #FEF5E7; color: #B9770E; border: 1px solid #FAE5C3; }
    .severity-low { background: #E8F8F0; color: #1E8449; border: 1px solid #C5E8D3; }

    /* ===== 对比框 ===== */
    .compare-box {
        background: var(--gray-50); border-radius: var(--radius-sm); padding: 14px 18px;
        margin: 8px 0; font-size: 13px; line-height: 1.6;
        border-left: 3px solid var(--gold); color: var(--gray-700);
    }
    .compare-box .lang-tag { font-size: 10px; font-weight: 700; color: var(--navy-mid); text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 4px; }

    /* ===== 按钮 ===== */
    .stButton > button { border-radius: var(--radius-md) !important; font-weight: 700 !important; font-size: 16px !important; letter-spacing: 0.3px !important; transition: all 0.2s ease !important; padding: 12px 32px !important; background: linear-gradient(135deg, var(--navy-mid), var(--navy-light)) !important; border: none !important; color: #FFFFFF !important; }
    .stButton > button:hover { transform: translateY(-2px); box-shadow: 0 8px 24px rgba(15,31,61,0.3) !important; background: linear-gradient(135deg, var(--navy-light), #3B6EC4) !important; }
    .stButton > button:disabled { background: var(--gray-300) !important; color: #FFFFFF !important; transform: none !important; box-shadow: none !important; }
    .stDownloadButton > button { border-radius: var(--radius-md) !important; font-weight: 600 !important; border: 2px solid var(--navy-mid) !important; background: #FFFFFF !important; color: var(--navy-mid) !important; transition: all 0.2s ease !important; }
    .stDownloadButton > button:hover { background: var(--navy-mid) !important; color: #FFFFFF !important; transform: translateY(-1px); }

    /* ===== 侧边栏 ===== */
    section[data-testid="stSidebar"] { background: linear-gradient(180deg, #FAFBFC 0%, #FFFFFF 100%); border-right: 1px solid var(--gray-100); }
    .streamlit-expanderHeader { font-weight: 600 !important; font-size: 14px !important; border-radius: var(--radius-sm) !important; }
    .streamlit-expanderHeader:hover { background: var(--gray-50) !important; }

    /* ===== 响应式 ===== */
    @media (max-width: 1024px) {
        .feature-grid, .dimension-strip { grid-template-columns: repeat(2, 1fr); }
        .stat-card .stat-number { font-size: 32px; }
        .brand-header { padding: 14px 18px; }
        .brand-header h1 { font-size: 24px; }
    }
    @media (max-width: 768px) {
        /* 网格和卡片 */
        .feature-grid, .dimension-strip { grid-template-columns: 1fr; gap: 10px; }
        .feature-card { padding: 20px 16px; }
        .feature-card .feature-icon-wrap { width: 44px; height: 44px; font-size: 22px; }
        .feature-card .feature-title { font-size: 14px; }
        .stat-card { padding: 14px 10px; }
        .stat-card .stat-number { font-size: 28px; }
        .stat-card .stat-label { font-size: 10px; }
        /* Header */
        .brand-header { flex-direction: column; gap: 10px; text-align: center; }
        .brand-header h1 { font-size: 20px; }
        .brand-header .badge { font-size: 10px; padding: 6px 14px; }
        .brand-header .logo-icon { width: 48px; height: 48px; font-size: 32px; }
        /* 仪表盘 */
        .ring-gauge-container { flex-direction: column; align-items: center; }
        /* 加载动画 */
        .scan-overlay { padding: 24px 12px; }
        .scan-steps { gap: 16px; }
        .scan-title { font-size: 16px; }
        /* Hero */
        .hero-strip { flex-direction: column; text-align: center; padding: 16px 14px; }
        .hero-strip .hero-icon { font-size: 36px; }
        .hero-strip .hero-text h3 { font-size: 16px; }
        /* 按钮全宽 */
        .stButton > button { width: 100% !important; }
        /* 缩小间距 */
        .block-container { padding-top: 1rem !important; padding-bottom: 0.5rem !important; }
        .section-divider { margin: 10px 0 8px 0; }
        /* 标签页字号 */
        .stTabs [data-baseweb="tab"] { font-size: 12px !important; }
    }
    @media (max-width: 480px) {
        .brand-header { border-radius: var(--radius-md); }
        .brand-header h1 { font-size: 18px; }
        .brand-header .tagline { font-size: 11px; }
        .stat-card { padding: 10px 6px; }
        .stat-card .stat-number { font-size: 24px; }
        .stat-card .stat-label { font-size: 9px; }
        .dimension-mini { padding: 10px 12px; }
        .dimension-mini .dim-count { font-size: 20px; }
        .scan-steps { flex-direction: column; gap: 8px; }
        .scan-overlay { padding: 16px 8px; }
        /* 侧边栏收起时调整 */
        .st-emotion-cache-1cypcdb { font-size: 12px; }
    }

    .app-footer { text-align: center; color: var(--gray-500); font-size: 11px; padding: 10px 0 14px 0; letter-spacing: 0.2px; }
    .app-footer b { color: var(--navy); }

    /* ===== 扫描加载动画 ===== */
    .scan-overlay {
        text-align: center; padding: 32px 20px;
        background: linear-gradient(180deg, #FDF8F2 0%, #FFFFFF 100%);
        border-radius: var(--radius-xl); border: 1px solid #E8D5B8;
        margin: 24px 0;
    }
    .pulse-container {
        position: relative; width: 100px; height: 100px;
        margin: 0 auto 32px auto;
    }
    .pulse-ring {
        position: absolute; top: 0; left: 0; width: 100px; height: 100px;
        border-radius: 50%; border: 3px solid var(--gold);
        animation: pulseExpand 1.8s cubic-bezier(0.25, 0.46, 0.45, 0.94) infinite;
        opacity: 0;
    }
    .pulse-ring:nth-child(2) { animation-delay: 0.6s; }
    .pulse-ring:nth-child(3) { animation-delay: 1.2s; }
    .pulse-center {
        position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%);
        width: 48px; height: 48px; border-radius: 50%;
        background: linear-gradient(135deg, var(--navy-mid), var(--navy-light));
        display: flex; align-items: center; justify-content: center;
        box-shadow: 0 4px 20px rgba(15,31,61,0.35);
    }
    .pulse-center .pulse-icon { font-size: 24px; animation: iconPulse 1.8s ease-in-out infinite; }
    @keyframes pulseExpand {
        0% { transform: scale(0.5); opacity: 1; }
        100% { transform: scale(1.8); opacity: 0; }
    }
    @keyframes iconPulse {
        0%, 100% { transform: scale(1); }
        50% { transform: scale(1.2); }
    }
    .scan-title {
        font-size: 20px; font-weight: 700; color: var(--navy);
        margin-bottom: 8px;
    }
    .scan-subtitle {
        font-size: 13px; color: var(--gray-500);
    }
    .scan-steps {
        display: flex; justify-content: center; gap: 40px;
        margin-top: 28px; flex-wrap: wrap;
    }
    .scan-step {
        text-align: center; opacity: 0.6;
        animation: stepFade 2.4s ease-in-out infinite;
    }
    .scan-step:nth-child(2) { animation-delay: 0.8s; }
    .scan-step:nth-child(3) { animation-delay: 1.6s; }
    .scan-step .step-num {
        width: 32px; height: 32px; border-radius: 50%;
        background: var(--gray-100); color: var(--gray-500);
        display: flex; align-items: center; justify-content: center;
        font-weight: 700; font-size: 14px; margin: 0 auto 8px auto;
    }
    .scan-step .step-label {
        font-size: 11px; color: var(--gray-500); font-weight: 500;
    }
    @keyframes stepFade {
        0%, 100% { opacity: 0.4; }
        50% { opacity: 1; }
    }

    /* ===== 功能卡片增强 ===== */
    .feature-card .feature-icon-wrap {
        width: 56px; height: 56px; border-radius: var(--radius-md);
        display: flex; align-items: center; justify-content: center;
        margin: 0 auto 14px auto;
        font-size: 28px;
    }
    .feature-card:nth-child(1) .feature-icon-wrap { background: #EEF2FA; }
    .feature-card:nth-child(2) .feature-icon-wrap { background: #FDF3E8; }
    .feature-card:nth-child(3) .feature-icon-wrap { background: #EDF0F4; }
    .feature-card:nth-child(4) .feature-icon-wrap { background: #E8F0FA; }

    /* ===== 空状态区域增强 ===== */
    .hero-strip {
        display: flex; align-items: center; gap: 20px;
        padding: 20px 28px; margin-bottom: 8px;
        background: linear-gradient(135deg, #FDF8F2, #FFFFFF);
        border-radius: var(--radius-lg); border: 1px solid #E8D5B8;
    }
    .hero-strip .hero-icon {
        font-size: 48px; filter: drop-shadow(0 2px 4px rgba(201,160,108,0.3));
    }
    .hero-strip .hero-text h3 {
        font-size: 18px; font-weight: 700; color: var(--navy); margin: 0;
    }
    .hero-strip .hero-text p {
        font-size: 13px; color: var(--gray-500); margin: 4px 0 0 0;
    }

    /* ===== 顶部毛玻璃效果 ===== */
    .top-frosted-bar {
        position: fixed;
        top: 0;
        left: 0;
        right: 0;
        height: 110px;
        z-index: 9998;
        backdrop-filter: blur(18px);
        -webkit-backdrop-filter: blur(18px);
        background: rgba(248, 249, 251, 0.5);
        mask-image: linear-gradient(to bottom, black 25%, black 55%, transparent 95%);
        -webkit-mask-image: linear-gradient(to bottom, black 25%, black 55%, transparent 95%);
        pointer-events: none;
    }
</style>
""", unsafe_allow_html=True)

# ========== 第3.5步：顶部毛玻璃遮罩 ==========
st.markdown('<div class="top-frosted-bar"></div>', unsafe_allow_html=True)

# ========== 第4步：品牌Header ==========
st.markdown(f"""
<div class="brand-header">
    <div class="logo-area">
        <div class="logo-icon">⚖️</div>
        <div>
            <h1>CrossLingo</h1>
            <div class="tagline">{t('header_tagline')}</div>
        </div>
    </div>
    <div class="badge">AI-Powered LegalTech</div>
</div>
""", unsafe_allow_html=True)

# ========== 第5步：合同模板库 ==========
CONTRACT_TEMPLATES = {
    "采购框架协议 (Procurement)": """【中英双语采购框架协议 / Bilateral Procurement Framework Agreement】

第1条 交货期限 / Article 1 Delivery Period
中文：卖方应在合同生效后30日内完成交货。
English: The Seller shall complete delivery within a reasonable time after the contract takes effect.

第2条 违约责任 / Article 2 Liability for Breach
中文：若卖方逾期交货，每逾期一日，应向买方支付合同总金额0.5%的违约金，违约金总额不超过合同总金额的20%。
English: If the Seller delays delivery, the Seller shall pay a penalty to the Buyer.

第3条 争议解决 / Article 3 Dispute Resolution
中文：双方因本合同产生的任何争议，应首先友好协商解决；协商不成的，任何一方均有权向买方所在地有管辖权的人民法院提起诉讼。
English: Any dispute arising from this contract shall be resolved through arbitration in London, UK.

第4条 不可抗力 / Article 4 Force Majeure
中文：（本条内容缺失）

第5条 保密条款 / Article 5 Confidentiality
中文：双方应对合作过程中知悉的对方商业秘密严格保密，保密义务在合同终止后五年内继续有效。
English: Both parties shall keep confidential information secret for a reasonable period.

第6条 验收标准 / Article 6 Acceptance Criteria
中文：货物应符合附件A所列技术规格，买方有权在收货后15个工作日内进行验收，验收不合格的，买方有权拒收并要求卖方在10个工作日内补货或退款。
English: The goods should meet the specifications. The Buyer may inspect the goods upon receipt.""",

    "服务外包协议 (Service Agreement)": """【中英双语服务外包协议 / Bilingual Service Outsourcing Agreement】

第1条 服务范围 / Article 1 Scope of Services
中文：乙方应根据附件A中列明的服务内容和标准，向甲方提供软件开发与运维支持服务。
English: Party B shall provide software development and maintenance support services to Party A in accordance with the contents and standards set forth in Appendix A.

第2条 服务费用与支付 / Article 2 Service Fees and Payment
中文：甲方应按月向乙方支付服务费人民币50万元整，每月5日前支付当月费用。逾期支付的，每逾期一日按应付金额的千分之一支付滞纳金。
English: Party A shall pay Party B a monthly service fee of RMB 500,000. Payment shall be made within a reasonable time each month.

第3条 知识产权归属 / Article 3 Intellectual Property Rights
中文：乙方在履行本合同过程中产生的所有工作成果的知识产权，自产生之日起即归属于甲方。
English: Intellectual property rights relating to the deliverables shall be discussed by both parties.

第4条 数据保护 / Article 4 Data Protection
中文：乙方在处理甲方提供的个人数据时，应遵守《中华人民共和国个人信息保护法》及相关法律法规，不得将数据转移至境外。
English: Party B shall handle personal data provided by Party A in accordance with applicable laws.

第5条 责任限制 / Article 5 Limitation of Liability
中文：乙方因履行本合同给甲方造成损失的，赔偿总额不超过甲方已支付的服务费总额。
English: Party B's total liability shall not in any event exceed the total fees paid by Party A, except for cases of gross negligence or willful misconduct.""",

    "保密协议 (NDA)": """【中英双语保密协议 / Bilateral Non-Disclosure Agreement】

第1条 保密信息定义 / Article 1 Definition of Confidential Information
中文：保密信息是指一方向另一方披露的、与本协议目的相关的所有非公开信息，包括但不限于商业计划、客户名单、技术资料、财务数据及任何标明为"保密"的信息。
English: Confidential Information means any non-public information disclosed by one party to the other, including business plans, customer lists, and technical data.

第2条 保密义务 / Article 2 Obligations of Confidentiality
中文：接收方应对保密信息严格保密，未经披露方书面同意，不得向任何第三方披露。保密义务在本协议终止后五年内继续有效。
English: The Receiving Party shall keep the Confidential Information confidential for a reasonable period after termination of this Agreement.

第3条 使用限制 / Article 3 Use Restriction
中文：接收方仅可为评估双方潜在合作之目的使用保密信息，不得用于任何其他目的。
English: The Receiving Party may use the Confidential Information for evaluation purposes and other reasonable business purposes.

第4条 违约责任 / Article 4 Breach of Agreement
中文：若接收方违反本协议任何条款，披露方有权要求接收方赔偿因此造成的全部损失，并有权寻求禁令救济。
English: If the Receiving Party breaches this Agreement, the Disclosing Party may seek appropriate remedies.

第5条 管辖法律 / Article 5 Governing Law
中文：本协议适用中华人民共和国法律。因本协议产生的争议，提交披露方所在地有管辖权的人民法院管辖。
English: This Agreement shall be governed by the laws of a jurisdiction to be agreed by the parties.""",

    "劳动合同 (Employment Contract)": """【中英双语劳动合同 / Bilingual Employment Contract】

第1条 聘用期限 / Article 1 Term of Employment
中文：本合同为固定期限劳动合同，自2026年1月1日起至2028年12月31日止，试用期为三个月。
English: This Contract is a fixed-term employment contract commencing on January 1, 2026 and ending on December 31, 2028, with a probation period to be determined.

第2条 工作内容与地点 / Article 2 Job Duties and Work Location
中文：乙方担任市场总监职务，工作地点为深圳市南山区。甲方可根据业务需要，在与乙方协商后调整工作岗位。
English: Party B shall serve as Marketing Director. The work location shall be determined by Party A.

第3条 薪酬与福利 / Article 3 Compensation and Benefits
中文：乙方税前月薪为人民币30,000元。甲方依法为乙方缴纳社会保险和住房公积金。年终奖金根据公司当年经营状况及乙方绩效考核结果确定。
English: Party B's monthly salary is RMB 30,000 before tax. Party A shall provide social insurance and housing fund in accordance with laws. Bonus is discretionary.

第4条 竞业限制 / Article 4 Non-Competition
中文：乙方离职后两年内，不得在与甲方有竞争关系的企业任职或为其提供服务。甲方在竞业限制期内按月向乙方支付经济补偿，补偿标准为离职前十二个月平均工资的30%。
English: Party B shall not work for any competitor of Party A for a reasonable period after termination of employment.

第5条 终止与解除 / Article 5 Termination
中文：任何一方解除本合同，应提前三十日以书面形式通知对方。甲方违法解除合同的，应向乙方支付赔偿金。
English: Either party may terminate this Contract by giving notice to the other party. Party A may terminate this Contract at any time with appropriate compensation.""",

    "数据跨境传输协议 (Data Transfer)": """【中英双语数据跨境传输协议 / Bilateral Cross-Border Data Transfer Agreement】

第1条 数据范围与目的 / Article 1 Scope and Purpose
中文：本协议适用于数据提供方为向境外接收方提供客户个人信息而进行的跨境数据传输活动。传输目的仅限于履行双方签订的IT技术支持服务合同。
English: This Agreement applies to cross-border data transfers from the Data Provider to the Overseas Recipient for the purpose of providing IT support services and other related business activities.

第2条 数据安全措施 / Article 2 Data Security Measures
中文：境外接收方应采取与中华人民共和国《个人信息保护法》要求同等级别的安全保护措施，包括但不限于加密传输、访问控制、去标识化处理。发生数据泄露时，应在24小时内通知数据提供方。
English: The Overseas Recipient shall implement reasonable security measures to protect the data. In the event of a data breach, the Overseas Recipient shall notify the Data Provider in a timely manner.

第3条 数据主体权利 / Article 3 Data Subject Rights
中文：境外接收方应协助数据提供方在15个工作日内响应数据主体的查阅、更正、删除及数据可携权请求。不得以任何理由拒绝数据主体行使法律赋予的权利。
English: The Overseas Recipient shall assist the Data Provider in responding to data subject requests within a reasonable time frame.

第4条 再转移限制 / Article 4 Restrictions on Onward Transfer
中文：未经数据提供方书面同意，境外接收方不得将数据转移至任何第三方。经同意的再转移，境外接收方应与第三方签订同等保护水平的书面协议，并就该第三方的违约行为承担连带责任。
English: The Overseas Recipient shall not transfer the data to any third party without consent. Where onward transfer is permitted, appropriate safeguards shall be in place.

第5条 管辖法律与争议解决 / Article 5 Governing Law and Dispute Resolution
中文：本协议适用中华人民共和国法律。因本协议产生的任何争议，提交深圳市有管辖权的人民法院管辖。
English: This Agreement shall be governed by the laws of a neutral jurisdiction. Any disputes shall be resolved through arbitration.""",

    "合资经营协议 (Joint Venture)": """【中英双语合资经营协议 / Bilingual Joint Venture Agreement】

第1条 合资公司设立 / Article 1 Establishment of JV Company
中文：甲乙双方共同出资在中国深圳设立合资经营企业，甲方出资比例为60%，乙方出资比例为40%。合资公司注册资本为人民币1,000万元，双方应在营业执照签发之日起90日内缴清各自认缴的出资额。
English: Party A and Party B shall establish a joint venture company in Shenzhen, China. Party A shall contribute 60% and Party B 40% of the registered capital. Contributions shall be made within a reasonable period after the business license is issued.

第2条 利润分配 / Article 2 Profit Distribution
中文：合资公司税后利润在提取法定公积金后，按双方实缴出资比例进行分配。每会计年度结束后三个月内完成利润分配。
English: After-tax profits of the JV Company shall be distributed between the parties in accordance with their respective contribution ratios. Distribution shall be completed in a timely manner after the end of each fiscal year.

第3条 股权转让限制 / Article 3 Transfer Restrictions
中文：任何一方转让合资公司全部或部分股权，须经另一方书面同意。另一方在同等条件下享有优先购买权。未经另一方同意的转让无效。转让价格不得低于届时经第三方评估机构评估的公允价值。
English: Either party may transfer its equity interests in the JV Company with the consent of the other party. The other party shall have a right of first refusal.

第4条 不竞争义务 / Article 4 Non-Competition
中文：合资各方在合资公司存续期间及退出后三年内，不得在中国境内直接或间接从事与合资公司业务相竞争的活动。
English: During the term of the JV Company, each party shall not engage in any business that competes with the JV Company. After exiting, the non-competition obligation shall continue for a reasonable period.

第5条 清算与解散 / Article 5 Liquidation and Dissolution
中文：合资公司解散时，清算组应在成立后60日内完成清算。剩余财产按双方实缴出资比例分配。因一方违约导致合资公司解散的，违约方应赔偿守约方因此遭受的全部损失。
English: Upon dissolution of the JV Company, liquidation shall be carried out in accordance with applicable laws. Remaining assets shall be distributed to the parties.""",

    "授权委托书 (Power of Attorney)": """【中英双语授权委托书 / Bilingual Power of Attorney】

第1条 授权事项 / Article 1 Scope of Authorization
中文：委托人兹授权受托人代表委托人办理以下事项：代表委托人参加在中国广州举办的国际进出口商品交易会，签署意向性采购协议，以及处理与交易会相关的展位租赁、广告宣传等事宜。授权期限自2026年1月1日起至2026年12月31日止。
English: The Principal hereby authorizes the Agent to attend the International Import and Export Trade Fair on behalf of the Principal, sign procurement agreements, and handle other related matters. This Power of Attorney shall remain in effect until revoked.

第2条 代理人权限 / Article 2 Authority of Agent
中文：受托人在授权范围内签署的法律文件，对委托人具有约束力。受托人不得将授权事项转委托给第三方，除非获得委托人书面同意。单笔采购金额超过人民币100万元的，须事先取得委托人书面确认。
English: The Agent may sign documents on behalf of the Principal within the scope of this authorization. The Agent may delegate its authority to third parties as necessary. For purchases exceeding a certain amount, the Agent shall seek the Principal's confirmation.

第3条 代理人义务 / Article 3 Obligations of Agent
中文：受托人应以善良管理人的注意义务处理委托事务，及时向委托人报告处理情况，并妥善保管与委托事务相关的所有文件和资金。
English: The Agent shall handle the authorized matters with due care, report to the Principal on a regular basis, and properly keep all relevant documents and funds.

第4条 责任与赔偿 / Article 4 Liability and Indemnification
中文：受托人因超越授权范围或故意或重大过失给委托人造成损失的，应承担全部赔偿责任。委托人应赔偿受托人在授权范围内行事所产生的合理费用和损失。
English: The Agent shall be liable for losses caused by acting beyond the scope of authorization. The Principal shall indemnify the Agent for losses incurred while acting within the scope of authorization.

第5条 终止 / Article 5 Termination
中文：委托人可随时书面通知受托人撤销本授权。受托人可随时书面通知委托人辞去授权。撤销或辞去授权自通知送达对方之日起生效。
English: This Power of Attorney may be terminated by either party. Termination shall take effect upon notice to the other party."""
}

# ========== 第6步：文件解析函数 ==========
def extract_text_from_file(uploaded_file):
    """从上传的文件中提取文本内容（支持 .txt / .docx / .pdf）"""
    filename = uploaded_file.name.lower()

    if filename.endswith(".txt"):
        return uploaded_file.getvalue().decode("utf-8", errors="ignore")

    elif filename.endswith(".docx"):
        doc = Document(io.BytesIO(uploaded_file.getvalue()))
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])

    elif filename.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(uploaded_file.getvalue()))
        text_parts = []
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text_parts.append(t)
        return "\n".join(text_parts)

    else:
        return None  # 不支持的文件格式

# ========== 第7步：侧边栏 ==========
with st.sidebar:
    # 语言切换
    lang_label = "🌐 中文" if st.session_state.lang == "zh" else "🌐 English"
    if st.button(lang_label, use_container_width=True, key="lang_toggle"):
        st.session_state.lang = "en" if st.session_state.lang == "zh" else "zh"
        st.rerun()

    st.markdown(f"### {t('ai_engine')}")

    # 优先 Streamlit Cloud secrets，其次 .env 本地文件
    try:
        saved_key = st.secrets["DEEPSEEK_API_KEY"]
    except Exception:
        saved_key = os.getenv("DEEPSEEK_API_KEY", "")

    if saved_key:
        masked = saved_key[:5] + "****" + saved_key[-4:]
        st.success(f"{t('key_configured')}（{masked}）")
        api_key = saved_key
        with st.expander(t("change_key")):
            manual_key = st.text_input(t("enter_new_key"), type="password", placeholder="sk-xxxxxxxxxxxxxxxx")
            if manual_key:
                api_key = manual_key
                st.info(t("temp_key_hint"))
    else:
        st.caption(t("key_tip"))
        api_key = st.text_input("DeepSeek API Key", type="password", placeholder="sk-xxxxxxxxxxxxxxxx", help=t("key_placeholder"))

    st.markdown("---")
    st.markdown(f"### {t('scan_dims')}")
    st.markdown(f"""
    <div style="font-size:13px; line-height:2.2;">
    🔤 &nbsp;<b>{t('dim_mismatch')}</b><br>
    <span style="color:#6B7B8D; font-size:11px;">{t('dim_mismatch_desc')}</span><br>
    🌫️ &nbsp;<b>{t('dim_ambiguous')}</b><br>
    <span style="color:#6B7B8D; font-size:11px;">{t('dim_ambiguous_desc')}</span><br>
    🕳️ &nbsp;<b>{t('dim_missing')}</b><br>
    <span style="color:#6B7B8D; font-size:11px;">{t('dim_missing_desc')}</span><br>
    🌐 &nbsp;<b>{t('dim_compliance')}</b><br>
    <span style="color:#6B7B8D; font-size:11px;">{t('dim_compliance_desc')}</span>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
    st.caption("⚖️ CrossLingo v5.0 | Powered by DeepSeek")

# ========== 第8步：主区域 — 合同输入 ==========
st.markdown(f"### {t('contract_input')}")

# ---- 工具栏：模板切换 + 文件上传 ----
tool_col1, tool_col2 = st.columns([1, 1])

with tool_col1:
    template_choice = st.selectbox(
        t("select_template"),
        list(CONTRACT_TEMPLATES.keys()),
        index=0
    )

with tool_col2:
    uploaded_file = st.file_uploader(
        t("upload_file"),
        type=["txt", "docx", "pdf"],
        help=t("upload_hint")
    )

# 决定文本输入框的内容
if uploaded_file is not None:
    extracted = extract_text_from_file(uploaded_file)
    if extracted:
        st.success(f"✅ {t('parse_success')}「{uploaded_file.name}」—— {len(extracted)} {t('parse_chars')}")
        default_text = extracted
    else:
        st.error(f"❌ {t('parse_error')}「{uploaded_file.name}」，{t('parse_only')}")
        default_text = CONTRACT_TEMPLATES[template_choice]
else:
    default_text = CONTRACT_TEMPLATES[template_choice]

contract_text = st.text_area(
    t("contract_label"),
    value=default_text,
    height=300,
    placeholder=t("contract_placeholder")
)

# ========== 第9步：扫描按钮 ==========
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    scan_button = st.button(
        t("scan_btn"),
        type="primary",
        use_container_width=True,
        disabled=(not api_key)
    )

if not api_key:
    st.warning(t("no_key_warn"))

# ========== 第10步：辅助函数 ==========

def calculate_risk_score(categories):
    total_findings = 0
    weighted_sum = 0
    for key, cat in categories.items():
        for f in cat.get("findings", []):
            total_findings += 1
            sev = f.get("severity", "低")
            weighted_sum += {"高": 3, "中": 2, "低": 1}.get(sev, 1)
    if total_findings == 0:
        return 0
    return min(int((weighted_sum / (total_findings * 3)) * 100), 100)

def count_by_severity(categories, severity):
    return sum(1 for cat in categories.values()
               for f in cat.get("findings", []) if f.get("severity") == severity)

def build_ring_gauge_svg(score, size=160):
    r, stroke_w = 52, 14
    circumference = 2 * 3.14159265 * r
    offset = circumference * (1 - score / 100)

    if score >= 66:
        gid, sc, ec = "gaugeRed", "#E74C3C", "#C0392B"
    elif score >= 33:
        gid, sc, ec = "gaugeAmber", "#F39C12", "#E67E22"
    else:
        gid, sc, ec = "gaugeGreen", "#2ECC71", "#27AE60"

    return f"""
    <svg width="{size}" height="{size}" viewBox="0 0 {size} {size}">
        <defs>
            <linearGradient id="{gid}" x1="0%" y1="0%" x2="100%" y2="0%">
                <stop offset="0%" stop-color="{sc}" /><stop offset="100%" stop-color="{ec}" />
            </linearGradient>
            <filter id="glow"><feGaussianBlur stdDeviation="2.5" result="blur"/>
                <feMerge><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge>
            </filter>
        </defs>
        <circle cx="{size//2}" cy="{size//2}" r="{r}" fill="none" stroke="#EDF0F4" stroke-width="{stroke_w}"/>
        <circle cx="{size//2}" cy="{size//2}" r="{r}" fill="none" stroke="url(#{gid})" stroke-width="{stroke_w}"
                stroke-dasharray="{circumference:.1f}" stroke-dashoffset="{offset:.1f}"
                stroke-linecap="round" transform="rotate(-90 {size//2} {size//2})" filter="url(#glow)"
                style="transition: stroke-dashoffset 1.2s cubic-bezier(0.25, 0.46, 0.45, 0.94);"/>
        <text x="{size//2}" y="{size//2}" text-anchor="middle" dominant-baseline="central"
              fill="#0F1F3D" font-size="42" font-weight="900" letter-spacing="-1">{score}</text>
        <text x="{size//2}" y="{size//2 + 26}" text-anchor="middle" dominant-baseline="central"
              fill="#6B7B8D" font-size="11" font-weight="600" letter-spacing="0.3">{t('gauge_subtitle')}</text>
    </svg>"""

def generate_pdf_report(result, risk_score, score_label, total_issues, high_count, medium_count, low_count):
    """生成 PDF 报告。本地 macOS 支持中文，云端回退英文标签 + 过滤中文内容。"""
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()

    FONT = "/System/Library/Fonts/STHeiti Medium.ttc"
    try:
        pdf.add_font("cjk", "", FONT)
    except Exception:
        pass
    has_cjk = "cjk" in pdf.fonts
    F = "cjk" if has_cjk else "Helvetica"

    # 无中文字体时，过滤中文内容防止 Unicode 崩溃
    def safe(s):
        if has_cjk:
            return s
        return re.sub(r'[一-鿿　-〿＀-￯]+', '[CN]', str(s))

    # 标签文本
    TITLE   = t("pdf_report_title") if has_cjk else "CrossLingo Contract Risk Scan Report"
    RISK    = t("pdf_risk_score") if has_cjk else "Risk Index"
    STATS   = t("pdf_stats_title") if has_cjk else "Statistics"
    SL = [("总数","Total"),("高危","High"),("中危","Medium"),("低危","Low")]
    DL = [(t("dim_mismatch") if has_cjk else "Translation Mismatch"),
          (t("dim_ambiguous") if has_cjk else "Ambiguous Clauses"),
          (t("dim_missing") if has_cjk else "Missing Clauses"),
          (t("dim_compliance") if has_cjk else "Cross-Border Compliance")]
    RT = "风险" if has_cjk else "Risk"
    IT = "问题" if has_cjk else "Issue"
    CT = "个" if has_cjk else ""
    PT = "问题" if has_cjk else "Problem"
    ST = "建议" if has_cjk else "Suggestion"
    CRT = "中文修订" if has_cjk else "CN Revised"
    ERT = "英文修订" if has_cjk else "EN Revised"
    GT  = t("pdf_dim_good") if has_cjk else "No significant risks found"

    # ---- 页眉 ----
    pdf.set_fill_color(15, 31, 61)
    pdf.rect(0, 0, 210, 34, "F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font(F, "", 20)
    pdf.set_xy(14, 9)
    pdf.cell(0, 10, TITLE)
    pdf.set_font(F, "", 8)
    pdf.set_xy(14, 22)
    pdf.cell(0, 6, f"{datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Powered by DeepSeek AI")
    pdf.ln(28)

    # ---- 风险指数 ----
    pdf.set_text_color(15, 31, 61)
    pdf.set_font(F, "", 16)
    pdf.cell(0, 10, f"{RISK}：{risk_score}/100（{safe(score_label)}）", ln=True)
    pdf.ln(4)
    bar_x, bar_y, bar_w, bar_h = 14, pdf.get_y(), 180, 5
    bc = (192,57,43) if risk_score>=66 else ((230,126,34) if risk_score>=33 else (39,174,96))
    pdf.set_fill_color(*bc)
    pdf.rect(bar_x, bar_y, bar_w*risk_score/100, bar_h, "F")
    pdf.set_draw_color(220, 220, 220)
    pdf.rect(bar_x, bar_y, bar_w, bar_h, "D")
    pdf.ln(12)

    # ---- 总体评价 ----
    pdf.set_fill_color(248, 249, 251)
    pdf.set_text_color(58, 74, 92)
    pdf.set_font(F, "", 10)
    pdf.set_x(14)
    pdf.multi_cell(0, 6, safe(result.get("overall_summary", "")), fill=True)
    pdf.ln(4)

    # ---- 统计卡片 ----
    pdf.set_text_color(15, 31, 61)
    pdf.set_font(F, "", 13)
    pdf.cell(0, 8, STATS, ln=True)
    pdf.ln(4)
    for i, (label, value, color) in enumerate([
        (SL[0][0] if has_cjk else SL[0][1], total_issues, (58,74,92)),
        (SL[1][0] if has_cjk else SL[1][1], high_count, (192,57,43)),
        (SL[2][0] if has_cjk else SL[2][1], medium_count, (230,126,34)),
        (SL[3][0] if has_cjk else SL[3][1], low_count, (39,174,96)),
    ]):
        x = 14 + i*46
        pdf.set_fill_color(248, 249, 251)
        pdf.set_xy(x, pdf.get_y())
        pdf.set_text_color(*color)
        pdf.set_font(F, "", 20)
        pdf.cell(42, 10, str(value), align="C")
        pdf.set_text_color(107, 123, 141)
        pdf.set_font(F, "", 7)
        pdf.set_xy(x, pdf.get_y()+10)
        pdf.cell(42, 5, label, align="C")
    pdf.ln(24)

    # ---- 详情 ----
    for key, dim_name in zip(
        ["translation_mismatch","ambiguous_clauses","missing_clauses","compliance_risk"], DL
    ):
        cat = result.get("categories", {}).get(key, {})
        findings = cat.get("findings", [])
        rl = cat.get("risk_level", "低")

        pdf.set_fill_color(245, 246, 248)
        pdf.set_text_color(15, 31, 61)
        pdf.set_font(F, "", 11)
        pdf.cell(0, 8, f"  {dim_name}  |  {RT}：{rl}  |  {IT}：{len(findings)}{CT}", ln=True, fill=True)
        pdf.ln(3)

        if findings:
            for f in findings:
                if pdf.get_y() > 255:
                    pdf.add_page()
                sev = f.get("severity", "低")
                sc = {"高":(192,57,43),"中":(230,126,34),"低":(39,174,96)}.get(sev,(0,0,0))
                pdf.set_text_color(*sc)
                pdf.set_font(F, "", 9)
                pdf.cell(0, 6, f"[{sev}{RT}] {safe(f.get('clause',''))}", ln=True)
                pdf.set_text_color(58, 74, 92)
                pdf.set_font(F, "", 8)
                pdf.set_x(14)
                pdf.multi_cell(0, 4.5, f"{PT}：{safe(f.get('issue',''))}")
                pdf.set_text_color(201, 160, 108)
                pdf.set_x(14)
                pdf.multi_cell(0, 4.5, f"{ST}：{safe(f.get('suggestion',''))}")
                for tag, txt in [(CRT, f.get("revised_cn","")), (ERT, f.get("revised_en",""))]:
                    if txt:
                        pdf.set_text_color(39, 174, 96)
                        pdf.set_x(14)
                        pdf.multi_cell(0, 4.5, f"{tag}：{safe(txt)}")
                pdf.ln(5)
        else:
            pdf.set_text_color(39, 174, 96)
            pdf.set_font(F, "", 8)
            pdf.set_x(14)
            pdf.cell(0, 5, GT, ln=True)
            pdf.ln(4)

    # ---- 页脚 ----
    pdf.set_y(-22)
    pdf.set_text_color(176, 186, 197)
    pdf.set_font(F, "", 7)
    pdf.cell(0, 8, "CrossLingo v5.0 | Chen @ 2026 | Powered by DeepSeek AI", align="C")

    return bytes(pdf.output())

def scan_contract(api_key, contract_text):
    """调用 DeepSeek AI 扫描合同"""
    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

    system_prompt = """你是一位资深涉外法务专家，拥有15年中英双语合同审核经验。
你的任务是对用户提供的中英双语合同进行四维风险扫描，并严格按JSON格式返回结果。

=== 分析标准 ===

1. 中英文不对应：逐条比对中英文条款语义，重点关注数字/时间/金额/范围是否对应、责任主体是否一致、限定条件是否有增减。

2. 模糊条款：识别 vague terms（reasonable, appropriate, timely, to the extent possible, substantially 等），指出缺乏客观量化标准的条款。

3. 缺失条款：中英文各自缺失的关键条款；完全缺失的重要条款类型（保密期限、知识产权归属、数据保护、不可抗力范围、管辖法律、赔偿上限等）。

4. 跨文化合规风险：争议解决条款中西差异（中国法院 vs 国际仲裁）、违约金制度差异（大陆法 vs 普通法对 penalty 的认定）、数据跨境传输合规要求（中国《个人信息保护法》 vs GDPR等）、中西方商业惯例差异。

=== 输出格式 ===
严格返回JSON，不要有任何额外文字：
{
  "overall_risk_level": "高风险/中等风险/低风险",
  "overall_summary": "一句话总体评价（中文，50字以内）",
  "categories": {
    "translation_mismatch": {
      "label": "中英文不对应", "risk_level": "高/中/低",
      "findings": [{"clause": "条款名称", "severity": "高/中/低", "issue": "问题描述（讲清楚为什么有问题、有什么后果）", "cn_text": "中文原文", "en_text": "英文原文", "suggestion": "修改建议说明", "revised_cn": "修改后的中文条款措辞（给出可直接使用的完整条款文本）", "revised_en": "修改后的英文条款措辞（给出可直接使用的完整条款文本）"}]
    },
    "ambiguous_clauses": {
      "label": "模糊条款", "risk_level": "高/中/低", "findings": [{ "同上格式，包含 revised_cn 和 revised_en" }]
    },
    "missing_clauses": {
      "label": "缺失条款", "risk_level": "高/中/低", "findings": [{ "同上格式。若中/英文条款缺失，对应的 revised 字段给出建议补充的完整条款文本" }]
    },
    "compliance_risk": {
      "label": "跨文化合规风险", "risk_level": "高/中/低", "findings": [{ "同上格式，包含 revised_cn 和 revised_en" }]
    }
  }
}
注意：如果某维度无问题，findings 返回空数组 []，risk_level 设为"低"。"""

    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"请分析以下中英双语合同：\n\n{contract_text}"}
            ],
            temperature=0.3, max_tokens=4096
        )
        result_text = response.choices[0].message.content
        if result_text.startswith("```"):
            result_text = "\n".join(result_text.split("\n")[1:-1])

        # ---- JSON 容错解析 ----
        text = result_text
        # 提取 {}
        s, e = text.find("{"), text.rfind("}")
        if s >= 0 and e > s:
            text = text[s:e+1]

        # 逐级修复
        for attempt, fixed in enumerate([
            text,                                                       # 1) 原始
            re.sub(r",(\s*[}\]])", r"\1", text),                         # 2) 去尾逗号
            re.sub(r'"\s*\n\s*"', r'",\n"', re.sub(r",(\s*[}\]])", r"\1", text)),  # 3) 补缺失逗号
        ]):
            try:
                return json.loads(fixed), None
            except json.JSONDecodeError:
                if attempt == 2:  # 最后一次尝试失败，让 AI 修复
                    try:
                        fix_resp = client.chat.completions.create(
                            model="deepseek-chat",
                            messages=[
                                {"role": "system", "content": "Fix the following JSON to be strictly valid. Return ONLY the fixed JSON, no markdown."},
                                {"role": "user", "content": text[:6000]}
                            ],
                            temperature=0, max_tokens=4096
                        )
                        fixed3 = fix_resp.choices[0].message.content
                        if fixed3.startswith("```"):
                            fixed3 = "\n".join(fixed3.split("\n")[1:-1])
                        return json.loads(fixed3), None
                    except Exception:
                        pass
        return None, "AI 返回格式异常，请重试。"

    except Exception as e:
        return None, f"扫描出错：{str(e)}"

# ========== 第11步：执行扫描 ==========
if "scan_result" not in st.session_state:
    st.session_state.scan_result = None

if scan_button and api_key and contract_text:
    # ★ 自定义扫描加载动画 — 脉冲环 + 三步骤提示
    loading_spot = st.empty()
    with loading_spot.container():
        st.markdown(f"""
        <div class="scan-overlay">
            <div class="pulse-container">
                <div class="pulse-ring"></div>
                <div class="pulse-ring"></div>
                <div class="pulse-ring"></div>
                <div class="pulse-center">
                    <span class="pulse-icon">⚖️</span>
                </div>
            </div>
            <div class="scan-title">{t('scanning_title')}</div>
            <div class="scan-subtitle">{t('scanning_sub')}</div>
            <div class="scan-steps">
                <div class="scan-step">
                    <div class="step-num">1</div>
                    <div class="step-label">{t('scan_step1')}</div>
                </div>
                <div class="scan-step">
                    <div class="step-num">2</div>
                    <div class="step-label">{t('scan_step2')}</div>
                </div>
                <div class="scan-step">
                    <div class="step-num">3</div>
                    <div class="step-label">{t('scan_step3')}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

    result, error = scan_contract(api_key, contract_text)
    loading_spot.empty()  # 清除加载动画

    if error:
        st.error(f"❌ {error}")
        st.session_state.scan_result = None
    elif result:
        st.session_state.scan_result = result

# ========== 第12步：渲染结果 ==========
if st.session_state.scan_result:
    result = st.session_state.scan_result
    categories = result.get("categories", {})

    # 严重等级双语映射（AI 返回中文值，UI 按语言显示）
    SD = {"高": t("severity_high"), "中": t("severity_medium"), "低": t("severity_low")}

    total_issues = sum(len(cat.get("findings", [])) for cat in categories.values())
    high_count = count_by_severity(categories, "高")
    medium_count = count_by_severity(categories, "中")
    low_count = count_by_severity(categories, "低")
    risk_score = calculate_risk_score(categories)

    if risk_score >= 66:
        score_color, score_label = "#C0392B", "高风险"
    elif risk_score >= 33:
        score_color, score_label = "#E67E22", "中等风险"
    else:
        score_color, score_label = "#27AE60", "低风险"

    # ---- 分割线 ----
    st.markdown('<div class="section-divider"><div class="divider-line"></div><div class="divider-diamond">◆</div><div class="divider-line"></div></div>', unsafe_allow_html=True)

    st.markdown(f"## {t('risk_dashboard')}")

    # ---- 环形仪表 + 综合评价 ----
    dash_col1, dash_col2 = st.columns([1, 3])
    with dash_col1:
        st.markdown(f'<div style="text-align:center;">{build_ring_gauge_svg(risk_score)}</div>', unsafe_allow_html=True)
    with dash_col2:
        risk_label = result.get("overall_risk_level", "未知")
        risk_emoji = {"高风险": "🔴", "中等风险": "🟡", "低风险": "🟢"}.get(risk_label, "⚪")
        st.markdown(f"""
        <div style="padding:8px 0;"><span style="font-size:18px;">{risk_emoji}</span> <b style="font-size:18px;color:#0F1F3D;">{t('ai_verdict')}：{risk_label}</b></div>
        <div style="background:#F8F9FB;border-radius:12px;padding:20px 24px;margin:12px 0;border-left:4px solid {score_color};">
            <p style="font-size:15px;color:#3A4A5C;margin:0;line-height:1.7;">{result.get('overall_summary', '')}</p>
        </div>
        """, unsafe_allow_html=True)

    # ---- 统计卡片 ----
    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    for col, cls, num, label in [
        (c1, "stat-total", total_issues, t("stat_total")),
        (c2, "stat-high", high_count, t("stat_high")),
        (c3, "stat-medium", medium_count, t("stat_medium")),
        (c4, "stat-low", low_count, t("stat_low"))
    ]:
        with col:
            st.markdown(f'<div class="stat-card {cls}"><p class="stat-number">{num}</p><p class="stat-label">{label}</p></div>', unsafe_allow_html=True)

    # ---- 四维简报 ----
    st.markdown(f"### {t('dim_brief')}")
    category_keys = ["translation_mismatch", "ambiguous_clauses", "missing_clauses", "compliance_risk"]
    dim_icons = ["🔤", "🌫️", "🕳️", "🌐"]
    dim_names = [t("dim_mismatch"), t("dim_ambiguous"), t("dim_missing"), t("dim_compliance")]
    dim_colors = ["#2B5EA7", "#C9A06C", "#6B7B8D", "#B0BAC5"]

    dim_html = '<div class="dimension-strip">'
    for key, icon, name, color in zip(category_keys, dim_icons, dim_names, dim_colors):
        cat = categories.get(key, {})
        n = len(cat.get("findings", []))
        rl = cat.get("risk_level", "低")
        rl_emoji = {"高": "🔴", "中": "🟡", "低": "🟢"}.get(rl, "⚪")
        bar_pct = min(100, int((n / max(total_issues, 1)) * 100))
        dim_html += f"""
        <div class="dimension-mini">
            <div class="dim-header"><span class="dim-icon">{icon}</span><span style="font-size:11px;">{rl_emoji} {rl}</span></div>
            <span class="dim-count">{n}</span><span class="dim-name">&nbsp;{t('dim_unit')}</span>
            <div class="dim-bar-bg"><div class="dim-bar-fill" style="width:{bar_pct}%;background:{color};"></div></div>
        </div>"""
    dim_html += '</div>'
    st.markdown(dim_html, unsafe_allow_html=True)

    # ---- 分割线 ----
    st.markdown('<div class="section-divider"><div class="divider-line"></div><div class="divider-diamond">◆</div><div class="divider-line"></div></div>', unsafe_allow_html=True)

    # ---- 逐条详情 ----
    st.markdown(f"## {t('detail_title')}")

    # ★ 筛选器
    filter_opts = [t("filter_all"), t("filter_high"), t("filter_medium"), t("filter_low")]
    severity_filter = st.radio(
        t("filter_label"),
        filter_opts,
        horizontal=True,
        label_visibility="collapsed"
    )
    filter_map = {t("filter_all"): None, t("filter_high"): "高", t("filter_medium"): "中", t("filter_low"): "低"}
    active_filter = filter_map[severity_filter]

    tabs = st.tabs([t("tab_mismatch"), t("tab_ambiguous"), t("tab_missing"), t("tab_compliance")])

    for tab, key, icon, name in zip(tabs, category_keys, dim_icons, dim_names):
        with tab:
            cat_data = categories.get(key, {})
            all_findings = cat_data.get("findings", [])
            # 应用筛选
            findings = [f for f in all_findings if active_filter is None or f.get("severity") == active_filter]
            risk_level = cat_data.get("risk_level", "未知")
            dim_emoji = {"高": "🔴", "中": "🟡", "低": "🟢"}.get(risk_level, "⚪")

            count_text = f"{len(findings)}/{len(all_findings)}" if active_filter else str(len(findings))
            st.markdown(f"### {dim_emoji} {icon} {name} <small style='color:#6B7B8D;font-size:14px;'>— {t('label_risk_level')}：{risk_level} | {t('label_issue_count')}：{count_text}</small>", unsafe_allow_html=True)
            st.markdown("---")

            if findings:
                for i, finding in enumerate(findings):
                    sev = finding.get("severity", "")
                    sev_class = {"高": "severity-high", "中": "severity-medium", "低": "severity-low"}.get(sev, "")
                    sev_emoji = {"高": "🔴", "中": "🟡", "低": "🟢"}.get(sev, "⚪")
                    cn = finding.get("cn_text", "")
                    en = finding.get("en_text", "")
                    clause_name = finding.get("clause", "未命名条款")
                    issue = finding.get("issue", "")
                    suggestion = finding.get("suggestion", "")

                    sev_label = SD.get(sev, sev)
                    with st.expander(f"{sev_emoji} [{sev_label}] {clause_name}", expanded=(i == 0)):
                        st.markdown(f'<span class="severity-badge {sev_class}">{sev_label}</span>', unsafe_allow_html=True)
                        st.markdown("")
                        st.markdown(f"**{t('problem_analysis')}：** {issue}")

                        if cn or en:
                            st.markdown(f"**{t('clause_compare')}：**")
                            cmp1, cmp2 = st.columns(2)
                            with cmp1:
                                st.markdown(f'<div class="compare-box"><div class="lang-tag">{t("cn_original")}</div><div>{cn if cn else "（—）"}</div></div>', unsafe_allow_html=True)
                            with cmp2:
                                st.markdown(f'<div class="compare-box"><div class="lang-tag">{t("en_original")}</div><div>{en if en else "（—）"}</div></div>', unsafe_allow_html=True)

                        # 修改建议
                        st.markdown(f"**{t('suggestion_label')}：** {suggestion}")

                        # 修订前后对比
                        revised_cn = finding.get("revised_cn", "")
                        revised_en = finding.get("revised_en", "")

                        if revised_cn or revised_en:
                            st.markdown(f"**{t('revised_compare')}：**")

                            if revised_cn:
                                rev_col1, rev_col2 = st.columns(2)
                                with rev_col1:
                                    st.markdown(f"""
                                    <div class="compare-box" style="border-left-color:#C0392B; background:#FFF5F5;">
                                        <div class="lang-tag" style="color:#C0392B;">{t('cn_delete')}</div>
                                        <div style="text-decoration:line-through; text-decoration-color:#C0392B40;">{cn if cn else t('missing_clause_placeholder')}</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                                with rev_col2:
                                    st.markdown(f"""
                                    <div class="compare-box" style="border-left-color:#27AE60; background:#F0FAF4;">
                                        <div class="lang-tag" style="color:#27AE60;">{t('cn_revised')}</div>
                                        <div>{revised_cn}</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                                    st.code(revised_cn, language=None)

                            if revised_en:
                                rev_col3, rev_col4 = st.columns(2)
                                with rev_col3:
                                    st.markdown(f"""
                                    <div class="compare-box" style="border-left-color:#C0392B; background:#FFF5F5;">
                                        <div class="lang-tag" style="color:#C0392B;">{t('en_delete')}</div>
                                        <div style="text-decoration:line-through; text-decoration-color:#C0392B40;">{en if en else t('missing_clause_placeholder')}</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                                with rev_col4:
                                    st.markdown(f"""
                                    <div class="compare-box" style="border-left-color:#27AE60; background:#F0FAF4;">
                                        <div class="lang-tag" style="color:#27AE60;">{t('en_revised')}</div>
                                        <div>{revised_en}</div>
                                    </div>
                                    """, unsafe_allow_html=True)
                                    st.code(revised_en, language=None)
                        else:
                            st.code(suggestion, language=None)

                        st.markdown("---")
            else:
                st.success(f"✅ {t('dim_good')}")

    # ---- 报告下载 ----
    st.markdown("---")
    st.markdown(f"### {t('download_section')}")

    report_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    report_text = f"""╔══════════════════════════════════════════════════╗
║     ⚖️ CrossLingo 合同风险扫描报告              ║
╚══════════════════════════════════════════════════╝

生成时间：{report_time}
综合风险指数：{risk_score}/100（{score_label}）
总体评价：{result.get('overall_summary', '无')}
发现问题总数：{total_issues}（🔴高危 {high_count} | 🟡中危 {medium_count} | 🟢低危 {low_count}）

{'='*50}
"""
    for key, name in zip(category_keys, dim_names):
        cat_data = categories.get(key, {})
        report_text += f"\n{'─'*50}\n【{name}】— 风险等级：{cat_data.get('risk_level', '未知')}\n{'─'*50}\n"
        for i, f in enumerate(cat_data.get("findings", []), 1):
            report_text += f"\n  {i}. [{f.get('severity', '')}风险] {f.get('clause', '')}\n     问题：{f.get('issue', '')}\n     建议：{f.get('suggestion', '')}\n"
            rc = f.get("revised_cn", "")
            re_en = f.get("revised_en", "")
            if rc:
                report_text += f"     中文修订：{rc}\n"
            if re_en:
                report_text += f"     英文修订：{re_en}\n"
        if not cat_data.get("findings"):
            report_text += "\n  ✅ 该维度未发现明显风险\n"
    report_text += f"\n{'='*50}\n⚖️ CrossLingo v5.0 | Powered by DeepSeek AI | Chen @ 2026\n"

    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        st.download_button(
            label=t("download_txt"),
            data=report_text,
            file_name=f"CrossLingo_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
            mime="text/plain",
            use_container_width=True
        )
    with dl_col2:
        pdf_bytes = generate_pdf_report(
            result, risk_score, score_label,
            total_issues, high_count, medium_count, low_count
        )
        st.download_button(
            label=t("download_pdf"),
            data=pdf_bytes,
            file_name=f"CrossLingo_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
            use_container_width=True
        )

else:
    # ---- 未扫描时：功能展示 ----
    st.markdown('<div class="section-divider"><div class="divider-line"></div><div class="divider-diamond">◆</div><div class="divider-line"></div></div>', unsafe_allow_html=True)

    # Hero 条
    st.markdown(f"""
    <div class="hero-strip">
        <div class="hero-icon">⚡</div>
        <div class="hero-text">
            <h3>{t('hero_title')}</h3>
            <p>{t('hero_desc')}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"### {t('core_capabilities')}")

    st.markdown(f"""
    <div class="feature-grid">
        <div class="feature-card">
            <div class="feature-icon-wrap">🔤</div>
            <p class="feature-title">{t('feat1_title')}</p>
            <p class="feature-desc">{t('feat1_desc')}</p>
        </div>
        <div class="feature-card">
            <div class="feature-icon-wrap">🌫️</div>
            <p class="feature-title">{t('feat2_title')}</p>
            <p class="feature-desc">{t('feat2_desc')}</p>
        </div>
        <div class="feature-card">
            <div class="feature-icon-wrap">🕳️</div>
            <p class="feature-title">{t('feat3_title')}</p>
            <p class="feature-desc">{t('feat3_desc')}</p>
        </div>
        <div class="feature-card">
            <div class="feature-icon-wrap">🌐</div>
            <p class="feature-title">{t('feat4_title')}</p>
            <p class="feature-desc">{t('feat4_desc')}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 第五张卡片（文件上传）
    st.markdown(f"""
    <div style="text-align:center; margin: 8px 0 16px 0;">
        <div class="feature-card" style="display:inline-block; max-width:320px;">
            <div class="feature-icon-wrap" style="background:#FDF3E8;">📎</div>
            <p class="feature-title">{t('feat5_title')}</p>
            <p class="feature-desc">{t('feat5_desc')}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.info(t("empty_info"))

# ========== 页脚 ==========
st.markdown("---")
st.markdown(f"""
<div class="app-footer">
    ⚖️ <b>CrossLingo</b> v5.0 Flagship &nbsp;|&nbsp;
    {t('footer_text')} &nbsp;|&nbsp;
    Chen @ 2026 &nbsp;|&nbsp;
    Powered by DeepSeek AI
</div>
""", unsafe_allow_html=True)
