import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.cell.cell import MergedCell
from copy import copy
import os
import zipfile
import tempfile
from io import BytesIO

# ===================== 全局页面美化配置 =====================
st.set_page_config(
    page_title="FBA清关&截单资料工具",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

custom_css = """
<style>
.main-title {
    font-size: 32px;
    color: #165DFF;
    font-weight: 700;
    margin-bottom: 10px;
}
.card {
    background-color: #f8fafc;
    padding: 20px;
    border-radius: 12px;
    border: 1px solid #e2e8f0;
    margin-bottom: 20px;
}
.stButton>button {
    background-color: #165DFF;
    color: white;
    font-size: 16px;
    padding: 8px 24px;
    border-radius: 8px;
    border: none;
}
.stButton>button:hover {
    background-color: #0E4BDB;
}
.download-btn>button {
    background-color: #00B42A;
}
.download-btn>button:hover {
    background-color: #009A24;
}
.info-text {
    color: #4E5969;
    font-size: 14px;
}
.preview-box {
    background: #E8F3FF;
    padding: 16px;
    border-radius: 8px;
    border-left: 4px solid #165DFF;
}
.section-title {
    font-size: 24px;
    color: #165DFF;
    font-weight: 600;
    margin: 30px 0 20px 0;
}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ===================== 账号配置 =====================
ACCOUNT_INFO = {
    "39": {
        "shipper_name": "Shenzhen Longyuan Junjie Technology Co., Ltd.",
        "shipper_addr": "Room 502, No. 5, Ruiyuan Second Lane, Nanlian Community, Longgang Subdistrict, Longgang District, Shenzhen,Guangdong,China",
        "contact": "ZHOUJUNJIE",
        "phone": "+8613427679670"
    },
    "79": {
        "shipper_name": "Mingtongsheng (Shenzhen) E-commerce Co., Ltd.",
        "shipper_addr": "3A Zijing Pavilion, Building 4, Baizhu Garden, 249 Zhuguang Road, Longlian Community, Tao Yuan Street, Nanshan District, Shenzhen,Guangdong,China",
        "contact": "LIMIN",
        "phone": "+8613902478270"
    },
    "76": {
        "shipper_name": "Shenzhen Chengziwei Technology Co., Ltd.",
        "shipper_addr": "B410, Buildings 2 and 3, Mingliang Technology Park, No. 88 Zhuguang North Road, Pingshan Community, Tao Yuan Street, Nanshan District, Shenzhen,Guangdong,China",
        "contact": "CHENZIWEI",
        "phone": "+8615875396146"
    },
    "47.92": {
        "shipper_name": "Shenzhen Dongshan Jinhao Technology Co., Ltd.",
        "shipper_addr": "Room 108, Building 8, Maker Town, Xili Street, Nanshan District, Shenzhen,Guangdong,China",
        "contact": "CHENJINHAO",
        "phone": "+8613530567440"
    },
    "47.100": {
        "shipper_name": "Shenzhen Weizhite Technology Co., Ltd.",
        "shipper_addr": "410, Building 2-3, Bright Technology Park, No. 88 Zhuguang North Road, Pingshan Community, Taoyuan Street, Nanshan District, Shenzhen,Guangdong,China",
        "contact": "CHENSIFA",
        "phone": "+8615986680681"
    },
    "47.99": {
        "shipper_name": "Shenzhen Shunhuixiong Technology Co., Ltd.",
        "shipper_addr": "115, Building 10, Maker Town, No. 4109 Liuxian Avenue, Pingshan Community, Taoyuan Street, Nanshan District, Shenzhen,Guangdong,China",
        "contact": "LIANGRIXIONG",
        "phone": "+8618026938073"
    },
    "70": {
        "shipper_name": "Shenzhen Xingyuepan Technology Co., Ltd.",
        "shipper_addr": "B429, No. 22 Dakan Industrial 2nd Road, Daguan Community, Xili Street, Nanshan District, Shenzhen,Guangdong,China",
        "contact": "LIUBIN",
        "phone": "+8613530369614"
    },
    "71": {
        "shipper_name": "Shenzhen Chenghai Liufa Technology Co., Ltd.",
        "shipper_addr": "28C, Unit A, Building 3, Xiangshanli Phase 5, Wenchang Street Community, Shahe Street, Nanshan District, Shenzhen,Guangdong,China",
        "contact": "GULIUFA",
        "phone": "+8617744965296"
    },
    "8.1": {
        "shipper_name": "Shenzhen Shiqi Jiechao Technology Co., Ltd.",
        "shipper_addr": "Tianxi Xiaoju V211, No. 10 Ruihua North Lane, Nanlian Community, Longgang Street, Longgang District, Shenzhen,Guangdong,China",
        "contact": "LUJIECHAO",
        "phone": "+8613670528672"
    },
    "47.108": {
        "shipper_name": "Guangzhou Changyou Weilin Technology Co., Ltd.",
        "shipper_addr": "Shop 422, No. 15 Yihe Road, Liwan District, Guangzhou,Guangdong,China",
        "contact": "LUOSILIN",
        "phone": "+8615728285292"
    },
    "47.239": {
        "shipper_name": "Hong Kong LingLingQinLv Technology Limited",
        "shipper_addr": "UNIT F22,RM 6, 10/F, LEMMI CENTRE, 50 HOI YUEN ROAD,Kwun Tong,Hong Kong",
        "contact": "LUQINGLING",
        "phone": "+8619864368710"
    }
}

CELL_MAP = {
    "fba_no": "J8",
    "ship_name": "B7",
    "ship_addr": "B8",
    "ship_contact": "B9",
    "ship_tel": "B10",
    "imp_name": "E7",
    "imp_addr": "E8",
    "imp_contact": "E9",
    "imp_tel": "E10",
    "manu_name": "C38",
    "manu_addr": "C39",
    "data_start_row": 22,
    "data_end_clear_row": 35,
    "total_row": 36,
    "gross_weight_col": 14,
    "volume_col": 16
}

TEMPLATE_FILE = "AL0-SBU6B5D6EZU6S.xlsx"

# ===================== 页面标题 =====================
st.markdown('<div class="main-title">📦 FBA清关单批量生成工具</div>', unsafe_allow_html=True)
st.divider()

# ===================== 模块1：FBA清关单生成（独立上传文件） =====================
col_left, col_right = st.columns([0.48, 0.48], gap="medium")
# 左侧：清关数据源上传
with col_left:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📁 第一步：上传清单数据源Excel")
    st.markdown("<p class='info-text'>文件内需包含【FBA编号】列</p>", unsafe_allow_html=True)
    upload_clear = st.file_uploader("", type=["xlsx", "xls"], key="clear_file")
    if upload_clear is not None:
        st.success("✅ 已读取文件")
    st.markdown('</div>', unsafe_allow_html=True)

# 右侧：账号选择
with col_right:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("🏢 第二步：选择发货账号")
    acc_list = list(ACCOUNT_INFO.keys())
    select_acc = st.selectbox("", options=acc_list, format_func=lambda x: f"账号{x}", key="acc_sel")
    current_acc = ACCOUNT_INFO[select_acc]
    st.markdown("<div class='preview-box'>", unsafe_allow_html=True)
    st.write(f"**公司名称：** {current_acc['shipper_name']}")
    st.write(f"**地址：** {current_acc['shipper_addr']}")
    st.write(f"**联系人：** {current_acc['contact']} | 电话：{current_acc['phone']}")
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

# 清关生成按钮
st.markdown('<div class="card">', unsafe_allow_html=True)
col_btn, _ = st.columns([0.2, 0.8])
with col_btn:
    gen_clear_btn = st.button("🚀 开始批量生成清关文件", key="gen_clear")
st.markdown('</div>', unsafe_allow_html=True)

# 清关生成逻辑
if gen_clear_btn:
    if upload_clear is None:
        st.error("❌ 请上传清单数据源Excel！")
    elif not os.path.exists(TEMPLATE_FILE):
        st.error(f"❌ 模板{TEMPLATE_FILE}缺失！")
    else:
        with st.spinner("⏳ 正在生成清关单..."):
            df = pd.read_excel(upload_clear)
            groups = df.groupby("FBA编号")
            acc_data = ACCOUNT_INFO[select_acc]
            tmp_dir = tempfile.TemporaryDirectory()
            tmp_path = tmp_dir.name
            out_file_list = []

            for fba_id, group_df in groups:
                wb = load_workbook(TEMPLATE_FILE)
                ws = wb.active
                ws[CELL_MAP["ship_name"]].value = acc_data["shipper_name"]
                ws[CELL_MAP["ship_addr"]].value = acc_data["shipper_addr"]
                ws[CELL_MAP["ship_contact"]].value = f"Contact:{acc_data['contact']}"
                ws[CELL_MAP["ship_tel"]].value = f"Phone:{acc_data['phone']}"
                ws[CELL_MAP["imp_name"]].value = acc_data["shipper_name"]
                ws[CELL_MAP["imp_addr"]].value = acc_data["shipper_addr"]
                ws[CELL_MAP["imp_contact"]].value = f"Contact:{acc_data['contact']}"
                ws[CELL_MAP["imp_tel"]].value = f"Phone:{acc_data['phone']}"
                ws[CELL_MAP["manu_name"]].value = acc_data["shipper_name"]
                ws[CELL_MAP["manu_addr"]].value = acc_data["shipper_addr"]
                ws[CELL_MAP["fba_no"]].value = fba_id

                s_r = CELL_MAP["data_start_row"]
                e_r = CELL_MAP["data_end_clear_row"]
                for r in range(s_r, e_r + 1):
                    for c in range(2, 17):
                        ws.cell(row=r, column=c, value=None)

                data_rows = group_df.values.tolist()
                for idx, row_data in enumerate(data_rows):
                    curr_r = s_r + idx
                    ws.cell(row=curr_r, column=2, value=row_data[0])
                    ws.cell(row=curr_r, column=3, value=row_data[1])
                    ws.cell(row=curr_r, column=4, value=row_data[2])
                    ws.cell(row=curr_r, column=5, value=row_data[3])
                    ws.cell(row=curr_r, column=8, value="CN")
                    ws.cell(row=curr_r, column=9, value=row_data[7])
                    ws.cell(row=curr_r, column=10, value=row_data[8])
                    ws.cell(row=curr_r, column=11, value=f"=J{curr_r}*I{curr_r}")
                    ws.cell(row=curr_r, column=13, value=row_data[11])
                    ws.cell(row=curr_r, column=14, value=round(row_data[12],3))
                    ws.cell(row=curr_r, column=15, value=row_data[13])
                    ws.cell(row=curr_r, column=16, value=round(row_data[14],3))

                data_end_r = s_r + len(data_rows) - 1
                total_r = CELL_MAP["total_row"]
                ws.cell(row=total_r, column=11, value=f"=SUM(K{s_r}:K{data_end_r})")
                ws.cell(row=total_r, column=13, value=f"=SUM(M{s_r}:M{data_end_r})")
                ws.cell(row=total_r, column=14, value=f"=SUM(N{s_r}:N{data_end_r})")
                ws.cell(row=total_r, column=15, value=f"=SUM(O{s_r}:O{data_end_r})")
                ws.cell(row=total_r, column=16, value=f"=SUM(P{s_r}:P{data_end_r})")

                save_file = os.path.join(tmp_path, f"{fba_id}.xlsx")
                wb.save(save_file)
                wb.close()
                out_file_list.append(save_file)

            zip_buffer = BytesIO()
            zip_name = f"{select_acc}清关资料.zip"
            with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for f_path in out_file_list:
                    zf.write(f_path, os.path.basename(f_path))
            zip_buffer.seek(0)

            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.success(f"🎉 清关单生成完成！共{len(out_file_list)}份")
            st.markdown('<div class="download-btn">', unsafe_allow_html=True)
            st.download_button("📥 下载清关资料ZIP", zip_buffer, zip_name, "application/zip", key="dl_clear")
            st.markdown("</div>", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
            tmp_dir.cleanup()

# ===================== 模块2：截单重量体积调整（独立上传，完全隔离，修复min参数） =====================
st.markdown('<div class="section-title">📦 截单资料重量体积比例调整</div>', unsafe_allow_html=True)
st.markdown("<p class='info-text'>单独上传截单Excel，输入目标总重/总体积，自动按比例缩放，数值保留3位小数</p>", unsafe_allow_html=True)

st.markdown('<div class="card">', unsafe_allow_html=True)
# 独立文件上传
st.subheader("1. 上传截单Excel文件")
upload_cut = st.file_uploader("", type=["xlsx", "xls"], key="cut_file")
if upload_cut is not None:
    st.success("✅ 已读取截单文件")

# 双输入框（修复min→min_value）
col_w, col_v = st.columns([0.48, 0.48], gap="medium")
with col_w:
    target_w = st.number_input("目标总重量 kg", min_value=0.001, step=0.001, format="%.3f", key="t_w")
with col_v:
    target_v = st.number_input("目标总体积 cbm", min_value=0.001, step=0.001, format="%.3f", key="t_v")

adjust_btn = st.button("🔄 按比例重新计算重量体积", key="adj_btn")
st.markdown('</div>', unsafe_allow_html=True)

# 截单调整逻辑
if adjust_btn:
    if upload_cut is None:
        st.error("❌ 请先上传截单Excel！")
    elif target_w <=0 or target_v <=0:
        st.error("❌ 目标重量/体积必须大于0！")
    else:
        with st.spinner("⏳ 正在按比例重算数据..."):
            wb = load_workbook(upload_cut)
            ws = wb.active
            s_r = CELL_MAP["data_start_row"]
            e_r = CELL_MAP["data_end_clear_row"]
            w_col = CELL_MAP["gross_weight_col"]
            v_col = CELL_MAP["volume_col"]

            sum_w = 0.0
            sum_v = 0.0
            row_data_list = []

            for r in range(s_r, e_r + 1):
                w_cell = ws.cell(row=r, column=w_col)
                v_cell = ws.cell(row=r, column=v_col)
                try:
                    w_val = float(w_cell.value)
                    v_val = float(v_cell.value)
                    row_data_list.append([r, w_val, v_val])
                    sum_w += w_val
                    sum_v += v_val
                except:
                    continue

            if sum_w <= 0 or sum_v <=0:
                st.error("❌ 文件内原始总重/体积为0，无法缩放！")
                wb.close()
                st.stop()

            ratio_w = target_w / sum_w
            ratio_v = target_v / sum_v

            # 逐行重写，保留3位小数
            for r, ow, ov in row_data_list:
                new_w = round(ow * ratio_w, 3)
                new_v = round(ov * ratio_v, 3)
                ws.cell(row=r, column=w_col, value=new_w)
                ws.cell(row=r, column=v_col, value=new_v)

            # 更新合计公式
            total_r = CELL_MAP["total_row"]
            ws.cell(row=total_r, column=w_col, value=f"=SUM(N{s_r}:N{e_r})")
            ws.cell(row=total_r, column=v_col, value=f"=SUM(P{s_r}:P{e_r})")

            buf = BytesIO()
            wb.save(buf)
            buf.seek(0)
            wb.close()

            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.success("✅ 调整完成，所有重量体积已保留3位小数")
            st.write(f"原总重：{round(sum_w,3)} kg → 目标：{round(target_w,3)} kg")
            st.write(f"原总体积：{round(sum_v,3)} cbm → 目标：{round(target_v,3)} cbm")
            st.markdown('<div class="download-btn">', unsafe_allow_html=True)
            st.download_button("📥 下载调整后截单Excel", buf, "截单资料_调整后.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_cut")
            st.markdown("</div>", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

# 底部说明
st.divider()
st.markdown("<p class='info-text'>💡 上下模块独立上传文件；重量体积统一保留3位小数；清关单制造商信息完整不丢失</p>", unsafe_allow_html=True)
