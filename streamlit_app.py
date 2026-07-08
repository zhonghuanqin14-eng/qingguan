import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import zipfile
import tempfile
from io import BytesIO

# ===================== 全局极简样式 =====================
st.set_page_config(page_title="单证工具", page_icon="📦", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""
<style>
.main-title {font-size: 28px; font-weight: 600; margin: 10px 0 24px 0;}
.section-divider {margin: 40px 0; border-top:1px solid #eee;}
.card {background:#f9fafb; padding:22px; border-radius:10px; margin-bottom:16px;}
.btn-main {background:#2563eb; color:white; font-size:16px; padding:8px 26px; border:none; border-radius:8px;}
.btn-main:hover {background:#1d4ed8;}
.info-block {background:#eef6ff; padding:12px; border-radius:8px; margin-top:10px; border-left:4px #2563eb solid;}
</style>
""", unsafe_allow_html=True)

# ===================== 账号字典 =====================
ACCOUNT_INFO = {
    "39": {
        "shipper_name": "Shenzhen Longyuan Junjie Technology Co., Ltd.",
        "shipper_addr": "Room 502, No. 5, Ruiyuan Second Lane, Nanlian Community, Longgang Subdistrict, Longgang District, Shenzhen,Guangdong,China",
        "contact": "ZHOUJUNJIE",
        "phone": "+8613427679670"
    },
    "79": {
        "shipper_name": "Mingtongsheng (Shenzhen) E-commerce Co., Ltd.",
        "shipper_addr": "3A Zijing Pavilion, Building 4, Baizhu Garden, 249 Zhuguang Road, Longlian Community, Taoyuan Street, Nanshan District, Shenzhen,Guangdong,China",
        "contact": "LIMIN",
        "phone": "+8613902478270"
    },
    "76": {
        "shipper_name": "Shenzhen Chengziwei Technology Co., Ltd.",
        "shipper_addr": "B410, Buildings 2 and 3, Mingliang Technology Park, No. 88 Zhuguang North Road, Pingshan Community, Taoyuan Street, Nanshan District, Shenzhen,Guangdong,China",
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
        "shipper_name": "Hong Kong Linglingqinlv Technology Limited",
        "shipper_addr": "UNIT F22,RM 6, 10/F, LEMMI CENTRE, 50 HOI YUEN ROAD,Kwun Tong,Hong Kong",
        "contact": "LUQINGLING",
        "phone": "+8619864368710"
    }
}

# ===================== 清关模板坐标（完全匹配截图） =====================
CLEAR_MAP = {
    "fba_row": 7,
    "fba_col": 10,          # J列 FBA单号
    "ship_name_row": 4,
    "ship_name_col": 2,     # B列 发货人名称
    "ship_addr_row": 4,
    "ship_addr_col": 2,     # B列 发货人地址
    "ship_contact_row": 9,
    "ship_contact_col": 2,  # B列 发货人联系人
    "ship_tel_row": 10,
    "ship_tel_col": 2,      # B列 发货人电话
    "imp_name_row": 4,
    "imp_name_col": 5,      # E列 进口商名称
    "imp_addr_row": 4,
    "imp_addr_col": 5,      # E列 进口商地址
    "imp_contact_row": 9,
    "imp_contact_col": 5,   # E列 进口商联系人
    "imp_tel_row": 10,
    "imp_tel_col": 5,       # E列 进口商电话
    "manu_name_row": 38,
    "manu_name_col": 2,     # B列 制造商名称（截图B38开始）
    "manu_addr_row": 39,
    "manu_addr_col": 2,     # B列 制造商地址
    "manu_addr2_row": 40,
    "manu_addr2_col": 2,    # B列 制造商地址第二行
    "data_start_row": 22,   # 明细开始行
    "data_end_clear_row": 25, # 明细结束行（截图里22-25行有数据）
    "total_row": 36,        # 合计行
    "qty_col": 9,           # I列 数量
    "unit_price_col": 10,   # J列 单价
    "total_price_col": 11,  # K列 总价
    "ctns_col": 13,         # M列 箱数（核心修正：之前错写14，现在改13）
    "gw_col": 14,           # N列 毛重
    "nw_col": 15,           # O列 净重
    "vol_col": 16           # P列 体积
}

# ===================== LCL截单模板坐标 =====================
CUT_MAP = {
    "header_row":3,
    "data_start":4,
    "data_end":7,
    "weight_col":3,
    "vol_col":4,
    "weight_head":"Gross weight",
    "vol_head":"Volume"
}

# 模板文件名
TEMPLATE_FILE = "FBA US Combined Commercial Invoice Packing List.xlsx"

# ===================== 页面标题 =====================
st.markdown('<div class="main-title">📦 单证批量处理工具</div>', unsafe_allow_html=True)

# ===================== 模块1：FBA清关单生成 =====================
st.markdown('<div class="card">', unsafe_allow_html=True)
st.subheader("1. FBA清关单批量生成")
col1, col2 = st.columns([0.48, 0.48], gap="medium")
with col1:
    file_clear = st.file_uploader("上传数据源Excel", type=["xlsx","xls"], key="clear_file")
with col2:
    acc_list = list(ACCOUNT_INFO.keys())
    select_acc = st.selectbox("选择账号", options=acc_list, key="acc_sel")
    acc_detail = ACCOUNT_INFO[select_acc]
    st.markdown('<div class="info-block">', unsafe_allow_html=True)
    st.write(f"公司：{acc_detail['shipper_name']}")
    st.write(f"地址：{acc_detail['shipper_addr']}")
    st.write(f"联系人：{acc_detail['contact']} | 电话：{acc_detail['phone']}")
    st.markdown('</div>', unsafe_allow_html=True)

gen_clear = st.button("生成并下载清关资料", key="gen_clear", type="primary")
st.markdown('</div>', unsafe_allow_html=True)

# 清关生成逻辑
if gen_clear:
    if not file_clear:
        st.error("请上传数据源文件")
    elif not os.path.exists(TEMPLATE_FILE):
        st.error(f"模板文件{TEMPLATE_FILE}缺失，请上传至仓库根目录")
    else:
        with st.spinner("正在生成单据..."):
            df = pd.read_excel(file_clear)
            groups = df.groupby("FBA编号")
            acc_info = ACCOUNT_INFO[select_acc]
            tmp_dir = tempfile.TemporaryDirectory()
            tmp_path = tmp_dir.name
            file_list = []

            for fba_id, group in groups:
                wb = load_workbook(TEMPLATE_FILE)
                ws = wb.active

                # 发货人填充 加异常捕获
                try:
                    cell_ship_name = ws.cell(row=CLEAR_MAP["ship_name_row"], column=CLEAR_MAP["ship_name_col"])
                    cell_ship_name.value = acc_info["shipper_name"]
                except:
                    pass
                try:
                    cell_ship_addr = ws.cell(row=CLEAR_MAP["ship_addr_row"], column=CLEAR_MAP["ship_addr_col"])
                    cell_ship_addr.value = acc_info["shipper_addr"]
                except:
                    pass
                try:
                    cell_ship_contact = ws.cell(row=CLEAR_MAP["ship_contact_row"], column=CLEAR_MAP["ship_contact_col"])
                    cell_ship_contact.value = f"Contact:{acc_info['contact']}"
                except:
                    pass
                try:
                    cell_ship_tel = ws.cell(row=CLEAR_MAP["ship_tel_row"], column=CLEAR_MAP["ship_tel_col"])
                    cell_ship_tel.value = f"Phone:{acc_info['phone']}"
                except:
                    pass

                # 进口商填充 加异常捕获
                try:
                    cell_imp_name = ws.cell(row=CLEAR_MAP["imp_name_row"], column=CLEAR_MAP["imp_name_col"])
                    cell_imp_name.value = acc_info["shipper_name"]
                except:
                    pass
                try:
                    cell_imp_addr = ws.cell(row=CLEAR_MAP["imp_addr_row"], column=CLEAR_MAP["imp_addr_col"])
                    cell_imp_addr.value = acc_info["shipper_addr"]
                except:
                    pass
                try:
                    cell_imp_contact = ws.cell(row=CLEAR_MAP["imp_contact_row"], column=CLEAR_MAP["imp_contact_col"])
                    cell_imp_contact.value = f"Contact:{acc_info['contact']}"
                except:
                    pass
                try:
                    cell_imp_tel = ws.cell(row=CLEAR_MAP["imp_tel_row"], column=CLEAR_MAP["imp_tel_col"])
                    cell_imp_tel.value = f"Phone:{acc_info['phone']}"
                except:
                    pass

                # 制造商强制填充（匹配截图B38-B40区域，双层兜底）
                try:
                    # 先解除制造商区域合并单元格，避免赋值失败
                    for merged_range in ws.merged_cells.ranges:
                        if (merged_range.min_row <= 40 and merged_range.max_row >= 38) and (merged_range.min_col <= 2 and merged_range.max_col >= 2):
                            ws.unmerge_cells(range_string=str(merged_range))
                            break
                    # 写入制造商名称、地址
                    cell_m_name = ws.cell(row=CLEAR_MAP["manu_name_row"], column=CLEAR_MAP["manu_name_col"])
                    cell_m_name.value = acc_info["shipper_name"]
                    cell_m_addr = ws.cell(row=CLEAR_MAP["manu_addr_row"], column=CLEAR_MAP["manu_addr_col"])
                    cell_m_addr.value = acc_info["shipper_addr"]
                    cell_m_addr2 = ws.cell(row=CLEAR_MAP["manu_addr2_row"], column=CLEAR_MAP["manu_addr2_col"])
                    cell_m_addr2.value = ""
                except Exception:
                    # 兜底方案，直接写入字母坐标
                    try:
                        ws["B38"].value = acc_info["shipper_name"]
                        ws["B39"].value = acc_info["shipper_addr"]
                    except:
                        pass

                # FBA单号
                try:
                    cell_fba = ws.cell(row=CLEAR_MAP["fba_row"], column=CLEAR_MAP["fba_col"])
                    cell_fba.value = fba_id
                except:
                    pass

                # 清空旧明细区域
                s_r = CLEAR_MAP["data_start_row"]
                e_r = CLEAR_MAP["data_end_clear_row"]
                for r in range(s_r, e_r+1):
                    for c in range(2, 18):
                        ws.cell(row=r, column=c, value=None)

                # 写入产品明细（核心修正：箱数列号改为13=M列）
                rows = group.values.tolist()
                for idx, row in enumerate(rows):
                    r = s_r + idx
                    ws.cell(r, 2, row[0])  # 零件号 B列
                    ws.cell(r, 3, row[1])  # 品名 C列
                    ws.cell(r, 4, row[2])  # 材质 D列
                    ws.cell(r, 5, row[3])  # 关税分类 E列
                    ws.cell(r, 8, "CN")   # 原产国 H列
                    ws.cell(r, 9, row[7])  # 数量 I列
                    ws.cell(r, 10, row[8]) # 单价 J列
                    ws.cell(r, 11, f"=J{r}*I{r}") # 总价 K列
                    ws.cell(r, 13, row[11]) # 箱数 M列（修正列号）
                    ws.cell(r, 14, round(row[12],3)) # 毛重 N列
                    ws.cell(r, 15, row[13]) # 净重 O列
                    ws.cell(r, 16, round(row[14],3)) # 体积 P列

                # 合计公式（修正列号，对应M/N/O/P列）
                end_data = s_r + len(rows) - 1
                total_r = CLEAR_MAP["total_row"]
                ws.cell(total_r, 11, f"=SUM(K{s_r}:K{end_data})") # 总价合计
                ws.cell(total_r, 13, f"=SUM(M{s_r}:M{end_data})") # 箱数合计
                ws.cell(total_r, 14, f"=SUM(N{s_r}:N{end_data})") # 毛重合计
                ws.cell(total_r, 15, f"=SUM(O{s_r}:O{end_data})") # 净重合计
                ws.cell(total_r, 16, f"=SUM(P{s_r}:P{end_data})") # 体积合计

                save_path = os.path.join(tmp_path, f"{fba_id}.xlsx")
                wb.save(save_path)
                wb.close()
                file_list.append(save_path)

            # 打包压缩包
            zip_buf = BytesIO()
            zip_name = f"{select_acc}清关资料.zip"
            with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for fp in file_list:
                    zf.write(fp, os.path.basename(fp))
            zip_buf.seek(0)
            st.download_button(
                label="点击下载清关压缩包",
                data=zip_buf,
                file_name=zip_name,
                mime="application/zip",
                key="dl_clear_auto"
            )
            st.success(f"{zip_name} 已生成，请点击上方按钮下载！")
            tmp_dir.cleanup()

# 分割线
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

# ===================== 模块2：LCL截单重量体积换算 =====================
st.markdown('<div class="card">', unsafe_allow_html=True)
st.subheader("2. LCL截单重量体积调整")
upload_cut = st.file_uploader("上传LCL截单Excel", type=["xlsx","xls"], key="cut_file")
colw, colv = st.columns([0.48,0.48], gap="medium")
with colw:
    target_w = st.number_input("目标总重量 kg", min_value=0.001, step=0.001, format="%.3f", key="tw")
with colv:
    target_v = st.number_input("目标总体积 CBM", min_value=0.001, step=0.001, format="%.3f", key="tv")
adjust_btn = st.button("调整并生成截单文件", key="adj_btn", type="primary")
st.markdown('</div>', unsafe_allow_html=True)

# 截单处理逻辑
if adjust_btn:
    if not upload_cut:
        st.error("请上传截单Excel文件")
    elif target_w <= 0 or target_v <= 0:
        st.error("重量、体积必须大于0")
    else:
        with st.spinner("计算并调整体积重量..."):
            fname = upload_cut.name
            name_no_ext = os.path.splitext(fname)[0]
            al0_code = ""
            for part in name_no_ext.split("_"):
                if part.startswith("AL0"):
                    al0_code = part
                    break
            if not al0_code:
                al0_code = "未知单号"
            out_name = f"截单资料{al0_code}.xlsx"

            wb = load_workbook(upload_cut)
            ws = wb.active
            s_r = CUT_MAP["data_start"]
            e_r = CUT_MAP["data_end"]
            w_col = CUT_MAP["weight_col"]
            v_col = CUT_MAP["vol_col"]
            raw_data = []
            sum_w_ori = 0.0
            sum_v_ori = 0.0
            for r in range(s_r, e_r + 1):
                wc = ws.cell(r, w_col)
                vc = ws.cell(r, v_col)
                try:
                    wv = float(wc.value)
                    vv = float(vc.value)
                    raw_data.append([r, wv, vv])
                    sum_w_ori += wv
                    sum_v_ori += vv
                except:
                    continue
            if sum_w_ori <= 0 or sum_v_ori <= 0:
                st.error("原始单据总重量/体积为0，无法调整")
                wb.close()
                st.stop()
            ratio_w = target_w / sum_w_ori
            ratio_v = target_v / sum_v_ori
            data_list = []
            for r, ow, ov in raw_data:
                ew = ow * ratio_w
                ev = ov * ratio_v
                data_list.append([r, ew, ev])
            # 精准分摊消除小数尾差
            target_w_int = int(round(target_w * 1000))
            w_int_list = []
            sum_wi = 0
            for _, ew, _ in data_list:
                i = int(round(ew * 1000))
                w_int_list.append(i)
                sum_wi += i
            w_int_list[-1] += target_w_int - sum_wi

            target_v_int = int(round(target_v * 1000))
            v_int_list = []
            sum_vi = 0
            for _, _, ev in data_list:
                i = int(round(ev * 1000))
                v_int_list.append(i)
                sum_vi += i
            v_int_list[-1] += target_v_int - sum_vi
            # 写入单元格
            for idx, (row_num, _, _) in enumerate(data_list):
                final_w = w_int_list[idx] / 1000
                final_v = v_int_list[idx] / 1000
                ws.cell(row_num, w_col, value=final_w)
                ws.cell(row_num, v_col, value=final_v)
            buf = BytesIO()
            wb.save(buf)
            buf.seek(0)
            wb.close()
            st.download_button(
                label="下载调整后截单文件",
                data=buf,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_cut_auto"
            )
            st.success(f"{out_name} 已生成，请点击按钮下载！")

# 底部说明文字
st.markdown("<p style='color:#666; font-size:13px; margin-top:30px;'>上方模块生成FBA清关打包文件，选择账号可预览公司信息；下方调整LCL截单重量体积，数值保留三位小数</p>", unsafe_allow_html=True)
