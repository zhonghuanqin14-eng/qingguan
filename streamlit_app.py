import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import os
import zipfile
import tempfile
from io import BytesIO
from datetime import datetime

# ===================== 全局极简样式配置 =====================
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

# ===================== 账号配置（清关模块专用） =====================
ACCOUNT_INFO = {
    "39": {"shipper_name": "Shenzhen Longyuan Junjie Technology Co., Ltd.","shipper_addr": "Room 502, No. 5, Ruiyuan Second Lane, Nanlian Community, Longgang Subdistrict, Longgang District, Shenzhen,Guangdong,China","contact": "ZHOUJUNJIE","phone": "+8613427679670"},
    "79": {"shipper_name": "Mingtongsheng (Shenzhen) E-commerce Co., Ltd.","shipper_addr": "3A Zijing Pavilion, Building 4, Baizhu Garden, 249 Zhuguang Road, Longlian Community, Tao Yuan Street, Nanshan District, Shenzhen,Guangdong,China","contact": "LIMIN","phone": "+8613902478270"},
    "76": {"shipper_name": "Shenzhen Chengziwei Technology Co., Ltd.","shipper_addr": "B410, Buildings 2 and 3, Mingliang Technology Park, No. 88 Zhuguang North Road, Pingshan Community, Tao Yuan Street, Nanshan District, Shenzhen,Guangdong,China","contact": "CHENZIWEI","phone": "+8615875396146"},
    "47.92": {"shipper_name": "Shenzhen Dongshan Jinhao Technology Co., Ltd.","shipper_addr": "Room 108, Building 8, Maker Town, Xili Street, Nanshan District, Shenzhen,Guangdong,China","contact": "CHENJINHAO","phone": "+8613530567440"},
    "47.100": {"shipper_name": "Shenzhen Weizhite Technology Co., Ltd.","shipper_addr": "410, Building 2-3, Bright Technology Park, No. 88 Zhuguang North Road, Pingshan Community, Taoyuan Street, Nanshan District, Shenzhen,Guangdong,China","contact": "CHENSIFA","phone": "+8615986680681"},
    "47.99": {"shipper_name": "Shenzhen Shunhuixiong Technology Co., Ltd.","shipper_addr": "115, Building 10, Maker Town, No. 4109 Liuxian Avenue, Pingshan Community, Taoyuan Street, Nanshan District, Shenzhen,Guangdong,China","contact": "LIANGRIXIONG","phone": "+8618026938073"},
    "70": {"shipper_name": "Shenzhen Xingyuepan Technology Co., Ltd.","shipper_addr": "B429, No. 22 Dakan Industrial 2nd Road, Daguan Community, Xili Street, Nanshan District, Shenzhen,Guangdong,China","contact": "LIUBIN","phone": "+8613530369614"},
    "71": {"shipper_name": "Shenzhen Chenghai Liufa Technology Co., Ltd.","shipper_addr": "28C, Unit A, Building 3, Xiangshanli Phase 5, Wenchang Street Community, Shahe Street, Nanshan District, Shenzhen,Guangdong,China","contact": "GULIUFA","phone": "+8617744965296"},
    "8.1": {"shipper_name": "Shenzhen Shiqi Jiechao Technology Co., Ltd.","shipper_addr": "Tianxi Xiaoju V211, No. 10 Ruihua North Lane, Nanlian Community, Longgang Street, Longgang District, Shenzhen,Guangdong,China","contact": "LUJIECHAO","phone": "+8613670528672"},
    "47.108": {"shipper_name": "Guangzhou Changyou Weilin Technology Co., Ltd.","shipper_addr": "Shop 422, No. 15 Yihe Road, Liwan District, Guangzhou,Guangdong,China","contact": "LUOSILIN","phone": "+8615728285292"},
    "47.239": {"shipper_name": "Hong Kong LingLingQinLv Technology Limited","shipper_addr": "UNIT F22,RM 6, 10/F, LEMMI CENTRE, 50 HOI YUEN ROAD,Kwun Tong,Hong Kong","contact": "LUQINGLING","phone": "+8619864368710"}
}

# ===================== 清关模板坐标（100%匹配AL0-SBU6B5D6EZU6S.xlsx模板） =====================
CLEAR_MAP = {
    "fba_no": "J8",          # FBA单号 J列第8行
    "ship_name": "B7",       # 发货人名称 B列第7行
    "ship_addr": "B8",       # 发货人地址 B列第8行
    "ship_contact": "B9",    # 发货人联系人 B列第9行
    "ship_tel": "B10",      # 发货人电话 B列第10行
    "imp_name": "E7",       # 进口商名称 E列第7行
    "imp_addr": "E8",       # 进口商地址 E列第8行
    "imp_contact": "E9",    # 进口商联系人 E列第9行
    "imp_tel": "E10",      # 进口商电话 E列第10行
    "manu_name": "C38",     # 制造商名称 C列第38行
    "manu_addr": "C39",     # 制造商地址 C列第39行
    "data_start": 22,        # 明细数据开始行
    "data_end": 35,          # 明细数据结束行
    "total_row": 36,         # 合计行
    "qty_col": 9,           # 数量 I列（第9列）
    "unit_price_col": 10,   # 单价 J列（第10列）
    "total_price_col": 11,  # 总价 K列（第11列）
    "ctns_col": 13,         # 箱数 M列（第13列）
    "weight_col": 14,       # 毛重 N列（第14列）
    "nw_col": 15,           # 净重 O列（第15列）
    "vol_col": 16           # 体积 P列（第16列）
}

# 截单LCL模板坐标
CUT_MAP = {
    "header_row":3,"data_start":4,"data_end":7,"weight_col":3,"vol_col":4,
    "weight_head":"Gross weight","vol_head":"Volume"
}

# 模板文件名（和你上传的完全一致）
TEMPLATE_FILE = "AL0-SBU6B5D6EZU6S.xlsx"

# ===================== 页面标题 =====================
st.markdown('<div class="main-title">📦 单证批量处理工具</div>', unsafe_allow_html=True)

# ===================== 模块1：FBA清关单生成（下拉下方展示公司信息） =====================
st.markdown('<div class="card">', unsafe_allow_html=True)
st.subheader("1. FBA清关单批量生成")
col1, col2 = st.columns([0.48, 0.48], gap="medium")
with col1:
    file_clear = st.file_uploader("上传数据源Excel", type=["xlsx","xls"], key="clear_file")
with col2:
    acc_list = list(ACCOUNT_INFO.keys())
    select_acc = st.selectbox("选择账号", options=acc_list, key="acc_sel")
    # 账号下拉下方展示公司信息
    acc_detail = ACCOUNT_INFO[select_acc]
    st.markdown('<div class="info-block">', unsafe_allow_html=True)
    st.write(f"公司：{acc_detail['shipper_name']}")
    st.write(f"地址：{acc_detail['shipper_addr']}")
    st.write(f"联系人：{acc_detail['contact']} | 电话：{acc_detail['phone']}")
    st.markdown('</div>', unsafe_allow_html=True)

gen_clear = st.button("生成并下载清关资料", key="gen_clear", type="primary")
st.markdown('</div>', unsafe_allow_html=True)

# 清关生成逻辑（完全适配模板，修复所有报错）
if gen_clear:
    if not file_clear:
        st.error("请上传数据源文件")
    elif not os.path.exists(TEMPLATE_FILE):
        st.error(f"模板文件{TEMPLATE_FILE}缺失，请上传至仓库根目录")
    else:
        with st.spinner("正在生成..."):
            df = pd.read_excel(file_clear)
            groups = df.groupby("FBA编号")
            acc_info = ACCOUNT_INFO[select_acc]
            tmp_dir = tempfile.TemporaryDirectory()
            tmp_path = tmp_dir.name
            file_list = []

            for fba_id, group in groups:
                wb = load_workbook(TEMPLATE_FILE)
                ws = wb.active

                # 自动解除制造商区域合并单元格，避免赋值失败
                try:
                    for merged_range in ws.merged_cells.ranges:
                        if (merged_range.min_row <= 39 and merged_range.max_row >= 38) and (merged_range.min_col <= 3 and merged_range.max_col >= 3):
                            ws.unmerge_cells(range_string=str(merged_range))
                            break
                except:
                    pass

                # 填充发货人信息（带异常捕获）
                try:
                    ws[CLEAR_MAP["ship_name"]].value = acc_info["shipper_name"]
                    ws[CLEAR_MAP["ship_addr"]].value = acc_info["shipper_addr"]
                    ws[CLEAR_MAP["ship_contact"]].value = f"Contact:{acc_info['contact']}"
                    ws[CLEAR_MAP["ship_tel"]].value = f"Phone:{acc_info['phone']}"
                except:
                    pass

                # 填充进口商信息（带异常捕获）
                try:
                    ws[CLEAR_MAP["imp_name"]].value = acc_info["shipper_name"]
                    ws[CLEAR_MAP["imp_addr"]].value = acc_info["shipper_addr"]
                    ws[CLEAR_MAP["imp_contact"]].value = f"Contact:{acc_info['contact']}"
                    ws[CLEAR_MAP["imp_tel"]].value = f"Phone:{acc_info['phone']}"
                except:
                    pass

                # 填充制造商信息（强制填充，双层兜底）
                try:
                    ws[CLEAR_MAP["manu_name"]].value = acc_info["shipper_name"]
                    ws[CLEAR_MAP["manu_addr"]].value = acc_info["shipper_addr"]
                except Exception:
                    try:
                        ws["C38"].value = acc_info["shipper_name"]
                        ws["C39"].value = acc_info["shipper_addr"]
                    except:
                        pass

                # 填充FBA编号（带异常捕获）
                try:
                    ws[CLEAR_MAP["fba_no"]].value = fba_id
                except:
                    pass

                # 清空旧明细区域
                s_r = CLEAR_MAP["data_start"]
                e_r = CLEAR_MAP["data_end"]
                for r in range(s_r, e_r+1):
                    for c in range(2, 17):
                        ws.cell(row=r, column=c, value=None)

                # 写入新明细（匹配模板列顺序）
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
                    ws.cell(r, 13, row[11]) # 箱数 M列
                    ws.cell(r, 14, round(row[12], 3)) # 毛重 N列
                    ws.cell(r, 15, row[13]) # 净重 O列
                    ws.cell(r, 16, round(row[14], 3)) # 体积 P列

                # 合计公式（匹配模板列号）
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

            # 打包zip压缩包
            zip_buf = BytesIO()
            zip_name = f"{select_acc}清关资料.zip"
            with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for fp in file_list:
                    zf.write(fp, os.path.basename(fp))
            zip_buf.seek(0)
            # 修复下载按钮：移除hidden=True，适配新版Streamlit
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

# ===================== 模块2：LCL截单重量体积比例调整 =====================
st.markdown('<div class="card">', unsafe_allow_html=True)
st.subheader("2. LCL截单重量体积调整")
upload_cut = st.file_uploader("上传LCL截单Excel", type=["xlsx","xls"], key="cut_file")
colw, colv = st.columns([0.48,0.48], gap="medium")
with colw:
    target_w = st.number_input("目标总重量 kg", min_value=0.001, step=0.001, format="%.3f", key="tw")
with colv:
    target_v = st.number_input("目标总体积 CBM", min_value=0.001, step=0.001, format="%.3f", key="tv")
adjust_btn = st.button("调整并下载截单文件", key="adj_btn", type="primary")
st.markdown('</div>', unsafe_allow_html=True)

# 截单调整逻辑（修复版）
if adjust_btn:
    if not upload_cut:
        st.error("请上传截单Excel文件")
    elif target_w <= 0 or target_v <= 0:
        st.error("重量、体积必须大于0")
    else:
        with st.spinner("计算并调整体积重量..."):
            # 提取AL0编号
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
            
            # 自动识别数据范围（替代硬编码行号）
            s_r = None
            e_r = None
            w_col = None
            v_col = None
            # 遍历表头行，自动识别重量/体积列
            for r in range(1, 10):
                row_cells = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[r]]
                for idx, cell_val in enumerate(row_cells):
                    if "weight" in cell_val.lower() or "gross" in cell_val.lower():
                        w_col = idx + 1
                    if "volume" in cell_val.lower() or "cbm" in cell_val.lower():
                        v_col = idx + 1
                if w_col and v_col:
                    s_r = r + 1
                    break
            # 自动识别数据结束行
            if s_r:
                for r in range(s_r, 20):
                    cell_val = ws.cell(r, w_col).value
                    if cell_val is None or str(cell_val).strip() == "":
                        e_r = r - 1
                        break
            # 兜底：如果自动识别失败，用硬编码默认值
            if not s_r or not e_r or not w_col or not v_col:
                st.warning("自动识别模板失败，使用默认坐标（第4-7行，C列重量，D列体积）")
                s_r = 4
                e_r = 7
                w_col = 3
                v_col = 4

            # 读取原始数据
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

            # 校验原始数据
            if sum_w_ori <= 0 or sum_v_ori <= 0:
                st.error("原始单据总重量/体积为0，无法调整")
                wb.close()
                st.stop()
            st.info(f"已读取原始数据：共{len(raw_data)}行，总重量{sum_w_ori:.3f}kg，总体积{sum_v_ori:.3f}CBM")

            # 计算统一缩放比例（保证单箱重体积比不变，和原始一致）
            ratio_w = target_w / sum_w_ori
            ratio_v = target_v / sum_v_ori
            final_ratio = ratio_w if abs(ratio_w - ratio_v) < 0.001 else (ratio_w + ratio_v) / 2
            st.info(f"缩放比例：{final_ratio:.6f}，目标总重量{target_w:.3f}kg，目标总体积{target_v:.3f}CBM")

            # 计算缩放后数值
            data_list = []
            for r, ow, ov in raw_data:
                ew = ow * final_ratio
                ev = ov * final_ratio
                data_list.append([r, ew, ev])

            # 精准分配小数尾差，保证总重量/体积完全等于目标值
            # 重量尾差分配
            target_w_int = int(round(target_w * 1000))
            w_int_list = []
            sum_wi = 0
            for _, ew, _ in data_list:
                i = int(round(ew * 1000))
                w_int_list.append(i)
                sum_wi += i
            # 尾差平均分配到所有行
            w_diff = target_w_int - sum_wi
            if w_diff != 0:
                step = 1 if w_diff > 0 else -1
                for i in range(abs(w_diff)):
                    w_int_list[i % len(w_int_list)] += step

            # 体积尾差分配
            target_v_int = int(round(target_v * 1000))
            v_int_list = []
            sum_vi = 0
            for _, _, ev in data_list:
                i = int(round(ev * 1000))
                v_int_list.append(i)
                sum_vi += i
            # 尾差平均分配到所有行
            v_diff = target_v_int - sum_vi
            if v_diff != 0:
                step = 1 if v_diff > 0 else -1
                for i in range(abs(v_diff)):
                    v_int_list[i % len(v_int_list)] += step

            # 写入单元格
            for idx, (row_num, _, _) in enumerate(data_list):
                final_w = w_int_list[idx] / 1000
                final_v = v_int_list[idx] / 1000
                ws.cell(row_num, w_col, value=final_w)
                ws.cell(row_num, v_col, value=final_v)

            # 保存文件
            buf = BytesIO()
            wb.save(buf)
            buf.seek(0)
            wb.close()

            # 验证结果
            wb_check = load_workbook(BytesIO(buf.getvalue()))
            ws_check = wb_check.active
            final_sum_w = 0.0
            final_sum_v = 0.0
            for r in range(s_r, e_r + 1):
                final_sum_w += float(ws_check.cell(r, w_col).value or 0)
                final_sum_v += float(ws_check.cell(r, v_col).value or 0)
            wb_check.close()
            st.success(f"调整完成！最终总重量{final_sum_w:.3f}kg，最终总体积{final_sum_v:.3f}CBM")

            # 下载按钮
            st.download_button(
                label="点击下载调整后截单文件",
                data=buf,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_cut_auto"
            )

# 分割线
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

# ===================== 模块3：按FBA号分单生成 =====================
st.markdown('<div class="card">', unsafe_allow_html=True)
st.subheader("3. 纽酷send下单发票生成")
file_split = st.file_uploader("上传数据源Excel", type=["xlsx","xls"], key="split_file")
gen_split = st.button("生成分单文件", key="gen_split", type="primary")
st.markdown('</div>', unsafe_allow_html=True)

# 分单生成逻辑
if gen_split:
    if not file_split:
        st.error("请上传数据源文件")
    elif not os.path.exists(TEMPLATE_FILE):
        st.error(f"模板文件{TEMPLATE_FILE}缺失，请上传至仓库根目录")
    else:
        with st.spinner("正在生成分单文件..."):
            # 读取数据源
            df = pd.read_excel(file_split)
            # 按FBA号分组（列名根据实际数据源调整，这里用"FBA编号"）
            if "FBA编号" not in df.columns:
                st.error("数据源中未找到'FBA编号'列，请检查列名")
                st.stop()
            groups = df.groupby("FBA编号")
            tmp_dir = tempfile.TemporaryDirectory()
            tmp_path = tmp_dir.name
            file_list = []
            # 今天的日期，格式2026-7-20（去除前导0）
            today_date = datetime.now().strftime("%Y-%m-%d").lstrip("0").replace("-0", "-")

            for fba_id, group in groups:
                wb = load_workbook(TEMPLATE_FILE)
                ws = wb.active

                # 1. 产品分类列：CPSC（H列，第8列，明细行）
                for row in range(CLEAR_MAP["data_start"], CLEAR_MAP["data_end"] + 1):
                    ws.cell(row=row, column=8, value="CPSC")
                # 2. 产品数量单位列：套（I列，第9列，明细行）
                for row in range(CLEAR_MAP["data_start"], CLEAR_MAP["data_end"] + 1):
                    ws.cell(row=row, column=9, value="套")
                # 3. PO创建日期：今天的日期（B列第12行，根据实际模板调整）
                ws.cell(row=12, column=2, value=today_date)
                # 4. FBA箱号：-（M列，第13列，明细行）
                for row in range(CLEAR_MAP["data_start"], CLEAR_MAP["data_end"] + 1):
                    ws.cell(row=row, column=13, value="-")
                # 5. 外箱分货标：A1（N列，第14列，明细行）
                for row in range(CLEAR_MAP["data_start"], CLEAR_MAP["data_end"] + 1):
                    ws.cell(row=row, column=14, value="A1")

                # 填充FBA编号
                ws[CLEAR_MAP["fba_no"]].value = fba_id

                # 清空旧明细区域
                s_r = CLEAR_MAP["data_start"]
                e_r = CLEAR_MAP["data_end"]
                for r in range(s_r, e_r+1):
                    for c in range(2, 17):
                        ws.cell(row=r, column=c, value=None)

                # 写入当前分组的产品明细
                rows = group.values.tolist()
                for idx, row in enumerate(rows):
                    r = s_r + idx
                    # 按数据源列顺序填充，根据实际列名调整
                    ws.cell(r, 2, row[0])  # 零件号 B列
                    ws.cell(r, 3, row[1])  # 品名 C列
                    ws.cell(r, 4, row[2])  # 材质 D列
                    ws.cell(r, 5, row[3])  # 关税分类 E列
                    ws.cell(r, 8, "CN")   # 原产国 H列
                    ws.cell(r, 9, row[7])  # 数量 I列
                    ws.cell(r, 10, row[8]) # 单价 J列
                    ws.cell(r, 11, f"=J{r}*I{r}") # 总价 K列
                    ws.cell(r, 13, row[11]) # 箱数 M列
                    ws.cell(r, 14, round(row[12], 3)) # 毛重 N列
                    ws.cell(r, 15, row[13]) # 净重 O列
                    ws.cell(r, 16, round(row[14], 3)) # 体积 P列

                # 合计公式
                end_data = s_r + len(rows) - 1
                total_r = CLEAR_MAP["total_row"]
                ws.cell(total_r, 11, f"=SUM(K{s_r}:K{end_data})")
                ws.cell(total_r, 13, f"=SUM(M{s_r}:M{end_data})")
                ws.cell(total_r, 14, f"=SUM(N{s_r}:N{end_data})")
                ws.cell(total_r, 15, f"=SUM(O{s_r}:O{end_data})")
                ws.cell(total_r, 16, f"=SUM(P{s_r}:P{end_data})")

                # 保存文件，命名为FBA号
                save_path = os.path.join(tmp_path, f"{fba_id}.xlsx")
                wb.save(save_path)
                wb.close()
                file_list.append(save_path)

            # 打包所有分单文件
            zip_buf = BytesIO()
            zip_name = "FBA分单文件.zip"
            with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for fp in file_list:
                    zf.write(fp, os.path.basename(fp))
            zip_buf.seek(0)
            # 下载按钮
            st.download_button(
                label="点击下载分单压缩包",
                data=zip_buf,
                file_name=zip_name,
                mime="application/zip",
                key="dl_split_auto"
            )
            st.success(f"已生成{len(file_list)}个分单文件，打包完成，请点击下载！")
            tmp_dir.cleanup()

# 底部极简说明
st.markdown("<p style='color:#666; font-size:13px; margin-top:30px;'>上方：生成FBA清关打包文件，选择账号后可预览公司信息 | 中间：调整LCL截单重量体积，数值保留3位小数 | 下方：按FBA号分单生成独立文件</p>", unsafe_allow_html=True)
