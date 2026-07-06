# ===================== 模块2：截单重量体积调整（适配LCL AMS模板，解决四舍五入总和偏移） =====================
st.markdown('<div class="section-title">📦 截单资料重量体积比例调整</div>', unsafe_allow_html=True)
st.markdown("<p class='info-text'>适配LCL AMS upload template，自动修正四舍五入误差，最终合计严格等于输入的目标总重/总体积，数值保留3位小数</p>", unsafe_allow_html=True)

st.markdown('<div class="card">', unsafe_allow_html=True)
# 独立文件上传
st.subheader("1. 上传截单Excel文件")
upload_cut = st.file_uploader("", type=["xlsx", "xls"], key="cut_file")
if upload_cut is not None:
    st.success("✅ 已读取截单文件")

# 双输入框
col_w, col_v = st.columns([0.48, 0.48], gap="medium")
with col_w:
    target_w = st.number_input("目标总重量 kg", min_value=0.001, step=0.001, format="%.3f", key="t_w")
with col_v:
    target_v = st.number_input("目标总体积 cbm", min_value=0.001, step=0.001, format="%.3f", key="t_v")

adjust_btn = st.button("🔄 按比例重新计算重量体积", key="adj_btn")
st.markdown('</div>', unsafe_allow_html=True)

# 截单调整逻辑（修复总和偏差）
if adjust_btn:
    if upload_cut is None:
        st.error("❌ 请先上传截单Excel！")
    elif target_w <=0 or target_v <=0:
        st.error("❌ 目标重量/体积必须大于0！")
    else:
        with st.spinner("⏳ 正在按比例重算数据，自动修正四舍五入误差..."):
            wb = load_workbook(upload_cut)
            ws = wb.active
            s_r = CUTTING_CELL_MAP["data_start_row"]
            e_r = CUTTING_CELL_MAP["data_end_row"]
            w_col = CUTTING_CELL_MAP["gross_weight_col"]
            v_col = CUTTING_CELL_MAP["volume_col"]
            header_row = CUTTING_CELL_MAP["header_row"]

            header_w = ws.cell(row=header_row, column=w_col).value
            header_v = ws.cell(row=header_row, column=v_col).value
            if header_w != CUTTING_CELL_MAP["gross_weight_header"] or header_v != CUTTING_CELL_MAP["volume_header"]:
                st.warning(f"⚠️ 表头校验提示：预期毛重列名{CUTTING_CELL_MAP['gross_weight_header']}，体积列名{CUTTING_CELL_MAP['volume_header']}")

            raw_rows = []
            sum_w = 0.0
            sum_v = 0.0
            for r in range(s_r, e_r + 1):
                w_cell = ws.cell(row=r, column=w_col)
                v_cell = ws.cell(row=r, column=v_col)
                try:
                    w_val = float(w_cell.value)
                    v_val = float(v_cell.value)
                    raw_rows.append([r, w_val, v_val])
                    sum_w += w_val
                    sum_v += v_val
                except Exception as e:
                    continue

            if sum_w <= 0 or sum_v <=0:
                st.error(f"❌ 原始总重：{round(sum_w,3)}kg，总体积：{round(sum_v,3)}cbm，无法缩放！")
                wb.close()
                st.stop()

            ratio_w = target_w / sum_w
            ratio_v = target_v / sum_v

            adjusted_list = []
            total_exact_w = 0.0
            total_exact_v = 0.0
            for r, ow, ov in raw_rows:
                ew = ow * ratio_w
                ev = ov * ratio_v
                adjusted_list.append([r, ew, ev])
                total_exact_w += ew
                total_exact_v += ev

            # 计算舍入后产生的差值，最后一行补齐差值，保证总和精准
            round_total_w = round(total_exact_w, 3)
            round_total_v = round(total_exact_v, 3)
            delta_w = target_w - round_total_w
            delta_v = target_v - round_total_v

            # 逐行写入，最后一行补差值
            for idx, (row_num, exact_w, exact_v) in enumerate(adjusted_list):
                if idx == len(adjusted_list) - 1:
                    final_w = round(exact_w, 3) + delta_w
                    final_v = round(exact_v, 3) + delta_v
                else:
                    final_w = round(exact_w, 3)
                    final_v = round(exact_v, 3)
                ws.cell(row=row_num, column=w_col, value=final_w)
                ws.cell(row=row_num, column=v_col, value=final_v)

            buf = BytesIO()
            wb.save(buf)
            buf.seek(0)
            wb.close()

            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.success("✅ 调整完成！已自动修正小数舍入误差，最终合计严格等于目标值")
            st.write(f"📊 原始总重：{round(sum_w,3)} kg → 目标总重：{target_w:.3f} kg")
            st.write(f"📊 原始总体积：{round(sum_v,3)} cbm → 目标总体积：{target_v:.3f} cbm")
            st.markdown('<div class="download-btn">', unsafe_allow_html=True)
            st.download_button("📥 下载调整后截单Excel", buf, "截单资料_调整后.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_cut")
            st.markdown("</div>", unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
