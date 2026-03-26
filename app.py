import streamlit as st
import pandas as pd
import math
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
import io

# --- 엑셀 스타일링 라이브러리 ---
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 한글 폰트 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

def generate_excel_in_memory(raw_data):
    """엑셀 데이터를 메모리에 생성하여 다운로드할 수 있게 반환"""
    df = pd.DataFrame(raw_data)
    df_grouped = df.groupby(["구분", "품명", "재단기장(L)", "상단 가공각(°)", "하단 가공각(°)"]).size().reset_index(name='1대당 수량')
    
    sort_mapping = {
        "상현대(전체)": -2, "하현대(전체)": -1,
        "용마루": 1, "상단용마루": 2, "하단용마루": 3,
        "수평재": 4, "밑더블수평재": 5, 
        "다대": 6, "상단다대": 7, "하단다대": 8,
        "살대": 9, "상단살대": 10, "하단살대": 11,
        "수평내부다대": 12, "수평내부살대": 13, "서브다대": 14, "서브살대": 15
    }
    df_grouped['정렬키'] = df_grouped['구분'].map(sort_mapping).fillna(99)
    df_grouped = df_grouped.sort_values(by=["정렬키", "재단기장(L)"], ascending=[True, False]).drop('정렬키', axis=1)
    
    df_grouped.insert(0, '순번', range(1, len(df_grouped) + 1))
    df_grouped["총 소요 수량"] = ""
    df_grouped = df_grouped[["순번", "구분", "품명", "1대당 수량", "총 소요 수량", "재단기장(L)", "상단 가공각(°)", "하단 가공각(°)"]]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_grouped.to_excel(writer, sheet_name='통합 재단표', index=False, startrow=2)
        ws = writer.sheets['통합 재단표']
        
        ws.merge_cells('A1:C1')
        ws['A1'] = "👉 트러스 총 제작 수량 (EA) :"
        ws['A1'].font = Font(bold=True, size=12)
        ws['A1'].alignment = Alignment(horizontal="right", vertical="center")
        
        ws['D1'] = 1
        ws['D1'].fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        ws['D1'].font = Font(color="FF0000", bold=True, size=14)
        ws['D1'].alignment = Alignment(horizontal="center", vertical="center")
        
        thin_border = Border(left=Side(style='thin', color='A6A6A6'), right=Side(style='thin', color='A6A6A6'),
                             top=Side(style='thin', color='A6A6A6'), bottom=Side(style='thin', color='A6A6A6'))
        ws['D1'].border = thin_border

        color_map = {
            "상현대(전체)": PatternFill(start_color="D0CECE", end_color="D0CECE", fill_type="solid"),
            "하현대(전체)": PatternFill(start_color="AEAAAA", end_color="AEAAAA", fill_type="solid"),
            "용마루": PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
            "상단용마루": PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid"),
            "하단용마루": PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid"),
            "수평재": PatternFill(start_color="D2B4DE", end_color="D2B4DE", fill_type="solid"), 
            "밑더블수평재": PatternFill(start_color="E8DAEF", end_color="E8DAEF", fill_type="solid"),
            "다대": PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"),
            "상단다대": PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid"),
            "하단다대": PatternFill(start_color="8EA9DB", end_color="8EA9DB", fill_type="solid"),
            "살대": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
            "상단살대": PatternFill(start_color="C6E0B4", end_color="C6E0B4", fill_type="solid"),
            "하단살대": PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid"),
            "수평내부다대": PatternFill(start_color="A9DFBF", end_color="A9DFBF", fill_type="solid"), 
            "수평내부살대": PatternFill(start_color="F9E79F", end_color="F9E79F", fill_type="solid"), 
            "서브다대": PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),
            "서브살대": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        }
        
        for r_idx, row in enumerate(ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=ws.max_column), 3):
            gubun_val = ws.cell(row=r_idx, column=2).value if r_idx > 3 else None
            for c_idx, cell in enumerate(row, 1):
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
                
                if r_idx == 3:
                    cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                    cell.font = Font(color="FFFFFF", bold=True)
                else:
                    header = ws.cell(row=3, column=c_idx).value
                    if header == "순번": cell.font = Font(bold=True)
                    elif header == "구분":
                        cell.fill = color_map.get(gubun_val, PatternFill(fill_type=None))
                        cell.font = Font(bold=True)
                    elif header == "총 소요 수량":
                        cell.value = f'=$D$1*D{r_idx}' 
                        cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                        cell.font = Font(color="0070C0", bold=True)
                    elif header in ["1대당 수량", "재단기장(L)", "상단 가공각(°)", "하단 가공각(°)"]:
                        cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                        cell.font = Font(color="C00000", bold=True)
                    elif r_idx % 2 == 0: cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

        # 레이저 가공 싸이즈 표 로직
        start_col = 10 
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=start_col+4)
        title_cell = ws.cell(row=1, column=start_col, value="레이저 가공 싸이즈")
        title_cell.font = Font(bold=True, size=14)
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        title_cell.fill = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")
        title_cell.border = thin_border
        
        headers_new = ["구분", "재단기장\n(반올림)", "상단 가공각\n(올림)", "하단 가공각\n(올림)", "총수량"]
        for i, h in enumerate(headers_new):
            c = ws.cell(row=3, column=start_col+i, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = thin_border

        counters = {}
        row_idx_new = 4
        for _, row_data in df_grouped.iterrows():
            cat_base = row_data["구분"]
            if cat_base not in counters:
                counters[cat_base] = 1
            name_val = f"{cat_base}{counters[cat_base]}"
            counters[cat_base] += 1
            
            L_val = row_data["재단기장(L)"]
            top_a = row_data["상단 가공각(°)"]
            bot_a = row_data["하단 가공각(°)"]
            qty_val = row_data["1대당 수량"]
            
            L_rounded = round(float(L_val)) if pd.notnull(L_val) else 0
            top_ceil = math.ceil(float(top_a)) if pd.notnull(top_a) else 0
            bot_ceil = math.ceil(float(bot_a)) if pd.notnull(bot_a) else 0
            
            ws.cell(row=row_idx_new, column=start_col, value=name_val)
            ws.cell(row=row_idx_new, column=start_col+1, value=L_rounded)
            ws.cell(row=row_idx_new, column=start_col+2, value=top_ceil)
            ws.cell(row=row_idx_new, column=start_col+3, value=bot_ceil)
            
            ws.cell(row=row_idx_new, column=start_col+4, value=f"=$D$1*{qty_val}") 
            ws.cell(row=row_idx_new, column=start_col+4).font = Font(color="0070C0", bold=True)
            
            for i in range(5):
                c = ws.cell(row=row_idx_new, column=start_col+i)
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = thin_border
                if row_idx_new % 2 == 0: 
                    c.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            
            row_idx_new += 1

        for col_idx in range(1, ws.max_column + 1):
            col_letter = get_column_letter(col_idx)
            max_len = max([len(str(ws.cell(row=r, column=col_idx).value)) for r in range(3, ws.max_row + 1) if ws.cell(row=r, column=col_idx).value] + [0])
            ws.column_dimensions[col_letter].width = max(max_len + 6, 12)

    return output.getvalue()

def create_truss_figure(type_choice, span_cm, divs, h_outer_cm, h_center_cm, h_tie_cm, m_od, v_od, r_od, d_od, offset_mm):
    """트러스 도면 Figure와 Raw Data를 생성하여 반환"""
    
    # 대표님 기존 로직 매핑
    type_num_map = {
        "대칭삼각(평탄)": "1", "아치형(평탄)": "2", "반삼각(평탄)": "3",
        "A자형_삼각": "4", "A자형_아치": "5", "A자형_반삼각": "6",
        "밑더블_삼각": "7", "밑더블_아치": "8", "밑더블_반삼각": "9"
    }
    tc = type_num_map.get(type_choice, "7")
    
    has_tie = tc in ["4", "5", "6"]
    is_double_bot = tc in ["7", "8", "9"]
    is_half = tc in ["3", "6", "9"]
    is_arch = tc in ["2", "5", "8"]

    S, H_out, H_cen, H_tie = span_cm * 10, h_outer_cm * 10, h_center_cm * 10, h_tie_cm * 10

    yc, R = 0, 0
    if is_arch:
        if H_cen <= H_out: H_cen = H_out + 10
        yc = ((S/2)**2 + H_out**2 - H_cen**2) / (2 * (H_out - H_cen))
        R = H_cen - yc

    max_y_bot = H_cen - H_out
    if has_tie and H_tie >= max_y_bot:
        H_tie = 0

    def get_y_top(x):
        if x < 0: x = 0
        if x > S: x = S
        if tc in ["1", "4", "7"]:
            m = (H_cen - H_out) / (S/2)
            return H_out + m * x if x <= S/2 else H_out + m * (S - x)
        elif is_arch:
            val = max(R**2 - (x - S/2)**2, 0)
            return yc + math.sqrt(val)
        elif is_half:
            return H_out + (x / S) * (H_cen - H_out)

    def get_y_bot(x):
        if has_tie: return max(get_y_top(x) - H_out, 0.0)
        else: return 0.0

    def get_slope(func, x):
        dx = 0.1
        test_x = x if x + dx <= S else x - dx
        dy = func(test_x + dx) - func(test_x)
        return math.degrees(math.atan2(dy, dx))
        
    def get_cos(func, x):
        dx = 0.1
        test_x = x if x + dx <= S else x - dx
        dy = func(test_x + dx) - func(test_x)
        cos_val = math.cos(math.atan2(dy, dx))
        return cos_val if cos_val != 0 else 0.0001
        
    def get_thick(func, x, od):
        return od / get_cos(func, x)

    def draw_dim_text(ax, x, y, text, angle=0, color='black', fontsize=11.5):
        if angle > 90: angle -= 180
        elif angle < -90: angle += 180
        ax.text(x, y, text, color=color, fontsize=fontsize, fontweight='bold', ha='center', va='center', rotation=angle,
                bbox=dict(facecolor='white', alpha=0.85, edgecolor='none', pad=1.5))

    fig, ax = plt.subplots(figsize=(40, 18), dpi=100)
    mid_idx = divs if is_half else divs // 2
    raw_data = []

    v_centers_x = [v_od/2 if i==0 else (S - (v_od if is_half else r_od)/2 if i==divs else i*(S/divs)) for i in range(divs + 1)]

    end_thick = get_thick(get_y_top, 0, m_od)
    H_mid_top = H_out - end_thick if is_double_bot else H_out
    H_mid_bot = H_mid_top - m_od if is_double_bot else H_out - m_od

    total_top_chord_len = sum(math.hypot(v_centers_x[i+1] - v_centers_x[i], get_y_top(v_centers_x[i+1]) - get_y_top(v_centers_x[i])) for i in range(divs))
    total_bot_chord_len = sum(math.hypot(v_centers_x[i+1] - v_centers_x[i], get_y_bot(v_centers_x[i+1]) - get_y_bot(v_centers_x[i])) for i in range(divs))
    
    raw_data.append({
        "구분": "상현대(전체)", "품명": f"{m_od}mm 파이프",
        "재단기장(L)": round(total_top_chord_len, 1), "상단 가공각(°)": 0.0, "하단 가공각(°)": 0.0
    })
    raw_data.append({
        "구분": "하현대(전체)", "품명": f"{m_od}mm 파이프",
        "재단기장(L)": round(total_bot_chord_len, 1), "상단 가공각(°)": 0.0, "하단 가공각(°)": 0.0
    })

    if is_double_bot:
        total_mid_pipe_len = S
        ax.add_patch(patches.Rectangle((0, H_mid_bot), total_mid_pipe_len, m_od, facecolor='#9b59b6', edgecolor='black', zorder=6))
        draw_dim_text(ax, S/2, H_mid_bot + m_od/2, f"밑더블 수평재(전체) L:{total_mid_pipe_len:.1f}", angle=0, color='purple', fontsize=14)
        raw_data.append({
            "구분": "밑더블수평재", "품명": f"{m_od}mm 파이프",
            "재단기장(L)": round(total_mid_pipe_len, 1), "상단 가공각(°)": 0.0, "하단 가공각(°)": 0.0
        })

    is_diag = True 
    
    # 1. 상하현부 메인 다대 및 용마루
    for i in range(divs + 1):
        x = v_centers_x[i]
        is_ridge = (i == mid_idx) and not is_half
        curr_v_od = r_od if is_ridge else v_od
        
        x_l, x_r = x - curr_v_od/2, x + curr_v_od/2
        
        yt_l = get_y_top(x_l) - get_thick(get_y_top, x_l, m_od)
        yt_r = get_y_top(x_r) - get_thick(get_y_top, x_r, m_od)
        yb_l = get_y_bot(x_l) + get_thick(get_y_bot, x_l, m_od)
        yb_r = get_y_bot(x_r) + get_thick(get_y_bot, x_r, m_od)
        
        y_bot_c = get_y_bot(x) + get_thick(get_y_bot, x, m_od)
        y_top_c = get_y_top(x) - get_thick(get_y_top, x, m_od)
        
        t_angle = round(abs(get_slope(get_y_top, x)), 1)
        b_angle = round(abs(get_slope(get_y_bot, x)), 1)
        
        if is_ridge: v_cut_l = y_top_c - min(yb_l, yb_r)
        else: v_cut_l = max(yt_l, yt_r) - min(yb_l, yb_r) 

        if has_tie and is_ridge and H_tie > 0:
            if tc != "5":
                yb_l = yb_r = y_bot_c = H_tie + m_od/2
                v_cut_l = y_top_c - y_bot_c
                b_angle = 0.0

        if is_double_bot:
            u_len = y_top_c - H_mid_top
            if u_len > 0:
                u_cut_max = y_top_c - H_mid_top if is_ridge else max(yt_l, yt_r) - H_mid_top
                ax.add_patch(patches.Rectangle((x - curr_v_od/2, H_mid_top), curr_v_od, u_len, facecolor='#2980b9', edgecolor='black', zorder=5))
                g_name = "상단용마루" if is_ridge else "상단다대"
                raw_data.append({
                    "구분": g_name, "품명": f"{curr_v_od}mm 파이프",
                    "재단기장(L)": round(u_cut_max, 1), "상단 가공각(°)": t_angle, "하단 가공각(°)": 0.0
                })
                stagger_top = 600 if i % 2 == 0 else 900
                my_top = y_top_c + stagger_top
                ax.plot([x, x], [y_top_c + m_od/2, my_top - 180], color='blue', linestyle=':', lw=1.5, zorder=1)
                draw_dim_text(ax, x, my_top, f"{g_name}\nL:{u_cut_max:.1f}", angle=90, color='blue', fontsize=10)

            l_len = H_mid_bot - y_bot_c
            if l_len > 0:
                l_cut_max = H_mid_bot - min(yb_l, yb_r) if not is_ridge else H_mid_bot - y_bot_c
                ax.add_patch(patches.Rectangle((x - curr_v_od/2, y_bot_c), curr_v_od, l_len, facecolor='#34495e', edgecolor='black', zorder=5))
                g_name = "하단용마루" if is_ridge else "하단다대"
                raw_data.append({
                    "구분": g_name, "품명": f"{curr_v_od}mm 파이프",
                    "재단기장(L)": round(l_cut_max, 1), "상단 가공각(°)": 0.0, "하단 가공각(°)": b_angle
                })
                stagger_bot = 600 if i % 2 == 0 else 900
                my_bot = y_bot_c - stagger_bot
                ax.plot([x, x], [y_bot_c - m_od/2, my_bot + 180], color='darkblue', linestyle=':', lw=1.5, zorder=1)
                draw_dim_text(ax, x, my_bot, f"{g_name}\nL:{l_cut_max:.1f}", angle=90, color='darkblue', fontsize=10)
        else:
            v_draw_h = y_top_c - y_bot_c
            ax.add_patch(patches.Rectangle((x - curr_v_od/2, y_bot_c), curr_v_od, v_draw_h, facecolor='#2c3e50', edgecolor='black', zorder=5))
            text_color = 'red' if is_ridge else 'blue'
            stagger_offset = 600 if i % 2 == 0 else 900
            my = y_top_c + stagger_offset
            ax.plot([x, x], [y_top_c + m_od/2, my - 180], color=text_color, linestyle=':', lw=1.5, zorder=1)
            draw_dim_text(ax, x, my, f"L:{v_cut_l:.1f} (상:{t_angle}°/하:{b_angle}°)", angle=90, color=text_color)
            
            v_gubun = "용마루" if is_ridge else "다대"
            raw_data.append({
                "구분": v_gubun, "품명": f"{curr_v_od}mm 파이프",
                "재단기장(L)": round(v_cut_l, 1), "상단 가공각(°)": t_angle, "하단 가공각(°)": b_angle
            })

    # 2. 상/하현부 테두리 및 메인 살대
    for i in range(divs):
        x, nx = v_centers_x[i], v_centers_x[i+1]
        
        pb1, pb2 = (x, get_y_bot(x)), (nx, get_y_bot(nx))
        pb3, pb4 = (nx, get_y_bot(nx) + get_thick(get_y_bot, nx, m_od)), (x, get_y_bot(x) + get_thick(get_y_bot, x, m_od))
        ax.add_patch(patches.Polygon(np.array([pb1, pb2, pb3, pb4]), facecolor='#7f8c8d', alpha=0.5, zorder=2))
        
        pt1, pt2 = (x, get_y_top(x) - get_thick(get_y_top, x, m_od)), (nx, get_y_top(nx) - get_thick(get_y_top, nx, m_od))
        pt3, pt4 = (nx, get_y_top(nx)), (x, get_y_top(x))
        ax.add_patch(patches.Polygon(np.array([pt1, pt2, pt3, pt4]), facecolor='#7f8c8d', zorder=7))

        if is_diag:
            is_r_curr = (i == mid_idx) and not is_half
            is_r_next = (i+1 == mid_idx) and not is_half
            c_v_od, n_v_od = (r_od if is_r_curr else v_od), (r_od if is_r_next else v_od)

            wx_start = x + c_v_od/2 + offset_mm
            wx_end = nx - n_v_od/2 - offset_mm
            
            def draw_diag(px_bot, px_top):
                py_bot = get_y_bot(px_bot) + get_thick(get_y_bot, px_bot, m_od)
                py_top = get_y_top(px_top) - get_thick(get_y_top, px_top, m_od)
                diag_l = math.hypot(px_top - px_bot, py_top - py_bot)
                
                color = '#c0392b'
                ax.plot([px_bot, px_top], [py_bot, py_top], color=color, lw=2.5, zorder=3)
                
                dx_line, dy_line = px_top - px_bot, py_top - py_bot
                diag_ang = math.degrees(math.atan2(dy_line, dx_line))
                
                t_slope = get_slope(get_y_top, px_top)
                b_slope = get_slope(get_y_bot, px_bot)
                
                t_intersect = abs(diag_ang - t_slope) % 180
                if t_intersect > 90: t_intersect = 180 - t_intersect
                d_top_angle = round(abs(90.0 - t_intersect), 1)
                
                b_intersect = abs(diag_ang - b_slope) % 180
                if b_intersect > 90: b_intersect = 180 - b_intersect
                d_bot_angle = round(abs(90.0 - b_intersect), 1)

                mx, my = (px_bot + px_top) / 2, (py_bot + py_top) / 2
                draw_dim_text(ax, mx, my, f"L:{diag_l:.1f} ({d_top_angle}°/{d_bot_angle}°)", angle=diag_ang, color='#8B0000', fontsize=11)

                raw_data.append({
                    "구분": "살대", "품명": f"{d_od}mm 파이프",
                    "재단기장(L)": round(diag_l, 1), "상단 가공각(°)": d_top_angle, "하단 가공각(°)": d_bot_angle
                })

            if is_double_bot:
                def draw_custom_diag(x_bot, y_bot, x_top, y_top, color, gubun_name):
                    diag_l = math.hypot(x_top - x_bot, y_top - y_bot)
                    ax.plot([x_bot, x_top], [y_bot, y_top], color=color, lw=2.5, zorder=3)
                    
                    dx_line, dy_line = x_top - x_bot, y_top - y_bot
                    diag_ang = math.degrees(math.atan2(dy_line, dx_line))
                    
                    if gubun_name == "상단살대":
                        t_slope = get_slope(get_y_top, x_top)
                        b_slope = 0.0
                    else:
                        t_slope = 0.0 
                        b_slope = get_slope(get_y_bot, x_bot)
                        
                    t_intersect = abs(diag_ang - t_slope) % 180
                    if t_intersect > 90: t_intersect = 180 - t_intersect
                    d_top_angle = round(abs(90.0 - t_intersect), 1)
                    
                    b_intersect = abs(diag_ang - b_slope) % 180
                    if b_intersect > 90: b_intersect = 180 - b_intersect
                    d_bot_angle = round(abs(90.0 - b_intersect), 1)
                    
                    mx, my = (x_bot + x_top)/2, (y_bot + y_top)/2
                    draw_dim_text(ax, mx, my, f"L:{diag_l:.1f} ({d_top_angle}°/{d_bot_angle}°)", angle=diag_ang, color='#8B0000', fontsize=10)
                    
                    raw_data.append({
                        "구분": gubun_name, "품명": f"{d_od}mm 파이프",
                        "재단기장(L)": round(diag_l, 1), "상단 가공각(°)": d_top_angle, "하단 가공각(°)": d_bot_angle
                    })

                if is_half or i < mid_idx:
                    px_bot_u, px_top_u = wx_end, wx_start
                else:
                    px_bot_u, px_top_u = wx_start, wx_end
                    
                py_bot_u = H_mid_top
                py_top_u = get_y_top(px_top_u) - get_thick(get_y_top, px_top_u, m_od)
                if py_top_u > py_bot_u:
                    draw_custom_diag(px_bot_u, py_bot_u, px_top_u, py_top_u, '#c0392b', "상단살대")

                lower_y_top = H_mid_bot
                
                if is_half: is_forward = (i % 2 == 0)
                else:
                    if i < mid_idx: is_forward = ((mid_idx - 1 - i) % 2 != 0) 
                    else: is_forward = ((i - mid_idx) % 2 == 0)
                        
                if is_forward:
                    px_bot_l, px_top_l = wx_start, wx_end
                else:
                    px_bot_l, px_top_l = wx_end, wx_start
                    
                lower_y_bot_actual = get_y_bot(px_bot_l) + get_thick(get_y_bot, px_bot_l, m_od)
                if lower_y_top > lower_y_bot_actual:
                    draw_custom_diag(px_bot_l, lower_y_bot_actual, px_top_l, lower_y_top, '#d35400', "하단살대")

            else:
                if is_half or i < mid_idx: draw_diag(wx_end, wx_start)
                else: draw_diag(wx_start, wx_end)

    # 4. 전체 스판 및 등분 자간 / 누적 거리 표시
    dim_y = -1500 if is_double_bot else -350
    ax.plot([0, S], [dim_y, dim_y], color='black', lw=2, zorder=10)
    ax.plot([0, 0], [dim_y - 25, dim_y + 25], color='black', lw=2, zorder=10)
    ax.plot([S, S], [dim_y - 25, dim_y + 25], color='black', lw=2, zorder=10)
    ax.text(S/2, dim_y - 80, f"전체 스판 : {S:.1f} mm", ha='center', va='center', fontsize=18, fontweight='bold', color='black')

    interval_len = S / divs
    for i in range(divs):
        x, nx = v_centers_x[i], v_centers_x[i+1]
        cx = (x + nx) / 2
        
        if i > 0: ax.plot([x, x], [dim_y - 20, dim_y + 20], color='black', lw=1.5, zorder=10)
        ax.plot([x, x], [0, dim_y], color='gray', linestyle=':', lw=1.5, zorder=1)
        
        f_size = 12 if interval_len > 300 else 10
        ax.text(cx, dim_y + 40, f"{interval_len:.1f}", ha='center', va='center', fontsize=f_size, color='navy', fontweight='bold')

    ax.plot([S, S], [0, dim_y], color='gray', linestyle=':', lw=1.5, zorder=1)

    dim_y_cum = dim_y - 300 
    ax.plot([0, S], [dim_y_cum, dim_y_cum], color='black', lw=1.5, zorder=10)
    ax.plot([0, 0], [dim_y_cum - 25, dim_y_cum + 25], color='black', lw=1.5, zorder=10)
    
    ax.text(0, dim_y_cum - 120, "0", ha='center', va='center', fontsize=11, color='teal', fontweight='bold')
    ax.plot([0, 0], [dim_y, dim_y_cum], color='gray', linestyle=':', lw=1.5, zorder=1)

    for i in range(1, divs + 1):
        x = v_centers_x[i]
        ax.plot([x, x], [dim_y_cum - 20, dim_y_cum + 20], color='black', lw=1.5, zorder=10)
        ax.plot([x, x], [dim_y, dim_y_cum], color='gray', linestyle=':', lw=1.5, zorder=1)
        ax.text(x, dim_y_cum - 120, f"{x:.1f}", ha='center', va='center', fontsize=11, color='teal', fontweight='bold', rotation=90)

    ax.set_xlim(-200, S + 200)
    ax.set_ylim(dim_y_cum - 400, H_cen + 1200) 
    ax.set_aspect('equal')
    ax.axis('off') 
    
    info_text = f"스판: {span_cm}cm | 등분: {divs} (자간: {interval_len/10:.1f}cm)"
    if has_tie: info_text += f" | 수평재 높이: {h_tie_cm}cm"
    if is_double_bot: info_text += f" | 밑더블 외경 높이: {h_outer_cm}cm"
    
    plt.title(f"트러스 도면 ({type_choice})\n{info_text}", fontsize=24, fontweight='bold', pad=20)
    
    return fig, raw_data

# ==========================================
# Streamlit 웹 앱 UI 구성
# ==========================================
st.set_page_config(page_title="트러스 자동 산출기", layout="wide")
st.title("🛠️ 트러스 도면 & 재단표 자동 산출기")

with st.sidebar:
    st.header("1. 기본 설정")
    type_choice = st.selectbox("트러스 형태", [
        "밑더블_삼각", "밑더블_아치", "밑더블_반삼각", 
        "대칭삼각(평탄)", "아치형(평탄)", "반삼각(평탄)",
        "A자형_삼각", "A자형_아치", "A자형_반삼각"
    ])
    
    span_cm = st.number_input("전체 스판 (cm)", value=1200.0, step=10.0)
    divs = st.number_input("등분 수 (다대 개수)", value=34, step=2)
    
    st.header("2. 높이 설정")
    h_outer_cm = st.number_input("외경/시작단 높이 (cm)", value=80.0, step=5.0)
    h_center_cm = st.number_input("최고점 높이 (cm)", value=250.0, step=10.0)
    h_tie_cm = st.number_input("수평보 높이 (cm, A자형 전용)", value=150.0, step=5.0)
    
    st.header("3. 파이프 규격 (mm)")
    m_od = st.number_input("상/하현대 파이프 외경", value=59.9, step=1.0)
    v_od = st.number_input("다대 파이프 외경", value=38.1, step=1.0)
    r_od = st.number_input("용마루 파이프 외경", value=59.9, step=1.0)
    d_od = st.number_input("살대 파이프 외경", value=31.8, step=1.0)
    offset_mm = st.number_input("살대 이격 거리", value=20.0, step=5.0)
    
    generate_btn = st.button("🚀 도면 및 재단표 생성하기", use_container_width=True)

if generate_btn:
    with st.spinner("도면 및 수치를 계산하는 중입니다..."):
        try:
            fig, raw_data = create_truss_figure(
                type_choice, span_cm, divs, h_outer_cm, h_center_cm, h_tie_cm, 
                m_od, v_od, r_od, d_od, offset_mm
            )
            
            st.subheader("📊 생성된 트러스 도면")
            st.pyplot(fig)
            
            pdf_buffer = io.BytesIO()
            fig.savefig(pdf_buffer, format="pdf", bbox_inches="tight")
            pdf_data = pdf_buffer.getvalue()
            
            excel_data = generate_excel_in_memory(raw_data)
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📄 도면 PDF 다운로드",
                    data=pdf_data,
                    file_name=f"Truss_{type_choice}_{int(span_cm)}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            with col2:
                st.download_button(
                    label="📊 재단표 엑셀 다운로드",
                    data=excel_data,
                    file_name=f"Truss_{type_choice}_{int(span_cm)}_재단표.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
