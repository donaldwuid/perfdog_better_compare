import os
import ntpath
import sys
import argparse
import statistics
import re
import math
from datetime import datetime
try:
    import pandas as pd
except ImportError:
    os.system('pip install pandas')
    import pandas as pd
try:
    import json
except ImportError:
    os.system('pip install json')
    import json
try:
    import openpyxl
    from openpyxl.styles import Alignment, PatternFill
    from openpyxl.formatting.rule import DataBarRule
except ImportError:
    os.system('pip install openpyxl')
    import openpyxl
    from openpyxl.styles import Alignment, PatternFill
    from openpyxl.formatting.rule import DataBarRule
try:
    import numpy as np
except ImportError:
    os.system('pip install numpy')
    import numpy as np
try:
    import matplotlib as mpl
except ImportError:
    os.system('pip install matplotlib')
    import matplotlib as mpl




def process_data(input_data_list, input_perfdog_config, output_xlsx, divided_by_framerate, divided_by_resolution, compare_target_column_name, compare_target_name, show_only_columns_in_config, sort_vs_by_value=True):

    input_data_list = [os.path.normpath(path) for path in input_data_list]
    if not output_xlsx:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_xlsx = os.path.splitext(input_data_list[0])[0] + f"_better_compare_{timestamp}.xlsx"

    print("=" * 60)
    print("  PerfDog Better Compare")
    print("=" * 60)
    print(f"  Input  : {input_data_list}")
    print(f"  Output : {output_xlsx}")
    print(f"  Normalize by framerate  : {divided_by_framerate}")
    print(f"  Normalize by resolution : {divided_by_resolution}")
    print(f"  Compare column : {compare_target_column_name or '(auto)'}")
    print(f"  Compare target : {compare_target_name or '(auto)'}")
    print(f"  Show only config columns: {show_only_columns_in_config}")
    print(f"  Sort VS sheet  : {sort_vs_by_value}")

    if not input_perfdog_config:
        current_script_path = os.path.abspath(__file__)
        current_script_dir = os.path.dirname(current_script_path)
        input_perfdog_config = current_script_dir + "/perfdog_export_better_compare_config.json"
        print(f"  Config : (default) {input_perfdog_config}")
    input_perfdog_config = os.path.normpath(input_perfdog_config)
    print(f"  Config : {input_perfdog_config}")
    print("-" * 60)

    try:
        # read the json config
        with open(input_perfdog_config, encoding='utf-8') as json_file:
            config_data = json.load(json_file)
    except FileNotFoundError:
        print(f"Error: File '{input_perfdog_config}' not found.")
        sys.exit(1)
    except IOError:
        print(f"Error: Unable to open file '{input_perfdog_config}'.")
        sys.exit(1)

    if not compare_target_column_name:
        compare_target_column_name = config_data["test_case_column"]
        print(f"  [auto] compare_target_column_name -> '{compare_target_column_name}'")

    # save original value; resolve after data is loaded
    _auto_compare_target_name = compare_target_name

    # convert kilo/mega/giga prefixes to numeric values
    def convert_unit_prefix(cell_value):
        if 'kilo' in str(cell_value).lower():
            return float(cell_value.lower().replace('kilo', '')) * 1e3
        elif 'mega' in str(cell_value).lower():
            return float(cell_value.lower().replace('mega', '')) * 1e6
        elif 'giga' in str(cell_value).lower():
            return float(cell_value.lower().replace('giga', '')) * 1e9
        else:
            return cell_value

    input_df_list = list()
    for one_input_data_name in input_data_list:
        temp_df = pd.read_excel(one_input_data_name)
        temp_df.columns = temp_df.columns.str.replace('\n', ' ')
        for column in temp_df.columns:
            temp_df[column] = temp_df[column].apply(convert_unit_prefix)
        input_df_list.append(temp_df)

    # merge all input files
    df = input_df_list[0].copy()
    for input_df in input_df_list[1:]:
        df = pd.concat([df, input_df], ignore_index=True)

    # aggregate: mean for numeric columns, newline-join for text columns
    def custom_agg(x):
        if pd.to_numeric(x, errors='coerce').notna().all():
            return pd.to_numeric(x, errors='coerce').mean()
        else:
            return x.astype(str).str.cat(sep='\n')

    df = df.groupby(compare_target_column_name).agg(custom_agg).reset_index()

    # default compare target: first row in the grouped data
    if not _auto_compare_target_name and compare_target_column_name in df.columns and len(df) > 1:
        _auto_compare_target_name = df[compare_target_column_name].iloc[0]
        print(f"  [auto] compare_target_name -> '{_auto_compare_target_name}'")
    compare_target_name = _auto_compare_target_name

    # prompt user for missing important columns
    important_columns = [config_data["average_framerate_column"]]
    for column in important_columns:
        if column not in df.columns:
            print(f"  [warn] Missing important column: '{column}'")
            for index, row in df.iterrows():
                value = input(f"  Please enter value for test case '{row[compare_target_column_name]}' / column '{column}': ")
                df.at[index, column] = value

    resolution_columns = [config_data["resolution_width_column"], config_data["resolution_height_column"]]
    for column in resolution_columns:
        if column not in df.columns:
            if divided_by_resolution:
                print(f"  [warn] Missing resolution column: '{column}'")
            else:
                print(f"  [info] Missing resolution column: '{column}', reset to 1")
            for index, row in df.iterrows():
                if divided_by_resolution:
                    value = input(f"  Please enter value for test case '{row[compare_target_column_name]}' / column '{column}': ")
                    df.at[index, column] = value
                else:
                    df.at[index, column] = 1


    # make the normalized data sheet copy
    df_processed = df.copy()

    # only show specified columns
    if show_only_columns_in_config:
        all_config_columns = [config_data["average_framerate_column"], config_data["resolution_height_column"], config_data["resolution_width_column"]]
        all_config_columns += config_data['columns_always_shown'] + config_data['columns_divide_by_framerate'] + config_data['columns_divide_by_resolution'] + config_data['columns_divide_by_framerate_and_resolution']
        columns_to_keep = [column for column in df_processed.columns if any(config_col in column for config_col in all_config_columns)]
        df_processed = df_processed[columns_to_keep]

    # processing the data
    for column in df_processed.columns:
        if any(config_col in column for config_col in config_data['columns_divide_by_framerate']):
            if divided_by_framerate:
                df_processed[column] = df_processed[column].astype(float) / df_processed[config_data["average_framerate_column"]].astype(float)
                df_processed.rename(columns={column: column + '\n(/framerate)'}, inplace=True)
        elif any(config_col in column for config_col in config_data['columns_divide_by_resolution']):
            if divided_by_resolution:
                df_processed[column] = df_processed[column].astype(float) / (df_processed[config_data["resolution_height_column"]].astype(float) * df_processed[config_data["resolution_width_column"]].astype(float))
                df_processed.rename(columns={column: column + '\n(/resolution)'}, inplace=True)
        elif any(config_col in column for config_col in config_data['columns_divide_by_framerate_and_resolution']):
            if divided_by_framerate and divided_by_resolution:
                df_processed[column] = df_processed[column].astype(float) / (df_processed[config_data["average_framerate_column"]].astype(float) * df_processed[config_data["resolution_height_column"]].astype(float) * df_processed[config_data["resolution_width_column"]].astype(float))
                df_processed.rename(columns={column: column + '\n(/(framerate*resolution))'}, inplace=True)


    # df_compare_target is no longer used for sheet generation (merged into CombinedCompare)
    # kept for has_compare / sort logic reference
    df_compare_target = None
    raw_values_dict = {}
    target_project_compare_sheet_name = ""
    if compare_target_column_name and compare_target_name:
        target_row = df_processed.loc[df_processed[compare_target_column_name] == compare_target_name]
        if not target_row.empty:
            df_compare_target = True  # sentinel: target exists
            print(f"  [info] Compare target found: '{compare_target_name}'")
        else:
            print(f"  [warn] Cannot find compare target '{compare_target_name}', percentage columns will not be generated.")


    def path_leaf(path):
        head, tail = ntpath.split(path)
        return tail or ntpath.basename(head)

    # ------------------------------------------------------------------ #
    #  Build the combined transposed sheet directly into openpyxl          #
    #  Layout (after transpose):                                           #
    #    Row 1 (header): metric name | Target | Target% | A | A% | B | B% #
    #    Row 2+: values per case                                           #
    # ------------------------------------------------------------------ #

    frozen_n = config_data.get('forzen_column_num', 3)  # number of header rows to freeze (用例/项目/场景)
    important_background_fill = PatternFill(start_color="f9d3e3", end_color="f9d3e3", fill_type="solid")

    wb_out = openpyxl.Workbook()

    # ---- build CombinedCompare sheet ---------------------------------- #
    ws_combined = wb_out.active
    ws_combined.title = 'CombinedCompare'

    # determine column order in df_processed
    all_metric_cols = list(df_processed.columns)   # first frozen_n are header cols (用例/项目/场景)
    header_cols = all_metric_cols[:frozen_n]        # e.g. [用例, 项目, 场景]
    metric_cols_only = all_metric_cols[frozen_n:]   # actual metric columns

    # get all case names and target row
    all_case_names = list(df_processed[compare_target_column_name])
    has_compare = (df_compare_target is not None and compare_target_name)
    target_row_data = None
    if has_compare:
        target_mask = df_processed[compare_target_column_name] == compare_target_name
        target_row_data = df_processed[target_mask].iloc[0]

    # sort metric columns if requested (same logic as before, based on VS ratios)
    if has_compare and sort_vs_by_value:
        def col_sort_key_combined(col_name):
            if target_row_data is None:
                return float('-inf')
            t_val = target_row_data[col_name]
            ratios = []
            for _, r in df_processed.iterrows():
                if r[compare_target_column_name] == compare_target_name:
                    continue
                o_val = r[col_name]
                if isinstance(t_val, (int, float)) and isinstance(o_val, (int, float)) and not np.isclose(o_val, 0.0):
                    ratios.append(t_val / o_val)
            return statistics.mean(ratios) if ratios else float('-inf')

        numeric_metric_cols = sorted(
            [c for c in metric_cols_only if col_sort_key_combined(c) > float('-inf')],
            key=col_sort_key_combined, reverse=True
        )
        non_numeric_metric_cols = [c for c in metric_cols_only if col_sort_key_combined(c) == float('-inf')]
        sorted_metric_cols = numeric_metric_cols + non_numeric_metric_cols
    else:
        sorted_metric_cols = metric_cols_only

    # ordered list of all cases: target first, then others
    if has_compare:
        other_cases = [n for n in all_case_names if n != compare_target_name]
        ordered_cases = [compare_target_name] + other_cases
    else:
        ordered_cases = all_case_names

    # build lookup: case_name -> row series
    case_data = {r[compare_target_column_name]: r for _, r in df_processed.iterrows()}

    # ---- write header row (row=1 in transposed sheet = "metric name" label row) ----
    # col layout: A=metric_name | B=Target_raw | C=Target_% | D=A_raw | E=A_% | ...
    ws_combined.cell(row=1, column=1, value="Metric")
    data_col_start = 2  # first data column
    # case_col_map: case_name -> (raw_col, pct_col)  [1-indexed]
    case_col_map = {}
    col_idx = data_col_start
    for case_name in ordered_cases:
        is_target_case = has_compare and (case_name == compare_target_name)
        # raw value column header: mark target with [TARGET] prefix
        raw_label = f"[TARGET]\n{case_name}" if is_target_case else case_name
        hdr_raw = ws_combined.cell(row=1, column=col_idx, value=raw_label)
        hdr_raw.alignment = Alignment(horizontal='left', wrap_text=True)
        raw_col = col_idx
        col_idx += 1
        if has_compare:
            # percentage column header: show "case vs. target %"
            if is_target_case:
                pct_label = f"[TARGET]\n{case_name}\n%"
            else:
                pct_label = f"{case_name}\n÷ {compare_target_name}\n%"
            hdr_pct = ws_combined.cell(row=1, column=col_idx, value=pct_label)
            hdr_pct.alignment = Alignment(horizontal='left', wrap_text=True)
            pct_col = col_idx
            col_idx += 1
        else:
            pct_col = None
        case_col_map[case_name] = (raw_col, pct_col)

    # ---- write data rows ------------------------------------------------
    # rows: header_cols first (用例/项目/场景), then sorted metric cols
    all_rows_in_order = header_cols + sorted_metric_cols

    for row_i, metric in enumerate(all_rows_in_order, start=2):
        # write metric name in col A
        name_cell = ws_combined.cell(row=row_i, column=1, value=metric)
        is_header_row = metric in header_cols

        # check if this metric should have important background
        base_metric = metric.split('\n')[0] if '\n' in metric else metric
        is_important = any(cfg in base_metric for cfg in config_data['columns_important_background'])
        if is_important:
            name_cell.fill = important_background_fill

        if not is_header_row:
            name_cell.alignment = Alignment(wrap_text=False)
            # auto-fit col A width
            char_width = sum(2 if ord(c) > 127 else 1 for c in str(metric))
            current_w = ws_combined.column_dimensions['A'].width or 15
            ws_combined.column_dimensions['A'].width = max(current_w, char_width + 2)
        else:
            name_cell.alignment = Alignment(wrap_text=True)

        # write values for each case
        raw_values_for_databar = []  # collect raw numerics for data bar rule
        pct_values_for_color = []    # collect pct numerics for heatmap

        for case_name in ordered_cases:
            raw_col, pct_col = case_col_map[case_name]
            row_series = case_data.get(case_name)
            if row_series is None:
                continue
            cell_value = row_series[metric] if metric in row_series.index else None

            # raw value cell
            raw_cell = ws_combined.cell(row=row_i, column=raw_col, value=cell_value)
            if is_important:
                raw_cell.fill = important_background_fill
            if isinstance(cell_value, (int, float)):
                raw_values_for_databar.append(cell_value)

            # percentage cell
            if has_compare and pct_col is not None:
                if is_header_row:
                    # for header rows (用例/项目/场景): show raw value in pct col too, no formatting
                    pct_cell = ws_combined.cell(row=row_i, column=pct_col, value=cell_value)
                    pct_cell.alignment = Alignment(horizontal='left', wrap_text=True)
                else:
                    pct_ratio = None
                    if case_name == compare_target_name:
                        # target vs itself = 100%
                        if isinstance(cell_value, (int, float)):
                            pct_ratio = 1.0
                    else:
                        if target_row_data is not None:
                            t_val = target_row_data[metric]
                            o_val = cell_value
                            if isinstance(t_val, (int, float)) and isinstance(o_val, (int, float)) and not np.isclose(o_val, 0.0):
                                pct_ratio = t_val / o_val

                    if pct_ratio is not None:
                        pct_cell = ws_combined.cell(row=row_i, column=pct_col, value=pct_ratio)
                        pct_cell.number_format = "0.00%"
                        pct_cell.fill = color_cell(pct_ratio, 0, 1)
                        pct_values_for_color.append(pct_ratio)
                        # add comment with raw values
                        if case_name != compare_target_name and target_row_data is not None:
                            t_val = target_row_data[metric]
                            o_val = cell_value
                            if isinstance(t_val, (int, float)) and isinstance(o_val, (int, float)):
                                try:
                                    pct_cell.comment = openpyxl.comments.Comment(
                                        f"{t_val} / {o_val}",
                                        'PerfDog Better Compare'
                                    )
                                except Exception:
                                    pass
                    else:
                        pct_cell = ws_combined.cell(row=row_i, column=pct_col, value=f"{target_row_data[metric] if target_row_data is not None else '-'} / {cell_value}")
                        pct_cell.alignment = Alignment(horizontal='left')

        # apply data bar to raw value columns for this metric row (skip header rows)
        if not is_header_row and raw_values_for_databar:
            max_v = max(raw_values_for_databar)
            avg_v = statistics.mean(raw_values_for_databar)
            raw_col_letters = [
                openpyxl.utils.get_column_letter(case_col_map[c][0])
                for c in ordered_cases
                if isinstance(case_data.get(c, {}).get(metric) if hasattr(case_data.get(c, {}), 'get') else None, (int, float))
            ]
            # apply data bar across all raw columns in this row
            raw_cols_indices = [case_col_map[c][0] for c in ordered_cases]
            first_raw = openpyxl.utils.get_column_letter(min(raw_cols_indices))
            # build range string covering only raw columns (interleaved, so apply per-column)
            for case_name in ordered_cases:
                raw_col_i = case_col_map[case_name][0]
                col_letter = openpyxl.utils.get_column_letter(raw_col_i)
                cell_ref = f"{col_letter}{row_i}"
                rule = DataBarRule(
                    start_type="num", start_value=max_v - avg_v,
                    end_type="num", end_value=max_v,
                    color="999999"
                )
                ws_combined.conditional_formatting.add(cell_ref, rule)

    # ---- set column widths ----------------------------------------------
    # col A already auto-sized above; data columns
    for case_name in ordered_cases:
        raw_col, pct_col = case_col_map[case_name]
        ws_combined.column_dimensions[openpyxl.utils.get_column_letter(raw_col)].width = 20
        if pct_col is not None:
            ws_combined.column_dimensions[openpyxl.utils.get_column_letter(pct_col)].width = 30

    # ---- header row height (row 1) --------------------------------------
    max_lines = max(
        (str(ws_combined.cell(row=1, column=c).value or '').count('\n') + 1
         for c in range(1, ws_combined.max_column + 1)),
        default=1
    )
    ws_combined.row_dimensions[1].height = max(max_lines * 18, 80)

    # ---- info cell: last column of row 1 --------------------------------
    info_col = ws_combined.max_column + 1
    cmdline = 'python ' + ' '.join(sys.argv)
    info_text = (
        f"Generated by PerfDog Better Compare\n"
        f"https://github.com/donaldwuid/perfdog_better_compare\n"
        f"\n"
        f"Command:\n"
        f"{cmdline}\n"
        f"\n"
        f"Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    )
    info_cell = ws_combined.cell(row=1, column=info_col, value=info_text)
    info_cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws_combined.column_dimensions[openpyxl.utils.get_column_letter(info_col)].width = 60

    # ---- freeze panes: freeze col A + first frozen_n data rows ----------
    ws_combined.freeze_panes = ws_combined.cell(row=frozen_n + 2, column=2)

    # ---- copy CompareSource sheets as-is --------------------------------
    base_output_xlsx = os.path.splitext(output_xlsx)[0] + "_base.xlsx"
    with pd.ExcelWriter(base_output_xlsx) as writer:
        for i, one_df in enumerate(input_df_list):
            one_df.to_excel(writer, sheet_name='CompareSource' + str(i), index=False)

    wb_src = openpyxl.load_workbook(base_output_xlsx)
    for sheet_name in wb_src.sheetnames:
        if sheet_name.startswith('CompareSource'):
            ws_source = wb_src[sheet_name]
            ws_target = wb_out.create_sheet(sheet_name)
            for row in ws_source.rows:
                for cell in row:
                    new_cell = ws_target.cell(row=cell.row, column=cell.column, value=cell.value)
                    if cell.alignment:
                        new_cell.alignment = Alignment(
                            horizontal=cell.alignment.horizontal,
                            vertical=cell.alignment.vertical,
                            wrap_text=cell.alignment.wrap_text,
                        )
            for column in ws_source.columns:
                ws_target.column_dimensions[column[0].column_letter].width = ws_source.column_dimensions[column[0].column_letter].width

    os.remove(base_output_xlsx)

    # remove default empty sheet
    if 'Sheet' in wb_out.sheetnames:
        wb_out.remove(wb_out['Sheet'])

    wb_out.save(output_xlsx)

    print(f"  [done] Output saved: {output_xlsx}")
    print("=" * 60)

def interpolate_among_3color(percentage, color0, color1, color2):
    # build piecewise linear color map
    cmap_dict = {
        'red':   [(0.0, color0[0], color0[0]), (0.5, color1[0], color1[0]), (1.0, color2[0], color2[0])],
        'green': [(0.0, color0[1], color0[1]), (0.5, color1[1], color1[1]), (1.0, color2[1], color2[1])],
        'blue':  [(0.0, color0[2], color0[2]), (0.5, color1[2], color1[2]), (1.0, color2[2], color2[2])],
    }
    cmap = mpl.colors.LinearSegmentedColormap('custom_cmap', cmap_dict)
    return cmap(percentage)

def rgb_to_hex(rgb):
    return "{:02x}{:02x}{:02x}".format(int(rgb[0] * 255), int(rgb[1] * 255), int(rgb[2] * 255))

NEUTRAL_ZONE_LOW  = 0.90  # below this: pure blue gradient
NEUTRAL_ZONE_HIGH = 1.10  # above this: orange-red gradient
COLOR_NEUTRAL = (0.75, 0.75, 0.75)  # gray at 100%

def color_cell(value, min_value, max_value):
    """
    Smooth gradient using log scale centered at 1.0 (100%):
      value << 1  → blue
      value ≈ 0.9 → blue-gray blend
      value = 1.0 → gray
      value ≈ 1.1 → gray-orange blend
      value >> 1  → red
    No hard cutoffs; the ±10% neutral zone fades naturally through gray.
    """
    color_blue   = (0, 0.44, 0.75)
    color_gray   = (0.75, 0.75, 0.75)
    color_orange = (1, 0.75, 0)
    color_red    = (1, 0, 0)

    def make_fill(rgb):
        h = rgb_to_hex(rgb)
        return PatternFill(start_color=h, end_color=h, fill_type="solid")

    # map value to [0,1] using log scale; log(1.0)=0 maps to 0.5 (gray center)
    # ±log(1.1) ≈ ±0.095 maps to the ±10% neutral boundary
    # use log(2) ≈ 0.693 as the saturation scale so that 200% → ~1.0 and 50% → ~0.0
    log_val = math.log(max(value, 1e-9))
    scale = math.log(2)  # half-saturation at 2x or 0.5x
    ratio = 0.5 + log_val / (2 * scale)
    ratio = max(0.0, min(1.0, ratio))

    # 5-stop color map: blue → blue → gray → orange → red
    # stops at [0.0, 0.35, 0.5, 0.65, 1.0]
    if ratio <= 0.35:
        t = ratio / 0.35
        rgb = interpolate_among_3color(t, color_blue, color_blue, color_blue)
    elif ratio <= 0.5:
        t = (ratio - 0.35) / (0.5 - 0.35)
        rgb = interpolate_among_3color(t, color_blue, color_gray, color_gray)
    elif ratio <= 0.65:
        t = (ratio - 0.5) / (0.65 - 0.5)
        rgb = interpolate_among_3color(t, color_gray, color_gray, color_orange)
    else:
        t = (ratio - 0.65) / (1.0 - 0.65)
        rgb = interpolate_among_3color(t, color_orange, color_orange, color_red)

    return make_fill(rgb)

def main():
    parser = argparse.ArgumentParser(description='Process PerfDog exported Excel files and generate better comparison.')
    
    # 添加位置参数，用于单个文件输入的情况
    parser.add_argument('input_file', nargs='?', help='input a single PerfDog exported xlsx file')
    
    # 原有的列表参数，现在变为可选
    parser.add_argument('-i', '--input_data_list', nargs='+', help='input multiple PerfDog exported xlsx files. multiple xlsx stats will be averaged for each project.')
    parser.add_argument('-c', '--input_perfdog_config', help='Input PerfDog config file path (json format). You can specify the provided perfdog_export_better_compare_config.json')
    parser.add_argument('-o', '--output_xlsx', help='output file path. if not specified, OUTPUT_XLSX will be INPUT_DATA_LIST[0]_better_compare.xlsx')
    
    parser.add_argument('-f', '--divided_by_framerate', action='store_true', help='false by default, whether normalized some columns by the framerate, see also perfdog_export_better_compare_config.json')
    parser.add_argument('-r', '--divided_by_resolution', action='store_true', help='false by default, whether normalized some columns by the resolution, see also perfdog_export_better_compare_config.json')

    parser.add_argument('-C', '--compare_target_column_name', help='optional, default to 用例, you may change to 项目')
    parser.add_argument('-t', '--compare_target_name', help='optional, input one target name and generate the "Target VS. Others" sheet. target name is one of values in compare_target_column_name column (default is 用例, you may change to 项目 by the --compare_target_column_name param')
    
    parser.add_argument('-s', '--show_only_columns_in_config', help='Normalized sheet show only columns in config json', default=False)
    parser.add_argument('-n', '--no_sort_vs', action='store_true', help='disable descending sort by ratio value in the VS. sheet (sort is enabled by default)')

    args = parser.parse_args()

    # 处理输入文件参数
    input_data_list = []
    if args.input_file:
        input_data_list = [args.input_file]
    elif args.input_data_list:
        input_data_list = args.input_data_list
    else:
        parser.error('No input files specified. Please provide either a single input file or use --input_data_list for multiple files.')

    process_data(input_data_list, args.input_perfdog_config, args.output_xlsx, args.divided_by_framerate, args.divided_by_resolution, args.compare_target_column_name, args.compare_target_name, args.show_only_columns_in_config, sort_vs_by_value=not args.no_sort_vs)

if __name__ == '__main__':
    main()