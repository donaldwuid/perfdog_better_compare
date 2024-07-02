import os
import ntpath
import sys
import argparse
import statistics
import re
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




def process_data(input_data_list, input_perfdog_config, output_xlsx, divided_by_framerate, divided_by_resolution, compare_target_column_name, compare_target_name, show_only_columns_in_config):

    input_data_list = [os.path.normpath(path) for path in input_data_list]
    if not output_xlsx:
        output_xlsx = os.path.splitext(input_data_list[0])[0] + "_better_compare.xlsx"

    print(f"input_data: {input_data_list}")
    print(f"output_xlsx: {output_xlsx}")
    

    print(f"divided_by_framerate: {divided_by_framerate}, divided_by_resolution: {divided_by_resolution}")
    
        
    print(f"compare_target_column_name: {compare_target_column_name}, compare_target_name: {compare_target_name}")
    print(f"show_only_columns_in_config: {show_only_columns_in_config}")

    if not input_perfdog_config:
        current_script_path = os.path.abspath(__file__)
        current_script_dir = os.path.dirname(current_script_path)
        input_perfdog_config = current_script_dir + "/perfdog_export_better_compare_config.json"
        print(f"input_perfdog_config is empty, default to: {input_perfdog_config}")
    input_perfdog_config = os.path.normpath(input_perfdog_config)
    print(f"input_perfdog_config: {input_perfdog_config}")

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
        compare_target_column_name = config_data["project_column"]
        print(f"compare_target_column_name is empty, default to: {compare_target_column_name}")

    # converting the mega or giga data cells
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

    # read the first excel file
    df = input_df_list[0].copy()

    # read and merge other excel files
    for input_df in input_df_list[1:]:
        df = pd.concat([df, input_df], ignore_index=True)

    # 计算每个项目的每列数据的平均值（数值类型）或换行拼接（非数值类型）
    def custom_agg(x):
        if pd.to_numeric(x, errors='coerce').notna().all():
            return pd.to_numeric(x, errors='coerce').mean()
        else:
            return x.astype(str).str.cat(sep='\n')

    df = df.groupby(compare_target_column_name).agg(custom_agg).reset_index()


    # ask user to input the missing important columns
    important_columns = [config_data["average_framerate_column"]]
    for column in important_columns:
        if column not in df.columns:
            print(f"missing important columns: {column}")
            for index, row in df.iterrows():
                value = input(f"Please input Test Case {row['用例']}'s {column} value: ")
                df.at[index, column] = value
                
    resolution_columns = [config_data["resolution_width_column"], config_data["resolution_height_column"]]
    for column in resolution_columns:
        if column not in df.columns:
            if divided_by_resolution:
                print(f"missing resolution columns: {column}")
            else:
                print(f"missing resolution columns: {column}, will reset to 1")
            for index, row in df.iterrows():
                if divided_by_resolution:
                    value = input(f"Please input Test Case {row['用例']}'s {column} value: ")
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


    def fix_filename(filename):
        # Windows 下的非法字符，如果在其他操作系统下，可以做适当修改
        illegal_chars = r'[<>:："/\\|?*]'
        return re.sub(illegal_chars, '_', filename)
    df_compare_target = None
    target_project_compare_sheet_name = ""
    if compare_target_column_name and compare_target_name:
        # 找到目标项目
        target_row = df_processed.loc[df_processed[compare_target_column_name] == compare_target_name]
        if not target_row.empty:
            # 创建新的DataFrame
            df_compare_target = pd.DataFrame(columns=df_processed.columns)
            target_project_compare_sheet_name = fix_filename(compare_target_name) + " VS. Others"
            comparison_rows = []
            for index, row in df_processed.iterrows():
                if row[compare_target_column_name] == compare_target_name:
                    continue

                comparison_row = {compare_target_column_name: f"{compare_target_name}\n VS. \n{row[compare_target_column_name]}"}
                for col in df_processed.columns[1:]:
                    target_value = target_row[col].values[0]
                    other_value = row[col]

                    if isinstance(target_value, (int, float)) and isinstance(other_value, (int, float)) and not np.isclose(other_value, 0.0):
                        comparison_row[col] = target_value / other_value
                    else:
                        comparison_row[col] = f"{target_value} / {other_value}"

                comparison_rows.append(comparison_row)
                df_compare_target.loc[len(df_compare_target.index)] = comparison_row
        else:
            print(f"Warning! Cannot find compare target value for {compare_target_name}! Will not generate the {compare_target_name} VS. Others heatmap!!!")


    

    def path_leaf(path):
        head, tail = ntpath.split(path)
        return tail or ntpath.basename(head)
    # save data to excels
    with pd.ExcelWriter(output_xlsx) as writer:
        df_processed.to_excel(writer, sheet_name='BarCompare', index=False)
        if df_compare_target is not None:
            df_compare_target.to_excel(writer, sheet_name=target_project_compare_sheet_name, index=False)
        for i, one_df in enumerate(input_df_list):
            one_df.to_excel(writer, sheet_name='CompareSource' + str(i), index=False)

    # drawing data bars
    wb = openpyxl.load_workbook(output_xlsx)
    ws = wb['BarCompare']
    for i, column in enumerate(ws.columns):
        column = [cell for cell in column]
        if all(isinstance(cell.value, (int, float)) for cell in column[1:]):
            values = [cell.value for cell in column[1:]]
            max_value = max(values)
            average_value = statistics.mean(values)
            rule = DataBarRule(start_type="num", start_value=max_value - average_value, end_type="num", end_value=max_value, color="999999")
            ws.conditional_formatting.add(openpyxl.utils.get_column_letter(i + 1) + '2:' + openpyxl.utils.get_column_letter(i + 1) + str(ws.max_row), rule)
    
    def interpolate_among_3color(percentage, color0, color1, color2):
        # 创建颜色映射的分段
        cmap_dict = {
            'red': [(0.0, color0[0], color0[0]),
                    (0.5, color1[0], color1[0]),
                    (1.0, color2[0], color2[0])],

            'green': [(0.0, color0[1], color0[1]),
                    (0.5, color1[1], color1[1]),
                    (1.0, color2[1], color2[1])],

            'blue': [(0.0, color0[2], color0[2]),
                    (0.5, color1[2], color1[2]),
                    (1.0, color2[2], color2[2])]
        }

        # 创建颜色映射
        cmap = mpl.colors.LinearSegmentedColormap('custom_cmap', cmap_dict)

        # 获取插值颜色
        return cmap(percentage)
    
    def rgb_to_hex(rgb):
        return "{:02x}{:02x}{:02x}".format(int(rgb[0] * 255), int(rgb[1] * 255), int(rgb[2] * 255))
    def color_cell(value, min_value, max_value):
        color0 = (0, 0.44, 0.75)  # 蓝色
        color1 = (1, 0.75, 0)  # 橙色
        color2 = (1, 0, 0)  # 红色
        
        if max_value > min_value:
            ratio = (value - min_value) / (max_value - min_value)
            interpolated_color = interpolate_among_3color(ratio, color0, color1, color2)
            hex = rgb_to_hex(interpolated_color)
            return PatternFill(start_color=hex, end_color=hex, fill_type="solid")
        else:
            return PatternFill(start_color="ffffff", end_color="ffffff", fill_type="solid")


    if df_compare_target is not None:
        ws_target_project_compare_sheet = wb[target_project_compare_sheet_name]
        for row in ws_target_project_compare_sheet.iter_rows():
            values = []

            for cell in row:
                if isinstance(cell.value, (int, float)):
                    values.append(cell.value)

            if len(values) > 0:
                # 计算 Q1, Q3 和 IQR
                q1 = np.percentile(values, 25)
                q3 = np.percentile(values, 75)
                iqr = q3 - q1

                # 计算异常值的阈值
                lower_bound = q1 - (1.5 * iqr)
                upper_bound = q3 + (1.5 * iqr)

                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.fill = color_cell(cell.value, lower_bound, upper_bound)
                        cell.number_format = "0.00%"
    
    important_background_fill = openpyxl.styles.PatternFill(start_color="f9d3e3", end_color="f9d3e3", fill_type="solid")
    for one_ws_name in wb.sheetnames:
        one_ws = wb[one_ws_name]
        # 设置第一行的单元格为自动换行
        for cell in one_ws[1]:
            cell.alignment = Alignment(wrap_text=True)

        # 调整行高以适应内容
        max_num_lines = 0
        for cell in one_ws[1]:
            num_lines = cell.value.count('\n') + 1
            max_num_lines = max(max_num_lines, num_lines)

        one_ws.row_dimensions[1].height = max_num_lines * 80
        

        for column in one_ws.columns:
            one_ws.column_dimensions[column[0].column_letter].width = 15

        for i, column in enumerate(one_ws.columns):
            column_name = one_ws.cell(row=1, column=i + 1).value
            if column_name and any(config_col in column_name for config_col in config_data['columns_important_background']):
                column[0].fill = important_background_fill



    wb.save(output_xlsx)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input_data_list', help='required, input at least one (or multiple) perfdog exported xls. multiple xls\'s stats will be averaged for each project.', nargs='+', required=True)
    parser.add_argument('--input_perfdog_config', help='Input PerfDog config file path (json format). You can specify the provided perfdog_export_better_compare_config.json')
    parser.add_argument('--output_xlsx', help='output file path. if not specified, OUTPUT_XLSX will be INPUT_DATA_LIST[0]_better_compare.xlsx')
    
    parser.add_argument('--divided_by_framerate', action='store_true', help='false by default, whether normalized some columns by the framerate, see also perfdog_export_better_compare_config.json')
    parser.add_argument('--divided_by_resolution', action='store_true', help='false by default, whether normalized some columns by the resolution, see also perfdog_export_better_compare_config.json')

    parser.add_argument('--compare_target_column_name', help='optional, optional, default to "项目", you may change to "用例"')
    parser.add_argument('--compare_target_name', help='optional, input one target name and generate the "Target VS. Others" sheet. target name is one of values in compare_target_column_name column (default is "项目", you may change to "用例" by the --compare_target_column_name param')
    
    parser.add_argument('--show_only_columns_in_config', help='Normalized sheet show only columns in config json', default=True)

    args = parser.parse_args()

    process_data(args.input_data_list, args.input_perfdog_config, args.output_xlsx, args.divided_by_framerate, args.divided_by_resolution, args.compare_target_column_name, args.compare_target_name, args.show_only_columns_in_config)

if __name__ == '__main__':
    main()