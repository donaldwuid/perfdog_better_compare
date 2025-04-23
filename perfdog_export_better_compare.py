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
        compare_target_column_name = config_data["test_case_column"]
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
        filename = re.sub(illegal_chars, '_', filename)
        # 如果添加" VS. Others"后会超过31个字符，则截断原始名称
        max_name_length = 31 - len(" VS. Others")
        if len(filename) > max_name_length:
            filename = filename[:max_name_length]
        return filename
    df_compare_target = None
    target_project_compare_sheet_name = ""
    if compare_target_column_name and compare_target_name:
        # 找到目标项目
        target_row = df_processed.loc[df_processed[compare_target_column_name] == compare_target_name]
        if not target_row.empty:
            # 创建新的DataFrame
            df_compare_target = pd.DataFrame(columns=df_processed.columns)
            # raw_values_dict用于存储每个单元格的原始值
            raw_values_dict = {}
            target_project_compare_sheet_name = fix_filename(compare_target_name) + " VS. Others"
            comparison_rows = []
            
            for index, row in df_processed.iterrows():
                if row[compare_target_column_name] == compare_target_name:
                    continue

                comparison_row = {compare_target_column_name: f"{compare_target_name}\n VS. \n{row[compare_target_column_name]}"}
                row_raw_values = {}
                
                for col in df_processed.columns[1:]:
                    target_value = target_row[col].values[0]
                    other_value = row[col]

                    if isinstance(target_value, (int, float)) and isinstance(other_value, (int, float)) and not np.isclose(other_value, 0.0):
                        comparison_row[col] = target_value / other_value
                        row_raw_values[col] = f"{target_value} / {other_value}"  # 存储原始值
                    else:
                        comparison_row[col] = f"{target_value} / {other_value}"
                        row_raw_values[col] = f"{target_value} / {other_value}"

                comparison_rows.append(comparison_row)
                raw_values_dict[len(df_compare_target)] = row_raw_values
                df_compare_target.loc[len(df_compare_target.index)] = comparison_row
            
            df_compare_target.raw_values_dict = raw_values_dict
        else:
            print(f"Warning! Cannot find compare target value for {compare_target_name}! Will not generate the {compare_target_name} VS. Others heatmap!!!")

    

    def path_leaf(path):
        head, tail = ntpath.split(path)
        return tail or ntpath.basename(head)
        
    # 生成临时的标准格式文件
    base_output_xlsx = os.path.splitext(output_xlsx)[0] + "_base.xlsx"
    
    # 保存标准格式数据到Excel（供CompareSource工作表使用）
    with pd.ExcelWriter(base_output_xlsx) as writer:
        df_processed.to_excel(writer, sheet_name='BarCompare', index=False)
        if df_compare_target is not None:
            df_compare_target.to_excel(writer, sheet_name=target_project_compare_sheet_name, index=False)
        # CompareSource表格始终保持原始状态
        for i, one_df in enumerate(input_df_list):
            one_df.to_excel(writer, sheet_name='CompareSource' + str(i), index=False)

    # 处理标准格式的Excel样式（供CompareSource工作表使用）
    wb = openpyxl.load_workbook(base_output_xlsx)
    
    # 处理BarCompare工作表的条件格式
    ws = wb['BarCompare']
    for i, column in enumerate(ws.columns):
        column = [cell for cell in column]
        if all(isinstance(cell.value, (int, float)) for cell in column[1:]):
            values = [cell.value for cell in column[1:]]
            if values:  # 确保有数值
                max_value = max(values)
                average_value = statistics.mean(values)
                rule = DataBarRule(start_type="num", start_value=max_value - average_value, end_type="num", end_value=max_value, color="999999")
                ws.conditional_formatting.add(openpyxl.utils.get_column_letter(i + 1) + '2:' + openpyxl.utils.get_column_letter(i + 1) + str(ws.max_row), rule)

    # 处理对比表格的样式和注释
    if df_compare_target is not None:
        ws_target_project_compare_sheet = wb[target_project_compare_sheet_name]
        raw_values_dict = df_compare_target.raw_values_dict

        for i, row in enumerate(ws_target_project_compare_sheet.iter_rows(), 1):
            values = []
            for j, cell in enumerate(row, 1):
                if isinstance(cell.value, (int, float)):
                    values.append(cell.value)
                    # 添加原始值注释
                    try:
                        row_idx = i - 2  # 减2是因为Excel的标题行占一行，且字典从0开始
                        if row_idx >= 0 and j < len(df_compare_target.columns):
                            col_name = df_compare_target.columns[j]
                            if row_idx in raw_values_dict and col_name in raw_values_dict[row_idx]:
                                raw_value = raw_values_dict[row_idx][col_name]
                                if raw_value:
                                    cell.comment = openpyxl.comments.Comment(
                                        raw_value,
                                        'PerfDog Better Compare'
                                    )
                    except Exception as e:
                        print(f"Warning: Failed to add comment for cell at row {i}, column {j}: {str(e)}")
                        continue

            if len(values) > 0:
                q1 = np.percentile(values, 25)
                q3 = np.percentile(values, 75)
                iqr = q3 - q1
                lower_bound = q1 - (1.5 * iqr)
                upper_bound = q3 + (1.5 * iqr)

                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.fill = color_cell(cell.value, lower_bound, upper_bound)
                        cell.number_format = "0.00%"

    # 设置标准格式的单元格样式
    important_background_fill = openpyxl.styles.PatternFill(start_color="f9d3e3", end_color="f9d3e3", fill_type="solid")
    for one_ws_name in wb.sheetnames:
        if one_ws_name.startswith('CompareSource'):
            continue
            
        one_ws = wb[one_ws_name]
        # 设置第一行的单元格为自动换行
        for cell in one_ws[1]:
            cell.alignment = Alignment(wrap_text=True)
            if cell.value is not None:
                if any(config_col in str(cell.value) for config_col in config_data['columns_important_background']):
                    cell.fill = important_background_fill

        # 调整行高以适应内容
        max_lines = 0
        for cell in one_ws[1]:
            if cell.value is not None:
                lines = str(cell.value).count('\n') + 1
                max_lines = max(max_lines, lines)
        
        if max_lines > 0:
            one_ws.row_dimensions[1].height = max_lines * 80
        else:
            one_ws.row_dimensions[1].height = 80

        # 设置默认列宽
        for column in one_ws.columns:
            one_ws.column_dimensions[column[0].column_letter].width = 15

    wb.save(base_output_xlsx)

    # 创建转置格式的Excel
    wb_transposed = openpyxl.Workbook()
    wb = openpyxl.load_workbook(base_output_xlsx)
    
    # 处理需要转置的工作表
    sheets_to_transpose = ['BarCompare']
    if df_compare_target is not None:
        sheets_to_transpose.append(target_project_compare_sheet_name)
        
    for sheet_name in sheets_to_transpose:
        if sheet_name in wb.sheetnames:
            ws_source = wb[sheet_name]
            
            # 如果转置的工作表已存在，则删除它
            if sheet_name in wb_transposed.sheetnames:
                wb_transposed.remove(wb_transposed[sheet_name])
            
            # 创建新的工作表
            ws_target = wb_transposed.create_sheet(sheet_name)
            
            # 读取原始数据并转置
            data = [[cell.value for cell in row] for row in ws_source.rows]
            transposed_data = list(zip(*data))
            
            # 写入转置后的数据并处理格式
            for i, row in enumerate(transposed_data, 1):
                values = []  # 用于存储当前行的数值，以计算数据条和热图颜色
                for j, value in enumerate(row, 1):
                    cell = ws_target.cell(row=i, column=j, value=value)
                    
                    # 设置样式
                    if i == 1:
                        # 第一行（标题）的样式
                        cell.alignment = Alignment(wrap_text=True)
                        if value is not None and any(config_col in str(value) for config_col in config_data['columns_important_background']):
                            cell.fill = important_background_fill
                    else:
                        # 对于非标题行的数值单元格
                        if isinstance(value, (int, float)):
                            values.append(value)
                            # 如果是对比表格，设置百分比格式和注释
                            if sheet_name == target_project_compare_sheet_name:
                                cell.number_format = "0.00%"
                                # 获取原始值注释
                                try:
                                    row_idx = i - 2  # 减2是因为Excel的标题行占一行，且字典从0开始
                                    if j < len(df_compare_target.columns):
                                        col_name = df_compare_target.columns[j]
                                        if row_idx in raw_values_dict and col_name in raw_values_dict[row_idx]:
                                            raw_value = raw_values_dict[row_idx][col_name]
                                            if raw_value:
                                                cell.comment = openpyxl.comments.Comment(
                                                    raw_value,
                                                    'PerfDog Better Compare'
                                                )
                                except Exception as e:
                                    print(f"Warning: Failed to add comment for cell at row {i}, column {j}: {str(e)}")
                
                # 根据工作表类型设置格式
                if i > 1:  # 跳过标题行
                    if sheet_name == 'BarCompare' and values:
                        # 为 BarCompare 添加数据条
                        max_value = max(values)
                        average_value = statistics.mean(values)
                        rule = DataBarRule(
                            start_type="num",
                            start_value=max_value - average_value,
                            end_type="num",
                            end_value=max_value,
                            color="999999"
                        )
                        ws_target.conditional_formatting.add(
                            f"B{i}:{openpyxl.utils.get_column_letter(ws_target.max_column)}{i}",
                            rule
                        )
                    elif sheet_name == target_project_compare_sheet_name and values:
                        # 为对比表格添加热图颜色
                        q1 = np.percentile(values, 25)
                        q3 = np.percentile(values, 75)
                        iqr = q3 - q1
                        lower_bound = q1 - (1.5 * iqr)
                        upper_bound = q3 + (1.5 * iqr)
                        
                        for j, value in enumerate(row, 1):
                            if isinstance(value, (int, float)):
                                cell = ws_target.cell(row=i, column=j)
                                cell.fill = color_cell(value, lower_bound, upper_bound)
                
                # 特殊处理第一列
                if i > 1:  # 跳过标题行
                    first_cell = ws_target.cell(row=i, column=1)
                    if first_cell.value is not None:
                        # 只对项目和用例列允许自动换行
                        if str(first_cell.value) in [config_data["project_column"], config_data["test_case_column"]]:
                            first_cell.alignment = Alignment(wrap_text=True)
                            if '\n' in str(first_cell.value):
                                ws_target.row_dimensions[i].height = (str(first_cell.value).count('\n') + 1) * 80
                        else:
                            first_cell.alignment = Alignment(wrap_text=False)
                            # 根据内容调整列宽
                            current_width = ws_target.column_dimensions['A'].width if 'A' in ws_target.column_dimensions else 15
                            ws_target.column_dimensions['A'].width = max(current_width, len(str(first_cell.value)) * 2)
            
            # 设置其他列的固定宽度
            for col in range(2, ws_target.max_column + 1):
                ws_target.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 15
    
    # 复制 CompareSource 工作表（不转置）
    for sheet_name in wb.sheetnames:
        if sheet_name.startswith('CompareSource'):
            ws_source = wb[sheet_name]
            ws_target = wb_transposed.create_sheet(sheet_name)
            
            for row in ws_source.rows:
                for cell in row:
                    new_cell = ws_target.cell(row=cell.row, column=cell.column, value=cell.value)
                    # 复制单元格格式
                    if cell.alignment:
                        new_cell.alignment = Alignment(
                            horizontal=cell.alignment.horizontal,
                            vertical=cell.alignment.vertical,
                            text_rotation=cell.alignment.text_rotation,
                            wrap_text=cell.alignment.wrap_text,
                            shrink_to_fit=cell.alignment.shrink_to_fit,
                            indent=cell.alignment.indent
                        )
                    if cell.fill:
                        new_cell.fill = PatternFill(
                            start_color=cell.fill.start_color.rgb,
                            end_color=cell.fill.end_color.rgb,
                            fill_type=cell.fill.fill_type
                        )
                    if cell.comment:
                        new_cell.comment = openpyxl.comments.Comment(
                            text=cell.comment.text,
                            author=cell.comment.author
                        )
            
            # 复制列宽
            for column in ws_source.columns:
                ws_target.column_dimensions[column[0].column_letter].width = ws_source.column_dimensions[column[0].column_letter].width
            
            # 复制行高
            for row in ws_source.row_dimensions:
                if row in ws_source.row_dimensions:
                    ws_target.row_dimensions[row] = ws_source.row_dimensions[row]
    
    # 删除默认创建的空工作表
    if 'Sheet' in wb_transposed.sheetnames:
        wb_transposed.remove(wb_transposed['Sheet'])
    
    # 保存最终的Excel文件
    wb_transposed.save(output_xlsx)
    
    # 删除临时文件
    os.remove(base_output_xlsx)

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

def main():
    parser = argparse.ArgumentParser(description='Process PerfDog exported Excel files and generate better comparison.')
    
    # 添加位置参数，用于单个文件输入的情况
    parser.add_argument('input_file', nargs='?', help='input a single PerfDog exported xlsx file')
    
    # 原有的列表参数，现在变为可选
    parser.add_argument('--input_data_list', nargs='+', help='input multiple PerfDog exported xlsx files. multiple xlsx stats will be averaged for each project.')
    parser.add_argument('--input_perfdog_config', help='Input PerfDog config file path (json format). You can specify the provided perfdog_export_better_compare_config.json')
    parser.add_argument('--output_xlsx', help='output file path. if not specified, OUTPUT_XLSX will be INPUT_DATA_LIST[0]_better_compare.xlsx')
    
    parser.add_argument('--divided_by_framerate', action='store_true', help='false by default, whether normalized some columns by the framerate, see also perfdog_export_better_compare_config.json')
    parser.add_argument('--divided_by_resolution', action='store_true', help='false by default, whether normalized some columns by the resolution, see also perfdog_export_better_compare_config.json')

    parser.add_argument('--compare_target_column_name', help='optional, default to 用例, you may change to 项目')
    parser.add_argument('--compare_target_name', help='optional, input one target name and generate the "Target VS. Others" sheet. target name is one of values in compare_target_column_name column (default is 用例, you may change to 项目 by the --compare_target_column_name param')
    
    parser.add_argument('--show_only_columns_in_config', help='Normalized sheet show only columns in config json', default=False)

    args = parser.parse_args()

    # 处理输入文件参数
    input_data_list = []
    if args.input_file:
        input_data_list = [args.input_file]
    elif args.input_data_list:
        input_data_list = args.input_data_list
    else:
        parser.error('No input files specified. Please provide either a single input file or use --input_data_list for multiple files.')

    process_data(input_data_list, args.input_perfdog_config, args.output_xlsx, args.divided_by_framerate, args.divided_by_resolution, args.compare_target_column_name, args.compare_target_name, args.show_only_columns_in_config)

if __name__ == '__main__':
    main()