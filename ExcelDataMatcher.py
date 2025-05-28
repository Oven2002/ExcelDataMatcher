import pandas as pd
import xlrd
import xlwt
from xlutils.copy import copy

def load_and_prepare_data(file_a_path, file_b_path):
    """
    从指定的a表和b表路径加载数据，并进行初步准备。
    假设第一行是表头，数据从第二行开始。
    a表“告警开始时间”列的值会移除末尾的 ".0"。
    """
    try:
        # 读取 a 表格
        df_a = pd.read_excel(file_a_path,
                             engine='xlrd',
                             usecols=['告警开始时间', '内容']) # 默认 header=0, pandas 使用第一行做表头
        df_a.columns = ['a_alarm_start_time', 'a_content']
        # 清理字符串列可能存在的前后空格
        df_a['a_alarm_start_time'] = df_a['a_alarm_start_time'].astype(str).str.strip()
        # 移除 a_alarm_start_time 列值末尾的 ".0"
        df_a['a_alarm_start_time'] = df_a['a_alarm_start_time'].str.replace(r'\.0$', '', regex=True)
        df_a['a_content'] = df_a['a_content'].astype(str).str.strip()


        # 读取 b 表格
        df_b = pd.read_excel(file_b_path,
                             engine='xlrd',
                             usecols=['告警开始时间', '告警描述']) # 默认 header=0
        df_b.columns = ['b_alarm_start_time', 'b_alarm_description']
        # 清理字符串列可能存在的前后空格
        df_b['b_alarm_start_time'] = df_b['b_alarm_start_time'].astype(str).str.strip()
        df_b['b_alarm_description'] = df_b['b_alarm_description'].astype(str).str.strip()
        
        return df_a, df_b
    except FileNotFoundError:
        print(f"错误：一个或多个Excel文件未找到。请检查路径：\n{file_a_path}\n{file_b_path}")
        return None, None
    except ValueError as e:
        print(f"错误：读取Excel文件时，列名可能不匹配或文件格式问题。确保指定的列名存在于文件的第一行。错误详情: {e}")
        return None, None
    except Exception as e:
        print(f"加载和准备数据时发生未知错误: {e}")
        return None, None


def find_matching_rows(df_a, df_b):
    """
    比较 df_a 和 df_b, 找出 df_a 中匹配的行索引 (相对于 df_a 的0基索引)。
    """
    if df_a is None or df_b is None:
        return # 返回空列表

    # 使用 merge 方法查找匹配项
    df_a_indexed = df_a.reset_index() # 保留原始索引以便高亮, 'index' 列保存了df_a的原始索引
    merged_df = pd.merge(df_a_indexed, df_b,
                         left_on=['a_alarm_start_time', 'a_content'],
                         right_on=['b_alarm_start_time', 'b_alarm_description'],
                         how='inner') # 'inner' join 只保留双方都匹配的行
    
    matching_indices_in_a_df = merged_df['index'].unique().tolist()
    return matching_indices_in_a_df


def apply_highlight_and_save(original_file_path, output_file_path, matching_indices_df, df_a_row_count):
    """
    打开原始 a 表格，对匹配的行应用红色背景高亮，并保存到新文件。
    matching_indices_df 是 pandas DataFrame 的0基索引列表 (相对于 df_a)。
    df_a_row_count 是 df_a (即跳过表头后的数据行数) 的行数。
    """
    if not matching_indices_df and df_a_row_count == 0 : 
        print("a表格中没有数据，未生成高亮文件。")
        return
    if not matching_indices_df and df_a_row_count > 0:
        print("a表格中没有找到需要高亮的匹配行。将仅复制原始文件（如果需要此行为，请取消注释 shutil 相关代码）。")
        # import shutil
        # shutil.copy(original_file_path, output_file_path)
        # print(f"已将a表格复制到：{output_file_path} (无高亮)")
        return


    try:
        rb = xlrd.open_workbook(original_file_path, formatting_info=True) # [5, 21]
        wb_copy = copy(rb) # 创建一个可写的副本 [5]
        sheet_to_write = wb_copy.get_sheet(0) # 假设在第一个sheet操作

        # 定义红色背景样式
        red_style = xlwt.easyxf('pattern: pattern solid, fore_colour red;') # [10, 22, 23]

        original_sheet_for_read = rb.sheet_by_index(0) # 用于读取原始单元格值
        num_cols_original = original_sheet_for_read.ncols # 获取原始表格的总列数 [25, 26]

        # 遍历 df_a 的原始行 (这些行在Excel中是从第二行开始的)
        # matching_indices_df 包含的是 df_a 的0基索引
        for r_idx_df in matching_indices_df: # 只遍历需要高亮的行
            # actual_row_in_sheet_model 指的是在 xlrd/xlwt 工作表中的0基行索引
            # Excel文件中的实际行号是 r_idx_df + 2 (1-based)
            # xlrd/xlwt 读取的包含表头的sheet中的行索引是 r_idx_df + 1 (0-based)
            actual_row_in_sheet_model = r_idx_df + 1 

            # 确保行索引在原始工作表范围内 (数据行部分)
            if actual_row_in_sheet_model < original_sheet_for_read.nrows: # [25]
                for c_idx in range(num_cols_original): # 遍历该行的所有列
                    try:
                        original_value = original_sheet_for_read.cell_value(actual_row_in_sheet_model, c_idx) # [25]
                        sheet_to_write.write(actual_row_in_sheet_model, c_idx, original_value, red_style) # [10, 11]
                    except IndexError:
                        print(f"警告：尝试访问单元格 ({actual_row_in_sheet_model}, {c_idx}) 超出范围。")
                        continue # 跳过此单元格
            else:
                print(f"警告：匹配索引 {r_idx_df} (对应工作表行 {actual_row_in_sheet_model}) 超出原始工作表行数 {original_sheet_for_read.nrows}。")

        wb_copy.save(output_file_path) # [10]
        print(f"处理完成，高亮后的文件已保存至：{output_file_path}")

    except FileNotFoundError:
        print(f"错误：原始文件 '{original_file_path}' 未找到，无法进行高亮操作。")
    except Exception as e:
        print(f"在应用高亮和保存文件时发生错误：{e}")


if __name__ == "__main__":
    # 文件路径定义 (用户需根据实际情况修改)
    file_a_path_input = "D:\\Temp\\work\\新建文件夹\\调阅情况202505281813.xls"
    file_b_path_input = "D:\\Temp\\work\\新建文件夹\\运维检修.xls"
    # 输出文件路径 (建议使用新文件名以避免覆盖原始文件)
    file_a_output_path = "D:\\Temp\\work\\新建文件夹\\调阅情况202505281813_highlighted.xls"

    print("开始处理Excel文件...")
    df_a_data, df_b_data = load_and_prepare_data(file_a_path_input, file_b_path_input)

    if df_a_data is not None and df_b_data is not None:
        print(f"a表格加载了 {len(df_a_data)} 行数据 (不含表头)。")
        print(f"b表格加载了 {len(df_b_data)} 行数据 (不含表头)。")
        
        if df_a_data.empty:
            print("a表格中没有数据 (除了表头)。未生成高亮文件。")
        else:
            matching_rows_indices = find_matching_rows(df_a_data, df_b_data)
            print(f"在 a 表格中找到 {len(matching_rows_indices)} 行匹配项。")

            if matching_rows_indices: 
                apply_highlight_and_save(file_a_path_input, file_a_output_path, matching_rows_indices, len(df_a_data))
            else:
                 print("a表格中没有找到需要高亮的匹配行。未生成高亮文件。")
                 # 如果希望即使没有匹配项也复制一份a表，可以取消下面这几行的注释
                 # import shutil
                 # try:
                 #     shutil.copy(file_a_path_input, file_a_output_path)
                 #     print(f"已将a表格复制到：{output_file_path} (无高亮)")
                 # except Exception as e_copy:
                 #     print(f"复制原始a表时出错: {e_copy}")
    else:
        print("数据加载失败，脚本终止。")