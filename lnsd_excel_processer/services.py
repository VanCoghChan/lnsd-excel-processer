import pandas as pd
import os
import uuid
from django.conf import settings
import numpy as np

TARGET_COLUMNS = ['组织', '资源集', '地域']
EXCLUDE_SHEETS = ['项目']
TEMP_UPLOAD_DIR = os.path.join(settings.MEDIA_ROOT, 'temp_uploads')

def get_excel_metadata(excel_file):
    """保存上传的Excel文件，提取其元数据（工作表、列名、可用的非空列按工作表分组）。"""
    try:
        os.makedirs(TEMP_UPLOAD_DIR, exist_ok=True)
        temp_file_id = f"temp_{uuid.uuid4().hex[:12]}.xlsx"
        temp_file_path = os.path.join(TEMP_UPLOAD_DIR, temp_file_id)
        
        # 保存上传的文件到临时位置
        with open(temp_file_path, 'wb+') as destination:
            for chunk in excel_file.chunks():
                destination.write(chunk)

        xls = pd.ExcelFile(temp_file_path)
        sheet_names = [name for name in xls.sheet_names if name not in EXCLUDE_SHEETS]
        
        sheets_metadata = []
        all_columns_set = set()
        # 存储所有带前缀的字段 (用于兼容性或未来使用)
        prefixed_columns_set = set() 
        # 存储按工作表分组的、非全空的、带前缀的字段
        # 格式: {'Sheet1': ['Sheet1-ColA', 'Sheet1-ColB'], 'Sheet2': ['Sheet2-ColC']}
        grouped_available_columns = {}
        
        for sheet_name in sheet_names:
            try:
                # 读取少量数据以推断类型和检查空值，同时获取列名
                try:
                    df_sample = pd.read_excel(xls, sheet_name=sheet_name, nrows=5) 
                    df_headers = pd.read_excel(xls, sheet_name=sheet_name, nrows=0)
                    columns = df_headers.columns.tolist()
                except Exception as read_err:
                    # 如果读取少量数据失败（例如空表），只读表头
                    print(f"警告：无法读取工作表 {sheet_name} 的样本行，仅读取表头。错误: {read_err}")
                    df_sample = pd.read_excel(xls, sheet_name=sheet_name, nrows=0) 
                    columns = df_sample.columns.tolist()
                
                sheets_metadata.append({'name': sheet_name, 'columns': columns})
                all_columns_set.update(columns)
                grouped_available_columns[sheet_name] = [] # 初始化当前工作表的列表

                # 添加所有带前缀的字段
                prefixed_columns = [f"{sheet_name}-{col}" for col in columns]
                prefixed_columns_set.update(prefixed_columns)

                # 识别非全空的字段，并添加到分组字典中
                for col in columns:
                    prefixed_col = f"{sheet_name}-{col}"
                    # 检查 df_sample 中列是否存在且至少有一个非空值
                    if col in df_sample.columns and df_sample[col].notna().any():
                        grouped_available_columns[sheet_name].append(prefixed_col)
                        
            except Exception as e:
                print(f"处理工作表 {sheet_name} 时出错: {e}")
                sheets_metadata.append({'name': sheet_name, 'columns': [], 'error': str(e)})
                if sheet_name not in grouped_available_columns: #确保即使出错也有空列表
                     grouped_available_columns[sheet_name] = []


        if not sheets_metadata:
            os.remove(temp_file_path) # 如果未找到有效工作表，清理临时文件
            return {'success': False, 'message': '未在Excel文件中找到有效的工作表。'}

        # 清理没有可用列的工作表条目 (如果需要)
        # grouped_available_columns = {k: v for k, v in grouped_available_columns.items() if v}

        # 排序最终列表/字典值
        all_available_columns_sorted = sorted(list(all_columns_set))
        prefixed_available_columns_sorted = sorted(list(prefixed_columns_set))
        # 对字典的值（列列表）进行排序
        for sheet in grouped_available_columns:
            grouped_available_columns[sheet].sort()


        return {
            'success': True,
            'data': {
                'temp_file_id': temp_file_id,
                'sheets': sheets_metadata,
                'all_available_columns': all_available_columns_sorted, # 保留原始字段列表（向后兼容）
                'prefixed_available_columns': prefixed_available_columns_sorted, # 所有带前缀的字段列表
                # 'numeric_prefixed_columns': numeric_prefixed_columns_sorted # 不再使用
                'grouped_available_columns': grouped_available_columns # 新增：按工作表分组的可用（非空）字段
            }
        }
    except Exception as e:
        error_message = f"提取Excel元数据时出错: {str(e)}"
        print(error_message)
        # 如果创建了临时文件但发生错误，尝试清理
        if 'temp_file_path' in locals() and os.path.exists(temp_file_path):
            try:
                os.remove(temp_file_path)
            except Exception as cleanup_error:
                print(f"清理临时文件 {temp_file_path} 时出错: {cleanup_error}")
        return {'success': False, 'message': error_message}

def perform_final_analysis(temp_file_id, selected_sheet_names, additional_stat_configs):
    """根据选择的表、字段和聚合方式处理Excel，进行分组统计和额外统计。"""
    temp_file_path = os.path.join(TEMP_UPLOAD_DIR, temp_file_id)
    if not os.path.exists(temp_file_path):
        return {'success': False, 'message': '临时文件不存在或已过期。'}

    try:
        xls = pd.ExcelFile(temp_file_path)
        
        all_sheet_count_results = []
        # 存储聚合结果。键: 带聚合方式后缀的带前缀列名, 值: DataFrame
        aggregated_stat_results = {}

        for sheet_name in selected_sheet_names:
            if sheet_name not in xls.sheet_names or sheet_name in EXCLUDE_SHEETS:
                continue
            
            df = pd.read_excel(xls, sheet_name=sheet_name)
            sheet_df_columns = df.columns.tolist()

            # 1. 处理工作表计数 (原始逻辑)
            sheet_count_df = df.groupby(TARGET_COLUMNS).size().reset_index(name=sheet_name)
            all_sheet_count_results.append(sheet_count_df)

            # 2. 根据配置处理额外的统计列
            if additional_stat_configs:
                relevant_configs = []
                agg_dict = {}
                col_mapping = {}

                for config in additional_stat_configs:
                    prefixed_col_name = config['column']
                    agg_method = config['agg']

                    if prefixed_col_name.startswith(f"{sheet_name}-"):
                        original_col_name = prefixed_col_name.split('-', 1)[1]

                        # 检查原始列是否存在于DataFrame中
                        if original_col_name in sheet_df_columns:
                            # 映射Pandas聚合函数名称, 加入 'count'
                            valid_agg_methods = {'sum': 'sum', 'mean': 'mean', 'max': 'max', 'min': 'min', 'count': 'count'}

                            if agg_method in valid_agg_methods:
                                # 对 sum, mean, max, min 额外检查是否为数值型
                                is_numeric_required = agg_method in ['sum', 'mean', 'max', 'min']
                                is_numeric = pd.api.types.is_numeric_dtype(df[original_col_name])

                                # 检查：count 不需要数值型；其他需要数值型
                                if agg_method == 'count' or (is_numeric_required and is_numeric):
                                    pandas_agg_func = valid_agg_methods[agg_method]
                                    final_col_name = f"{prefixed_col_name}_{agg_method}"

                                    if original_col_name not in agg_dict:
                                        agg_dict[original_col_name] = pandas_agg_func
                                        col_mapping[original_col_name] = final_col_name
                                        relevant_configs.append({'original': original_col_name, 'final': final_col_name})
                                elif is_numeric_required and not is_numeric:
                                    print(f"警告：工作表 '{sheet_name}' 中的列 '{original_col_name}' 不是数值型，无法执行 '{agg_method}' 操作，已跳过。")


                if agg_dict:
                    # 使用 .agg() 执行聚合
                    current_sheet_agg_df = df.groupby(TARGET_COLUMNS, as_index=False).agg(agg_dict)
                    
                    # 将列重命名为带前缀和聚合后缀的最终名称
                    rename_map = {orig_col: col_mapping[orig_col] for orig_col in agg_dict.keys()}
                    current_sheet_agg_df.rename(columns=rename_map, inplace=True)
                    
                    # 将结果添加到全局聚合字典中
                    for config_info in relevant_configs:
                        final_col = config_info['final']
                        # 提取此特定聚合列的DataFrame
                        single_agg_df = current_sheet_agg_df[TARGET_COLUMNS + [final_col]].copy()
                        aggregated_stat_results[final_col] = single_agg_df
        
        if not all_sheet_count_results and not aggregated_stat_results:
            if os.path.exists(temp_file_path): os.remove(temp_file_path)
            return {'success': False, 'message': '没有选择任何有效的工作表或没有可处理的数据。'}

        # 首先合并所有工作表计数结果
        final_df = None
        if all_sheet_count_results:
            final_df = all_sheet_count_results[0]
            for i in range(1, len(all_sheet_count_results)):
                final_df = pd.merge(final_df, all_sheet_count_results[i], on=TARGET_COLUMNS, how='outer')
            # 计数列中的NaN表示该分组不存在，保留NaN，不填充0

        # 合并聚合统计结果
        # 对键进行排序以确保最终输出中列的顺序一致
        sorted_stat_cols = sorted(aggregated_stat_results.keys())
        for stat_col_name in sorted_stat_cols:
            stat_df = aggregated_stat_results[stat_col_name]
            if final_df is None:
                final_df = stat_df
            else:
                # 使用外连接合并，以包含两侧的所有分组
                final_df = pd.merge(final_df, stat_df, on=TARGET_COLUMNS, how='outer')
            # 聚合列中的NaN填充为0
            final_df[stat_col_name] = final_df[stat_col_name].fillna(0)

        if final_df is None or final_df.empty:
            if os.path.exists(temp_file_path): os.remove(temp_file_path)
            return {'success': False, 'message': '最终处理结果为空。'}
        
        # 确保 TARGET_COLUMNS 在最前面，然后是工作表计数列，最后是统计列
        if all_sheet_count_results:
            count_col_names = sorted([res.columns[-1] for res in all_sheet_count_results if res.columns[-1] not in TARGET_COLUMNS])
        else:
            count_col_names = []    
        
        # 统计列已通过 sorted_stat_cols 排序
        stat_col_names_sorted = sorted_stat_cols
        
        ordered_columns = TARGET_COLUMNS + count_col_names + stat_col_names_sorted
        # 过滤掉可能因某些部分为空而未出现在 final_df 中的列
        ordered_columns = [col for col in ordered_columns if col in final_df.columns]
        final_df = final_df[ordered_columns]

        output_filename = f"analysis_result_{uuid.uuid4().hex[:8]}.xlsx"
        output_file_path = os.path.join(settings.MEDIA_ROOT, output_filename)
        os.makedirs(settings.MEDIA_ROOT, exist_ok=True)
        final_df.to_excel(output_file_path, index=False)
        
        # 清理临时文件
        if os.path.exists(temp_file_path): os.remove(temp_file_path)

        return {
            'success': True,
            'data': {
                'rows': len(final_df),
                'columns': len(final_df.columns),
                'column_names': final_df.columns.tolist(),
                'download_filename': output_filename
            }
        }
    except Exception as e:
        error_message = f"执行最终分析时出错: {str(e)}"
        print(error_message)
        # 如果在错误处理期间临时文件存在，尝试清理
        if os.path.exists(temp_file_path):
            try: os.remove(temp_file_path)
            except Exception as cleanup_error: print(f"在错误处理期间清理临时文件 {temp_file_path} 时出错: {cleanup_error}")
        return {'success': False, 'message': error_message} 