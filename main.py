# main.py
import os
import openpyxl
from copy import copy
from config import CURRENT_DIR
from utils import find_xlsx_files_with_dates
from data_loader import load_sheet_data, calculate_weekly_data, analyze_sku_performance, calculate_acos_analysis,analyze_high_conversion_products, analyze_high_conversion_low_acos_products
from excel_writer import ExcelWriter
from sku_manager import manage_new_arrival_skus

def copy_sheet_contents(source_ws, target_ws):
    """将 source_ws 的内容、样式和常用布局复制到 target_ws。"""
    for row in source_ws.iter_rows():
        for source_cell in row:
            target_cell = target_ws.cell(row=source_cell.row, column=source_cell.column)
            target_cell.value = source_cell.value

            if source_cell.has_style:
                target_cell.font = copy(source_cell.font)
                target_cell.fill = copy(source_cell.fill)
                target_cell.border = copy(source_cell.border)
                target_cell.alignment = copy(source_cell.alignment)
                target_cell.protection = copy(source_cell.protection)
                target_cell.number_format = source_cell.number_format

            if source_cell.comment:
                target_cell.comment = copy(source_cell.comment)
            if source_cell.hyperlink:
                target_cell.hyperlink = copy(source_cell.hyperlink)

    for merged_range in source_ws.merged_cells.ranges:
        target_ws.merge_cells(str(merged_range))

    for key, dim in source_ws.column_dimensions.items():
        target_ws.column_dimensions[key] = copy(dim)
    for key, dim in source_ws.row_dimensions.items():
        target_ws.row_dimensions[key] = copy(dim)

    target_ws.freeze_panes = source_ws.freeze_panes
    target_ws.sheet_format = copy(source_ws.sheet_format)
    target_ws.sheet_properties = copy(source_ws.sheet_properties)
    target_ws.page_margins = copy(source_ws.page_margins)
    target_ws.page_setup = copy(source_ws.page_setup)
    target_ws.print_options = copy(source_ws.print_options)
    target_ws.auto_filter.ref = source_ws.auto_filter.ref

def overwrite_summary_sheet_from_draft(draft_file, target_file, sheet_name="总计"):
    """用 draft.xlsx 中的总计 sheet 覆盖目标每日销量文件中的总计 sheet。"""
    wb_draft = openpyxl.load_workbook(draft_file, data_only=False)
    if sheet_name not in wb_draft.sheetnames:
        print(f"未找到 {sheet_name} sheet，跳过覆盖每日销量文件。")
        return

    wb_target = openpyxl.load_workbook(target_file, data_only=False)
    source_ws = wb_draft[sheet_name]

    if sheet_name in wb_target.sheetnames:
        sheet_index = wb_target.sheetnames.index(sheet_name)
        target_ws_old = wb_target[sheet_name]
        if len(wb_target.worksheets) > 1:
            wb_target.remove(target_ws_old)
            target_ws = wb_target.create_sheet(sheet_name, sheet_index)
        else:
            target_ws = target_ws_old
            for merged_range in list(target_ws.merged_cells.ranges):
                target_ws.unmerge_cells(str(merged_range))
            target_ws.delete_rows(1, target_ws.max_row)
    else:
        target_ws = wb_target.create_sheet(sheet_name)

    copy_sheet_contents(source_ws, target_ws)
    wb_target.save(target_file)
    print(f"已将 draft.xlsx 的 {sheet_name} sheet 覆盖到最新每日销量文件: {target_file}")

def main():
    # 1. 查找并验证文件
    files_with_dates = find_xlsx_files_with_dates(CURRENT_DIR)
    if len(files_with_dates) < 3:
        print("错误：当前目录中至少需要 3 个 .xlsx 文件。")
        print("找到的文件：", [f for f,d in files_with_dates])
        return

    # 取最新三份文件
    file_prev2_name, date_prev2 = files_with_dates[-3]
    file_prev_name, date_prev = files_with_dates[-2]
    file_latest_name, date_latest = files_with_dates[-1]

    file_prev2 = os.path.join(CURRENT_DIR, file_prev2_name)
    file_prev = os.path.join(CURRENT_DIR, file_prev_name)
    file_latest = os.path.join(CURRENT_DIR, file_latest_name)

    print("使用文件 (按时间从早到晚)：")
    print("01:", file_prev2_name, date_prev2)
    print("02:", file_prev_name, date_prev)
    print("03:", file_latest_name, date_latest)

    # 2. 加载数据
    print("正在加载数据...")
    _, _, data_prev2 = load_sheet_data(file_prev2)
    _, _, data_prev = load_sheet_data(file_prev)
    _, _, data_latest = load_sheet_data(file_latest)

    # 3. 计算周度数据
    print("正在计算周度数据...")
    week1, week0 = calculate_weekly_data(
        data_prev2, data_prev, data_latest, 
        date_prev2, date_prev, date_latest
    )
    
    if week1 is None or week0 is None:
        return

    # 4. 分析SKU性能
    print("正在分析SKU性能...")
    top2_sales, bottom2_sales, top2_conv, bottom2_conv, sales_results, conv_results = analyze_sku_performance(week1, week0)

    # 5. 写入结果
    print("正在写入分析结果...")
    wb_out = openpyxl.load_workbook(file_latest, data_only=False)
    ws_out = wb_out.worksheets[0]
    
    # 初始化Excel写入器
    excel_writer = ExcelWriter(ws_out)

    # 第一步：调整原始数据格式和排序（必须在最前面调用）
    excel_writer.adjust_original_data_simple()
    
    # 第二步：查找SKU列和构建映射
    excel_writer.find_sku_column()
    excel_writer.build_sku_row_map()
    
    # # 写入ACOS列
    excel_writer.write_acos_column(data_latest)
    # 在main.py中修改输出信息
    print("正在分析ACOS...")
    top_acos_results = calculate_acos_analysis(week1, week0)
    print(f"ACOS分析完成，找到{len(top_acos_results)}个本周ACOS>10%且销量增长最快的SKU")
    for result in top_acos_results:
        print(f"  SKU: {result['sku']}, 销量增长: {result['sales_growth_rate']:+.2f}%, 本周ACOS: {result['acos2']:.2f}%, 广告出单占比: {result['ad_sales_ratio2']:.2f}%")

    # 新增：分析高转化率产品
    print("正在分析高转化率产品...")
    high_conversion_results = analyze_high_conversion_products(week1, week0, data_latest)
    
    # 新增：分析高效产品（高转化低ACOS）
    print("正在分析高效产品...")
    efficient_results = analyze_high_conversion_low_acos_products(week1, week0, data_latest)
    
    print(f"ACOS分析完成，找到{len(top_acos_results)}个本周ACOS>10%且销量增长最快的SKU")
    print(f"高转化率产品分析完成，找到{len(high_conversion_results)}个订单>10的高转化产品")
    print(f"高效产品分析完成，找到{len(efficient_results)}个转化率>10%且ACOS<5%的产品")

    # 添加求和行
    print("正在添加求和行...")
    excel_writer.write_sum_row()

    # 管理新上架SKU
    print("正在管理新上架SKU...")
    manage_new_arrival_skus(ws_out, excel_writer.sku_row_map, excel_writer.sku_col_idx, file_latest_name, date_latest)

    # 写入分析结果（现在改为表格形式）
    excel_writer.write_analysis_table(top2_sales, bottom2_sales, top2_conv, bottom2_conv)

    # 新增：写入ACOS分析表格
    excel_writer.write_acos_analysis_table(top_acos_results)
    excel_writer.write_high_conversion_table(high_conversion_results)  # 新增
    excel_writer.write_efficient_products_table(efficient_results)     # 新增
    
    # 写入汇总表
    print("正在生成汇总表...")
    excel_writer.write_summary_table(file_prev2, file_prev, file_latest, date_prev2, date_prev, date_latest)

    # 6. 保存结果
    try:
        draft_file = os.path.join(os.path.dirname(file_latest), "draft.xlsx")
        wb_out.save(draft_file)
        overwrite_summary_sheet_from_draft(draft_file, file_latest)
        print("分析完成！最新-----------")
        print(f"销量分析SKU数量: {len(sales_results)}")
        print(f"转化率分析SKU数量: {len(conv_results)}")
        print(f"临时结果文件: {draft_file}")
    except PermissionError as e:
        print(f"错误：无法保存文件，文件可能正在被其他程序使用")
        print(f"请关闭Excel后重试，或检查文件权限")
        return

if __name__ == "__main__":
    main()
