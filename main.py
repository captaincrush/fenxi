# main.py
import os
import subprocess
import openpyxl
from config import CURRENT_DIR
from utils import find_xlsx_files_with_dates
from data_loader import load_sheet_data, calculate_weekly_data, analyze_sku_performance, calculate_acos_analysis,analyze_high_conversion_products, analyze_high_conversion_low_acos_products
from excel_writer import ExcelWriter
from sku_manager import manage_new_arrival_skus

ANALYSIS_DIR_NAMES = ("分析", "分析文件夹")

def overwrite_summary_sheet_from_draft(draft_file, target_file, sheet_name="总计"):
    """用 Excel COM 复制整张总计 sheet，避免 openpyxl 重写目标文件导致图片/形状丢失。"""
    script = r'''
$ErrorActionPreference = "Stop"
$draftPath = $env:DRAFT_FILE
$targetPath = $env:TARGET_FILE
$sheetName = $env:SHEET_NAME

$excel = $null
$draftWb = $null
$targetWb = $null
$tempWs = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $missing = [System.Type]::Missing

    $draftWb = $excel.Workbooks.Open($draftPath)
    $targetWb = $excel.Workbooks.Open($targetPath)

    $sourceWs = $null
    foreach ($ws in @($draftWb.Worksheets)) {
        if ($ws.Name -eq $sheetName) {
            $sourceWs = $ws
            break
        }
    }

    if ($null -eq $sourceWs) {
        Write-Output "未找到 $sheetName sheet，跳过覆盖每日销量文件。"
        return
    }

    $targetWs = $null
    $targetIndex = $targetWb.Worksheets.Count + 1
    for ($i = 1; $i -le $targetWb.Worksheets.Count; $i++) {
        $ws = $targetWb.Worksheets.Item($i)
        if ($ws.Name -eq $sheetName) {
            $targetWs = $ws
            $targetIndex = $i
            break
        }
    }

    if ($null -ne $targetWs) {
        if ($targetWb.Worksheets.Count -eq 1) {
            $tempWs = $targetWb.Worksheets.Add()
        }
        $targetWs.Delete()
    }

    if ($targetIndex -le 1) {
        $sourceWs.Copy($targetWb.Worksheets.Item(1), $missing)
    } else {
        $afterIndex = [Math]::Min($targetIndex - 1, $targetWb.Worksheets.Count)
        $sourceWs.Copy($missing, $targetWb.Worksheets.Item($afterIndex))
    }

    $copiedWs = $targetWb.ActiveSheet
    $copiedWs.Name = $sheetName

    if ($null -ne $tempWs) {
        $tempWs.Delete()
    }

    $targetWb.Save()
}
finally {
    if ($null -ne $targetWb) { $targetWb.Close($false) }
    if ($null -ne $draftWb) { $draftWb.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}
'''
    env = os.environ.copy()
    env["DRAFT_FILE"] = os.path.abspath(draft_file)
    env["TARGET_FILE"] = os.path.abspath(target_file)
    env["SHEET_NAME"] = sheet_name

    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            check=True,
            capture_output=True,
            text=True,
            env=env
        )
    except subprocess.CalledProcessError as e:
        if e.stdout:
            print(e.stdout.strip())
        if e.stderr:
            print(e.stderr.strip())
        raise

    if result.stdout.strip():
        output = result.stdout.strip()
        print(output)
        if "跳过覆盖" in output:
            return
    print(f"已将 draft.xlsx 的 {sheet_name} sheet 覆盖到最新每日销量文件: {target_file}")

def find_analysis_directories(root_dir):
    """查找每个同级A文件夹里的分析目录；没有找到时兼容处理当前目录。"""
    analysis_dirs = []
    ignored_dirs = {".git", "__pycache__", "build", "dist", "SKU"}

    for name in sorted(os.listdir(root_dir)):
        folder_path = os.path.join(root_dir, name)
        if not os.path.isdir(folder_path) or name in ignored_dirs:
            continue

        for analysis_dir_name in ANALYSIS_DIR_NAMES:
            analysis_path = os.path.join(folder_path, analysis_dir_name)
            if os.path.isdir(analysis_path):
                analysis_dirs.append(analysis_path)
                break

    if analysis_dirs:
        return analysis_dirs

    return [root_dir]

def process_directory(current_dir):
    # 1. 查找并验证文件
    print(f"\n开始处理目录: {current_dir}")
    files_with_dates = find_xlsx_files_with_dates(current_dir)
    if len(files_with_dates) < 3:
        print("错误：当前目录中至少需要 3 个 .xlsx 文件。")
        print("找到的文件：", [f for f,d in files_with_dates])
        return

    # 取最新三份文件
    file_prev2_name, date_prev2 = files_with_dates[-3]
    file_prev_name, date_prev = files_with_dates[-2]
    file_latest_name, date_latest = files_with_dates[-1]

    file_prev2 = os.path.join(current_dir, file_prev2_name)
    file_prev = os.path.join(current_dir, file_prev_name)
    file_latest = os.path.join(current_dir, file_latest_name)

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
    manage_new_arrival_skus(ws_out, excel_writer.sku_row_map, excel_writer.sku_col_idx, file_latest_name, date_latest, current_dir)

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

def main():
    root_dir = os.path.dirname(os.path.abspath(__file__))
    analysis_dirs = find_analysis_directories(root_dir)

    print(f"共找到 {len(analysis_dirs)} 个待处理分析目录")
    for analysis_dir in analysis_dirs:
        try:
            process_directory(analysis_dir)
        except Exception as e:
            print(f"处理目录失败: {analysis_dir}")
            print(f"错误信息: {e}")

if __name__ == "__main__":
    main()
