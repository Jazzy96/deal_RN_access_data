import polars as pl
from datetime import datetime
from tqdm import tqdm
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

def format_duration_hours(seconds):
    """将秒数转换为小时（保留2位小数）"""
    hours = seconds / 3600  # 将秒转换为小时
    return round(hours, 2)  # 保留2位小数

def process_wifi_data(file_path):
    try:
        # 读取Excel文件
        print(f"\n正在处理 {file_path}...")
        df = pl.read_excel(file_path)
        total_rows = len(df)
        print(f"总行数: {total_rows}")

        if total_rows == 0:
            print("警告：文件为空！")
            return None

        # 1. 过滤掉user不为空的行，并保留需要的列
        df = df.filter(pl.col('user').is_null()).select(['serial_no', 'mac', 'signal', 'tx_rate', 'rx_rate', 'create_time'])
        valid_rows = len(df)
        print(f"有效行数: {valid_rows} (已过滤user不为空的行)")

        if valid_rows == 0:
            print("警告：没有有效数据行！")
            return None

        # 将create_time转换为datetime类型
        df = df.with_columns(pl.col('create_time').str.strptime(pl.Datetime))

        # 按MAC地址分组并排序
        df = df.sort(['mac', 'create_time'])
        unique_macs = df['mac'].n_unique()
        print(f"不同MAC地址数量: {unique_macs}")

        if unique_macs == 0:
            print("警告：没有有效的MAC地址！")
            return None

        results = []

        print("\n开始处理每个MAC地址的数据...")
        for mac, group in df.groupby('mac', maintain_order=True):
            if len(group) < 2:  # 如果某个MAC只有一条记录，跳过
                continue

            times = group['create_time'].to_list()
            segment_start_idx = 0

            for i in range(1, len(times)):
                time_diff = (times[i] - times[i-1]).total_seconds()
                if time_diff > 300:  # 5分钟阈值
                    # 处理当前段
                    segment = group.slice(segment_start_idx, i - segment_start_idx)
                    sn = segment['serial_no'][0]
                    avg_signal = segment['signal'].mean()
                    avg_tx_rate = segment['tx_rate'].mean()
                    avg_rx_rate = segment['rx_rate'].mean()
                    duration = (times[i-1] - times[segment_start_idx]).total_seconds()

                    results.append({
                        'SN': sn,
                        'mac': mac,
                        'avg_signal': round(float(avg_signal), 2),
                        'avg_tx_rate': round(float(avg_tx_rate), 2),
                        'avg_rx_rate': round(float(avg_rx_rate), 2),
                        'total_duration(hour)': format_duration_hours(duration),
                        'start_time': times[segment_start_idx]
                    })

                    segment_start_idx = i

            # 处理最后一段
            segment = group.slice(segment_start_idx)
            sn = segment['serial_no'][0]
            avg_signal = segment['signal'].mean()
            avg_tx_rate = segment['tx_rate'].mean()
            avg_rx_rate = segment['rx_rate'].mean()
            duration = (times[-1] - times[segment_start_idx]).total_seconds()

            results.append({
                'SN': sn,
                'mac': mac,
                'avg_signal': round(float(avg_signal), 2),
                'avg_tx_rate': round(float(avg_tx_rate), 2),
                'avg_rx_rate': round(float(avg_rx_rate), 2),
                'total_duration(hour)': format_duration_hours(duration),
                'start_time': times[segment_start_idx]
            })

        if not results:
            print("警告：没有生成任何结果！")
            return None

        # 创建结果DataFrame并按开始时间排序
        result_df = pl.DataFrame(results).sort('start_time')
        return result_df

    except Exception as e:
        print(f"处理文件 {file_path} 时发生错误: {str(e)}")
        return None

def format_worksheet(worksheet):
    # 设置所有单元格居中对齐
    for row in worksheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')

    # 获取每列的最大宽度
    column_widths = {}

    # 遍历所有行（包括表头和数据）
    for row in worksheet.rows:
        for cell in row:
            if cell.value is not None:  # 确保单元格有值
                # 获取列号
                column = cell.column_letter
                # 计算当前单元格内容的显示宽度（中英文分别处理）
                cell_length = 0
                for char in str(cell.value):
                    if ord(char) <= 127:  # ASCII字符
                        cell_length += 1
                    else:  # 非ASCII字符（如中文）
                        cell_length += 2

                # 更新该列的最大宽度
                current_width = column_widths.get(column, 0)
                column_widths[column] = max(current_width, cell_length)

    # 设置每列的宽度（添加一些padding）
    for column, width in column_widths.items():
        worksheet.column_dimensions[column].width = width + 4  # +4 作为边距