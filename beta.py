import re
from typing import Dict, List, Tuple
from dataclasses import dataclass
import xlsxwriter
import sys

@dataclass
class Member:
    id: int
    name: str
    city: str
    status: str
    practice_records: List[Tuple[int, str, str]]  # List of (minutes, content, date) tuples

def parse_member_list(content: str) -> Dict[int, Member]:
    members = {}
    lines = content.strip().split('\n')
    
    # Skip header lines (both the title and the column headers)
    for line in lines[2:]:  # Skip both "## 在群人员名单" and the column headers
        if not line.strip():
            continue
        # 使用制表符分割，并确保去除每个字段的空白字符
        parts = [part.strip() for part in line.split('\t')]
        if len(parts) >= 3:  # 只需要确保至少有ID、昵称和城市三个字段
            try:
                member_id = int(parts[0])
                members[member_id] = Member(
                    id=member_id,
                    name=parts[1],
                    city=parts[2],
                    status=parts[3] if len(parts) > 3 else "",
                    practice_records=[]
                )
            except ValueError:
                print(f"Warning: Could not parse member ID from line: {line}")
                continue
    return members

def parse_practice_records(content: str, members: Dict[int, Member]):
    lines = content.strip().split('\n')
    current_date = ""
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Check for date headers
        # 示例： 2025年3月18日 星期二
        date_header_pattern = re.match(r'^(\d{4})年(\d{1,2})月(\d{1,2})日\s*(?:星期[一二三四五六日])?', line)
        if date_header_pattern:
            # Extract date from header, e.g., "2025年3月18日 星期二" -> "3月18日"
            current_date = f"{date_header_pattern.group(2)}月{date_header_pattern.group(3)}日"
            continue
            
        if line.startswith('#'):
            continue
            
        # Skip example lines
        if "例" in line:
            continue
            
        # Match pattern: number. 。member_id。 name (city)。minutes。 content
        # 示例：1. 。140。VV（四川成都）。70。拜厄89.90.94，乐曲。交作业啦
        match = re.match(r'\d+\.\s*[。.]\s*(\d+)\s*[。.]\s*([^(]+?)\s*[（(]([^)）]+)[)）]\s*[。.]\s*(\d+)\s*[。.]\s*(.*)', line)
        if match:
            member_id = int(match.group(1))
            minutes = int(match.group(4))
            content = match.group(5).strip()
            
            if member_id in members:
                members[member_id].practice_records.append((minutes, content, current_date))
                # print(f"Successfully parsed record for member {member_id}: {minutes} minutes on {current_date}")
            else:
                print(f"Warning: Member ID {member_id} not found in member list")
        else:
            print(f"Warning: Could not parse line: {line}。本条记录不参与统计")

def calculate_statistics(members: Dict[int, Member]) -> List[Tuple[int, str, int, float, int, int, List[Tuple[int, str, str]]]]:
    # Filter members who need to check in (empty status)
    required_members = {k: v for k, v in members.items() if not v.status}
    
    # Calculate statistics for each member
    stats = []
    for member_id, member in required_members.items():
        total_minutes = sum(record[0] for record in member.practice_records)
        total_hours = round(total_minutes / 60, 2)
        days = len(member.practice_records)
        daily_records = member.practice_records

        stats.append((member_id, member.name, total_minutes, total_hours, days, 0, daily_records))
    
    # Sort by total minutes and assign rankings
    stats.sort(key=lambda x: x[2], reverse=True)
    
    # Assign rankings with same rank for equal values
    current_rank = 1
    previous_minutes = None
    
    for i, stat in enumerate(stats):
        if previous_minutes is None or stat[2] < previous_minutes:
            current_rank = i + 1
        previous_minutes = stat[2]
        # Preserve the daily_records when updating the ranking
        stats[i] = (*stat[:-1], current_rank, stat[6])
    
    return stats

def find_non_compliant_members(members: Dict[int, Member]) -> List[int]:
    non_compliant = []
    for member_id, member in members.items():
        if not member.status:  # Only check members who need to check in
            total_minutes = sum(record[0] for record in member.practice_records)
            days = len(member.practice_records)
            
            if total_minutes < 120 and days < 2:
                non_compliant.append(member_id)
    return non_compliant

def generate_attendance_excel(excel_data: list, start_year: str, start_month: str, start_day: str, end_month: str, end_day: str):
    """
    Generate Excel file for attendance records
    """
    # Create Excel writer with xlsxwriter engine
    output_filename = f'files/{start_year}{start_month.zfill(2)}月打卡（{start_month}.{start_day}-{end_month}.{end_day}) .xlsx'
    workbook = xlsxwriter.Workbook(output_filename)
    worksheet = workbook.add_worksheet('打卡记录')
    
    # Define formats
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'text_wrap': True  # Enable text wrapping for multi-line text
    })
    
    cell_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    warning_format = workbook.add_format({
        'bg_color': '#FFB6C1',  # Light pink
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    
    # Set column widths
    worksheet.set_column('A:A', 8)   # 月份列
    worksheet.set_column('B:B', 10)  # 入群编号
    worksheet.set_column('C:C', 15)  # 姓名
    for day_offset in range(7):
        col = 3 + day_offset * 2  # 从D列开始，每天占2列
        worksheet.set_column(col, col, 10)     # 分钟数列
        worksheet.set_column(col + 1, col + 1, 30)  # 内容列
    worksheet.set_column('Q:T', 15)  # 统计列
    
    # Define weekdays in Chinese
    weekdays = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    
    # Set row height for header
    worksheet.set_row(0, 20)  # Set height for date row
    worksheet.set_row(1, 20)  # Set height for weekday row
    
    # Write month in first column
    worksheet.merge_range(0, 0, 1, 0, f'{start_month}月', header_format)
    
    # Format other headers - starting from second column
    worksheet.merge_range(0, 1, 1, 1, '入群编号', header_format)
    worksheet.merge_range(0, 2, 1, 2, '姓名', header_format)
    
    # Merge cells for date headers
    for day_offset in range(7):
        col = 3 + day_offset * 2  # 从D列开始，每天占2列
        current_day = int(start_day) + day_offset
        current_month = start_month if current_day <= 31 else end_month
        current_day_str = str(current_day if current_day <= 31 else 1)
        date_str = f'{start_year}/{current_month}/{current_day_str}'
        weekday_str = weekdays[day_offset]
        # 日期和星期分别写入两行
        worksheet.merge_range(0, col, 0, col + 1, date_str, header_format)
        worksheet.merge_range(1, col, 1, col + 1, weekday_str, header_format)
    
    # Write statistics headers - starting from column 17 (Q)
    stats_headers = ['总时长（分钟）', '总时长（小时）', '总天数', '本周排名（总时长）']
    for i, header in enumerate(stats_headers):
        col = 17 + i  # Start from column Q (17)
        worksheet.merge_range(0, col, 1, col, header, header_format)
    
    # Write data rows
    for row_idx, row_data in enumerate(excel_data, start=2):
        # Write basic info with cell_format
        worksheet.write(row_idx, 0, row_data['月份'], cell_format)  # 月份
        worksheet.write(row_idx, 1, row_data['入群编号'], cell_format)  # 入群编号
        worksheet.write(row_idx, 2, row_data['姓名'], cell_format)  # 姓名
        
        # Write daily records with cell_format
        for day_offset in range(7):
            col = 3 + day_offset * 2  # 从D列开始，每天占2列
            current_day = int(start_day) + day_offset
            current_month = start_month if current_day <= 31 else end_month
            current_day_str = str(current_day if current_day <= 31 else 1)
            
            minutes_key = f'{current_month}月{current_day_str}日打卡分钟数'
            content_key = f'{current_month}月{current_day_str}日打卡内容'
            
            worksheet.write(row_idx, col, row_data[minutes_key], cell_format)
            worksheet.write(row_idx, col + 1, row_data[content_key], cell_format)
        
        # Write statistics with appropriate format
        total_minutes = row_data['总时长（分钟）']
        total_days = row_data['总天数']
        format_to_use = warning_format if total_minutes < 120 and total_days < 2 else cell_format
        
        worksheet.write(row_idx, 17, total_minutes, cell_format)  # 总时长（分钟）
        worksheet.write(row_idx, 18, row_data['总时长（小时）'], format_to_use)  # 总时长（小时）
        worksheet.write(row_idx, 19, total_days, format_to_use)  # 总天数
        worksheet.write(row_idx, 20, row_data['本周排名（总时长）'], cell_format)  # 排名
    
    # Save the Excel file
    workbook.close()
    print(f"\n统计数据已保存到 '{output_filename}'")

def generate_ranking_excel(stats: list, start_year: str, start_month: str, start_day: str, end_month: str, end_day: str):
    """
    Generate Excel file for ranking
    """
    output_filename = f'files/{start_year}{start_month.zfill(2)}月打卡排名（{start_month}.{start_day}-{end_month}.{end_day}) .xlsx'
    workbook = xlsxwriter.Workbook(output_filename)
    worksheet = workbook.add_worksheet('排名')

    # Define formats
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#D7E4BC',
        'border': 1
    })
    title_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'font_size': 12
    })
    cell_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })
    warning_format = workbook.add_format({
        'bg_color': '#FFB6C1',  # Light pink
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    # Set column widths
    worksheet.set_column('A:A', 10)  # 入群编号
    worksheet.set_column('B:B', 15)  # 姓名
    worksheet.set_column('C:E', 12)  # 统计数据列
    worksheet.set_column('F:F', 8)   # 排名

    # Write title - merge all columns
    title = f'周排名（{start_month}.{start_day}-{end_month}.{end_day}）'
    worksheet.merge_range(0, 0, 0, 5, title, title_format)

    # Write headers
    headers = ['入群编号', '姓名', '总时长（分钟）', '总时长（小时）', '总天数', '排名']
    for col, header in enumerate(headers):
        worksheet.write(1, col, header, header_format)

    # Write data
    for row, stat in enumerate(stats, start=2):
        total_minutes = stat[2]
        total_days = stat[4]
        format_to_use = warning_format if total_minutes < 120 and total_days < 2 else cell_format
        
        worksheet.write(row, 0, stat[0], cell_format)  # 入群编号
        worksheet.write(row, 1, stat[1], cell_format)  # 姓名
        worksheet.write(row, 2, total_minutes, cell_format)  # 总时长（分钟）
        worksheet.write(row, 3, stat[3], format_to_use)  # 总时长（小时）
        worksheet.write(row, 4, total_days, format_to_use)  # 总天数
        worksheet.write(row, 5, stat[6], cell_format)  # 排名

    workbook.close()
    print(f"\n排名表已保存到 '{output_filename}'")

def generate_warning_message(non_compliant: List[int], weekday: str) -> str:
    """
    Generate warning message for non-compliant members
    """
    message = "📣统计组预警提醒：\n\n"
    message += f"今天{weekday}啦！\n"
    message += "打卡群周最低线：天数≥2天或总时长≥2小时，二者满足其一即可。\n\n"
    message += "以下在打卡群（请假除外）参与本周打卡统计的伙伴还要差一丢丢，各位小伙伴周末加加油哦[嘿哈]\n\n"
    message += ",".join(map(str, sorted(non_compliant)))
    message += "\n\n（统计截至周五打卡数据，如有今天已经达标的，忽略即可~)"
    return message

def save_warning_message(message: str, start_date: str):
    """
    Save warning message to file
    """
    output_filename = "files/oncall_msg.txt"
    with open(output_filename, "w", encoding="utf-8") as f:
        f.write(message)
    print(f"\n预警消息已保存到 '{output_filename}'")

def process_data(member_list_content: str, practice_records_content: str, start_date: str):
    """
    Process the data with a given start date.
    start_date format: 'YYYYMMDD', e.g., '20250317'
    """
    # Parse start date
    year = start_date[:4]
    start_month = str(int(start_date[4:6]))  # Remove leading zero
    start_day = str(int(start_date[6:8]))    # Remove leading zero
    
    # Calculate end date
    start_day_int = int(start_day)
    end_day_int = start_day_int + 6
    
    # Handle month rollover
    days_in_month = 31  # Simplified version, you might want to add proper month length calculation
    end_month = start_month
    if end_day_int > days_in_month:
        end_day_int = end_day_int - days_in_month
        end_month = str(int(start_month) + 1)
    end_day = str(end_day_int)
    
    # Parse input data
    members = parse_member_list(member_list_content)
    parse_practice_records(practice_records_content, members)
    
    # Calculate statistics
    stats = calculate_statistics(members)
    
    # Find non-compliant members
    non_compliant = find_non_compliant_members(members)
    
    # Print results
    print("\n1. 统计在群人员名单中，需要打卡人员的本周打卡记录")
    # print("入群编号\t姓名\t总时长（分钟）\t总时长（小时）\t总天数\t本周排名（总时长）")
    # for stat in stats:
    #     print(f"{stat[0]}\t{stat[1]}\t{stat[2]}\t{stat[3]}\t{stat[4]}\t{stat[6]}")
    
    print("\n2. 统计在群人员名单中，本周打卡不达标的成员序号名单")
    print(",".join(map(str, non_compliant)))
    
    # Generate and save warning message
    weekdays = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    current_weekday = weekdays[5]  # 固定为星期六
    warning_message = generate_warning_message(non_compliant, current_weekday)
    save_warning_message(warning_message, start_date)
    
    # Prepare data for Excel
    excel_data = []
    for stat in stats:
        # Create a row with basic stats
        row = {
            '月份': '',  # Add empty column for month
            '入群编号': stat[0],
            '姓名': stat[1]
        }
        
        # Initialize all days with default values
        for day_offset in range(7):
            current_day = start_day_int + day_offset
            current_month = start_month if current_day <= days_in_month else end_month
            current_day_str = str(current_day if current_day <= days_in_month else current_day - days_in_month)
            row[f'{current_month}月{current_day_str}日打卡分钟数'] = 0
            row[f'{current_month}月{current_day_str}日打卡内容'] = ""
        
        # Add daily records
        daily_records = stat[7]  # Get daily records from stats
        
        # Fill in actual practice records
        for record in daily_records:
            minutes, content, date = record
            # Extract day number from date (e.g., "3月18日" -> 18)
            day_match = re.search(r'(\d+)月(\d+)日', date)
            if day_match:
                record_month = int(day_match.group(1))
                record_day = int(day_match.group(2))
                if record_month == int(start_month) or record_month == int(end_month):
                    day_offset = record_day - start_day_int
                    if 0 <= day_offset < 7:  # Ensure day is within valid range
                        current_month = str(record_month)
                        row[f'{current_month}月{record_day}日打卡分钟数'] = minutes
                        row[f'{current_month}月{record_day}日打卡内容'] = content
        
        # Add remaining stats
        row.update({
            '总时长（分钟）': stat[2],
            '总时长（小时）': stat[3],
            '总天数': stat[4],
            '本周排名（总时长）': stat[6]
        })
        
        excel_data.append(row)
    
    # Generate Excel files
    generate_attendance_excel(excel_data, year, start_month, start_day, end_month, end_day)
    generate_ranking_excel(stats, year, start_month, start_day, end_month, end_day)

if __name__ == "__main__":
    # Check if start_date is provided as command line argument
    if len(sys.argv) != 2:
        print("错误：请提供开始日期参数")
        print("用法：python3 beta.py YYYYMMDD")
        sys.exit(1)
    
    start_date = sys.argv[1]
    
    # Parse start date for file names
    year = start_date[:4]
    month = str(int(start_date[4:6])).zfill(2)  # Keep leading zero
    day = str(int(start_date[6:8])).zfill(2)    # Keep leading zero
    end_day = str(int(day) + 6).zfill(2)
    
    # Generate file names (without end date in input file names)
    member_list_file = f"files/{year}{month}{day}_在群人员名单.md"
    practice_records_file = f"files/{year}{month}{day}_打卡记录.md"
    
    try:
        # Read member list
        with open(member_list_file, "r", encoding="utf-8") as f:
            content = f.read()
            # 直接使用整个文件内容，因为成员名单文件只包含成员名单
            member_list_content = content
        
        # Read practice records
        with open(practice_records_file, "r", encoding="utf-8") as f:
            practice_records_content = f.read()
        
        # Process the data
        process_data(member_list_content, practice_records_content, start_date)
        
    except FileNotFoundError as e:
        print(f"错误：找不到必需的文件 - {str(e)}")
        sys.exit(1)
    except Exception as e:
        print(f"错误：程序运行出错 - {str(e)}")
        sys.exit(1)




