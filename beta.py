import re
from typing import Dict, List, Tuple
from dataclasses import dataclass
from datetime import datetime
import pandas as pd

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
    
    # Skip header
    for line in lines[1:]:
        if not line.strip():
            continue
        parts = line.split('\t')
        if len(parts) >= 4:
            member_id = int(parts[0])
            members[member_id] = Member(
                id=member_id,
                name=parts[1],
                city=parts[2],
                status=parts[3] if len(parts) > 3 else "",
                practice_records=[]
            )
    return members

def parse_practice_records(content: str, members: Dict[int, Member]):
    lines = content.strip().split('\n')
    current_date = ""
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Check for date headers
        if "年" in line:
            # Extract date from header, e.g., "2025年3月18日 星期二" -> "3月18日"
            date_match = re.search(r'(\d+)年(\d+)月(\d+)日', line)
            if date_match:
                current_date = f"{date_match.group(2)}月{date_match.group(3)}日"
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
                print(f"Successfully parsed record for member {member_id}: {minutes} minutes on {current_date}")
            else:
                print(f"Warning: Member ID {member_id} not found in member list")
        else:
            print(f"Warning: Could not parse line: {line}")

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
    for i, stat in enumerate(stats, 1):
        # Preserve the daily_records when updating the ranking
        stats[i-1] = (*stat[:-1], i, stat[6])
    
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

def process_data(member_list_content: str, practice_records_content: str, start_date: str):
    """
    Process the data with a given start date.
    start_date format: 'YYYYMMDD', e.g., '20250317'
    """
    # Parse start date
    year = start_date[:4]
    month = str(int(start_date[4:6]))  # Remove leading zero
    day = str(int(start_date[6:8]))    # Remove leading zero
    start_day = int(day)
    
    # Parse input data
    members = parse_member_list(member_list_content)
    parse_practice_records(practice_records_content, members)
    
    # Calculate statistics
    stats = calculate_statistics(members)
    
    # Find non-compliant members
    non_compliant = find_non_compliant_members(members)
    
    # Print results
    print("\n1. 统计在群人员名单中，需要打卡人员的本周打卡记录")
    print("入群编号\t姓名\t总时长（分钟）\t总时长（小时）\t总天数\t本周排名（总时长）")
    for stat in stats:
        print(f"{stat[0]}\t{stat[1]}\t{stat[2]}\t{stat[3]}\t{stat[4]}\t{stat[6]}")
    
    print("\n2. 统计在群人员名单中，本周打卡不达标的成员序号名单")
    print(",".join(map(str, non_compliant)))
    
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
            current_day = start_day + day_offset
            row[f'{month}月{current_day}日打卡分钟数'] = 0
            row[f'{month}月{current_day}日打卡内容'] = ""
        
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
                if record_month == int(month):
                    day_offset = record_day - start_day
                    if 0 <= day_offset < 7:  # Ensure day is within valid range
                        row[f'{month}月{record_day}日打卡分钟数'] = minutes
                        row[f'{month}月{record_day}日打卡内容'] = content
        
        # Add remaining stats
        row.update({
            '总时长（分钟）': stat[2],
            '总时长（小时）': stat[3],
            '总天数': stat[4],
            '本周排名（总时长）': stat[6]
        })
        
        excel_data.append(row)
    
    # Create DataFrame and save to Excel
    df = pd.DataFrame(excel_data)
    
    # Create Excel writer with xlsxwriter engine
    output_filename = f'{year}{month.zfill(2)}月打卡（{month}.{day}-{month}.{start_day+6}) .xlsx'
    writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='打卡记录', startrow=2, header=False)  # Start from row 3 and don't write headers
    
    # Get workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets['打卡记录']
    
    # Define formats
    header_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'border': 1,
        'text_wrap': True  # Enable text wrapping for multi-line text
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
    worksheet.merge_range(0, 0, 1, 0, f'{month}月', header_format)
    
    # Format other headers - starting from second column
    worksheet.merge_range(0, 1, 1, 1, '入群编号', header_format)
    worksheet.merge_range(0, 2, 1, 2, '姓名', header_format)
    
    # Merge cells for date headers
    for day_offset in range(7):
        col = 3 + day_offset * 2  # 从D列开始，每天占2列
        current_day = start_day + day_offset
        date_str = f'{year}/{month}/{current_day}'
        weekday_str = weekdays[day_offset]
        # 日期和星期分别写入两行
        worksheet.merge_range(0, col, 0, col + 1, date_str, header_format)
        worksheet.merge_range(1, col, 1, col + 1, weekday_str, header_format)
    
    # Write statistics headers - starting from column 17 (Q)
    stats_headers = ['总时长（分钟）', '总时长（小时）', '总天数', '本周排名（总时长）']
    for i, header in enumerate(stats_headers):
        col = 17 + i  # Start from column Q (17)
        worksheet.merge_range(0, col, 1, col, header, header_format)
    
    # Save the Excel file
    writer.close()
    print(f"\n统计数据已保存到 '{output_filename}'")

if __name__ == "__main__":
    start_date = '20250317'  # Format: YYYYMMDD
    
    # Parse start date for file names
    year = start_date[:4]
    month = str(int(start_date[4:6])).zfill(2)  # Keep leading zero
    day = str(int(start_date[6:8])).zfill(2)    # Keep leading zero
    end_day = str(int(day) + 6).zfill(2)
    
    # Generate file names
    member_list_file = f"{year}{month}{day}-{year}{month}{end_day}_在群人员名单.md"
    practice_records_file = f"{year}{month}{day}-{year}{month}{end_day}_打卡记录.md"
    
    # Read member list
    with open(member_list_file, "r", encoding="utf-8") as f:
        content = f.read()
        member_list_content = content.split("## 打卡记录")[0].split("## 在群人员名单")[1].strip()
    
    # Read practice records
    with open(practice_records_file, "r", encoding="utf-8") as f:
        practice_records_content = f.read()
    
    # Process the data
    process_data(member_list_content, practice_records_content, start_date)
