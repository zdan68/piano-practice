#!/bin/bash

# Check if date parameter is provided
if [ $# -eq 0 ]; then
    echo "请提供开始日期参数，格式：YYYYMMDD"
    echo "例如：./run.sh 20250317"
    exit 1
fi

start_date=$1

# Validate date format
if ! [[ $start_date =~ ^[0-9]{8}$ ]]; then
    echo "日期格式错误！请使用YYYYMMDD格式"
    echo "例如：./run.sh 20250317"
    exit 1
fi

# Extract date components
year=${start_date:0:4}
month=${start_date:4:2}
day=${start_date:6:2}

# Calculate end date (start_date + 6 days)
end_date=$(date -v+6d -j -f "%Y%m%d" "$start_date" "+%Y%m%d")

# Check if required files exist
member_list_file="${start_date}-${end_date}_在群人员名单.md"
practice_records_file="${start_date}-${end_date}_打卡记录.md"

if [ ! -f "$member_list_file" ]; then
    echo "错误：找不到文件 $member_list_file"
    exit 1
fi

if [ ! -f "$practice_records_file" ]; then
    echo "错误：找不到文件 $practice_records_file"
    exit 1
fi

# Run the Python script with the start date parameter
python3 beta.py "$start_date" 