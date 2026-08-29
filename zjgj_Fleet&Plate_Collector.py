# -*- coding: utf-8 -*-

import json
import os
import time
import threading
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
from openpyxl import Workbook, load_workbook


BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LINE_CONFIG_FILE = os.path.join(
    BASE_DIR,
    'lines.json'
)

EXCEL_FILE = os.path.join(
    BASE_DIR,
    'vehicle_id_code.xlsx'
)

INTERVAL_SECONDS = 10 * 60
MAX_WORKERS = 15
REQUEST_TIMEOUT = 10

API_URL = (
    'http://121.40.11.192:8080/lyx/lineBus/getLinePipe'
    '?lineId={}&upOrDownStr={}'
)

HEADERS = [
    '线路号',
    '上下行',
    '自编号',
    '车牌号',
    '发车时间',
    '工号'
]

_thread_local = threading.local()


def get_session():
    """获取当前线程的HTTP Session。"""

    if not hasattr(_thread_local, 'session'):
        session = requests.Session()
        session.headers.update({
            'User-Agent': (
                'Mozilla/5.0 '
                '(Windows NT 10.0; Win64; x64) '
                'AppleWebKit/537.36 '
                '(KHTML, like Gecko) '
                'Chrome/120 Safari/537.36'
            )
        })
        _thread_local.session = session

    return _thread_local.session


def load_lines():
    """读取线路配置。"""

    if not os.path.exists(LINE_CONFIG_FILE):
        raise FileNotFoundError(
            f'找不到线路配置文件：{LINE_CONFIG_FILE}'
        )

    with open(
        LINE_CONFIG_FILE,
        'r',
        encoding='utf-8'
    ) as f:
        config = json.load(f)

    if not isinstance(config, dict):
        raise ValueError(
            'lines.json格式错误：最外层必须是对象。'
        )

    lines = config.get('lines')

    if not isinstance(lines, list):
        raise ValueError(
            'lines.json中必须存在lines数组。'
        )

    result = []

    for index, item in enumerate(lines, start=1):

        if not isinstance(item, dict):
            print(
                f'警告：第{index}个线路配置不是对象，已跳过。'
            )
            continue

        line_id = item.get('id')
        line_name = item.get('name')
        company = item.get('company', '')

        if line_id is None or not str(line_id).strip():
            print(
                f'警告：第{index}个线路配置缺少id，已跳过。'
            )
            continue

        if line_name is None or not str(line_name).strip():
            print(
                f'警告：线路{line_id}缺少name，已跳过。'
            )
            continue

        result.append({
            'id': str(line_id).strip(),
            'name': str(line_name).strip(),
            'company': str(company).strip()
        })

    if not result:
        raise ValueError(
            'lines.json中没有有效线路。'
        )

    return result


def format_line_name(line):
    """返回JSON中配置的线路名称。"""

    return line['name']


def format_plan_run_time(plan_run_time, scan_date):
    """将planRunTime转换为Excel日期时间。"""

    if not plan_run_time:
        return None

    plan_run_time = str(plan_run_time).strip()

    for fmt in (
        '%Y-%m-%d %H:%M',
        '%Y-%m-%d %H:%M:%S'
    ):
        try:
            return datetime.strptime(
                f'{scan_date} {plan_run_time}',
                fmt
            )
        except ValueError:
            continue

    print(
        f'[{datetime.now().isoformat()}] '
        f'无法解析planRunTime：{plan_run_time}'
    )

    return None


def get_one_line_vehicle_info(line, mark, scan_date):
    """获取一条线路一个方向的车辆信息。"""

    line_id = line['id']
    line_name = format_line_name(line)

    direction = (
        '上行'
        if mark == 0
        else '下行'
    )

    session = get_session()

    try:
        response = session.get(
            API_URL.format(line_id, mark),
            timeout=REQUEST_TIMEOUT
        )

        response.raise_for_status()
        data = response.json()

    except Exception as e:
        print(
            f'[{datetime.now().isoformat()}] '
            f'{line_name} {direction}请求失败：{e}'
        )
        return []

    line_pipe_beans = data.get('linePipeBeans')

    if not isinstance(line_pipe_beans, list):
        return []

    result = []

    for item in line_pipe_beans:

        if not isinstance(item, dict):
            continue

        vehicle_id = item.get('vehicleId')
        vehicle_code = item.get('vehicleCode')
        plan_run_time = item.get('planRunTime')
        driver_code = item.get('driverCode')

        if (
            vehicle_id is None
            and vehicle_code is None
            and plan_run_time is None
            and driver_code is None
        ):
            continue

        departure_time = format_plan_run_time(
            plan_run_time,
            scan_date
        )

        result.append((
            line_name,
            direction,
            vehicle_id,
            vehicle_code,
            departure_time,
            driver_code
        ))

    return result


def get_vehicle_data(lines):
    """并发获取所有线路上下行数据。"""

    scan_date = datetime.now().strftime(
        '%Y-%m-%d'
    )

    tasks = []

    for line in lines:
        tasks.append((line, 0))
        tasks.append((line, 1))

    result = []
    total_tasks = len(tasks)
    completed = 0

    print()
    print('=' * 60)
    print(
        f'[{datetime.now().isoformat()}] 开始扫描'
    )
    print(
        f'线路：{len(lines)}条'
    )
    print(
        f'请求：{total_tasks}个'
    )
    print(
        f'并发：{MAX_WORKERS}线程'
    )
    print('=' * 60)

    with ThreadPoolExecutor(
        max_workers=MAX_WORKERS
    ) as executor:

        future_map = {
            executor.submit(
                get_one_line_vehicle_info,
                line,
                mark,
                scan_date
            ): (line, mark)
            for line, mark in tasks
        }

        for future in as_completed(future_map):

            line, mark = future_map[future]

            line_name = format_line_name(line)

            direction = (
                '上行'
                if mark == 0
                else '下行'
            )

            completed += 1

            try:
                data = future.result()
                result.extend(data)

                print(
                    f'[{completed}/{total_tasks}] '
                    f'{line_name} {direction}：'
                    f'{len(data)}条'
                )

            except Exception as e:
                print(
                    f'[{completed}/{total_tasks}] '
                    f'{line_name} {direction}失败：{e}'
                )

    print(
        f'[{datetime.now().isoformat()}] '
        f'扫描结束，共获取{len(result)}条。'
    )

    return result


def normalize_datetime(value):
    """统一Excel日期时间格式。"""

    if value is None:
        return None

    if isinstance(value, datetime):
        return value.replace(
            second=0,
            microsecond=0
        )

    value = str(value).strip()

    for fmt in (
        '%Y-%m-%d %H:%M',
        '%Y-%m-%d %H:%M:%S'
    ):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue

    return value


def make_record_key(record):
    """生成历史记录唯一键。"""

    departure_time = normalize_datetime(
        record[4]
    )

    values = [
        record[0],
        record[1],
        record[2],
        record[3],
        departure_time,
        record[5]
    ]

    return tuple(
        None if value is None
        else str(value).strip()
        if not isinstance(value, datetime)
        else value
        for value in values
    )


def setup_excel_format(ws):
    """设置Excel筛选、冻结和列宽。"""

    ws.auto_filter.ref = (
        f'A1:F{max(ws.max_row, 1)}'
    )

    ws.freeze_panes = 'A2'

    widths = {
        'A': 20,
        'B': 10,
        'C': 15,
        'D': 15,
        'E': 22,
        'F': 12
    }

    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    for row in range(2, ws.max_row + 1):
        cell = ws.cell(
            row=row,
            column=5
        )

        if isinstance(cell.value, datetime):
            cell.number_format = (
                'yyyy-mm-dd hh:mm'
            )


def create_excel(filepath):
    """创建新的历史数据Excel。"""

    wb = Workbook()
    ws = wb.active
    ws.title = '车辆信息'

    for column, header in enumerate(
        HEADERS,
        start=1
    ):
        ws.cell(
            row=1,
            column=column,
            value=header
        )

    setup_excel_format(ws)

    wb.save(filepath)
    wb.close()


def append_unique_to_excel(vehicle_list, filepath):
    """追加新记录"""

    if not os.path.exists(filepath):
        create_excel(filepath)

    wb = load_workbook(filepath)
    ws = wb.active

    for column, header in enumerate(
        HEADERS,
        start=1
    ):
        ws.cell(
            row=1,
            column=column,
            value=header
        )

    existing = set()

    for row in ws.iter_rows(
        min_row=2,
        max_col=6,
        values_only=True
    ):

        if all(value is None for value in row):
            continue

        existing.add(
            make_record_key(list(row))
        )

    to_append = []

    for item in vehicle_list:

        key = make_record_key(item)

        if key in existing:
            continue

        existing.add(key)
        to_append.append(item)

    for item in to_append:

        row = ws.max_row + 1

        for column, value in enumerate(
            item,
            start=1
        ):
            ws.cell(
                row=row,
                column=column,
                value=value
            )

        if isinstance(item[4], datetime):
            ws.cell(
                row=row,
                column=5
            ).number_format = (
                'yyyy-mm-dd hh:mm'
            )

    setup_excel_format(ws)

    wb.save(filepath)
    wb.close()

    return len(to_append)


def run_one_scan(lines):
    """执行一轮扫描并保存数据。"""

    vehicle_data = get_vehicle_data(lines)

    added = append_unique_to_excel(
        vehicle_data,
        EXCEL_FILE
    )

    print(
        f'[{datetime.now().isoformat()}] '
        f'本轮获取{len(vehicle_data)}条，'
        f'新增{added}条。'
    )

    return added


if __name__ == '__main__':

    print('=' * 60)
    print(
        'zjgj_Fleet&Plate_Collector'
    )
    print(
        'zjgj_Fleet&Plate_Collector v3.0'
    )
    print('=' * 60)

    try:
        lines = load_lines()

    except Exception as e:

        print(
            f'线路配置读取失败：{e}'
        )

        input('按Enter键退出……')
        raise SystemExit(1)

    print(
        f'线路配置：{LINE_CONFIG_FILE}'
    )

    print(
        f'线路数量：{len(lines)}'
    )

    print(
        f'扫描间隔：{INTERVAL_SECONDS // 60}分钟'
    )

    print(
        f'并发线程：{MAX_WORKERS}'
    )

    print(
        f'历史数据：{EXCEL_FILE}'
    )

    try:
        run_one_scan(lines)

    except Exception as e:

        print(
            f'首次扫描失败：{e}'
        )

    try:

        while True:

            print()
            print(
                f'[{datetime.now().isoformat()}] '
                f'等待10分钟后进行下一轮扫描……'
            )

            time.sleep(
                INTERVAL_SECONDS
            )

            try:
                run_one_scan(lines)

            except Exception as e:

                print(
                    f'周期扫描失败：{e}'
                )

    except KeyboardInterrupt:

        print()
        print(
            f'[{datetime.now().isoformat()}] '
            f'程序已退出。'
        )