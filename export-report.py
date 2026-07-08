import os
import pyodbc
import pandas as pd
import numpy as np
from datetime import datetime
import xlsxwriter

# ==================== 配置区 ====================
connStr = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=localhost,1444;"
    "DATABASE=test;"
    "UID=sa;"
    "PWD=bigStrongPwd@123;"
)

EXPORT_DIR = os.path.join(os.getcwd(), 'exports')

# ==================== 辅助函数 ====================

def apply_mapping_and_fillna(df, header_mapping, fillna_dict):
    """基于原始字段名填充空值，再应用列头映射"""
    if df.empty:
        return df
    # 1. 空值填充（基于原始字段名）
    if fillna_dict:
        valid_fill = {col: val for col, val in fillna_dict.items() if col in df.columns}
        if valid_fill:
            df = df.fillna(valid_fill)
    # 2. 列头映射
    if header_mapping:
        valid_map = {k: v for k, v in header_mapping.items() if k in df.columns}
        if valid_map:
            df = df.rename(columns=valid_map)
    return df


def safe_write_cell(worksheet, row, col, value, cell_format=None):
    """
    安全写入单元格，根据值的类型显式调用对应的 xlsxwriter 方法，
    避免自动类型判断导致的 'float has no len()' 等错误。
    """
    # 处理 pandas/numpy 的缺失值
    if value is None or (isinstance(value, float) and np.isnan(value)):
        worksheet.write_blank(row, col, None, cell_format)
        return
    # 处理 numpy 数值类型
    if isinstance(value, (np.integer,)):
        value = int(value)
    elif isinstance(value, (np.floating,)):
        value = float(value)
    # 写入
    if isinstance(value, bool):
        worksheet.write_boolean(row, col, value, cell_format)
    elif isinstance(value, (int, float, np.number)):
        worksheet.write_number(row, col, value, cell_format)
    elif isinstance(value, str):
        worksheet.write_string(row, col, value, cell_format)
    elif isinstance(value, pd.Timestamp):
        # 写入为 Excel 日期时间格式（可自定义）
        worksheet.write_datetime(row, col, value.to_pydatetime(), cell_format)
    else:
        # 其他类型统一转为字符串写入
        worksheet.write_string(row, col, str(value), cell_format)


def write_batch_to_excel(worksheet, df, start_row, formats_dict=None):
    """
    将 DataFrame 逐行逐列安全写入 xlsxwriter 工作表。
    返回写入的行数。
    """
    if df.empty:
        return 0
    # 使用 .values 获取 numpy 数组，避免 iterrows 的性能问题，同时确保类型稳定
    rows = df.values.tolist()  # 转为 Python 原生 list of lists
    for i, row in enumerate(rows):
        excel_row = start_row + i
        for j, value in enumerate(row):
            cell_format = formats_dict.get(j) if formats_dict else None
            safe_write_cell(worksheet, excel_row, j, value, cell_format)
    return len(rows)


def generate_output_filename(prefix="export"):
    """生成带时间戳的文件名"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if not os.path.exists(EXPORT_DIR):
        os.makedirs(EXPORT_DIR)
    return os.path.join(EXPORT_DIR, f"{prefix}_{timestamp}.xlsx")


# ==================== 核心导出函数（流式） ====================

def export_large_data_to_excel(
    sql1,
    output_path,
    sql2=None,
    header_mapping1=None,
    header_mapping2=None,
    fillna_dict1=None,
    fillna_dict2=None,
    batch_size=10000,
    max_batches=200,
    sheet_name='Sheet1',
    default_col_width=20
):
    """
    流式导出大数据量查询到 Excel，支持双 SQL 同步分页拼接。
    内存友好，适用于百万级数据。
    """
    conn = None
    try:
        conn = pyodbc.connect(connStr)
        workbook = xlsxwriter.Workbook(output_path, {'constant_memory': True})
        worksheet = workbook.add_worksheet(sheet_name)

        dual = (sql2 is not None)
        offset = 0
        page_size = batch_size
        current_data_row = 1          # 0 行为表头
        header_written = False
        columns = []                  # 最终写入的列名列表

        for batch_num in range(1, max_batches + 1):
            # 1. 分页查询
            batch_sql1 = sql1.format(offset=offset, page_size=page_size)
            df1 = pd.read_sql(batch_sql1, conn)
            df2 = pd.DataFrame()
            if dual:
                batch_sql2 = sql2.format(offset=offset, page_size=page_size)
                df2 = pd.read_sql(batch_sql2, conn)
                if len(df1) != len(df2):
                    raise RuntimeError(
                        f"批次 {batch_num} 行数不一致：第一部分 {len(df1)}，第二部分 {len(df2)}。"
                    )

            # 2. 如果完全没有数据则退出
            if df1.empty and (not dual or df2.empty):
                break

            # 3. 处理第一批以确定表头
            if not header_written:
                df1_head = apply_mapping_and_fillna(df1.head(1).copy(), header_mapping1, fillna_dict1)
                df2_head = apply_mapping_and_fillna(df2.head(1).copy(), header_mapping2, fillna_dict2) if dual else pd.DataFrame()
                combined_head = pd.concat([df1_head, df2_head], axis=1)
                columns = list(combined_head.columns)
                # 写入表头
                for col_idx, col_name in enumerate(columns):
                    worksheet.write_string(0, col_idx, str(col_name))
                header_written = True

            # 4. 处理当前批次数据并写入
            if header_written:
                df1_proc = apply_mapping_and_fillna(df1.copy(), header_mapping1, fillna_dict1)
                df2_proc = apply_mapping_and_fillna(df2.copy(), header_mapping2, fillna_dict2) if dual else pd.DataFrame()

                if dual:
                    batch_combined = pd.concat([df1_proc, df2_proc], axis=1)
                else:
                    batch_combined = df1_proc

                # 确保列顺序与表头一致（处理列名重复或顺序不同）
                batch_combined = batch_combined[columns]
                rows_written = write_batch_to_excel(worksheet, batch_combined, current_data_row)
                current_data_row += rows_written

            # 如果不足一页，说明已是最后一批
            if len(df1) < page_size:
                break

            offset += page_size

        if not header_written:
            worksheet.write_string(0, 0, "无数据")

        # 设置统一列宽
        if columns:
            for col_idx in range(len(columns)):
                worksheet.set_column(col_idx, col_idx, default_col_width)

        workbook.close()
        total_rows = current_data_row - 1
        print(f"✅ 导出完成！共 {total_rows} 行数据，文件: {output_path}")
        return total_rows

    except Exception as e:
        print(f"❌ 导出失败: {e}")
        import traceback
        traceback.print_exc()
        return -1
    finally:
        if conn:
            conn.close()


# ==================== 主程序示例 ====================

if __name__ == "__main__":
    print("=" * 70)
    print("大数据量流式导出工具（修复 float has no len 问题）")
    print(f"开始时间: {datetime.now()}")
    print("=" * 70)

    sql_part1 = """
        SELECT
            BANK_IN_DATE,
            BANK_ACCOUNT_1,
            CREDIT_DEBIT,
            AMOUNT
        FROM
            test.manulife.From_Bank_Statement_CMU
        ORDER BY BANK_IN_DATE
        OFFSET {offset} ROWS 
        FETCH NEXT {page_size} ROWS ONLY
    """
    sql_part2 = """
        SELECT
            RECORD_DETAILS,
            ADDITIONAL_INFORMATION_1,
            FILE_NAME,
            FILE_DATE,
            VERSION
        FROM
            test.manulife.From_Bank_Statement_CMU
        ORDER BY BANK_IN_DATE
        OFFSET {offset} ROWS 
        FETCH NEXT {page_size} ROWS ONLY
    """

    mapping1 = {
        'BANK_IN_DATE': '入账日期',
        'BANK_ACCOUNT_1': '银行账户',
        'CREDIT_DEBIT': '借贷标识',
        'AMOUNT': '金额'
    }
    mapping2 = {
        'RECORD_DETAILS': '交易详情',
        'ADDITIONAL_INFORMATION_1': '补充信息',
        'FILE_NAME': '文件名',
        'FILE_DATE': '文件日期',
        'VERSION': '版本'
    }
    fillna1 = {
        'BANK_ACCOUNT_1': '未知账户',
        'AMOUNT': 0
    }
    fillna2 = {
        'RECORD_DETAILS': '无明细'
    }

    output_path = generate_output_filename("large_safe_export")

    export_large_data_to_excel(
        sql1=sql_part1,
        output_path=output_path,
        sql2=sql_part2,        # 若为 None 则只导出第一部分
        header_mapping1=mapping1,
        header_mapping2=mapping2,
        fillna_dict1=fillna1,
        fillna_dict2=fillna2,
        batch_size=10000,
        max_batches=500,
        sheet_name='数据导出',
        default_col_width=20
    )

    print("\n" + "=" * 70)
    print(f"完成时间: {datetime.now()}")
    print("=" * 70)
