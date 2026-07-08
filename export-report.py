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

def apply_mapping_and_fillna(df, header_mapping, fillna_dict, tag="第一部分"):
    """
    基于原始字段名填充空值，再应用列头映射。
    返回处理后的DataFrame，以及可能产生的重复列名处理映射。
    """
    if df.empty:
        return df, {}

    # 1. 空值填充（基于原始字段名）
    if fillna_dict:
        # 精确匹配，键必须与 df.columns 完全一致（区分大小写）
        valid_fill = {}
        for col, val in fillna_dict.items():
            if col in df.columns:
                valid_fill[col] = val
            else:
                print(f"⚠️ [{tag}] 填充键 '{col}' 在SQL结果中不存在，已忽略。可用字段: {list(df.columns)}")
        if valid_fill:
            print(f"[{tag}] 应用空值填充: {valid_fill}")
            df = df.fillna(valid_fill)

    # 2. 列头映射（保留未映射的字段原名）
    if header_mapping:
        # 只映射存在的列
        rename_map = {k: v for k, v in header_mapping.items() if k in df.columns}
        if rename_map:
            df = df.rename(columns=rename_map)

    # 3. 处理映射后可能出现的重复列名（自动添加后缀避免混乱）
    cols = df.columns.tolist()
    seen = {}
    new_cols = []
    for col in cols:
        if col in seen:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            new_cols.append(col)
    if new_cols != cols:
        df.columns = new_cols
        print(f"[{tag}] 检测到重复列名，已自动重命名: {dict(zip(cols, new_cols))}")

    return df, {}


def safe_write_cell(worksheet, row, col, value, cell_format=None):
    """安全写入单元格，避免 'float has no len()' 错误"""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        worksheet.write_blank(row, col, None, cell_format)
        return
    if isinstance(value, (np.integer,)):
        value = int(value)
    elif isinstance(value, (np.floating,)):
        value = float(value)
    if isinstance(value, bool):
        worksheet.write_boolean(row, col, value, cell_format)
    elif isinstance(value, (int, float, np.number)):
        worksheet.write_number(row, col, value, cell_format)
    elif isinstance(value, str):
        worksheet.write_string(row, col, value, cell_format)
    elif isinstance(value, pd.Timestamp):
        worksheet.write_datetime(row, col, value.to_pydatetime(), cell_format)
    else:
        worksheet.write_string(row, col, str(value), cell_format)


def write_batch_to_excel(worksheet, df, start_row, formats_dict=None):
    """流式写入批次数据"""
    if df.empty:
        return 0
    rows = df.values.tolist()
    for i, row in enumerate(rows):
        excel_row = start_row + i
        for j, value in enumerate(row):
            cell_format = formats_dict.get(j) if formats_dict else None
            safe_write_cell(worksheet, excel_row, j, value, cell_format)
    return len(rows)


def generate_output_filename(prefix="export"):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if not os.path.exists(EXPORT_DIR):
        os.makedirs(EXPORT_DIR)
    return os.path.join(EXPORT_DIR, f"{prefix}_{timestamp}.xlsx")


# ==================== 核心流式导出 ====================

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
    conn = None
    try:
        conn = pyodbc.connect(connStr)
        workbook = xlsxwriter.Workbook(output_path, {'constant_memory': True})
        worksheet = workbook.add_worksheet(sheet_name)

        dual = (sql2 is not None)
        offset = 0
        page_size = batch_size
        current_data_row = 1
        header_written = False
        columns = []

        for batch_num in range(1, max_batches + 1):
            # 分页查询
            batch_sql1 = sql1.format(offset=offset, page_size=page_size)
            df1 = pd.read_sql(batch_sql1, conn)
            df2 = pd.DataFrame()
            if dual:
                batch_sql2 = sql2.format(offset=offset, page_size=page_size)
                df2 = pd.read_sql(batch_sql2, conn)
                if len(df1) != len(df2):
                    raise RuntimeError(f"批次 {batch_num} 行数不一致：第一部分 {len(df1)}，第二部分 {len(df2)}。")

            if df1.empty and (not dual or df2.empty):
                break

            # ---------- 处理第一批以确定表头 ----------
            if not header_written and (not df1.empty or (dual and not df2.empty)):
                # 处理第一部分表头
                df1_head, _ = apply_mapping_and_fillna(df1.head(1).copy(), header_mapping1, fillna_dict1, "第一部分")
                # 处理第二部分表头（如果存在）
                df2_head = pd.DataFrame()
                if dual:
                    df2_head, _ = apply_mapping_and_fillna(df2.head(1).copy(), header_mapping2, fillna_dict2, "第二部分")
                combined_head = pd.concat([df1_head, df2_head], axis=1)
                columns = list(combined_head.columns)
                # 写入表头
                for col_idx, col_name in enumerate(columns):
                    worksheet.write_string(0, col_idx, str(col_name))
                header_written = True
                print(f"表头已写入，共 {len(columns)} 列: {columns}")

            # ---------- 写入当前批次数据 ----------
            if header_written:
                df1_proc, _ = apply_mapping_and_fillna(df1.copy(), header_mapping1, fillna_dict1, "第一部分")
                df2_proc = pd.DataFrame()
                if dual:
                    df2_proc, _ = apply_mapping_and_fillna(df2.copy(), header_mapping2, fillna_dict2, "第二部分")

                if dual:
                    batch_combined = pd.concat([df1_proc, df2_proc], axis=1)
                else:
                    batch_combined = df1_proc

                # 严格按表头列顺序和数量对齐，防止意外多列/少列
                try:
                    batch_combined = batch_combined.reindex(columns=columns, fill_value=None)
                except KeyError as e:
                    print(f"列对齐失败，可能映射有误: {e}")
                    raise

                rows_written = write_batch_to_excel(worksheet, batch_combined, current_data_row)
                current_data_row += rows_written

            if len(df1) < page_size:
                break
            offset += page_size

        if not header_written:
            worksheet.write_string(0, 0, "无数据")
        else:
            # 设置列宽
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


# ==================== 示例 ====================
if __name__ == "__main__":
    print("=" * 70)
    print("大数据流式导出工具（自动处理重复列名，精确匹配填充）")
    print(f"开始时间: {datetime.now()}")
    print("=" * 70)

    sql_part1 = """
        SELECT
            BANK_IN_DATE,
            BANK_ACCOUNT_1,
            CREDIT_DEBIT,
            AMOUNT
        FROM test.manulife.From_Bank_Statement_CMU
        ORDER BY BANK_IN_DATE
        OFFSET {offset} ROWS FETCH NEXT {page_size} ROWS ONLY
    """
    # 单表导出：注释掉或设置为 None
    sql_part2 = None

    mapping1 = {
        'BANK_IN_DATE': '入账日期',
        'BANK_ACCOUNT_1': '银行账户',
        'CREDIT_DEBIT': '借贷标识',
        'AMOUNT': '金额'
    }
    # fillna_dict 键必须是 SQL 返回的原始字段名，严格一致
    fillna1 = {
        'BANK_ACCOUNT_1': '未知账户',
        'AMOUNT': 0
    }

    output_path = generate_output_filename("single_export")

    export_large_data_to_excel(
        sql1=sql_part1,
        output_path=output_path,
        sql2=sql_part2,               # 单表模式
        header_mapping1=mapping1,
        fillna_dict1=fillna1,
        batch_size=10000,
        sheet_name='单表数据',
        default_col_width=20
    )

    print("\n" + "=" * 70)
    print(f"完成时间: {datetime.now()}")
    print("=" * 70)
