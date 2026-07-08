import os
import pyodbc
import pandas as pd
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
    """对DataFrame应用列头映射和基于原始字段的默认值填充"""
    if df.empty:
        return df
    # 1. 空值填充（原始字段名）
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


def write_batch_to_excel(worksheet, df, start_row, formats_dict=None):
    """
    将DataFrame写入xlsxwriter的worksheet指定起始行
    返回写入的行数
    """
    if df.empty:
        return 0
    # 如果还未写入表头，需在外部先写入
    for row_idx, row in df.iterrows():
        # row_idx是df的索引（0开始的相对索引），写入时加上start_row偏移
        excel_row = start_row + row_idx
        for col_idx, value in enumerate(row):
            # 尝试使用预设格式，否则默认
            cell_format = formats_dict.get(col_idx) if formats_dict else None
            worksheet.write(excel_row, col_idx, value, cell_format)
    return len(df)


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
    default_col_width=18       # 固定列宽，避免全量计算
):
    """
    流式导出大数据量查询到Excel，支持双SQL同步分页拼接
    """
    conn = None
    try:
        conn = pyodbc.connect(connStr)
        workbook = xlsxwriter.Workbook(output_path, {'constant_memory': True})
        worksheet = workbook.add_worksheet(sheet_name)

        # 双模式标志
        dual = (sql2 is not None)

        # 第一次先拉第一批数据，确定列头并写入
        offset = 0
        page_size = batch_size

        # 用于记录当前Excel已写入的行数（表头占第0行，数据从第1行开始）
        current_data_row = 1
        header_written = False
        columns = []

        for batch_num in range(1, max_batches + 1):
            # 查询第一部分
            batch_sql1 = sql1.format(offset=offset, page_size=page_size)
            df1 = pd.read_sql(batch_sql1, conn)
            df2 = pd.DataFrame()

            if dual:
                batch_sql2 = sql2.format(offset=offset, page_size=page_size)
                df2 = pd.read_sql(batch_sql2, conn)
                # 检查行数是否一致（关键：确保拼接正确）
                if len(df1) != len(df2):
                    raise RuntimeError(
                        f"批次 {batch_num} 两个查询行数不一致：第一部分 {len(df1)} 行，第二部分 {len(df2)} 行。"
                        f"请确保两个SQL返回相同行数且顺序一致。"
                    )

            # 如果第一批，处理表头
            if not header_written and (not df1.empty or (dual and not df2.empty)):
                # 应用映射和填充
                df1_processed = apply_mapping_and_fillna(df1.copy(), header_mapping1, fillna_dict1)
                df2_processed = apply_mapping_and_fillna(df2.copy(), header_mapping2, fillna_dict2) if dual else pd.DataFrame()

                # 拼接以获取最终列名
                if dual:
                    combined_head = pd.concat([df1_processed.head(0), df2_processed.head(0)], axis=1)
                else:
                    combined_head = df1_processed.head(0)
                columns = list(combined_head.columns)

                # 写入表头
                for col_idx, col_name in enumerate(columns):
                    worksheet.write(0, col_idx, col_name)
                header_written = True

            # 如果没有数据，退出循环
            if df1.empty and (not dual or df2.empty):
                break

            # 处理当前批次数据
            if header_written:
                # 应用映射和填充
                df1_proc = apply_mapping_and_fillna(df1.copy(), header_mapping1, fillna_dict1)
                df2_proc = apply_mapping_and_fillna(df2.copy(), header_mapping2, fillna_dict2) if dual else pd.DataFrame()

                if dual:
                    batch_combined = pd.concat([df1_proc, df2_proc], axis=1)
                else:
                    batch_combined = df1_proc

                # 确保列顺序与表头一致（避免因映射导致顺序问题）
                batch_combined = batch_combined[columns]

                # 写入数据
                rows_written = write_batch_to_excel(
                    worksheet, batch_combined, current_data_row, formats_dict=None
                )
                current_data_row += rows_written

            # 判断是否为最后一批
            if len(df1) < page_size:
                break

            offset += page_size

        # 如果循环结束仍未写入任何头，说明全为空
        if not header_written:
            # 写入一个空的表头占位，避免文件损坏
            worksheet.write(0, 0, "无数据")

        # 设置列宽（统一固定宽度，简单高效）
        if columns:
            for col_idx in range(len(columns)):
                worksheet.set_column(col_idx, col_idx, default_col_width)

        workbook.close()
        conn.close()
        conn = None

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
    print("大数据量流式导出工具（内存友好）")
    print(f"开始时间: {datetime.now()}")
    print("=" * 70)

    # SQL 示例（必须包含 {offset} 和 {page_size} 占位符，且两SQL行对应）
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

    # 列头映射（原始字段 -> 显示名称）
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

    # 基于原始字段名的空值填充
    fillna1 = {
        'BANK_ACCOUNT_1': '未知账户',
        'AMOUNT': 0
    }
    fillna2 = {
        'RECORD_DETAILS': '无明细'
    }

    output_path = generate_output_filename("large_export")

    # 单SQL导出模式：将 sql2 设为 None
    export_large_data_to_excel(
        sql1=sql_part1,
        output_path=output_path,
        sql2=sql_part2,              # 若为 None 则仅导出第一部分
        header_mapping1=mapping1,
        header_mapping2=mapping2,
        fillna_dict1=fillna1,
        fillna_dict2=fillna2,
        batch_size=10000,            # 每批1万行，可按内存调整
        max_batches=500,             # 最大批次数，防止死循环
        sheet_name='数据导出',
        default_col_width=20
    )

    print("\n" + "=" * 70)
    print(f"完成时间: {datetime.now()}")
    print("=" * 70)
