import os
import pyodbc
import pandas as pd
from datetime import datetime
from openpyxl.utils import get_column_letter

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

def fetch_paginated_data(sql, batch_size, max_batches, conn):
    """
    循环分页查询，返回合并后的DataFrame
    参数:
        sql: 带 {offset} 和 {page_size} 占位符的SQL
        batch_size: 每批行数
        max_batches: 最大循环次数
        conn: 数据库连接
    返回:
        DataFrame，若无数据则返回空DataFrame
    """
    df_list = []
    offset = 0
    for batch_num in range(1, max_batches + 1):
        batch_sql = sql.format(offset=offset, page_size=batch_size)
        batch_df = pd.read_sql(batch_sql, conn)
        row_count = len(batch_df)
        if row_count == 0:
            break
        df_list.append(batch_df)
        if row_count < batch_size:   # 尾页检查：本次行数小于批次大小，结束
            break
        offset += batch_size

    if df_list:
        return pd.concat(df_list, ignore_index=True)
    else:
        return pd.DataFrame()


def adjust_column_width(worksheet, df):
    """自动调整Excel列宽"""
    for col_idx, column in enumerate(df.columns, start=1):
        max_len = max(df[column].astype(str).map(len).max(), len(str(column))) + 2
        max_len = min(max_len, 50)  # 限制最大宽度
        col_letter = get_column_letter(col_idx)
        worksheet.column_dimensions[col_letter].width = max_len


# ==================== 核心导出函数 ====================

def export_two_parts_to_excel(
    sql1, sql2,
    output_path,
    header_mapping1=None,
    header_mapping2=None,
    batch_size1=1000,
    batch_size2=1000,
    max_batches=100,
    sheet_name='Sheet1'
):
    """
    导出两部分查询数据到同一个Excel Sheet，第一部分在左，第二部分紧挨着右侧

    参数:
        sql1: 第一部分SQL，须包含 {offset} 和 {page_size} 占位符
        sql2: 第二部分SQL，同样须包含分页占位符
        output_path: 输出Excel文件路径
        header_mapping1: 第一部分列名映射字典 {原字段名: 显示名称}
        header_mapping2: 第二部分列名映射字典
        batch_size1: 第一部分每批行数
        batch_size2: 第二部分每批行数
        max_batches: 最大循环批次数（防止死循环）
        sheet_name: Excel工作表名称
    """
    conn = None
    try:
        conn = pyodbc.connect(connStr)

        print("开始分批查询第一部分...")
        df1 = fetch_paginated_data(sql1, batch_size1, max_batches, conn)
        print(f"第一部分共 {len(df1)} 行，{len(df1.columns)} 列")

        print("开始分批查询第二部分...")
        df2 = fetch_paginated_data(sql2, batch_size2, max_batches, conn)
        print(f"第二部分共 {len(df2)} 行，{len(df2.columns)} 列")

        # 应用列头映射
        if header_mapping1 and not df1.empty:
            valid_map1 = {k: v for k, v in header_mapping1.items() if k in df1.columns}
            if valid_map1:
                df1 = df1.rename(columns=valid_map1)

        if header_mapping2 and not df2.empty:
            valid_map2 = {k: v for k, v in header_mapping2.items() if k in df2.columns}
            if valid_map2:
                df2 = df2.rename(columns=valid_map2)

        # 横向拼接（按索引对齐，行数不同会自动填充NaN）
        if df1.empty and df2.empty:
            print("⚠️ 两部分查询结果均为空，无数据导出")
            return 0

        df_combined = pd.concat([df1, df2], axis=1)
        print(f"合并后总列数: {len(df_combined.columns)}")

        # 写入Excel
        print(f"正在写入 {output_path} ...")
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_combined.to_excel(writer, sheet_name=sheet_name, index=False)
            adjust_column_width(writer.sheets[sheet_name], df_combined)

        print(f"✅ 导出成功！保存至: {output_path}")
        return len(df_combined)

    except Exception as e:
        print(f"❌ 导出失败: {e}")
        import traceback
        traceback.print_exc()
        return -1
    finally:
        if conn:
            conn.close()


# ==================== 便捷生成文件名 ====================

def generate_output_filename(prefix="export"):
    """生成带时间戳的文件名"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if not os.path.exists(EXPORT_DIR):
        os.makedirs(EXPORT_DIR)
    return os.path.join(EXPORT_DIR, f"{prefix}_{timestamp}.xlsx")


# ==================== 主程序示例 ====================

if __name__ == "__main__":
    print("=" * 70)
    print("两部分查询横向导出工具（支持分页）")
    print(f"开始时间: {datetime.now()}")
    print("=" * 70)

    # 示例SQL - 第一部分（必须含 {offset} 和 {page_size}）
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

    # 示例SQL - 第二部分
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

    # 列头映射（第一部分）
    mapping1 = {
        'BANK_IN_DATE': '入账日期',
        'BANK_ACCOUNT_1': '银行账户',
        'CREDIT_DEBIT': '借贷标识',
        'AMOUNT': '金额'
    }

    # 列头映射（第二部分）
    mapping2 = None
    # mapping2 = {
    #     'RECORD_DETAILS': '交易详情',
    #     'ADDITIONAL_INFORMATION_1': '补充信息',
    #     'FILE_NAME': '文件名',
    #     'FILE_DATE': '文件日期',
    #     'VERSION': '版本'
    # }

    # 生成输出路径
    output_path = generate_output_filename("two_parts_export")

    # 调用导出函数
    export_two_parts_to_excel(
        sql1=sql_part1,
        sql2=sql_part2,
        output_path=output_path,
        header_mapping1=mapping1,
        header_mapping2=mapping2,
        batch_size1=10000,   # 第一部分每批1万行
        batch_size2=10000,   # 第二部分每批1万行
        max_batches=200,
        sheet_name='合并数据'
    )

    print("\n" + "=" * 70)
    print(f"完成时间: {datetime.now()}")
    print("=" * 70) 