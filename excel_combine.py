import pandas as pd
import os
import glob


def merge_excel_files(input_folder, output_filepath, drop_duplicates=True):
    print(f"正在扫描文件夹: {input_folder} ...")

    # 查找文件夹下所有的 .xls 和 .xlsx 文件
    search_pattern = os.path.join(input_folder, "*.xls*")
    file_list = glob.glob(search_pattern)

    # 排除掉输出文件自身（如果输出文件也放在同一个文件夹内的话），防止无限嵌套读取
    file_list = [f for f in file_list if os.path.abspath(f) != os.path.abspath(output_filepath)]

    if not file_list:
        print("未找到任何 Excel 文件，请检查文件夹路径！")
        return

    df_list = []
    total_rows = 0

    # 逐个读取 Excel 文件
    for file in file_list:
        filename = os.path.basename(file)
        print(f"正在读取: {filename} ...")
        try:
            # 同样必须使用 dtype=str，防止WOS中的年份(PY)、卷期号等变成浮点数
            df = pd.read_excel(file, dtype=str)
            df_list.append(df)
            total_rows += len(df)
        except Exception as e:
            print(f"读取 {filename} 时发生错误: {e}")

    if not df_list:
        print("没有成功读取到任何数据。")
        return

    print("正在合并数据...")
    # 将所有读取到的 DataFrame 纵向拼接
    merged_df = pd.concat(df_list, ignore_index=True)

    # 将 NaN 替换为空字符串
    merged_df = merged_df.fillna('')

    # ================= 核心附加功能：根据WOS唯一标识符(UT)去重 =================
    if drop_duplicates and 'UT' in merged_df.columns:
        before_dedup = len(merged_df)
        # 根据 'UT' (Unique Tracking Number) 去重，保留第一次出现的记录
        merged_df = merged_df.drop_duplicates(subset=['UT'], keep='first')
        after_dedup = len(merged_df)
        dup_count = before_dedup - after_dedup
        print(f"执行去重：共发现并清除了 {dup_count} 条重复文献！")
    # =========================================================================

    # 整理列排序，确保 PT 在最前，ER 在最后 (如果存在的话)
    cols = merged_df.columns.tolist()
    if 'PT' in cols:
        cols.remove('PT')
        cols.insert(0, 'PT')
    if 'ER' in cols:
        cols.remove('ER')
        cols.append('ER')
    merged_df = merged_df[cols]

    print(f"正在保存合并后的文件至: {output_filepath} ...")
    merged_df.to_excel(output_filepath, index=False)

    print("-" * 30)
    print(f"✅ 合并完成！")
    print(f"共读取了 {len(file_list)} 个文件。")
    print(f"合并前总行数: {total_rows}")
    print(f"合并去重后最终保存行数: {len(merged_df)}")


if __name__ == "__main__":
    # ================= 配置区 =================
    # 存放需要合并的 excel 文件的文件夹路径（可以使用相对路径，如 '.' 代表当前文件夹）
    INPUT_FOLDER = './wos_excel_files'

    # 合并后生成的最终 Excel 文件名
    OUTPUT_FILE = 'merged_wos_records.xlsx'

    # 是否根据 UT(WOS入藏号) 自动去除重复文献？ True 为去重，False 为不去重
    AUTO_DEDUPLICATE = True
    # ==========================================

    # 如果文件夹不存在，自动创建（方便你把文件放进去）
    if not os.path.exists(INPUT_FOLDER):
        os.makedirs(INPUT_FOLDER)
        print(f"已自动创建文件夹 '{INPUT_FOLDER}'，请把你想要合并的 Excel 文件放进去后，重新运行本程序。")
    else:
        merge_excel_files(INPUT_FOLDER, OUTPUT_FILE, drop_duplicates=AUTO_DEDUPLICATE)