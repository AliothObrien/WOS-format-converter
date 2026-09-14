import pandas as pd
import os


def excel_to_wos_txt(input_excel_path, output_txt_path):
    print(f"正在读取清洗后的 Excel 文件: {input_excel_path} ...")

    try:
        df = pd.read_excel(input_excel_path, dtype=str)
        df = df.fillna('')
    except Exception as e:
        print(f"读取 Excel 文件失败: {e}")
        return

    print("正在转换为 WOS 原生 txt 格式 ...")

    with open(output_txt_path, 'w', encoding='utf-8') as f:
        # 1. 写入 WOS 文件的固定文件头
        f.write("FN Clarivate Analytics Web of Science\n")
        f.write("VR 1.0\n")
        columns = df.columns.tolist()

        # 2. 遍历
        for index, row in df.iterrows():
            if 'PT' in columns and str(row['PT']).strip() != '':
                f.write(f"PT {row['PT'].strip()}\n")
            for col in columns:
                if col in ['PT', 'ER'] or len(col) != 2:
                    continue

                value = str(row[col]).strip()
                if value == '':
                    continue

                lines = value.split('\n')

                f.write(f"{col} {lines[0].strip()}\n")

                if len(lines) > 1:
                    for line in lines[1:]:
                        if line.strip() != '':
                            f.write(f"   {line.strip()}\n")

            # 3. 写入每条文献的结束符 ER
            f.write("ER\n")

            f.write("\n")

        # 4. 写入 WOS 文件的固定文件尾
        f.write("EF\n")

    print(f"转换成功！清洗后的原生文献库已保存至: {output_txt_path}")
    print(f"本次共成功转换 {len(df)} 条文献记录。现在你可以将它直接导入 VOSviewer 了！")


if __name__ == "__main__":
    # 指定刚才清洗完的 Excel 文件名，和希望输出的 txt 文件名
    INPUT_FILE = 'merged_wos_records.xlsx'  # 替换成你清洗后的excel文件名
    OUTPUT_FILE = 'wos_uncleaned.txt'  # 准备放进VOSviewer的txt文件名

    if os.path.exists(INPUT_FILE):
        excel_to_wos_txt(INPUT_FILE, OUTPUT_FILE)
    else:
        print(f"错误：找不到文件 '{INPUT_FILE}'，请检查文件名或路径。")
