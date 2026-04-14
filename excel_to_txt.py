import pandas as pd
import os


def excel_to_wos_txt(input_excel_path, output_txt_path):
    print(f"正在读取清洗后的 Excel 文件: {input_excel_path} ...")

    # 强制将所有列读取为字符串格式，防止年份(PY)或卷号(VL)变成浮点数(如 2018.0)
    # 并将所有的空值(NaN)替换为空字符串
    try:
        df = pd.read_excel(input_excel_path, dtype=str)
        df = df.fillna('')
    except Exception as e:
        print(f"读取 Excel 文件失败: {e}")
        return

    print("正在转换为 WOS 原生 txt 格式 ...")

    with open(output_txt_path, 'w', encoding='utf-8') as f:
        # 1. 写入 WOS 文件的固定文件头（VOSviewer等软件识别文件类型的关键）
        f.write("FN Clarivate Analytics Web of Science\n")
        f.write("VR 1.0\n")

        # 获取所有列名
        columns = df.columns.tolist()

        # 2. 遍历每一行（每一篇文献）
        for index, row in df.iterrows():
            # 确保每条记录第一行是 PT
            if 'PT' in columns and str(row['PT']).strip() != '':
                f.write(f"PT {row['PT'].strip()}\n")

            # 遍历其他所有字段
            for col in columns:
                # 跳过 PT（已写完）和 ER（最后写），以及清洗数据时可能自己加的辅助列（限制WOS标签为2个字符）
                if col in ['PT', 'ER'] or len(col) != 2:
                    continue

                value = str(row[col]).strip()
                if value == '':  # 跳过没有内容的空字段
                    continue

                # 核心逻辑：处理换行内容（如多个作者、多个参考文献）
                # 按照前一个程序保留的换行符 \n 进行切割
                lines = value.split('\n')

                # 字段的第一行： 标签 + 1个空格 + 内容
                f.write(f"{col} {lines[0].strip()}\n")

                # 字段的延续行： 行首3个空格 + 内容 (严格符合WOS格式)
                if len(lines) > 1:
                    for line in lines[1:]:
                        if line.strip() != '':
                            f.write(f"   {line.strip()}\n")

            # 3. 写入每条文献的结束符 ER
            f.write("ER\n")
            # WOS原生格式中，相邻两篇文献之间通常有一个空行
            f.write("\n")

        # 4. 写入 WOS 文件的固定文件尾
        f.write("EF\n")

    print(f"转换成功！清洗后的原生文献库已保存至: {output_txt_path}")
    print(f"本次共成功转换 {len(df)} 条文献记录。现在你可以将它直接导入 VOSviewer 了！")


if __name__ == "__main__":
    # 指定刚才清洗完的 Excel 文件名，和希望输出的 txt 文件名
    INPUT_FILE = 'merged_wos_records.xlsx'  # 这里替换成你清洗后的excel文件名
    OUTPUT_FILE = 'wos_uncleaned.txt'  # 准备放进VOSviewer的txt文件名

    if os.path.exists(INPUT_FILE):
        excel_to_wos_txt(INPUT_FILE, OUTPUT_FILE)
    else:
        print(f"错误：找不到文件 '{INPUT_FILE}'，请检查文件名或路径。")