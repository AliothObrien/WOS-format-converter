import pandas as pd
import os


def wos_txt_to_excel(input_txt_path, output_excel_path):
    records = []  # 存放所有文献记录
    current_record = {}  # 存放当前正在解析的单条文献
    current_tag = ""  # 当前正在解析的字段标签

    print(f"正在读取文件: {input_txt_path} ...")

    with open(input_txt_path, 'r', encoding='utf-8-sig') as f:
        for line in f:
            # 去除行末的回车换行符，但保留行首的空格以判断是否为延续行
            line = line.rstrip('\n')

            # 跳过空行（如果没有处在任何文献记录解析中）
            if not line.strip() and not current_record:
                continue

            # 跳过WOS文件的开头标识
            if line.startswith("FN Clarivate") or line.startswith("VR "):
                continue

            # WOS文件的结束标识
            if line == "EF":
                break

            # 'ER' 标识一条文献记录的结束
            if line == "ER":
                current_record['ER'] = ''  # 满足需求：以 ER 作为列标题
                records.append(current_record)
                current_record = {}
                current_tag = ""
                continue

            # 判断是否为多行延续的内容
            if line.startswith("   ") and current_tag:
                # 遇到多行内容，使用换行符 '\n' 拼接，完美保留各字段下的独立内容（如多作者、多参考文献）
                current_record[current_tag] += "\n" + line.strip()

            # 判断是否为新的字段标签
            elif len(line) >= 3 and line[2] == " " and line[0].isupper() and line[1].isalnum():
                tag = line[:2]
                value = line[3:]
                current_tag = tag
                # 记录该标签对应的内容
                current_record[current_tag] = value
            else:
                if current_tag:
                    current_record[current_tag] += " " + line.strip()

    # 转换为 pandas DataFrame
    df = pd.DataFrame(records)

    if df.empty:
        print("未解析到任何数据，请检查 txt 文件格式！")
        return

    cols = df.columns.tolist()
    if 'PT' in cols:
        cols.remove('PT')
        cols.insert(0, 'PT')
    if 'ER' in cols:
        cols.remove('ER')
        cols.append('ER')

    df = df[cols]

    # 导出为 Excel 文件 (xlsx格式)
    print(f"正在生成 Excel 文件: {output_excel_path} ...")
    df.to_excel(output_excel_path, index=False)
    print(f"转换成功！共解析了 {len(records)} 条文献记录。")


if __name__ == "__main__":
    # 指定输入输出文件路径
    INPUT_FILE = 'savedrecs (5).txt'
    OUTPUT_FILE = 'savedrecs_converted_5.xlsx'

    if os.path.exists(INPUT_FILE):
        wos_txt_to_excel(INPUT_FILE, OUTPUT_FILE)
    else:
        print(f"错误：找不到文件 '{INPUT_FILE}'，请确保文件与脚本在同一目录下。")
