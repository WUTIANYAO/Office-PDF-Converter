import os
import json

def parse_line(line):
    # 使用制表符拆分行
    fields = line.strip().split('\t')
    
    # 检查字段数量是否足够
    if len(fields) < 2:
        print(f"Warning: Line skipped due to insufficient fields: {line}")
        return None
    
    # 处理第一个字段，按照空格拆分
    sub_fields = fields[0].split()
    
    # 提取 subclass 和 smallclass
    subclass = sub_fields[0] if len(sub_fields) > 0 else ""
    smallclass = sub_fields[1] if len(sub_fields) > 1 else ""
    smallclass = smallclass.replace(":","/")
    
    # 处理 fi_code
    fi_code = ""
    if len(sub_fields) > 2:
        fi_code = sub_fields[2]
        if len(sub_fields) > 3 and sub_fields[3] != "\\":
            fi_code += sub_fields[3]
    
    # 提取描述字段
    ja_description = fields[-2] if len(fields) >= 2 else ""
    en_description = fields[-1] if len(fields) >= 1 else ""
    id = subclass+smallclass.replace("/","_")+fi_code.replace("\\","emptyfi")
    
    return {
        "id": id,
        "subclass": subclass,
        "smallclass": smallclass,
        "fi_code": fi_code,
        "ja_description": ja_description,
        "en_description": en_description
    }

def process_files(input_directory, output_file):
    data = []
    total_lines = 0  # 总行数计数器
    skipped_lines = 0  # 跳过的行数计数器

    # 遍历目录中的文件
    for root, dirs, files in os.walk(input_directory):
        for filename in files:
            file_path = os.path.join(root, filename)

            # 检查是否为文件且文件名长度大于9
            if os.path.isfile(file_path) and len(filename) > 9 and filename.endswith('.txt'):
                with open(file_path, 'r', encoding='EUC-JP') as file:
                    for line in file:
                        # 解析每行并添加到数据列表
                        parsed_data = parse_line(line)
                        if parsed_data:
                            data.append(parsed_data)
                            total_lines += 1
                        else:
                            skipped_lines += 1

    # 将数据写入 JSON 文件
    with open(output_file, 'w', encoding='utf-8') as json_file:
        json.dump(data, json_file, ensure_ascii=False, indent=4)

    # 打印总共转换的行数和跳过的行数
    print(f"Total lines processed: {total_lines}")
    print(f"Total lines skipped: {skipped_lines}")

if __name__ == "__main__":
    input_directory = r"C:\Users\Ken\Desktop\Test\IPCtest\data_20240808\data_fi"
    output_file = r"C:\Users\Ken\Desktop\Test\IPCtest\output_fi.json"
    process_files(input_directory, output_file)
