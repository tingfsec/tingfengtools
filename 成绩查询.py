from colorama import Fore, Style, init
from art import text2art
import random
import openpyxl
import requests
import os

init(autoreset=True)


def random_color():
    colors = [Fore.RED, Fore.GREEN, Fore.YELLOW, Fore.BLUE, Fore.MAGENTA, Fore.CYAN, Fore.WHITE]
    return random.choice(colors)


def display_header():
    """显示工具名字和版权信息"""
    print(random_color() + Style.BRIGHT + text2art("Ting Feng Tools"))
    print(random_color() + Style.BRIGHT + "版权所有 © 2025 听风网络安全实验室 All Rights Reserved")
    print(random_color() + Style.BRIGHT + "Qq：2262937477")
    print(random_color() + "支持单人查询和 Excel 批量查询\n")


def extract_data_from_excel(file_path, name_col, sfz_col, start_row):
    """根据用户输入的列和起始行提取 Excel 数据"""
    if not os.path.exists(file_path):
        print(Fore.RED + f"文件不存在：{file_path}")
        return []

    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active

    def col_to_index(col):
        return ord(col.upper()) - 65

    name_idx = col_to_index(name_col)
    sfz_idx = col_to_index(sfz_col)

    data_list = []
    for row in sheet.iter_rows(min_row=start_row):
        name = row[name_idx].value
        sfz = row[sfz_idx].value
        if name and sfz:
            data_list.append((name, sfz))
    return data_list


def query(name, sfz):
    """查询单个人的四六级成绩"""
    url = "https://cachecloud.neea.cn/latest/results/cet"
    params = {"km": "1", "xm": name, "no": sfz, "source": "pc"}
    headers = {
        "accept": "*/*",
        "origin": "https://cjcx.neea.edu.cn",
        "referer": "https://cjcx.neea.edu.cn/",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }

    try:
        response = requests.get(url, headers=headers, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()
        if data.get("code") == 0:
            score = int(data.get("score", 0))
            result = "通过✅" if score >= 425 else "未通过❌"
            color_result = Fore.GREEN if score >= 425 else Fore.RED
            print(f"{name} ({sfz}) -> {color_result}{result} 分数: {score}")
            return {"name": name, "sfz": sfz, "score": score, "result": result}
        else:
            print(Fore.RED + f"{name} ({sfz}) 查询失败")
            return {"name": name, "sfz": sfz, "score": None, "result": "查询失败"}
    except Exception as e:
        print(Fore.RED + f"{name} ({sfz}) 查询异常: {e}")
        return {"name": name, "sfz": sfz, "score": None, "result": f"查询异常: {e}"}


def batch_query(data_list):
    results = []
    for name, sfz in data_list:
        result = query(name, sfz)
        results.append(result)
    return results


def main():
    display_header()

    mode = input("请选择模式（1-单人查询，2-Excel批量查询）：").strip()

    if mode == "1":
        name = input("请输入姓名: ").strip()
        sfz = input("请输入身份证号: ").strip()
        query(name, sfz)
    elif mode == "2":
        file_path = input("请输入 Excel 文件路径: ").strip()
        name_col = input("请输入姓名所在列 (如 C): ").strip()
        sfz_col = input("请输入身份证所在列 (如 E): ").strip()
        start_row = int(input("请输入起始行号 (如 2): ").strip())
        data_list = extract_data_from_excel(file_path, name_col, sfz_col, start_row)
        if not data_list:
            print(Fore.RED + "没有数据可查询！")
            return
        results = batch_query(data_list)

        save_file = input("是否保存结果到文件？(y/n)：").strip().lower()
        if save_file == "y":
            sort_choice = input("是否按分数排序？(y/n)：").strip().lower()
            if sort_choice == "y":
                results = sorted(results, key=lambda x: x["score"] or 0, reverse=True)
            output_file = "results.txt"
            with open(output_file, "w", encoding="utf-8") as f:
                f.write("姓名,身份证,成绩,结果\n")
                for r in results:
                    f.write(f"{r['name']},{r['sfz']},{r['score']},{r['result']}\n")
            print(Fore.GREEN + f"结果已保存到 {output_file}")
    else:
        print(Fore.RED + "无效选择！")


if __name__ == "__main__":
    main()
