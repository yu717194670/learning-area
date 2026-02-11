import pandas as pd
import re
import os

EXCEL_PATH = r"Single_word.xlsx"

# 合并两份txt的所有内容（直接复制，避免路径问题）
TXT_ALL_CONTENT = """"""

def extract_equal_from_txt():
    """逐行解析txt，提取所有单词的“=”后内容"""
    equal_info = {}
    current_word = None  # 当前绑定的单词
    lines = TXT_ALL_CONTENT.split("\n")  # 按行分割
    
    print("📝 开始逐行提取（含“=”的行）：")
    print("-"*80)
    
    for line_num, line in enumerate(lines, 1):
        line = line.strip()
        # 跳过空行和版权行
        if not line or "西西（抖音xixi787）" in line or "Copyright" in line:
            continue
        
        # 定位单词行（格式：单词 [音标]）
        word_match = re.match(r"(\w+)\s+\[.+\]", line)
        if word_match:
            current_word = word_match.group(1).lower().strip()
            equal_info[current_word] = []  # 初始化该单词的Equal列表
            print(f"行{line_num}：找到单词 → {current_word}")
            continue
        
        # 提取“=”后内容（当前单词已定位，且行含“=”）
        if current_word and "=" in line and not line.startswith("联想"):
            # 截取“=”后内容，支持一行多个“=”
            parts = line.split("=")
            for part in parts[1:]:  # 从第二个元素开始（跳过“=”前内容）
                part = part.strip()
                if part:
                    equal_info[current_word].append(part)
                    print(f"行{line_num}：{current_word} → 提取“{part}”")
    
    # 整理结果：列表转换行字符串
    for word in equal_info:
        if equal_info[word]:
            equal_info[word] = "\n".join(list(set(equal_info[word])))  # 去重
        else:
            del equal_info[word]  # 删除无内容的单词
    
    print("-"*80)
    print(f"\n✅ 提取完成：共 {len(equal_info)} 个单词")
    print(f"提取的单词及Equal内容：")
    for word, content in equal_info.items():
        print(f"  {word}：\n{content}\n")
    return equal_info

def update_excel_equal(equal_info):
    """更新Excel的Equal列，精准匹配"""
    # 读取Excel
    df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
    df.columns = [col.strip().lower() for col in df.columns]
    # 清理English列，确保匹配
    df["english_clean"] = df["english"].astype(str).str.strip().str.lower()
    df_valid = df[(df["english_clean"] != "") & (df["english_clean"] != "nan")]
    
    updated_count = 0
    updated_words = []
    for idx, row in df_valid.iterrows():
        excel_word = row["english_clean"]
        if excel_word in equal_info:
            # 找到原始行并更新
            original_idx = df[df["english_clean"] == excel_word].index[0]
            df.loc[original_idx, "equal"] = equal_info[excel_word]
            updated_count += 1
            updated_words.append(excel_word)
    
    # 保存文件
    output_path = EXCEL_PATH.replace(".xlsx", "_ultimate_updated.xlsx")
    df = df.drop(columns=["english_clean"])
    df.to_excel(output_path, index=False, engine="openpyxl")
    
    # 打印最终结果
    print("="*80)
    print(f"📊 Excel更新结果：")
    print(f"   - Excel有效单词行数：{len(df_valid)}")
    print(f"   - TXT提取单词数：{len(equal_info)}")
    print(f"   - 成功更新Equal列：{updated_count} 个单词")
    print(f"   - 已更新单词：{', '.join(updated_words)}")
    print(f"✅ 最终文件：{output_path}")
    print("="*80)

def main():
    try:
        print("🚀 开始处理：逐行提取txt → 更新Excel Equal列")
        # 1. 提取txt中所有“=”后内容
        equal_info = extract_equal_from_txt()
        # 2. 更新Excel
        update_excel_equal(equal_info)
    except Exception as e:
        print(f"\n❌ 出错：{str(e)}")
        import traceback
        traceback.print_exc()
        input("按回车键退出...")

if __name__ == "__main__":
    # 自动安装依赖
    for pkg in ["pandas", "openpyxl"]:
        try:
            __import__(pkg)
        except ImportError:
            print(f"⚠️  安装依赖 {pkg}...")
            os.system(f"pip install {pkg}")
    main()