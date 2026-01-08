import pandas as pd
import random
import json
import os
import time
from datetime import datetime

class EnglishSpellingChecker:
    def __init__(self):
        self.COLORS = {
            'reset': '\033[0m',
            'red': '\033[31m',
            'green': '\033[32m',
            'yellow': '\033[33m',
            'blue': '\033[34m',
            'purple': '\033[35m',
            'cyan': '\033[36m',
            'white': '\033[37m',
            'bold': '\033[1m',
            'underline': '\033[4m'
        }
        
        self.file_path = r"D:\Data\Single_word.xlsx"
        self.records_file = "learning_records.json"
        
        self.words_data = self._load_and_prepare_data()
        self.learning_records = self._load_learning_records()
        
    def _load_and_prepare_data(self):
        """加载并预处理Excel表格数据"""
        try:
            if not os.path.exists(self.file_path):
                raise FileNotFoundError(f"文件不存在: {self.file_path}")
            
            df = pd.read_excel(self.file_path, engine='openpyxl')
            
            df.columns = [col.strip() for col in df.columns]
            df.columns = [col.lower() for col in df.columns]
            
            required_columns = ['english', 'chinese', 'example']
            available_columns = list(df.columns)
            
            column_mapping = {}
            for req_col in required_columns:
                if req_col in available_columns:
                    column_mapping[req_col] = req_col
                else:
                    for avail_col in available_columns:
                        if avail_col.lower() == req_col:
                            column_mapping[req_col] = avail_col
                            break
            
            new_df = pd.DataFrame()
            for col in ['english', 'chinese', 'example']:
                if col in column_mapping:
                    new_df[col] = df[column_mapping[col]]
                else:
                    new_df[col] = ""
                    print(f"{self.COLORS['yellow']}⚠️  警告：未找到'{col}'列，创建空列{self.COLORS['reset']}")
            
            new_df.columns = ['English', 'Chinese', 'Example']
            
            new_df['English'] = new_df['English'].astype(str).str.strip().str.lower()
            new_df['Chinese'] = new_df['Chinese'].astype(str).str.strip()
            new_df['Example'] = new_df['Example'].astype(str).str.strip()
            
            new_df = new_df[new_df['English'] != 'nuil']
            new_df = new_df[new_df['English'] != '']
            new_df = new_df[new_df['English'].str.len() > 0]
            
            words_data = new_df.to_dict('records')
            
            seen = set()
            unique_words = []
            for word in words_data:
                if word['English'] not in seen:
                    seen.add(word['English'])
                    unique_words.append(word)
            
            print(f"{self.COLORS['green']}✓ 成功加载 {len(unique_words)} 个英语单词（去重后）{self.COLORS['reset']}")
            return unique_words
            
        except FileNotFoundError as e:
            print(f"{self.COLORS['red']}✗ 文件未找到: {str(e)}{self.COLORS['reset']}")
            print(f"{self.COLORS['yellow']}💡 请检查：{self.COLORS['reset']}")
            print(f"   1. 文件路径是否正确")
            print(f"   2. 文件名是否正确（区分大小写）")
            print(f"   3. 文件是否被其他程序占用")
        except Exception as e:
            print(f"{self.COLORS['red']}✗ 加载数据失败: {type(e).__name__}: {str(e)}{self.COLORS['reset']}")
        
        print(f"{self.COLORS['yellow']}⚠️ 使用备用单词数据继续运行{self.COLORS['reset']}")
        backup_words = [
            {"English": "ride", "Example": "", "Chinese": "骑行"},
            {"English": "sand", "Example": "", "Chinese": "沙滩"},
            {"English": "bed", "Example": "", "Chinese": "床"},
            {"English": "door", "Example": "", "Chinese": "门"},
            {"English": "mad", "Example": "", "Chinese": "疯狂"},
            {"English": "bag", "Example": "", "Chinese": "包"},
            {"English": "bar", "Example": "爸爸在草地上开了个酒吧", "Chinese": "酒吧"},
            {"English": "air", "Example": "矮人在呼吸新鲜空气", "Chinese": "空气"},
            {"English": "go", "Example": "小狗喜欢到外边到处走", "Chinese": "走"},
            {"English": "leg", "Example": "可乐洒在了哥哥的腿上", "Chinese": "腿"}
        ]
        return backup_words
    
    def _load_learning_records(self):
        """加载学习记录（兼容旧版本）"""
        default_records = {
            "words": {},
            "stats": {
                "total_practiced": 0,
                "total_correct": 0,
                "last_practice": "",
                "quiz_records": []  # 新增测验记录
            }
        }
        
        if os.path.exists(self.records_file):
            try:
                with open(self.records_file, 'r', encoding='utf-8') as f:
                    loaded_records = json.load(f)
                
                # 兼容旧版本数据 - 确保所有字段都存在
                if 'stats' not in loaded_records:
                    loaded_records['stats'] = default_records['stats']
                else:
                    # 检查stats中的字段
                    for key, value in default_records['stats'].items():
                        if key not in loaded_records['stats']:
                            loaded_records['stats'][key] = value
                
                # 确保words字段存在
                if 'words' not in loaded_records:
                    loaded_records['words'] = {}
                
                return loaded_records
                
            except Exception as e:
                print(f"{self.COLORS['yellow']}⚠️ 学习记录文件损坏或格式不兼容，将创建新记录：{str(e)}{self.COLORS['reset']}")
        
        return default_records
    
    def _save_learning_records(self):
        """保存学习记录"""
        try:
            self.learning_records['stats']['last_practice'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(self.records_file, 'w', encoding='utf-8') as f:
                json.dump(self.learning_records, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"{self.COLORS['red']}✗ 保存学习记录失败: {str(e)}{self.COLORS['reset']}")
            return False
    
    def _get_example_hint(self, word_data):
        """获取Example列的提示内容"""
        example_hint = word_data['Example'].strip()
        if example_hint and example_hint != "":
            return f"{self.COLORS['cyan']}提示：{example_hint}{self.COLORS['reset']}"
        else:
            # 如果没有例句，显示备用提示（首字母+字母数）
            english_word = word_data['English'].lower()
            return f"{self.COLORS['cyan']}提示：单词以 '{english_word[0]}' 开头，共 {len(english_word)} 个字母{self.COLORS['reset']}"
    
    def _get_correct_spelling(self, word_data):
        """获取正确的拼写"""
        english_word = word_data['English'].lower()
        return f"{self.COLORS['green']}正确拼写：{self.COLORS['bold']}{english_word}{self.COLORS['reset']}"
    
    def _update_learning_record(self, english_word, is_correct):
        """更新单词学习记录"""
        if english_word not in self.learning_records['words']:
            self.learning_records['words'][english_word] = {
                "attempts": 0,
                "correct": 0,
                "last_attempt": "",
                "success_rate": 0.0
            }
        
        record = self.learning_records['words'][english_word]
        record['attempts'] += 1
        if is_correct:
            record['correct'] += 1
        
        record['success_rate'] = round(record['correct'] / record['attempts'] * 100, 1)
        record['last_attempt'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        self.learning_records['stats']['total_practiced'] += 1
        if is_correct:
            self.learning_records['stats']['total_correct'] += 1
    
    def _get_difficult_words(self, count=10):
        """获取需要加强练习的单词"""
        difficult_words = []
        
        for word, record in self.learning_records['words'].items():
            if record['success_rate'] < 70 and record['attempts'] > 0:
                for wd in self.words_data:
                    if wd['English'].lower() == word.lower():
                        difficult_words.append(wd)
                        break
        
        if len(difficult_words) < count:
            all_words_copy = [wd for wd in self.words_data if wd not in difficult_words]
            random.shuffle(all_words_copy)
            difficult_words.extend(all_words_copy[:count - len(difficult_words)])
        
        return difficult_words[:count]
    
    def check_spelling(self, word_data, quiz_mode=False):
        """
        检查单个单词的拼写
        :param word_data: 单词数据
        :param quiz_mode: 是否为测验模式（测验模式规则：
                          1. 第一次错误显示Example提示
                          2. 第二次错误显示正确拼写
                          3. 共2次机会）
        :return: 是否正确
        """
        english_word = word_data['English'].lower()
        chinese_meaning = word_data['Chinese']
        example = word_data['Example']
        
        print(f"\n{self.COLORS['blue']}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{self.COLORS['reset']}")
        print(f"{self.COLORS['bold']}请拼写单词：{self.COLORS['yellow']}{chinese_meaning}{self.COLORS['reset']}")
        # 测验模式下，首次不显示例句，普通模式正常显示
        if example and not quiz_mode:
            print(f"{self.COLORS['cyan']}例句：{example}{self.COLORS['reset']}")
        
        attempts = 0
        is_correct = False
        
        # 区分测验模式和普通练习模式的规则
        if quiz_mode:
            max_attempts = 2  # 测验模式2次机会
        else:
            max_attempts = 3  # 普通模式3次机会
        
        while attempts < max_attempts and not is_correct:
            user_input = input(f"\n{self.COLORS['purple']}请输入拼写: {self.COLORS['reset']}").strip().lower()
            
            attempts += 1
            
            if user_input == english_word:
                is_correct = True
                print(f"{self.COLORS['green']}✓ 恭喜！拼写正确！{self.COLORS['reset']}")
            else:
                if attempts < max_attempts:
                    print(f"{self.COLORS['red']}✗ 拼写错误！{self.COLORS['reset']}")
                    # 测验模式下：第一次错误就显示Example提示
                    if quiz_mode:
                        print(self._get_example_hint(word_data))
                        print(f"{self.COLORS['yellow']}⚠️  还有1次尝试机会{self.COLORS['reset']}")
                    else:
                        # 普通模式保持原有提示逻辑
                        english_word = word_data['English'].lower()
                        if attempts == 1:
                            hint = f"{self.COLORS['cyan']}提示：单词以 '{english_word[0]}' 开头{self.COLORS['reset']}"
                        elif attempts == 2:
                            hint = f"{self.COLORS['cyan']}提示：单词以 '{english_word[0]}' 开头，共 {len(english_word)} 个字母{self.COLORS['reset']}"
                        print(hint)
                else:
                    print(f"{self.COLORS['red']}✗ 拼写错误！{self.COLORS['reset']}")
                    # 第二次错误显示正确拼写
                    print(self._get_correct_spelling(word_data))
        
        self._update_learning_record(english_word, is_correct)
        return is_correct
    
    def _redo_wrong_words(self, wrong_words):
        """重做错题功能"""
        if not wrong_words:
            print(f"{self.COLORS['green']}✓ 没有错误单词需要重做！{self.COLORS['reset']}")
            return
        
        print(f"\n{self.COLORS['bold']}{self.COLORS['red']}🔄 开始错题重做 ({len(wrong_words)} 个单词){self.COLORS['reset']}")
        print("=" * 50)
        print(f"{self.COLORS['yellow']}重做规则：{self.COLORS['reset']}")
        print("1. 每题有3次尝试机会")
        print("2. 提供例句作为拼写提示")
        print("3. 完成后显示重做成绩")
        print("=" * 50)
        
        input(f"{self.COLORS['cyan']}按回车键开始重做...{self.COLORS['reset']}")
        
        redo_correct = 0
        start_time = time.time()
        
        # 转换错题格式为原始数据格式
        redo_words = []
        for wrong_word in wrong_words:
            # 从原始数据中找到完整的单词信息
            for word in self.words_data:
                if word['English'] == wrong_word['english']:
                    redo_words.append(word)
                    break
        
        # 开始重做
        for i, word_data in enumerate(redo_words, 1):
            print(f"\n{self.COLORS['bold']}{self.COLORS['red']}【重做】单词 {i}/{len(redo_words)}{self.COLORS['reset']}")
            if self.check_spelling(word_data, quiz_mode=False):  # 重做用普通模式
                redo_correct += 1
        
        # 重做统计
        end_time = time.time()
        redo_elapsed = round(end_time - start_time, 2)
        redo_accuracy = round(redo_correct / len(redo_words) * 100, 1) if redo_words else 100
        
        print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}🏆 错题重做完成！{self.COLORS['reset']}")
        print("=" * 60)
        print(f"错题总数：{len(redo_words)}")
        print(f"重做正确：{self.COLORS['green']}{redo_correct}{self.COLORS['reset']}")
        print(f"重做错误：{self.COLORS['red']}{len(redo_words) - redo_correct}{self.COLORS['reset']}")
        print(f"重做正确率：{self.COLORS['bold']}{redo_accuracy}%{self.COLORS['reset']}")
        print(f"重做用时：{redo_elapsed} 秒")
        
        # 计算最终掌握情况
        total_wrong = len(redo_words)
        final_correct = total_wrong - (len(redo_words) - redo_correct)
        mastery_rate = round(final_correct / total_wrong * 100, 1) if total_wrong > 0 else 100
        
        print(f"\n{self.COLORS['blue']}📊 最终掌握情况：{self.COLORS['reset']}")
        print(f"原始错误：{total_wrong} 个")
        print(f"重做后掌握：{self.COLORS['green']}{final_correct} 个{self.COLORS['reset']}")
        print(f"最终掌握率：{self.COLORS['bold']}{mastery_rate}%{self.COLORS['reset']}")
        
        if mastery_rate >= 90:
            mastery_grade = f"{self.COLORS['green']}优秀{self.COLORS['reset']}"
        elif mastery_rate >= 80:
            mastery_grade = f"{self.COLORS['green']}良好{self.COLORS['reset']}"
        elif mastery_rate >= 70:
            mastery_grade = f"{self.COLORS['yellow']}中等{self.COLORS['reset']}"
        else:
            mastery_grade = f"{self.COLORS['red']}仍需加强{self.COLORS['reset']}"
        
        print(f"掌握等级：{mastery_grade}")
        
        # 保存记录
        self._save_learning_records()
        
        input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
    
    def start_100_quiz(self):
        """开始100题随机测验（规则：
           1. 每题2次拼写机会
           2. 第一次错误显示Example提示
           3. 第二次错误显示正确拼写
           4. 完成后显示详细成绩
           5. 可选择重做所有错误单词）
        """
        if not self.words_data:
            print(f"{self.COLORS['red']}✗ 错误：没有可用的单词数据！{self.COLORS['reset']}")
            input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
            return
        
        # 检查单词数量是否足够
        quiz_count = 100
        if len(self.words_data) < quiz_count:
            print(f"{self.COLORS['yellow']}⚠️  警告：单词库只有 {len(self.words_data)} 个单词，将使用全部单词进行测验{self.COLORS['reset']}")
            quiz_count = len(self.words_data)
        
        print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}📝 100题英语拼写测验 {self.COLORS['reset']}")
        print("=" * 50)
        print(f"{self.COLORS['yellow']}测验规则：{self.COLORS['reset']}")
        print("1. 每题有2次拼写机会")
        print("2. 第一次错误显示例句提示")
        print("3. 第二次错误显示正确拼写")
        print("4. 完成后显示详细成绩")
        print("5. 可选择重做所有错误单词")
        print("=" * 50)
        
        input(f"{self.COLORS['cyan']}按回车键开始测验...{self.COLORS['reset']}")
        
        # 随机选择单词
        practice_words = random.sample(self.words_data, quiz_count)
        
        correct_count = 0
        wrong_words = []  # 记录错误的单词
        start_time = time.time()
        
        for i, word_data in enumerate(practice_words, 1):
            print(f"\n{self.COLORS['bold']}{self.COLORS['blue']}【测验】单词 {i}/{quiz_count}{self.COLORS['reset']}")
            # 调用check_spelling时传入quiz_mode=True，启用测验模式规则
            is_correct = self.check_spelling(word_data, quiz_mode=True)
            if is_correct:
                correct_count += 1
            else:
                wrong_words.append({
                    'english': word_data['English'],
                    'chinese': word_data['Chinese'],
                    'example': word_data['Example']
                })
        
        # 测验结束统计
        end_time = time.time()
        elapsed_time = round(end_time - start_time, 2)
        accuracy = round(correct_count / quiz_count * 100, 1) if quiz_count > 0 else 0
        
        print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}🏆 测验完成！{self.COLORS['reset']}")
        print("=" * 60)
        print(f"总题数：{quiz_count}")
        print(f"正确拼写：{self.COLORS['green']}{correct_count}{self.COLORS['reset']}")
        print(f"错误拼写：{self.COLORS['red']}{quiz_count - correct_count}{self.COLORS['reset']}")
        print(f"正确率：{self.COLORS['bold']}{accuracy}%{self.COLORS['reset']}")
        print(f"总用时：{elapsed_time} 秒")
        print(f"平均每题用时：{round(elapsed_time/quiz_count, 2)} 秒")
        
        # 评级
        if accuracy >= 90:
            grade = f"{self.COLORS['green']}优秀 (S级){self.COLORS['reset']}"
        elif accuracy >= 80:
            grade = f"{self.COLORS['green']}良好 (A级){self.COLORS['reset']}"
        elif accuracy >= 70:
            grade = f"{self.COLORS['yellow']}中等 (B级){self.COLORS['reset']}"
        elif accuracy >= 60:
            grade = f"{self.COLORS['yellow']}及格 (C级){self.COLORS['reset']}"
        else:
            grade = f"{self.COLORS['red']}不及格 (D级){self.COLORS['reset']}"
        print(f"综合评级：{grade}")
        
        # 保存测验记录（安全处理）
        try:
            # 清理颜色代码后的纯文本评级
            clean_grade = grade.replace('\033[32m', '').replace('\033[33m', '').replace('\033[31m', '').replace('\033[0m', '')
            
            quiz_record = {
                "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "total": quiz_count,
                "correct": correct_count,
                "accuracy": accuracy,
                "time_used": elapsed_time,
                "grade": clean_grade,
                "wrong_count": len(wrong_words)
            }
            
            # 确保quiz_records字段存在
            if 'quiz_records' not in self.learning_records['stats']:
                self.learning_records['stats']['quiz_records'] = []
            
            self.learning_records['stats']['quiz_records'].append(quiz_record)
            self._save_learning_records()
            print(f"{self.COLORS['green']}✓ 测验记录已保存{self.COLORS['reset']}")
            
        except Exception as e:
            print(f"{self.COLORS['yellow']}⚠️ 保存测验记录失败：{str(e)}{self.COLORS['reset']}")
        
        # 显示错误单词
        if wrong_words:
            print(f"\n{self.COLORS['red']}❌ 错误单词列表 ({len(wrong_words)}个):{self.COLORS['reset']}")
            print("-" * 50)
            display_count = min(10, len(wrong_words))
            for idx, word in enumerate(wrong_words[:display_count], 1):
                print(f"{idx}. {word['english']} - {word['chinese']}")
            if len(wrong_words) > 10:
                print(f"   ... 还有 {len(wrong_words)-10} 个错误单词")
        
        print("\n" + "=" * 60)
        
        # 询问是否重做错题
        if wrong_words:
            while True:
                redo_choice = input(f"\n{self.COLORS['yellow']}是否要重做这些错误单词？(y/n): {self.COLORS['reset']}").strip().lower()
                if redo_choice in ['y', 'n']:
                    break
                print(f"{self.COLORS['red']}✗ 请输入 y 或 n{self.COLORS['reset']}")
            
            if redo_choice == 'y':
                self._redo_wrong_words(wrong_words)
        else:
            print(f"\n{self.COLORS['green']}🎉 恭喜！所有单词拼写都正确！{self.COLORS['reset']}")
            input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
    
    def start_practice(self):
        """开始拼写练习"""
        if not self.words_data:
            print(f"{self.COLORS['red']}✗ 错误：没有可用的单词数据！{self.COLORS['reset']}")
            input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
            return
        
        print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}📚 开始英语拼写练习 {self.COLORS['reset']}")
        print("=" * 50)
        
        print(f"\n{self.COLORS['yellow']}请选择练习模式：{self.COLORS['reset']}")
        print("1. 随机单词练习")
        print("2. 困难单词练习（正确率低的单词）")
        
        while True:
            mode_choice = input(f"\n{self.COLORS['purple']}请选择模式 (1/2): {self.COLORS['reset']}").strip()
            if mode_choice in ['1', '2']:
                break
            print(f"{self.COLORS['red']}✗ 无效选择，请输入 1 或 2{self.COLORS['reset']}")
        
        while True:
            try:
                count = int(input(f"\n{self.COLORS['purple']}请输入要练习的单词数量: {self.COLORS['reset']}"))
                if 1 <= count <= len(self.words_data):
                    break
                print(f"{self.COLORS['red']}✗ 请输入1到{len(self.words_data)}之间的数字{self.COLORS['reset']}")
            except ValueError:
                print(f"{self.COLORS['red']}✗ 请输入有效的数字{self.COLORS['reset']}")
        
        if mode_choice == '1':
            practice_words = random.sample(self.words_data, count)
            print(f"\n{self.COLORS['green']}✅ 已随机选择 {count} 个单词进行练习{self.COLORS['reset']}")
        else:
            practice_words = self._get_difficult_words(count)
            print(f"\n{self.COLORS['green']}✅ 已选择 {count} 个需要加强的单词进行练习{self.COLORS['reset']}")
        
        correct_count = 0
        start_time = time.time()
        
        for i, word_data in enumerate(practice_words, 1):
            print(f"\n{self.COLORS['bold']}{self.COLORS['blue']}单词 {i}/{count}{self.COLORS['reset']}")
            # 普通练习模式传入quiz_mode=False，保持原有3次机会+每次提示的规则
            if self.check_spelling(word_data, quiz_mode=False):
                correct_count += 1
        
        end_time = time.time()
        elapsed_time = round(end_time - start_time, 2)
        accuracy = round(correct_count / count * 100, 1) if count > 0 else 0
        
        print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}📊 练习完成！{self.COLORS['reset']}")
        print("=" * 50)
        print(f"总单词数：{count}")
        print(f"正确拼写：{self.COLORS['green']}{correct_count}{self.COLORS['reset']}")
        print(f"错误拼写：{self.COLORS['red']}{count - correct_count}{self.COLORS['reset']}")
        print(f"正确率：{self.COLORS['bold']}{accuracy}%{self.COLORS['reset']}")
        print(f"用时：{elapsed_time} 秒")
        
        self._save_learning_records()
        
        while True:
            choice = input(f"\n{self.COLORS['yellow']}是否继续练习？(y/n): {self.COLORS['reset']}").strip().lower()
            if choice in ['y', 'n']:
                break
            print(f"{self.COLORS['red']}✗ 请输入 y 或 n{self.COLORS['reset']}")
        
        if choice == 'y':
            self.start_practice()
    
    def show_word_list(self):
        """显示单词列表"""
        print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}📖 英语单词列表 {self.COLORS['reset']}")
        print("=" * 60)
        
        if not self.words_data:
            print(f"{self.COLORS['red']}✗ 暂无单词数据可显示！{self.COLORS['reset']}")
            input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
            return
        
        sorted_words = sorted(self.words_data, key=lambda x: x['English'])
        
        page_size = 15
        total_pages = (len(sorted_words) + page_size - 1) // page_size
        current_page = 1
        
        while True:
            os.system('cls' if os.name == 'nt' else 'clear')
            print(f"{self.COLORS['bold']}{self.COLORS['purple']}📖 英语单词列表 (共{len(sorted_words)}个单词) {self.COLORS['reset']}")
            print("=" * 80)
            print(f"{self.COLORS['blue']}第 {current_page}/{total_pages} 页{self.COLORS['reset']}")
            print(f"{self.COLORS['cyan']}{'序号':<4} {'英语单词':<12} {'中文释义':<15} {'学习状态':<8}{self.COLORS['reset']}")
            print("-" * 80)
            
            start_idx = (current_page - 1) * page_size
            end_idx = min(start_idx + page_size, len(sorted_words))
            
            for i in range(start_idx, end_idx):
                idx = i + 1
                word = sorted_words[i]
                english = word['English'][:10] + '...' if len(word['English']) > 10 else word['English']
                chinese = word['Chinese'][:13] + '...' if len(word['Chinese']) > 13 else word['Chinese']
                
                if english in self.learning_records['words']:
                    record = self.learning_records['words'][english]
                    if record['success_rate'] >= 90:
                        status = f"{self.COLORS['green']}已掌握{self.COLORS['reset']}"
                    elif record['success_rate'] >= 60:
                        status = f"{self.COLORS['yellow']}学习中{self.COLORS['reset']}"
                    else:
                        status = f"{self.COLORS['red']}需加强{self.COLORS['reset']}"
                else:
                    status = f"{self.COLORS['white']}未练习{self.COLORS['reset']}"
                
                print(f"{idx:<4} {english:<12} {chinese:<15} {status:<8}")
            
            print("\n" + "=" * 80)
            print(f"{self.COLORS['purple']}操作：{self.COLORS['reset']}")
            print("n - 下一页 | p - 上一页 | q - 返回主菜单 | s - 搜索单词")
            
            choice = input(f"\n{self.COLORS['yellow']}请选择操作: {self.COLORS['reset']}").strip().lower()
            
            if choice == 'n' and current_page < total_pages:
                current_page += 1
            elif choice == 'p' and current_page > 1:
                current_page -= 1
            elif choice == 's':
                keyword = input(f"\n{self.COLORS['purple']}请输入要搜索的单词或中文释义: {self.COLORS['reset']}").lower()
                results = [word for word in sorted_words if keyword in word['English'].lower() or keyword in word['Chinese'].lower()]
                
                print(f"\n{self.COLORS['cyan']}🔍 搜索结果 (共{len(results)}个):{self.COLORS['reset']}")
                if results:
                    for i, word in enumerate(results[:10], 1):
                        print(f"   {i}. {word['English']} - {word['Chinese']}")
                else:
                    print(f"   {self.COLORS['red']}未找到匹配的单词{self.COLORS['reset']}")
                input(f"\n{self.COLORS['yellow']}按回车键继续...{self.COLORS['reset']}")
            elif choice == 'q':
                break
            else:
                print(f"{self.COLORS['red']}✗ 无效操作或已到边界{self.COLORS['reset']}")
                time.sleep(1)
    
    def show_learning_stats(self):
        """显示学习统计"""
        print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}📈 学习进度统计 {self.COLORS['reset']}")
        print("=" * 60)
        
        total_words = len(self.words_data)
        practiced_words = len(self.learning_records['words'])
        practiced_percent = round(practiced_words / total_words * 100, 1) if total_words > 0 else 0
        
        mastered = 0
        learning = 0
        need_practice = 0
        
        for record in self.learning_records['words'].values():
            if record['success_rate'] >= 90:
                mastered += 1
            elif record['success_rate'] >= 60:
                learning += 1
            else:
                need_practice += 1
        
        total_practiced = self.learning_records['stats']['total_practiced']
        total_correct = self.learning_records['stats']['total_correct']
        overall_accuracy = round(total_correct / total_practiced * 100, 1) if total_practiced > 0 else 0
        
        print(f"{self.COLORS['blue']}📋 总体情况：{self.COLORS['reset']}")
        print(f"总单词数：{total_words}")
        print(f"已练习：{practiced_words} ({practiced_percent}%)")
        print(f"未练习：{total_words - practiced_words}")
        print(f"\n{self.COLORS['blue']}📊 掌握情况：{self.COLORS['reset']}")
        print(f"已掌握：{self.COLORS['green']}{mastered} 个{self.COLORS['reset']} (成功率≥90%)")
        print(f"学习中：{self.COLORS['yellow']}{learning} 个{self.COLORS['reset']} (60%≤成功率<90%)")
        print(f"需加强：{self.COLORS['red']}{need_practice} 个{self.COLORS['reset']} (成功率<60%)")
        print(f"\n{self.COLORS['blue']}📈 练习统计：{self.COLORS['reset']}")
        print(f"总练习次数：{total_practiced}")
        print(f"总正确次数：{total_correct}")
        print(f"总体正确率：{self.COLORS['bold']}{overall_accuracy}%{self.COLORS['reset']}")
        
        if self.learning_records['stats']['last_practice']:
            print(f"最后练习：{self.learning_records['stats']['last_practice']}")
        
        # 显示测验记录（安全处理）
        quiz_records = self.learning_records['stats'].get('quiz_records', [])
        if quiz_records and len(quiz_records) > 0:
            print(f"\n{self.COLORS['blue']}📝 测验记录：{self.COLORS['reset']}")
            recent_quizzes = quiz_records[-3:]  # 显示最近3次
            for i, quiz in enumerate(recent_quizzes, 1):
                wrong_count = quiz.get('wrong_count', '未知')
                print(f"{i}. {quiz['date']} | {quiz['correct']}/{quiz['total']} | {quiz['accuracy']}% | "
                      f"{self.COLORS['red']}错误：{wrong_count}{self.COLORS['reset']} | {quiz['grade']}")
        
        if self.learning_records['words']:
            print(f"\n{self.COLORS['blue']}🔍 最近练习的单词：{self.COLORS['reset']}")
            recent_words = sorted(
                self.learning_records['words'].items(),
                key=lambda x: x[1]['last_attempt'],
                reverse=True
            )[:5]
            
            for word, record in recent_words:
                print(f"{word:<10} | 正确率：{record['success_rate']}% | 练习次数：{record['attempts']}")
        
        input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
    
    def main_menu(self):
        """主菜单"""
        while True:
            os.system('cls' if os.name == 'nt' else 'clear')
            print(f"{self.COLORS['bold']}{self.COLORS['purple']}🎯 英语拼写检查练习程序 {self.COLORS['reset']}")
            print("=" * 50)
            print(f"📝 当前单词库：{self.COLORS['green']}{len(self.words_data)} 个单词{self.COLORS['reset']}")
            print("=" * 50)
            print("1. 开始拼写练习")
            print("2. 100题随机测验")
            print("3. 查看单词列表")
            print("4. 查看学习进度")
            print("5. 退出程序")
            print("=" * 50)
            
            choice = input(f"\n{self.COLORS['yellow']}请选择操作 (1-5): {self.COLORS['reset']}").strip()
            
            if choice == '1':
                self.start_practice()
            elif choice == '2':  # 新增选项处理
                self.start_100_quiz()
            elif choice == '3':
                self.show_word_list()
            elif choice == '4':
                self.show_learning_stats()
            elif choice == '5':
                print(f"\n{self.COLORS['green']}👋 感谢使用！学习记录已保存{self.COLORS['reset']}")
                self._save_learning_records()
                break
            else:
                print(f"{self.COLORS['red']}✗ 无效选择，请输入1-5之间的数字{self.COLORS['reset']}")
                time.sleep(1)

if __name__ == "__main__":
    try:
        checker = EnglishSpellingChecker()
        checker.main_menu()
    except ImportError as e:
        print("缺少必要的依赖包，请运行以下命令安装：")
        print("pip install pandas openpyxl fsspec")
        print(f"\n错误详情：{e}")
        input("按回车键退出...")
    except Exception as e:
        print(f"程序运行出错：{type(e).__name__}: {str(e)}")
        import traceback
        traceback.print_exc()
        input("按回车键退出...")