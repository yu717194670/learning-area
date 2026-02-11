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
        
        # ========== macOS 路径配置 ==========
        # 方式1：用户文档目录（推荐）
        # self.file_path = os.path.expanduser("~/Documents/Single_word.xlsx")
        # 方式2：桌面路径
        # self.file_path = os.path.expanduser("~/Desktop/Single_word.xlsx")
        # 方式3：相对路径（代码同目录）
        # self.file_path = "./Single_word.xlsx"
        # ========== windows 路径配置 ==========
        self.file_path = os.path.expanduser("Single_word.xlsx")
    
        self.records_file = "learning_records.json"
        self.has_mark_column = False
        self.target_marks = ["1", 1]  # 同时兼容字符串1和数字1
        
        # 初始化数据
        self.words_data = self._load_and_prepare_data()
        self.learning_records = self._load_learning_records()
        
    def _load_and_prepare_data(self):
        """加载并预处理Excel数据，包含详细调试日志"""
        try:
            # 检查文件是否存在
            if not os.path.exists(self.file_path):
                raise FileNotFoundError(f"文件不存在: {self.file_path}")
            
            # 读取Excel文件
            df = pd.read_excel(self.file_path, engine='openpyxl')
            # 保存原始DataFrame（用于后续更新mark列）
            self.original_df = df.copy()
            
            # 列名统一处理：转小写 + 去除前后空格
            df.columns = [col.strip().lower() for col in df.columns]
            
            # 映射必要列（新增equal列）
            required_columns = ['english', 'chinese', 'example', 'equal']
            column_mapping = {}
            
            for req_col in required_columns:
                if req_col in df.columns:
                    column_mapping[req_col] = req_col
                else:
                    print(f"{self.COLORS['yellow']}⚠️  未找到{req_col}列，创建空列{self.COLORS['reset']}")
                    column_mapping[req_col] = None
            
            # 创建新的DataFrame
            new_df = pd.DataFrame()
            for col in ['english', 'chinese', 'example', 'equal']:
                if column_mapping[col] is not None:
                    new_df[col] = df[column_mapping[col]]
                else:
                    new_df[col] = ""
            
            # 处理mark列
            if 'mark' in df.columns:
                self.has_mark_column = True
                new_df['mark'] = df['mark']
                
                # 打印mark列的所有唯一值，方便排查
                mark_unique_values = new_df['mark'].unique()
                # print(f"\n✅ mark列所有唯一值: {list(mark_unique_values)}")
                
                # 统计目标标记（1）的数量
                target_mask = new_df['mark'].isin(self.target_marks)
                target_count = len(new_df[target_mask])
                # print(f"✅ 匹配到标记为1的单词数量: {target_count}")
            else:
                self.has_mark_column = False
                # 如果没有mark列，创建空的mark列
                new_df['mark'] = ""
                self.has_mark_column = True
                print(f"{self.COLORS['yellow']}⚠️  未找到mark列，已自动创建空mark列{self.COLORS['reset']}")
            
            # 数据清洗
            print(f"\n✅ 开始数据清洗...")
            # 处理英文列：转小写、去空格、过滤空值和nan
            new_df['english'] = new_df['english'].astype(str).str.strip().str.lower()
            new_df = new_df[new_df['english'] != '']
            new_df = new_df[new_df['english'] != 'nan']
            new_df = new_df[new_df['english'] != 'null']
            
            # 处理其他列（新增equal列清洗）
            new_df['chinese'] = new_df['chinese'].astype(str).str.strip()
            new_df['example'] = new_df['example'].astype(str).str.strip()
            new_df['equal'] = new_df['equal'].astype(str).str.strip()
            
            print(f"✅ 清洗后剩余单词总数: {len(new_df)}")
            
            # 转换为字典列表
            words_data = new_df.to_dict('records')
            
            # 去重（基于英文单词）
            seen = set()
            unique_words = []
            for word in words_data:
                if word['english'] not in seen:
                    seen.add(word['english'])
                    unique_words.append(word)
            
            print(f"✅ 去重后最终单词数: {len(unique_words)}")
            print(f"{self.COLORS['blue']}===== 数据加载完成 ====={self.COLORS['reset']}\n")
            
            return unique_words
            
        except FileNotFoundError as e:
            print(f"{self.COLORS['red']}❌ 错误: {e}{self.COLORS['reset']}")
            print(f"{self.COLORS['yellow']}💡 请检查文件路径是否正确{self.COLORS['reset']}")
        except Exception as e:
            print(f"{self.COLORS['red']}❌ 数据加载失败: {type(e).__name__} - {str(e)}{self.COLORS['reset']}")

        self.has_mark_column = True

    def _load_learning_records(self):
        """加载学习记录（兼容旧版本）"""
        default_records = {
            "words": {},
            "stats": {
                "total_practiced": 0,
                "total_correct": 0,
                "last_practice": "",
                "quiz_records": []
            }
        }
        
        if os.path.exists(self.records_file):
            try:
                with open(self.records_file, 'r', encoding='utf-8') as f:
                    loaded_records = json.load(f)
                
                # 兼容旧版本数据结构
                if 'stats' not in loaded_records:
                    loaded_records['stats'] = default_records['stats']
                if 'words' not in loaded_records:
                    loaded_records['words'] = {}
                
                return loaded_records
                
            except Exception as e:
                print(f"{self.COLORS['yellow']}⚠️  学习记录文件损坏: {str(e)}，将创建新记录{self.COLORS['reset']}")
        
        return default_records
    
    def _save_learning_records(self):
        """保存学习记录"""
        try:
            self.learning_records['stats']['last_practice'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(self.records_file, 'w', encoding='utf-8') as f:
                json.dump(self.learning_records, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"{self.COLORS['red']}❌ 保存学习记录失败: {str(e)}{self.COLORS['reset']}")
            return False
    
    def _update_word_mark(self, english_word):
        """
        更新单词的mark标记为1（重点单词）
        :param english_word: 要标记的英文单词（小写）
        """
        # 1. 更新内存中的words_data
        for word in self.words_data:
            if word['english'] == english_word:
                if word.get('mark') not in self.target_marks:
                    word['mark'] = 1
                    print(f"{self.COLORS['purple']}🔖 已将单词 '{english_word}' 标记为重点单词（mark=1）{self.COLORS['reset']}")
                break
        
        # 2. 同步更新Excel文件中的mark列
        try:
            # 找到原始Excel中对应的行
            english_col = None
            for col in self.original_df.columns:
                if col.strip().lower() == 'english':
                    english_col = col
                    break
            
            if english_col:
                # 将原始Excel中的英文列转小写，匹配目标单词
                mask = self.original_df[english_col].astype(str).str.strip().str.lower() == english_word
                if mask.any():
                    # 更新mark列（如果没有mark列则创建）
                    if 'mark' not in self.original_df.columns:
                        self.original_df['mark'] = ""
                    self.original_df.loc[mask, 'mark'] = 1
                    # 保存回Excel文件
                    self.original_df.to_excel(self.file_path, index=False, engine='openpyxl')
        except Exception as e:
            print(f"{self.COLORS['yellow']}⚠️  同步更新Excel mark列失败: {str(e)}{self.COLORS['reset']}")
    
    def _get_example_hint(self, word_data, attempt_num=1):
        """
        获取例句/Equal/字母提示（新增Equal列内容展示）
        :param word_data: 单词数据字典
        :param attempt_num: 当前尝试次数（1表示第一次错误）
        """
        # 第一次错误时优先展示Equal列内容
        equal_content = word_data['equal'].strip()
        if attempt_num == 1 and equal_content and equal_content != "nan":
            hint = f"{self.COLORS['cyan']}提示：\n{equal_content}{self.COLORS['reset']}"
            # 如果有例句，补充展示
            example_hint = word_data['example'].strip()
            if example_hint and example_hint != "nan":
                hint += f"\n{self.COLORS['cyan']}联想：{example_hint}{self.COLORS['reset']}"
            return hint
        # 非第一次错误或无Equal内容时，展示原有提示
        else:
            example_hint = word_data['example'].strip()
            if example_hint and example_hint != "nan":
                return f"{self.COLORS['cyan']}提示：{example_hint}{self.COLORS['reset']}"
            else:
                english_word = word_data['english'].lower()
                return f"{self.COLORS['cyan']}提示：单词以 '{english_word[0]}' 开头，共 {len(english_word)} 个字母{self.COLORS['reset']}"
    
    def _get_correct_spelling(self, word_data):
        """返回正确拼写提示"""
        english_word = word_data['english'].lower()
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
        
        # 拼写错误时自动标记为重点单词
        if not is_correct:
            self._update_word_mark(english_word)
    
    def _get_all_target_mark_words(self):
        """获取所有标记为1的单词（兼容数字和字符串）"""
        if not self.has_mark_column:
            return []
        return [wd for wd in self.words_data if wd.get('mark') in self.target_marks]
    
    def _get_difficult_words(self, count=10):
        """获取需要加强练习的单词（成功率<70%）"""
        difficult_words = []
        
        # 筛选成功率低的单词
        for word, record in self.learning_records['words'].items():
            if record['success_rate'] < 70 and record['attempts'] > 0:
                for wd in self.words_data:
                    if wd['english'].lower() == word.lower():
                        difficult_words.append(wd)
                        break
        
        # 如果困难单词不足，补充随机单词
        if len(difficult_words) < count:
            all_words_copy = [wd for wd in self.words_data if wd not in difficult_words]
            random.shuffle(all_words_copy)
            difficult_words.extend(all_words_copy[:count - len(difficult_words)])
        
        return difficult_words[:count]
    
    def check_spelling(self, word_data, quiz_mode=False):
        """
        检查单个单词拼写（修复：仅大写 E 退出）
        :param word_data: 单词数据字典
        :param quiz_mode: 是否为测验模式（2次机会）
        :return: (是否拼写正确, 是否停止测验)
        """
        english_word = word_data['english'].lower()
        chinese_meaning = word_data['chinese']

        # 显示题目
        print(f"\n{self.COLORS['blue']}━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━{self.COLORS['reset']}")
        print(f"{self.COLORS['bold']}请拼写单词：{self.COLORS['yellow']}{chinese_meaning}{self.COLORS['reset']}")

        attempts = 0
        is_correct = False
        stop_quiz = False
        max_attempts = 2 if quiz_mode else 3  # 测验模式2次机会，普通模式3次

        while attempts < max_attempts and not is_correct and not stop_quiz:
            user_input = input(f"\n{self.COLORS['purple']}请输入拼写: {self.COLORS['reset']}").strip()

            # 仅输入大写 E 才退出，其他都正常判题
            if user_input == 'E':
                stop_quiz = True
                print(f"{self.COLORS['yellow']}⚠️  已停止当前测验/练习{self.COLORS['reset']}")
                continue

            # 转小写判题，不影响正常拼写
            user_input = user_input.lower()

            attempts += 1

            if user_input == english_word:
                is_correct = True
                print(f"{self.COLORS['green']}✓ 恭喜！拼写正确！{self.COLORS['reset']}")
            else:
                if attempts < max_attempts:
                    print(f"{self.COLORS['red']}✗ 拼写错误！{self.COLORS['reset']}")
                    print(self._get_example_hint(word_data, attempts))
                    # print(f"{self.COLORS['yellow']}⚠️  还有 {max_attempts - attempts} 次尝试机会{self.COLORS['reset']}")
                else:
                    print(f"{self.COLORS['red']}✗ 拼写错误！{self.COLORS['reset']}")
                    print(self._get_correct_spelling(word_data))

        # 只有非停止状态才更新学习记录
        if not stop_quiz:
            self._update_learning_record(english_word, is_correct)

        return is_correct, stop_quiz
    
    def _redo_wrong_words(self, wrong_words):
        """重做错题功能（新增停止逻辑）"""
        if not wrong_words:
            print(f"{self.COLORS['green']}✓ 没有错误单词需要重做！{self.COLORS['reset']}")
            return
        
        print(f"\n{self.COLORS['bold']}{self.COLORS['red']}🔄 开始错题重做 ({len(wrong_words)} 个单词){self.COLORS['reset']}")
        print("=" * 50)
        print(f"{self.COLORS['yellow']}重做规则：{self.COLORS['reset']}")
        print("1. 每题有3次尝试机会")
        print("2. 首次错误后提供提示（含Equal列），第二次错误显示正确拼写")
        print("3. 错误单词会自动标记为重点单词（mark=1）")
        print("4. 完成后显示重做成绩")
        print(f"{self.COLORS['cyan']}💡 输入'E'可停止重做{self.COLORS['reset']}")
        print("=" * 50)
        
        input(f"{self.COLORS['cyan']}按回车键开始重做...{self.COLORS['reset']}")
        
        redo_correct = 0
        stop_redo = False
        start_time = time.time()
        
        # 逐个练习错误单词
        for i, word_data in enumerate(wrong_words, 1):
            if stop_redo:
                break
                
            print(f"\n{self.COLORS['bold']}{self.COLORS['red']}【重做】单词 {i}/{len(wrong_words)}{self.COLORS['reset']}")
            is_correct, stop_redo = self.check_spelling(word_data, quiz_mode=False)
            if is_correct and not stop_redo:
                redo_correct += 1
        
        # 统计重做结果
        end_time = time.time()
        redo_elapsed = round(end_time - start_time, 2)
        completed_count = i if stop_redo else len(wrong_words)
        redo_accuracy = round(redo_correct / completed_count * 100, 1) if completed_count > 0 else 0
        
        if stop_redo:
            print(f"\n{self.COLORS['yellow']}⚠️  错题重做已停止{self.COLORS['reset']}")
        else:
            print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}🏆 错题重做完成！{self.COLORS['reset']}")
        
        print("=" * 60)
        print(f"已完成单词数：{completed_count}/{len(wrong_words)}")
        print(f"正确拼写：{self.COLORS['green']}{redo_correct}{self.COLORS['reset']}")
        print(f"错误拼写：{self.COLORS['red']}{completed_count - redo_correct}{self.COLORS['reset']}")
        print(f"正确率：{self.COLORS['bold']}{redo_accuracy}%{self.COLORS['reset']}")
        print(f"用时：{redo_elapsed} 秒")
        
        # 保存记录
        self._save_learning_records()
        
        input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
    
    def start_100_quiz(self):
        """100题随机测验功能（新增停止测验逻辑）"""
        if not self.words_data:
            print(f"{self.COLORS['red']}✗ 错误：没有可用的单词数据！{self.COLORS['reset']}")
            input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
            return
        
        # 确定测验题数
        quiz_count = 100
        if len(self.words_data) < quiz_count:
            print(f"{self.COLORS['yellow']}⚠️  警告：单词库只有 {len(self.words_data)} 个单词，将使用全部单词测验{self.COLORS['reset']}")
            quiz_count = len(self.words_data)
        
        print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}📝 100题英语拼写测验 {self.COLORS['reset']}")
        print("=" * 50)
        print(f"{self.COLORS['yellow']}测验规则：{self.COLORS['reset']}")
        print("1. 每题有2次拼写机会")
        print("2. 第一次错误显示提示（含Equal列），第二次错误显示正确拼写")
        print("3. 错误单词会自动标记为重点单词（mark=1）")
        print("4. 完成后显示详细成绩")
        print("5. 可选择重做所有错误单词")
        print(f"{self.COLORS['cyan']}💡 输入'E'可随时停止测验{self.COLORS['reset']}")
        print("=" * 50)
        
        input(f"{self.COLORS['cyan']}按回车键开始测验...{self.COLORS['reset']}")
        
        # 随机选择单词
        practice_words = random.sample(self.words_data, quiz_count)
        
        correct_count = 0
        wrong_words = []
        stop_quiz = False
        start_time = time.time()
        
        # 开始测验
        for i, word_data in enumerate(practice_words, 1):
            if stop_quiz:
                break
                
            print(f"\n{self.COLORS['bold']}{self.COLORS['blue']}【测验】单词 {i}/{quiz_count}{self.COLORS['reset']}")
            is_correct, stop_quiz = self.check_spelling(word_data, quiz_mode=True)
            if is_correct and not stop_quiz:
                correct_count += 1
            elif not is_correct and not stop_quiz:
                wrong_words.append(word_data)
        
        # 测验统计
        end_time = time.time()
        elapsed_time = round(end_time - start_time, 2)
        completed_count = i if stop_quiz else quiz_count
        accuracy = round(correct_count / completed_count * 100, 1) if completed_count > 0 else 0
        
        if stop_quiz:
            print(f"\n{self.COLORS['yellow']}⚠️  100题测验已停止{self.COLORS['reset']}")
        else:
            print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}🏆 测验完成！{self.COLORS['reset']}")
        
        print("=" * 60)
        print(f"已完成题数：{completed_count}/{quiz_count}")
        print(f"正确拼写：{self.COLORS['green']}{correct_count}{self.COLORS['reset']}")
        print(f"错误拼写：{self.COLORS['red']}{completed_count - correct_count}{self.COLORS['reset']}")
        print(f"正确率：{self.COLORS['bold']}{accuracy}%{self.COLORS['reset']}")
        print(f"总用时：{elapsed_time} 秒")
        if completed_count > 0:
            print(f"平均每题用时：{round(elapsed_time/completed_count, 2)} 秒")
        
        # 评级（仅完成时显示）
        if not stop_quiz:
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
        
        # 保存测验记录（仅完成时保存完整记录）
        try:
            if not stop_quiz:
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
                
                if 'quiz_records' not in self.learning_records['stats']:
                    self.learning_records['stats']['quiz_records'] = []
                self.learning_records['stats']['quiz_records'].append(quiz_record)
                print(f"{self.COLORS['green']}✓ 测验记录已保存{self.COLORS['reset']}")
            else:
                print(f"{self.COLORS['yellow']}⚠️  测验未完成，不保存完整测验记录（学习记录已保存）{self.COLORS['reset']}")
            
            self._save_learning_records()
            
        except Exception as e:
            print(f"{self.COLORS['yellow']}⚠️  保存记录失败：{str(e)}{self.COLORS['reset']}")
        
        # 显示错误单词（仅完成时）
        if not stop_quiz and wrong_words:
            print(f"\n{self.COLORS['red']}❌ 错误单词列表 ({len(wrong_words)}个):{self.COLORS['reset']}")
            print("-" * 50)
            display_count = min(10, len(wrong_words))
            for idx, word in enumerate(wrong_words[:display_count], 1):
                print(f"{idx}. {word['english']} - {word['chinese']}")
            if len(wrong_words) > 10:
                print(f"   ... 还有 {len(wrong_words)-10} 个错误单词")
        
        print("\n" + "=" * 60)
        
        # 询问是否重做错题（仅完成且有错题时）
        if not stop_quiz and wrong_words:
            while True:
                redo_choice = input(f"\n{self.COLORS['yellow']}是否要重做这些错误单词？(y/n): {self.COLORS['reset']}").strip().lower()
                if redo_choice in ['y', 'n']:
                    break
                print(f"{self.COLORS['red']}✗ 请输入 y 或 n{self.COLORS['reset']}")
            
            if redo_choice == 'y':
                self._redo_wrong_words(wrong_words)
        else:
            if not stop_quiz:
                print(f"\n{self.COLORS['green']}🎉 恭喜！所有单词拼写都正确！{self.COLORS['reset']}")
            
            input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
    
    def start_practice(self):
        """开始拼写练习（新增停止逻辑+返回上一页功能）"""
        if not self.words_data:
            print(f"{self.COLORS['red']}✗ 错误：没有可用的单词数据！{self.COLORS['reset']}")
            input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
            return
        
        print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}📚 开始英语拼写练习 {self.COLORS['reset']}")
        print("=" * 50)
        
        # 显示练习模式选项（新增返回上一页选项）
        print(f"\n{self.COLORS['yellow']}请选择练习模式：{self.COLORS['reset']}")
        print("1. 随机单词练习")
        print("2. 困难单词练习（正确率低的单词）")
        if self.has_mark_column:
            print("3. 重点单词练习（标记为1的单词）")
        print("0. 返回上一页（主菜单）")
        
        # 验证模式选择
        valid_choices = ['0', '1', '2']
        if self.has_mark_column:
            valid_choices.append('3')
        
        while True:
            mode_choice = input(f"\n{self.COLORS['purple']}请选择模式 ({'/'.join(valid_choices)}): {self.COLORS['reset']}").strip()
            if mode_choice in valid_choices:
                break
            print(f"{self.COLORS['red']}✗ 无效选择，请输入 {'、'.join(valid_choices)}{self.COLORS['reset']}")
        
        # 返回上一页（主菜单）
        if mode_choice == '0':
            print(f"{self.COLORS['blue']}🔙 返回主菜单...{self.COLORS['reset']}")
            return
        
        # 处理重点单词练习模式
        if mode_choice == '3' and self.has_mark_column:
            practice_words = self._get_all_target_mark_words()
            
            if not practice_words:
                print(f"\n{self.COLORS['yellow']}⚠️  没有找到标记为1的单词！{self.COLORS['reset']}")
                # 练习模式为空时，询问是否返回或重新选择
                while True:
                    back_choice = input(f"\n{self.COLORS['yellow']}是否返回上一页重新选择？(y/n): {self.COLORS['reset']}").strip().lower()
                    if back_choice in ['y', 'n']:
                        break
                    print(f"{self.COLORS['red']}✗ 请输入 y 或 n{self.COLORS['reset']}")
                
                if back_choice == 'y':
                    self.start_practice()
                else:
                    input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
                return
            
            # 打乱单词顺序
            random.shuffle(practice_words)
            count = len(practice_words)
            print(f"\n{self.COLORS['green']}✅ 已选择所有标记为1的重点单词，共 {count} 个{self.COLORS['reset']}")
        
        # 处理随机单词练习
        elif mode_choice == '1':
            while True:
                try:
                    count_input = input(f"\n{self.COLORS['purple']}请输入要练习的单词数量 (输入0返回上一页): {self.COLORS['reset']}")
                    # 支持输入0返回上一页
                    if count_input.strip() == '0':
                        print(f"{self.COLORS['blue']}🔙 返回上一页...{self.COLORS['reset']}")
                        self.start_practice()
                        return
                    
                    count = int(count_input)
                    if 1 <= count <= len(self.words_data):
                        break
                    print(f"{self.COLORS['red']}✗ 请输入1到{len(self.words_data)}之间的数字{self.COLORS['reset']}")
                except ValueError:
                    print(f"{self.COLORS['red']}✗ 请输入有效的数字（输入0可返回上一页）{self.COLORS['reset']}")
            
            practice_words = random.sample(self.words_data, count)
            print(f"\n{self.COLORS['green']}✅ 已随机选择 {count} 个单词进行练习{self.COLORS['reset']}")
            print(f"{self.COLORS['yellow']}💡 提示：错误单词会自动标记为重点单词（mark=1）{self.COLORS['reset']}")
        
        # 处理困难单词练习
        elif mode_choice == '2':
            while True:
                try:
                    count_input = input(f"\n{self.COLORS['purple']}请输入要练习的单词数量 (输入0返回上一页): {self.COLORS['reset']}")
                    # 支持输入0返回上一页
                    if count_input.strip() == '0':
                        print(f"{self.COLORS['blue']}🔙 返回上一页...{self.COLORS['reset']}")
                        self.start_practice()
                        return
                    
                    count = int(count_input)
                    if 1 <= count <= len(self.words_data):
                        break
                    print(f"{self.COLORS['red']}✗ 请输入1到{len(self.words_data)}之间的数字{self.COLORS['reset']}")
                except ValueError:
                    print(f"{self.COLORS['red']}✗ 请输入有效的数字（输入0可返回上一页）{self.COLORS['reset']}")
            
            practice_words = self._get_difficult_words(count)
            print(f"\n{self.COLORS['green']}✅ 已选择 {count} 个需要加强的单词进行练习{self.COLORS['reset']}")
        
        # 开始练习（新增停止逻辑）
        print(f"\n{self.COLORS['cyan']}💡 练习过程中输入'E'可随时停止练习{self.COLORS['reset']}")
        input(f"{self.COLORS['cyan']}按回车键开始练习...{self.COLORS['reset']}")
        
        correct_count = 0
        wrong_words = []
        stop_practice = False
        start_time = time.time()
        
        for i, word_data in enumerate(practice_words, 1):
            if stop_practice:
                break
                
            print(f"\n{self.COLORS['bold']}{self.COLORS['blue']}单词 {i}/{count}{self.COLORS['reset']}")
            is_correct, stop_practice = self.check_spelling(word_data, quiz_mode=False)
            if is_correct and not stop_practice:
                correct_count += 1
            elif not is_correct and not stop_practice:
                wrong_words.append(word_data)
        
        # 练习统计
        end_time = time.time()
        elapsed_time = round(end_time - start_time, 2)
        completed_count = i if stop_practice else count
        accuracy = round(correct_count / completed_count * 100, 1) if completed_count > 0 else 0
        
        if stop_practice:
            print(f"\n{self.COLORS['yellow']}⚠️  拼写练习已停止{self.COLORS['reset']}")
        else:
            print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}📊 练习完成！{self.COLORS['reset']}")
        
        print("=" * 50)
        print(f"已完成单词数：{completed_count}/{count}")
        print(f"正确拼写：{self.COLORS['green']}{correct_count}{self.COLORS['reset']}")
        print(f"错误拼写：{self.COLORS['red']}{completed_count - correct_count}{self.COLORS['reset']}")
        print(f"正确率：{self.COLORS['bold']}{accuracy}%{self.COLORS['reset']}")
        print(f"用时：{elapsed_time} 秒")
        
        self._save_learning_records()
        
        # 重点单词练习完成后，询问是否重做错题（仅完成时）
        if mode_choice == '3' and not stop_practice and wrong_words:
            while True:
                redo_choice = input(f"\n{self.COLORS['yellow']}是否要重做这些错误的重点单词？(y/n): {self.COLORS['reset']}").strip().lower()
                if redo_choice in ['y', 'n']:
                    break
                print(f"{self.COLORS['red']}✗ 请输入 y 或 n{self.COLORS['reset']}")
            
            if redo_choice == 'y':
                self._redo_wrong_words(wrong_words)
        else:
            # 其他模式询问是否继续练习（仅完成时）
            if not stop_practice:
                while True:
                    choice = input(f"\n{self.COLORS['yellow']}请选择：1-继续练习  2-返回上一页  3-返回主菜单: {self.COLORS['reset']}").strip()
                    if choice in ['1', '2', '3']:
                        break
                    print(f"{self.COLORS['red']}✗ 请输入 1、2 或 3{self.COLORS['reset']}")
                
                if choice == '1':
                    self.start_practice()
                elif choice == '2':
                    print(f"{self.COLORS['blue']}🔙 返回练习模式选择页...{self.COLORS['reset']}")
                    self.start_practice()
                else:
                    print(f"{self.COLORS['blue']}🔙 返回主菜单...{self.COLORS['reset']}")
                    input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
            else:
                # 停止练习时直接返回主菜单
                input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
    
    def show_word_list(self):
        """查看单词列表（支持分页、搜索、筛选，新增Equal列展示）"""
        print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}📖 英语单词列表 {self.COLORS['reset']}")
        print("=" * 60)
        
        if not self.words_data:
            print(f"{self.COLORS['red']}✗ 暂无单词数据可显示！{self.COLORS['reset']}")
            input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
            return
        
        # 按字母排序
        sorted_words = sorted(self.words_data, key=lambda x: x['english'])
        
        page_size = 15
        total_pages = (len(sorted_words) + page_size - 1) // page_size
        current_page = 1
        
        while True:
            # macOS 清屏命令适配
            if os.name == 'nt':
                os.system('cls')
            else:  # macOS/Linux
                os.system('clear')
            
            print(f"{self.COLORS['bold']}{self.COLORS['purple']}📖 英语单词列表 (共{len(sorted_words)}个单词，第{current_page}/{total_pages}页) {self.COLORS['reset']}")
            print("=" * 80)
            
            # 表头（新增Equal列）
            header_parts = [
                f"{self.COLORS['cyan']}{'序号':<4}",
                f"{'英语单词':<12}",
                f"{'中文释义':<15}",
                f"{'Equal':<15}"
            ]
            if self.has_mark_column:
                header_parts.append(f"{'标记状态':<12}")
            print("".join(header_parts) + self.COLORS['reset'])
            print("-" * 80)
            
            # 显示当前页单词
            start_idx = (current_page - 1) * page_size
            end_idx = min(start_idx + page_size, len(sorted_words))
            
            for i in range(start_idx, end_idx):
                idx = i + 1
                word = sorted_words[i]
                english = word['english'][:10] + '...' if len(word['english']) > 10 else word['english']
                chinese = word['chinese'][:13] + '...' if len(word['chinese']) > 13 else word['chinese']
                equal = word['equal'][:13] + '...' if len(word['equal']) > 13 else word['equal']
                
                # 构建行内容
                row_parts = [f"{idx:<4}", f"{english:<12}", f"{chinese:<15}", f"{equal:<15}"]
                
                # 标记状态
                if self.has_mark_column:
                    if word.get('mark') in self.target_marks:
                        mark_status = f"{self.COLORS['red']}重点单词{self.COLORS['reset']}"
                    else:
                        mark_status = "普通单词"
                    row_parts.append(f"{mark_status:<12}")
                
                print("".join(row_parts))
            
            print("\n" + "=" * 80)
            # 操作提示
            operation_parts = [f"{self.COLORS['purple']}操作：{self.COLORS['reset']}"]
            operation_parts.append("n - 下一页 | p - 上一页 | q - 返回主菜单 | s - 搜索单词")
            if self.has_mark_column:
                operation_parts.append(" | m - 筛选重点单词")
            
            print("".join(operation_parts))
            
            # 处理用户输入
            choice = input(f"\n{self.COLORS['yellow']}请选择操作: {self.COLORS['reset']}").strip().lower()
            
            if choice == 'n' and current_page < total_pages:
                current_page += 1
            elif choice == 'p' and current_page > 1:
                current_page -= 1
            elif choice == 's':
                keyword = input(f"\n{self.COLORS['purple']}请输入要搜索的单词或中文释义: {self.COLORS['reset']}").lower()
                results = [word for word in sorted_words if keyword in word['english'].lower() or keyword in word['chinese'].lower() or keyword in word['equal'].lower()]
                
                print(f"\n{self.COLORS['cyan']}🔍 搜索结果 (共{len(results)}个):{self.COLORS['reset']}")
                if results:
                    for i, word in enumerate(results[:10], 1):
                        mark_tag = f" [{self.COLORS['red']}重点{self.COLORS['reset']}]" if word.get('mark') in self.target_marks else ""
                        print(f"   {i}. {word['english']} - {word['chinese']} - {word['equal']}{mark_tag}")
                    if len(results) > 10:
                        print(f"   ... 还有 {len(results)-10} 个匹配结果")
                else:
                    print(f"   {self.COLORS['red']}未找到匹配的单词{self.COLORS['reset']}")
                input(f"\n{self.COLORS['yellow']}按回车键继续...{self.COLORS['reset']}")
            elif choice == 'm' and self.has_mark_column:
                # 筛选重点单词
                focus_words = [word for word in sorted_words if word.get('mark') in self.target_marks]
                print(f"\n{self.COLORS['cyan']}🔍 重点单词筛选结果 (共{len(focus_words)}个):{self.COLORS['reset']}")
                if focus_words:
                    for i, word in enumerate(focus_words[:10], 1):
                        print(f"   {i}. {word['english']} - {word['chinese']} - {word['equal']} [{self.COLORS['red']}重点{self.COLORS['reset']}]")
                    if len(focus_words) > 10:
                        print(f"   ... 还有 {len(focus_words)-10} 个重点单词")
                else:
                    print(f"   {self.COLORS['red']}未找到标记为1的重点单词{self.COLORS['reset']}")
                input(f"\n{self.COLORS['yellow']}按回车键继续...{self.COLORS['reset']}")
            elif choice == 'q':
                break
            else:
                print(f"{self.COLORS['red']}✗ 无效操作或已到边界{self.COLORS['reset']}")
                time.sleep(1)
    
    def show_learning_stats(self):
        """查看学习统计数据"""
        print(f"\n{self.COLORS['bold']}{self.COLORS['purple']}📈 学习进度统计 {self.COLORS['reset']}")
        print("=" * 60)
        
        total_words = len(self.words_data)
        practiced_words = len(self.learning_records['words'])
        practiced_percent = round(practiced_words / total_words * 100, 1) if total_words > 0 else 0
        
        # 统计掌握情况
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
        
        # 总体练习统计
        total_practiced = self.learning_records['stats']['total_practiced']
        total_correct = self.learning_records['stats']['total_correct']
        overall_accuracy = round(total_correct / total_practiced * 100, 1) if total_practiced > 0 else 0
        
        # 显示总体情况
        print(f"{self.COLORS['blue']}📋 总体情况：{self.COLORS['reset']}")
        print(f"总单词数：{total_words}")
        print(f"已练习：{practiced_words} ({practiced_percent}%)")
        print(f"未练习：{total_words - practiced_words}")
        
        # 显示重点单词统计
        if self.has_mark_column:
            focus_count = len(self._get_all_target_mark_words())
            practiced_focus = 0
            for word in self._get_all_target_mark_words():
                if word['english'] in self.learning_records['words']:
                    practiced_focus += 1
            
            print(f"\n{self.COLORS['blue']}🏷️  重点单词统计（标记为1）：{self.COLORS['reset']}")
            print(f"重点单词总数：{self.COLORS['red']}{focus_count} 个{self.COLORS['reset']}")
            print(f"已练习重点单词：{practiced_focus} 个")
            print(f"未练习重点单词：{focus_count - practiced_focus} 个")
        
        # 显示掌握情况
        print(f"\n{self.COLORS['blue']}📊 掌握情况：{self.COLORS['reset']}")
        print(f"已掌握：{self.COLORS['green']}{mastered} 个{self.COLORS['reset']} (成功率≥90%)")
        print(f"学习中：{self.COLORS['yellow']}{learning} 个{self.COLORS['reset']} (60%≤成功率<90%)")
        print(f"需加强：{self.COLORS['red']}{need_practice} 个{self.COLORS['reset']} (成功率<60%)")
        
        # 显示练习统计
        print(f"\n{self.COLORS['blue']}📈 练习统计：{self.COLORS['reset']}")
        print(f"总练习次数：{total_practiced}")
        print(f"总正确次数：{total_correct}")
        print(f"总体正确率：{self.COLORS['bold']}{overall_accuracy}%{self.COLORS['reset']}")
        
        if self.learning_records['stats']['last_practice']:
            print(f"最后练习：{self.learning_records['stats']['last_practice']}")
        
        # 显示最近测验记录
        quiz_records = self.learning_records['stats'].get('quiz_records', [])
        if quiz_records and len(quiz_records) > 0:
            print(f"\n{self.COLORS['blue']}📝 最近测验记录：{self.COLORS['reset']}")
            recent_quizzes = quiz_records[-3:]
            for i, quiz in enumerate(recent_quizzes, 1):
                wrong_count = quiz.get('wrong_count', '未知')
                print(f"{i}. {quiz['date']} | {quiz['correct']}/{quiz['total']} | {quiz['accuracy']}% | "
                      f"{self.COLORS['red']}错误：{wrong_count}{self.COLORS['reset']} | {quiz['grade']}")
        
        # 显示最近练习的单词
        if self.learning_records['words']:
            print(f"\n{self.COLORS['blue']}🔍 最近练习的单词：{self.COLORS['reset']}")
            recent_words = sorted(
                self.learning_records['words'].items(),
                key=lambda x: x[1]['last_attempt'],
                reverse=True
            )[:5]
            
            for word, record in recent_words:
                mark_tag = ""
                if self.has_mark_column:
                    for wd in self.words_data:
                        if wd['english'] == word and wd.get('mark') in self.target_marks:
                            mark_tag = f" {self.COLORS['red']}[重点]{self.COLORS['reset']}"
                            break
                
                print(f"   {word:<10} | 正确率：{record['success_rate']}% | 练习次数：{record['attempts']}{mark_tag}")
        
        input(f"\n{self.COLORS['yellow']}按回车键返回主菜单...{self.COLORS['reset']}")
    
    def main_menu(self):
        """主菜单"""
        while True:
            # macOS 清屏命令适配
            if os.name == 'nt':
                os.system('cls')
            else:
                os.system('clear')
            
            # 显示主菜单
            print(f"{self.COLORS['bold']}{self.COLORS['purple']}🎯 英语拼写检查练习程序 {self.COLORS['reset']}")
            print("=" * 50)
            print(f"📝 当前单词库：{self.COLORS['green']}{len(self.words_data)} 个单词{self.COLORS['reset']}")
            
            # 显示重点单词数量
            if self.has_mark_column:
                focus_count = len(self._get_all_target_mark_words())
                print(f"🏷️  重点单词（标记为1）：{self.COLORS['red']}{focus_count} 个{self.COLORS['reset']}")
                print(f"💡 提示：拼写错误的单词会自动标记为重点单词，练习时输入'E'可停止")
            
            print("=" * 50)
            print("1. 开始拼写练习")
            print("2. 100题随机测验")
            print("3. 查看单词列表")
            print("4. 查看学习进度")
            print("5. 退出程序")
            print("=" * 50)
            
            # 处理用户选择
            choice = input(f"\n{self.COLORS['yellow']}请选择操作 (1-5): {self.COLORS['reset']}").strip()
            
            if choice == '1':
                self.start_practice()
            elif choice == '2':
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

if __name__ == "__main__":
    try:
        checker = EnglishSpellingChecker()
        checker.main_menu()
    except ImportError as e:
        print("❌ 缺少必要的依赖包，请运行以下命令安装：")
        print("pip3 install pandas openpyxl")
        print(f"\n错误详情：{e}")
        input("按回车键退出...")
    except Exception as e:
        print(f"❌ 程序运行出错：{type(e).__name__}: {str(e)}")
        import traceback
        traceback.print_exc()
        input("按回车键退出...")