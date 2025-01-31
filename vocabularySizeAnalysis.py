from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from nltk.stem import WordNetLemmatizer

import nltk
print(nltk.__version__)

import csv
csv_file_path = r'C:\ning\dev\dev\tool\dic.csv'  # 替换为你的 CSV 文件路径

# print(lemma) 
# exit()
# 下载需要的NLTK资源
# nltk.download('punkt', download_dir='/')
# nltk.download('punkt')
# nltk.download('wordnet')
# nltk.download('stopwords')
# nltk.download('punkt_tab')

# tokenizer = nltk.data.load('nltk:tokenizers/punkt/english.pickle')

# 验证
# text = "This is a test sentence."
# tokens = word_tokenize(text)
# print(tokens)

def load_csv(file_path):
    try:
        with open(file_path, mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            next(reader)  # 跳过表头
            
            word_dic = {}
            for row in reader:
                if len(row) >= 2:  # 确保行中有至少两列
                    word_dic[row[0].strip()] = row[1].strip()
            
        return word_dic
    
    except FileNotFoundError:
        return f"File '{file_path}' not found."
    except Exception as e:
        return f"An error occurred: {str(e)}"
csv_dict = load_csv(csv_file_path)

def find_value_in_csv(word_dict, key_to_find):
    if key_to_find in word_dict.keys():
        return word_dict[key_to_find]
    else:
        return "unknow"    

# 示例用法
# result = find_value_in_csv(csv_dic, "key")
# print(result)

def calculate_required_vocabulary(vocab_counts):
    # unknow 除く
    _t = vocab_counts.pop()
    # 定义词汇量级别
    vocab_levels = [500, 1000, 2000, 4000, 8000, 16000]  # A1, A2, B1, B2, C1, C2
    level_names = ['A1', 'A2', 'B1', 'B2', 'C1', 'C2']
    
    # 将输入词汇数量存储为列表
    # vocab_counts = [a1, a2, b1, b2, c1, c2]
    sum_vocab = sum(vocab_counts)
    
    # 从最高级别（C2）到最低级别（A1）依次判断
    for end_idx in range(len(vocab_levels) - 1, -1, -1):
        # 计算从当前级别到最高级别的词汇量总和
        total = sum(vocab_counts[end_idx:])
        if total / sum_vocab > 0.05:
            # 计算需要掌握的比例
            required_ratio = (total - sum_vocab * 0.05) / vocab_counts[end_idx]
            # 计算所需词汇量
            lower_level = vocab_levels[end_idx - 1] if end_idx > 0 else 0
            current_level = vocab_levels[end_idx]
            return lower_level + (current_level - lower_level) * required_ratio
    
    # 如果所有级别都未超过阈值，则返回 A1 水平
    return vocab_levels[0]

def analyze_vocab(text):
    # 1. 分词处理
    words = word_tokenize(text, "english")
    
    # 2. 去除标点符号和非单词字符
    words = [word.lower() for word in words if word.isalpha()]
    
    # 3. 词形还原
    lemmatizer = WordNetLemmatizer()
    lemmatized_word1 = [lemmatizer.lemmatize(word) for word in words]
    lemmatized_words = [lemmatizer.lemmatize(word, pos='v') for word in lemmatized_word1]
    
    # 4. 统计词汇量
    unique_words = set(lemmatized_words)
    
    # 5. 去除停用词（可选）
    stop_words = set(stopwords.words('english'))
    filtered_words = [word for word in unique_words if word not in stop_words]
    
    # 唯一单词処理
    level_dict = {"A1":0,"A2":0,"B1":0,"B2":0,"C1":0,"C2":0,"unknow":0}
    unknow_word_list = []
    for w in unique_words:
        w_level = find_value_in_csv(csv_dict, w)
        level_dict[w_level]+=1
        print(f"{w_level} {w} ({level_dict[w_level]})")
        if w_level=="unknow":
            unknow_word_list.append(w)
    print(unknow_word_list)
    unique_words_clean = len(unique_words) - len(unknow_word_list)

    # 比例计算
    level_ratios = {}
    for level, count in level_dict.items():
        level_ratios[level] = count / unique_words_clean
        print(f"{level}: {count}  {level_ratios[level] * 100:.4f}%")

    """
    A1 = 500
    A2 = 1,000
    B1 = 2,000
    B2 = 4,000
    C1 = 8,000
    C2 = 16,000
    """

    # 词汇量计算
    required_vocab = calculate_required_vocabulary(list(level_dict.values()))
    print(f"理解文章所需词汇量: {required_vocab}")

    return {
        "total_words": len(words),               # 总单词数
        "unique_words": len(unique_words),       # 唯一单词数
        "unique_non_stopwords": len(filtered_words)  # 唯一非停用词数量
    }

# 测试示例
text = """
The absence of parents may lead to feelings of loneliness, insecurity, or a lack of emotional support in children. 
They might lack role models, and their emotional needs may not be met promptly, which can affect their self-esteem and trust. 
Long-term absence of parental involvement can also cause issues in emotional and social development.
"""
text2 = """
Focusing on the specific situation where parents frequently compare their children with others, this article will address 
the emotional toll it takes on a student’s confidence and mental health. It will offer practical advice for talking to parents about 
feeling overwhelmed by comparisons, and how to set realistic expectations that align with one’s personal interests and strengths, 
rather than external measures of success."""
text3 = """
2. Title: Breaking Free from the Comparison Trap: Embrace Your Own Journey
Content Overview:
This piece will focus on the negative effects of being constantly compared to others, whether by parents or peers, and 
how it can create an unhealthy sense of competition. It will emphasize the importance of accepting your own path, learning to 
appreciate your unique strengths, and using healthy coping mechanisms like mindfulness and positive self-talk to counter feelings of 
inadequacy."""
result = analyze_vocab(text2)
print("分析结果:", result)
