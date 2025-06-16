import os
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity



# === 設定 ===
os.chdir(os.path.dirname(os.path.abspath(__file__)))
csv_path = 'fuguai_list.csv'  # ここにCSVのファイルパスを指定
try:
    df = pd.read_csv(csv_path)
except FileNotFoundError:
    print(f"エラー：ファイル '{csv_path}' が見つかりません。")
    exit()

keyword = '骨折'  # 例: '骨'
top_n = 10  # 類似レコードの表示件数

# === 検索対象カラムと出力カラム ===
search_columns = ['definition', 'inclusion', 'preferred', 'pronunciation']
output_columns = ['code', 'preferred']

# === CSV読み込み ===
df = pd.read_csv(csv_path)

# 欠損値を空文字に変換し、検索対象列のテキストを結合
df['__combined_text'] = df[search_columns].fillna('').astype(str).agg(' '.join, axis=1)

# TF-IDF ベクトル化
vectorizer = TfidfVectorizer()
X = vectorizer.fit_transform(df['__combined_text'])

# クエリをベクトル化
query_vec = vectorizer.transform([keyword])

# コサイン類似度を計算
similarities = cosine_similarity(query_vec, X).flatten()

# 類似度が高い順に並べる
top_indices = similarities.argsort()[-top_n:][::-1]
top_df = df.iloc[top_indices].copy()
top_df['similarity'] = similarities[top_indices]

# 結果表示
print(top_df[output_columns + ['similarity']])
