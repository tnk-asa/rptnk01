import os
import pandas as pd
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

# === 設定 ===
os.chdir(os.path.dirname(os.path.abspath(__file__)))
csv_path = 'fuguai_list.csv'  # ここにCSVのファイルパスを指定
try:
    df = pd.read_csv(csv_path)
except FileNotFoundError:
    print(f"エラー：ファイル '{csv_path}' が見つかりません。")
    exit()
    
keyword = '心臓が止まる'  # 例: '骨'
top_n = 50

output_csv = os.path.join(os.getcwd(), 'search_results.csv')



# === 対象カラム ===
search_columns = ['definition', 'inclusion', 'preferred', 'pronunciation']
output_columns = ['code', 'preferred', 'definition', 'similarity']

# === データ読み込みと整形 ===
df = pd.read_csv(csv_path)
df['__combined_text'] = df[search_columns].fillna('').astype(str).agg(' '.join, axis=1)
texts = df['__combined_text'].tolist()

# === モデル読み込み（日本語BERT）===
model = SentenceTransformer('cl-tohoku/bert-base-japanese')

# ベクトル変換
text_embeddings = model.encode(texts, convert_to_tensor=True)
query_embedding = model.encode([keyword], convert_to_tensor=True)

# 類似度計算
raw_scores = cosine_similarity(query_embedding.cpu().numpy(), text_embeddings.cpu().numpy())[0]

# === ブースト補正処理 ===
boosted_scores = []
for i, text in enumerate(texts):
    score = raw_scores[i]
    if keyword == text:
        score += 0.5  # 完全一致ブースト
    elif keyword in text:
        score += 0.2  # 部分一致ブースト
    boosted_scores.append(score)

# 上位抽出（ブースト後スコア）
top_indices = sorted(range(len(boosted_scores)), key=lambda i: boosted_scores[i], reverse=True)[:top_n]
top_df = df.iloc[top_indices].copy()
top_df['similarity'] = [round(boosted_scores[i], 4) for i in top_indices]

# 出力列の整形
result_df = top_df[output_columns]

# === CSV出力（UTF-8 BOM付きでExcel対応）===
result_df.to_csv(output_csv, index=False, encoding='utf-8-sig')

print(f"上位 {top_n} 件を '{output_csv}' に出力しました。")
