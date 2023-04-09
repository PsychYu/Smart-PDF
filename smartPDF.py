import os  # OSモジュールをインポート
import openai  # OpenAIモジュールをインポート
import pdfplumber  # pdfplumberモジュールをインポート
from pathlib import Path  # pathlibモジュールからPathをインポート
import re  # 正規表現モジュールをインポート
import tkinter as tk  # tkinterモジュールをtkとしてインポート
from tkinter import messagebox  # tkinterモジュールからmessageboxをインポート

openai.api_key = os.environ["OPENAI_API_KEY"]  # 環境変数からOpenAI APIキーを取得

def preprocess_text(text):  # テキストの前処理を行う関数
    text = re.sub(r'[!-@[-`{-~]', '', text)  # 記号を削除
    text = re.sub(r'\d', '', text)  # 数字を削除
    return text

def extract_text_from_pdf(pdf_path):  # PDFファイルからテキストを抽出する関数
    with pdfplumber.open(pdf_path) as pdf:  # pdfファイルを開く
        text = ""
        for page in pdf.pages:  # ページごとにテキストを抽出
            text += page.extract_text()
    return text

def generate_title_with_chatgpt(prompt, model='gpt-3.5-turbo'):  # ChatGPTを使って題名を生成する関数
    response = openai.ChatCompletion.create(  # ChatGPT APIにリクエストを送る
        model=model,
        messages=[
            {"role": "user", "content": prompt},
        ]
    )

    title = response.choices[0].message['content'].strip()  # レスポンスから題名を取得
    return title

def rename_pdf_files(input_folder, output_folder):  # PDFファイルのファイル名をリネームする関数
    os.makedirs(output_folder, exist_ok=True)  # 出力フォルダを作成（既に存在する場合は何もしない）
    renamed_files = []  # リネームされたファイルのリストを初期化

    for pdf_file in Path(input_folder).glob("*.pdf"):  # 入力フォルダ内のPDFファイルを順番に処理
        text = extract_text_from_pdf(str(pdf_file)).replace('\n', '')  # テキストを抽出し、改行を削除
        title_prompt = f"この文書のファイル名を考えてください。理由等は不要です。体言止めで一言で鍵括弧は付けずにファイル名のみを返してください。: {text[:1000]}"  # プロンプトを生成
        title = generate_title_with_chatgpt(title_prompt)  # ChatGPTを使って題名を生成
        new_filename = f"{pdf_file.stem}_{title}{pdf_file.suffix}"  # 新しいファイル名を生成
        invalid_chars = set('<>:\"/\\|?*')  # 無効な文字のセットを定義
        new_filename = ''.join(c if c not in invalid_chars else '_' for c in new_filename)  # 無効な文字をアンダースコアに置き換え
        new_filepath = Path(output_folder, new_filename)  # 新しいファイルパスを生成
        pdf_file.rename(new_filepath)  # ファイル名を変更
        renamed_files.append((pdf_file.name, new_filename))  # リネームされたファイルのリストに追加

    return renamed_files

def display_summary(renamed_files):  # リネームのサマリーを表示する関数
    summary = "Renamed Files:\n"
    for old_name, new_name in renamed_files:  # リネームされたファイルのリストから、古いファイル名と新しいファイル名を取得
        summary += f"{old_name} -> {new_name}\n"

    summary += "\nProcess completed successfully!"

    # ポップアップウィンドウを表示
    root = tk.Tk()  # tkinterのルートウィンドウを作成
    root.withdraw()  # ルートウィンドウを非表示にする
    messagebox.showinfo("Summary", summary)  # サマリーメッセージをポップアップウィンドウで表示

if __name__ == "__main__":
    input_folder = "Write the input path here"  # 入力フォルダのパスを設定
    output_folder = "Write the output path here"  # 出力フォルダのパスを設定
    renamed_files = rename_pdf_files(input_folder, output_folder)  # PDFファイルのファイル名をリネーム
    display_summary(renamed_files)  # リネームのサマリーを表示

