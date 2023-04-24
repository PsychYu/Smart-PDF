import os  # OSモジュールをインポート
import openai  # OpenAIモジュールをインポート
import pdfplumber  # pdfplumberモジュールをインポート
from pathlib import Path  # pathlibモジュールからPathをインポート
import re  # 正規表現モジュールをインポート
import tkinter as tk  # tkinterモジュールをtkとしてインポート
from tkinter import messagebox  # tkinterモジュールからmessageboxをインポート
from tkinter import filedialog  # フォルダ選択ダイアログをインポート

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

def generate_title_with_chatgpt(system_prompt, user_prompt, model='gpt-3.5-turbo'):  # ChatGPTを使って題名を生成する関数
    response = openai.ChatCompletion.create(  # ChatGPT APIにリクエストを送る
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ]
    )

    title = response.choices[0].message['content'].strip()  # レスポンスから題名を取得
    return title

def rename_pdf_files(input_folder, output_folder):  # PDFファイルのファイル名をリネームする関数
    os.makedirs(output_folder, exist_ok=True)  # 出力フォルダを作成（既に存在する場合は何もしない）
    renamed_files = []  # リネームされたファイルのリストを初期化

    for pdf_file in Path(input_folder).glob("*.pdf"):  # 入力フォルダ内のPDFファイルを順番に処理
        text = extract_text_from_pdf(str(pdf_file)).replace('\n', '')  # テキストを抽出し、改行を削除
        system_prompt = f"あなたは、短くてわかりやすいファイル名を提案する専門家です。ファイル名は日本語で、体言止めで表現し、簡潔にしてください。"  # システムプロンプトを生成
        user_prompt = f"このPDF文書の内容に基づいて、適切なファイル名を提案してください。内容は次のとおりです: {text[:1000]}"  # ユーザープロンプトを生成
        print("ChatGPT APIを呼び出し中です・・・")
        title = generate_title_with_chatgpt(system_prompt, user_prompt)  # ChatGPTを使って題名を生成
        title = title.replace('「', '').replace('」', '')  # 鍵括弧を削除
        print("ChatGPT APIの呼び出しが完了しました。title: " + title)
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

def browse_directory(label):
    folder_path = filedialog.askdirectory()
    label.config(text=folder_path)
    return folder_path

def create_main_window():
    root = tk.Tk()  # ルートウィンドウを作成
    root.title("SmartPDF")  # ウィンドウのタイトルを"SmartPDF"に変更

    # 入力フォルダ選択用のラベルとボタンを作成
    input_label = tk.Label(root, text="入力フォルダを選択してください")
    input_label.grid(row=0, column=0, sticky='w', padx=(10, 5), pady=(10, 5))
    input_folder_button = tk.Button(root, text="参照", command=lambda: browse_directory(input_folder_path_label))
    input_folder_button.grid(row=0, column=1, sticky='w', padx=(5, 10), pady=(10, 5))
    input_folder_path_label = tk.Label(root, text="")
    input_folder_path_label.grid(row=0, column=2, sticky='w', padx=(10, 10), pady=(10, 5))

    # 出力フォルダ選択用のラベルとボタンを作成
    output_label = tk.Label(root, text="出力フォルダを選択してください")
    output_label.grid(row=1, column=0, sticky='w', padx=(10, 5), pady=(5, 5))
    output_folder_button = tk.Button(root, text="参照", command=lambda: browse_directory(output_folder_path_label))
    output_folder_button.grid(row=1, column=1, sticky='w', padx=(5, 10), pady=(5, 5))
    output_folder_path_label = tk.Label(root, text="")
    output_folder_path_label.grid(row=1, column=2, sticky='w', padx=(10, 10), pady=(5, 5))

    # 開始ボタンを作成し、start_process関数をコマンドとして設定
    start_button = tk.Button(root, text="開始", command=lambda: start_process(input_folder_path_label, output_folder_path_label))
    start_button.grid(row=2, column=0, columnspan=3, padx=(10, 10), pady=(10, 10))

    root.mainloop()  # ウィンドウのイベントループを開始

def start_process(input_label, output_label):
    input_folder = input_label.cget("text")  # 入力フォルダのパスを取得
    output_folder = output_label.cget("text")  # 出力フォルダのパスを取得
    renamed_files = rename_pdf_files(input_folder, output_folder)  # PDFファイルのファイル名をリネーム
    display_summary(renamed_files)  # リネームのサマリーを表示

if __name__ == "__main__":
    create_main_window()  # メインウィンドウを作成してイベントループを開始
