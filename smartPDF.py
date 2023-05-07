import os  # OSモジュールをインポート
import openai  # OpenAIモジュールをインポート
import pdfplumber  # pdfplumberモジュールをインポート
from pathlib import Path  # pathlibモジュールからPathをインポート
import re  # 正規表現モジュールをインポート
import tkinter as tk  # tkinterモジュールをtkとしてインポート
from tkinter import messagebox  # tkinterモジュールからmessageboxをインポート
from tkinter import filedialog  # フォルダ選択ダイアログをインポート
import pytesseract # OCRモジュールをインポート
from pdf2image import convert_from_path # PDFから画像を生成するモジュールをインポート
import PyPDF4 # PyPDF4モジュールをインポート
import configparser # configparserモジュールをインポート

config = configparser.ConfigParser() # configparserをインスタンス化

if os.path.exists('settings.ini'): # settings.iniが存在する場合は読み込む
    config.read('settings.ini')
else: # settings.iniが存在しない場合は作成する
    config['DEFAULT'] = {'OpenAI_API_KEY': '', 'input_folder_path': '', 'output_folder_path': '', 'path_tesseract': 'C:\Program Files\Tesseract-OCR', 'path_poppler': Path(__file__).parent.absolute() / "poppler/bin"}
    with open('settings.ini', 'w') as configfile:
        config.write(configfile)

# OpenAI APIキーが設定されていない場合は設定を促すメッセージダイアログを表示
if config.get('DEFAULT', 'OpenAI_API_KEY') == '':
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(title='お知らせ', message='OpenAI APIキーが設定されていません。settings.iniにAPIキーを設定してください。')
    # settings.iniを開く
    os.system('notepad.exe settings.ini')
    exit()
openai.api_key = config.get('DEFAULT', 'OpenAI_API_KEY') # OpenAI APIキーを設定

# poppler/binを環境変数PATHに追加する
poppler_dir = config.get('DEFAULT', 'path_poppler')
os.environ["PATH"] += os.pathsep + str(poppler_dir)

# Tesseract OCRをインストールしたパスを環境変数PATHに追加する
path_tesseract = config.get('DEFAULT', 'path_tesseract')
os.environ["PATH"] += os.pathsep + path_tesseract

def preprocess_text(text):  # テキストの前処理を行う関数
    text = re.sub(r'[!-@[-`{-~]', '', text)  # 記号を削除
    text = re.sub(r'\d', '', text)  # 数字を削除
    return text

def is_searchable_pdf(pdf_path): # PDFがサーチャブルかどうかを判定する関数
    with open(pdf_path, 'rb') as f:
        reader = PyPDF4.PdfFileReader(f)
        for page_num in range(reader.getNumPages()):
            page = reader.getPage(page_num)
            text = page.extractText()
            if text:  # ページからテキストが抽出できた場合、サーチャブルなPDFと判断
                return True
    return False  # テキストが抽出できなかった場合、非サーチャブルなPDFと判断

def extract_text_from_serchable_pdf(pdf_path):  # サーチャブルPDFからテキストを抽出する関数
    with pdfplumber.open(pdf_path) as pdf:  # pdfファイルを開く
        text = ""
        for page in pdf.pages:  # ページごとにテキストを抽出
            text += page.extract_text()
    return text

def extract_text_from_image_pdf(pdf_path): # OCRを使用して日本語テキストを抽出する関数
    images = convert_from_path(pdf_path)  # PDFを画像に変換
    text = ""
    for image in images:
        text += pytesseract.image_to_string(image, lang='jpn')  # 画像から日本語テキストを抽出
    return text

def extract_text_from_pdf(pdf_path):
    if is_searchable_pdf(pdf_path):  # PDFがサーチャブルな場合
        return extract_text_from_serchable_pdf(pdf_path)  # サーチャブルPDFからテキストを抽出
    else:  # PDFが非サーチャブルな場合
        return extract_text_from_image_pdf(pdf_path)  # OCRを使用して日本語テキストを抽出

def generate_title_with_chatgpt(system_prompt, user_prompt, model='gpt-3.5-turbo'):  # ChatGPTを使って題名を生成する関数
    try:
        response = openai.ChatCompletion.create(  # ChatGPT APIにリクエストを送る
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ]
        )
    except openai.error.RateLimitError:
        messagebox.showerror(title='お知らせ', message='ファイルの読み込みが上限に達しました。しばらくしてから再度実行してください。')
        return
    title = response.choices[0].message['content'].strip()  # レスポンスから題名を取得
    return title

def rename_pdf_files(input_folder, output_folder):  # PDFファイルのファイル名をリネームする関数
    os.makedirs(output_folder, exist_ok=True)  # 出力フォルダを作成（既に存在する場合は何もしない）
    renamed_files = []  # リネームされたファイルのリストを初期化
    words_count = config.get('DEFAULT', 'words_count')  # 送信する文字数を取得

    for pdf_file in Path(input_folder).glob("*.pdf"):  # 入力フォルダ内のPDFファイルを順番に処理
        text = extract_text_from_pdf(str(pdf_file)).replace('\n', '')  # テキストを抽出し、改行を削除
        # textの冒頭から指定した文字数までを抽出
        text_to_send = text[:int(words_count)]
        system_prompt = f"あなたは、短くてわかりやすいファイル名を提案する専門家です。ファイル名は日本語で、体言止めで表現し、簡潔にしてください。"  # システムプロンプトを生成
        user_prompt = f"このPDF文書の内容に基づいて、適切なファイル名を提案してください。内容は次のとおりです: {text_to_send}..."  # ユーザープロンプトを生成
        print("ChatGPT APIを呼び出し中です・・・")
        title = generate_title_with_chatgpt(system_prompt, user_prompt)  # ChatGPTを使って題名を生成
        if(title is None):  # 題名が生成できなかった場合
            break  # 処理を終了
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
    summary = "変更済みのファイル:\n"
    for old_name, new_name in renamed_files:  # リネームされたファイルのリストから、古いファイル名と新しいファイル名を取得
        summary += f"{old_name} \n-> {new_name}\n\n"

    summary += "\nファイル名の変更が完了しました"

    # ポップアップウィンドウを表示
    root = tk.Tk()  # tkinterのルートウィンドウを作成
    root.withdraw()  # ルートウィンドウを非表示にする
    messagebox.showinfo("Summary", summary)  # サマリーメッセージをポップアップウィンドウで表示

def browse_input_directory(label):
    folder_path = filedialog.askdirectory()
    label.config(text=folder_path) 
    # 設定ファイルに入力フォルダのパスを保存
    config.set('DEFAULT', 'input_folder_path', folder_path)
    with open('settings.ini', 'w') as config_file:
        config.write(config_file)
    return folder_path

def browse_output_directory(label):
    folder_path = filedialog.askdirectory()
    label.config(text=folder_path) 
    # 設定ファイルに出力フォルダのパスを保存
    config.set('DEFAULT', 'output_folder_path', folder_path)
    with open('settings.ini', 'w') as config_file:
        config.write(config_file)
    return folder_path

def create_main_window():
    root = tk.Tk()  # ルートウィンドウを作成
    root.title("SmartPDF")  # ウィンドウのタイトルを"SmartPDF"に変更

    # 入力フォルダ選択用のラベルとボタンを作成
    input_label = tk.Label(root, text="入力フォルダを選択してください")
    input_label.grid(row=0, column=0, sticky='w', padx=(10, 5), pady=(10, 5))
    input_folder_button = tk.Button(root, text="参照", command=lambda: browse_input_directory(input_folder_path_label))
    input_folder_button.grid(row=0, column=1, sticky='w', padx=(5, 10), pady=(10, 5))
    if(config.get('DEFAULT', 'input_folder_path')): # 設定ファイルに入力フォルダのパスが保存されている場合は、そのパスをラベルに表示
        input_folder_path_label = tk.Label(root, text=config.get('DEFAULT', 'input_folder_path'))
    else: # 設定ファイルに入力フォルダのパスが保存されていない場合は、空のラベルを表示
        input_folder_path_label = tk.Label(root, text="")
    input_folder_path_label.grid(row=0, column=2, sticky='w', padx=(10, 10), pady=(10, 5))

    # 出力フォルダ選択用のラベルとボタンを作成
    output_label = tk.Label(root, text="出力フォルダを選択してください")
    output_label.grid(row=1, column=0, sticky='w', padx=(10, 5), pady=(5, 5))
    output_folder_button = tk.Button(root, text="参照", command=lambda: browse_output_directory(output_folder_path_label))
    output_folder_button.grid(row=1, column=1, sticky='w', padx=(5, 10), pady=(5, 5))
    if(config.get('DEFAULT', 'output_folder_path')): # 設定ファイルに出力フォルダのパスが保存されている場合は、そのパスをラベルに表示
        output_folder_path_label = tk.Label(root, text=config.get('DEFAULT', 'output_folder_path'))
    else: # 設定ファイルに出力フォルダのパスが保存されていない場合は、空のラベルを表示
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
