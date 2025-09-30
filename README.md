バージョン
    1.0

SAP RPA
    SAPの定型業務を自動化するツール
    SAP GUI Scripting engineに接続し、SAPのクライアントソフトを操作する
    ・複数のSAP操作を記述することで、連続してRPAを実行する
    ・RPAの実行順、ファイルの保存先・名称、検索条件をカスタマイズできる
    ・異常発生時にはアラートメールを発信し、管理者が即座に気づける仕組みにする

利用前提
    ・RPAはデータの検索・ダウンロードのみに使用してください、登録・更新は禁止です
    ・RPAの実行時は他のSAPセッションを削除して起動してください、複数RPAを実行しないでください
        ※RPA実行時に既存のセッションがある場合、エラーが発生します
    ・RPAを自動で起動する際は、RPA専用のSAPアカウントを発行する必要があります
    ・SAPアカウントは2カ月に1回パスワード変更を強制するため、2か月以内にハンドでパスワード変更してください
        ※RPAによるパスワード入力ミスを繰り返すとパスワードが凍結されてしまいますのでご注意ください

ファイル構成
    sap_rpa
    ├── .venv                           ：python仮想環境
    ├── log                             ：RPA実行ログ
        └── sap_rpa_result.log
    ├── output                          ：SAPダウンロードデータの置場
        └── export.XLSX
    ├── python
        ├── config.ini                  ：プログラムの設定ファイル
        └── sap_rpa_by_gui_scripting.py ：pythonプログラム
    ├── scripts                         ：GUI scriptファイルの置場
        └── script.vsb
    ├── README.md
    └── requirements.txt                ：pythonライブラリリスト

SAP事前準備
    ・GUI Scriptをインストールする
        「GUI Script_740.exe」
        https://jppanasonic.sharepoint.com/:f:/s/TLA87/EpCNR0AJZbJKvS5-hhgHo-cBHZrPedTC4V8WvLfpdyUllw?e=BPDFh6
        ※Sharepoint onlineに保存しておりますが、万が一保存先が変更・削除されている場合はインターネットから取得ください
    ・SAP GUIにてGUI Scriptを下記の通り設定
        Options
        └── Accesibility & Scripting
            └── Scripting
                └── Enable scripting: on
                        Notify when script attaches to SAP GUI: off
                        Notify when script opens a connection: off
                        Show native Microsoft Windows dialogs: off
    ・SAPログインユーザ情報を環境変数に設定
        SETX SAP_RPA_USER "*****"
        SETX SAP_RPA_PASSWORD "*****"

Python環境構築
    ・Pythonのバージョンは「3.13.5」で動作確認しています（おそらく「3.12」以上必須）
        Python：https://www.python.org/
    ・sap_rpaフォルダ直下に「.venv」の名称で仮想環境を作成し、アクティベートしてください
        python -m venv .venv
        .venv\Scripts\activate
    ・Pythonライブラリはrequirements.txtでまとめてインストールしてください
        pip install -r requirements.txt --proxy=http://proxy.mei.co.jp:8080

config.ini設定
    ・LOGGER_MAIL           ：エラー発生時のアラートメールの設定
    ・SAP_CONNECTION        ：SAPの接続情報
    ・COMMON                ：RPA全体の共通設定
    ・SCRIPTxx              ：スクリプトごとの個別の設定、この順番に実行されます
        ・SCRIPT_FILE       ：Scriptsフォルダ内のスクリプトファイル名
        ・OUTPUT_DIRECTORY  ：出力先フォルダ
        ・OUTPUT_FILE       ：出力ファイル名

GUI scriptファイル作成
    Scriptファイルの作成
        始めに、SAP標準のGUI script機能でスクリプトのvbsファイルを作成し、scriptsフォルダに配置してください。
    動的要素の記述
        scriptファイル内で動的に変更したい箇所は下記を参考に記述してください。
        相対日付の日数はデフォルトで30日。
            ・出力ディレクトリ  ：WScript.Arguments(0)
            ・出力ファイル名    ：WScript.Arguments(1)
            ・当日              ：WScript.Arguments(2)
            ・相対日付（開始）   ：WScript.Arguments(3)
            ・相対日付（終了）   ：WScript.Arguments(4)

利用方法
    Execute.batから実行ください。pythonファイルを直接実行すると不具合の原因となります。

資料
    Python仮想環境の構築
    https://qiita.com/fiftystorm36/items/b2fd47cf32c7694adc2e
    PySapScriptドキュメント
    https://pypi.org/project/pysapscript/
    SAP GUI Script勉強会資料
    https://jppanasonic.sharepoint.com/:f:/s/TLA87/EpCNR0AJZbJKvS5-hhgHo-cBHZrPedTC4V8WvLfpdyUllw?e=BPDFh6