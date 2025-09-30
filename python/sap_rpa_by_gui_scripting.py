#ライブラリはアルファベット順に記載する
import configparser
import datetime
import os
import psutil
import subprocess
import sys
import time
import traceback

from logging import Formatter, getLogger, INFO, ERROR, StreamHandler
from logging.handlers import TimedRotatingFileHandler, SMTPHandler
from pathlib import Path

#カレントディレクトリをpythonファイルの階層に移動する
os.chdir(os.path.dirname(os.path.abspath(__file__)))

#ログ
LOG_DIR = os.path.dirname(__file__)+'/../logs/'
LOG_BACKUP_COUNT = 30
LOG_WHEN = 'MIDNIGHT'
LOG_ENCODING = 'utf-8'

#Configファイル名
ETL_RESULT_CONFIG_FILE = 'config.ini'

def _get_configparser(config_file:str, encoding:str='utf8'):
    """ConfigParserを取得する。
    ・大文字/小文字は区別しない

    Args:
        config_file: Configファイル
        encoding: 文字コード. Defaults to 'utf8'.

    Returns:
        conf: ConfigParser
    """
    conf = configparser.ConfigParser()
    #conf.optionxform = str #大文字/小文字を区別する場合はコメントアウトを外す
    conf.read(config_file, encoding=encoding)
    return conf

def _connect_to_SAP(params:dict):

    try:
        #SAP GUIを起動してログイン
        user_id = params["USER"]
        user_password = params["PASSWORD"]
        language = params["LANGUAGE"]
        sys_id = params["SID"]
        client = params["CLIENT"]

        # SAPshcut.exeのpath
        #sapshcut_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe"
        sapshcut_path = params["SAPSHCUT_PATH"]

        # コマンド実行してSAPログオン
        subprocess.run([sapshcut_path, f'-system={sys_id}', f'-client={client}', f'-language={language}', f'-user={user_id}', f'-pw={user_password}'] )

    except Exception as e:
        logger.error('Failed to connect to SAP')
        raise

def _get_vbscript_path(params:dict):
    """vbscriptのpathを取得する。
    ・相対パスを絶対パスに変換する。

    Args:
        params: 設定値
    Returns:
        script_path: vbscriptのpath
    """
    try:
        file_name = params["SCRIPT_FILE"]
        abs_path = os.path.abspath(__file__)
        parent_dir = os.path.dirname(os.path.dirname(abs_path))
        script_directory = str(parent_dir).replace(os.sep, '/')
        os.makedirs(script_directory, exist_ok=True)
        script_path = f'{script_directory}/scripts/{file_name}'
        return script_path

    except Exception as e:
        logger.error('Failed to get vbscript path')
        raise

def _get_output_directory(params:dict):
    """出力フォルダを取得する。
    ・相対パスを絶対パスに変換する。
    ・対象のフォルダが無い場合は作成する。

    Args:
        params: 設定値
    Returns:
        output_directory: 出力フォルダ
    """
    try:
        output_directory = params["OUTPUT_DIRECTORY"]
        abs_path = Path(output_directory).resolve()
        output_directory = str(abs_path).replace(os.sep, '/')
        os.makedirs(output_directory, exist_ok=True)
        return output_directory

    except Exception as e:
        logger.error('Failed to get output directory')
        raise

def _get_output_file_name(params:dict):
    """出力ファイル名を取得する。
    ・設定値に応じて、ファイル名の末尾に日時を追記する。
    ・拡張子は固定でXLSX

    Args:
        params: 設定値
    Returns:
        file_name: 出力ファイル名
    """
    try:
        extension = 'XLSX'
        if params["DOWNLOAD_WITH_DATE"]:
            file_name = f'{params["OUTPUT_FILE"]}_{datetime.datetime.now().strftime("%Y%m%d")}.{extension}'
        else:
            file_name = f'{params["OUTPUT_FILE"]}.{extension}'
        return file_name

    except Exception as e:
        logger.error('Failed to get output file name')
        raise

def _close_SAP_proc():
    """SAP Loginウィンドウを終了させる
    """
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and 'SAPLOGON.EXE' in proc.info['name'].upper():
            try:
                proc.terminate()
                proc.wait(timeout=3)
            except psutil.NoSuchProcess:
                pass
            except (psutil.AccessDenied, psutil.TimeoutError):
                proc.kill()
    logger.info('SAP Login process is closed')

def _close_Excel_proc():
    """起動中のExcelを停止する
    """
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and 'EXCEL.EXE' in proc.info['name'].upper():
            try:
                proc.terminate()  # 正常終了を試みる
                proc.wait(timeout=3)
            except psutil.NoSuchProcess:
                pass
            except psutil.AccessDenied:
                # アクセス権限がない場合は強制終了
                proc.kill()
    logger.info('All Excel process is closed')

#Configファイルを読み込む
config_file = os.path.dirname(__file__)+'/'+ETL_RESULT_CONFIG_FILE
conf = _get_configparser(config_file)
params = {}

def main():
    """SAPのRPA処理を実行する。
    ・スクリプトファイルが複数ある場合、1つ実行に失敗したら次のスクリプトファイルを実行する。
    """
    logger.info('Start SAP RPA')

    #[COMMON]
    section = 'COMMON'
    params["DOWNLOAD_WITH_DATE"] = conf.getboolean(section, "DOWNLOAD_WITH_DATE")

    #[SAP_CONNECTION]
    section = 'SAP_CONNECTION'
    params["SAPSHCUT_PATH"] = conf.get(section, "SAPSHCUT_PATH")
    params["SID"] = conf.get(section, "SID")
    params["CLIENT"] = conf.get(section, "CLIENT")
    params["LANGUAGE"] = conf.get(section, "LANGUAGE")
    params["USER"] = os.getenv("SAP_RPA_USER") # 環境変数で設定
    params["PASSWORD"] = os.getenv("SAP_RPA_PASSWORD") # 環境変数で設定

    # スクリプトを順に実行
    for i in range(10):
        #[SCRIPTxx]
        xx = i + 1
        section = 'SCRIPT%02d'%(xx)

        if conf.has_section(section) == False:
            continue

        params["SCRIPT_FILE"] = conf.get(section, "SCRIPT_FILE")
        params["OUTPUT_DIRECTORY"] = conf.get(section, "OUTPUT_DIRECTORY")
        params["OUTPUT_FILE"] = conf.get(section, "OUTPUT_FILE")

        try:
            # SAPを起動してログイン
            _connect_to_SAP(params)
            time.sleep(10)

            # スクリプトファイルの読み込み
            script_path = _get_vbscript_path(params)

            # 出力ディレクトリ、出力ファイル名、日付の取得
            # これらはVBスクリプト内で任意で使用する変数
            output_directory = _get_output_directory(params)
            file_name = _get_output_file_name(params)
            today = datetime.datetime.now().strftime("%d.%m.%Y")
            date_start = (datetime.datetime.now()-datetime.timedelta(days=30)).strftime("%d.%m.%Y")
            date_end = (datetime.datetime.now()+datetime.timedelta(days=30)).strftime("%d.%m.%Y")

            # スクリプトの実行
            # cscriptコマンドでVBスクリプトを実行する　※構文：cscript myscript.vbs param1 param2 param3
            # vbsのパス以降はコマンドライン引数を渡している
            # vbsスクリプト内では受け取った引数を「WScript.Arguments(n)」で使用する、nは0から始まる整数
            logger.info(f'Execute script {params["SCRIPT_FILE"]}')
            subprocess.run(['cscript', '//B', '//Nologo', script_path, output_directory, file_name, today, date_start, date_end])

            time.sleep(10)

            # SAP Loginウィンドウを閉じる
            _close_SAP_proc()

            # 開かれたExcelファイルを削除
            _close_Excel_proc()

            logger.info(f'Finshed script {params["SCRIPT_FILE"]}')
        except Exception as e:
            # スクリプトファイルの実行に失敗した場合、次のスクリプトを実行する。
            logger.error(f'Script execution failed\r\n{params["SCRIPT_FILE"]}')
            continue
        finally:
            time.sleep(1)

    # 終了カウント
    count = 5
    for i in range(count):
        time.sleep(1)
        print(f'finish: {count - i}sec')

    print('end')

def create_logger():
    """ログを設定する。
    ・ログ出力フォルダが存在しない場合、作成する。

    Returns:
        logger: ロガー
    """

    #[LOGGER_MAIL]
    section = 'LOGGER_MAIL'
    params["SMTP_SERVER"] = conf.get(section, "SMTP_SERVER")
    params["SMTP_PORT"] = conf.getint(section, "SMTP_PORT")
    params["FROM_ADDRESS"] = conf.get(section, "FROM_ADDRESS")
    params["TO_ADDRESSES"] = conf.get(section, "TO_ADDRESSES").split(',')
    params["SUBJECT"] = conf.get(section, "SUBJECT")

    os.makedirs(LOG_DIR, exist_ok=True)
    log_filename = LOG_DIR+'sap_rpa_result.log'

    #タイムローテーションファイルハンドラの設定
    handler = TimedRotatingFileHandler(log_filename,
                                        when=LOG_WHEN,
                                        backupCount=LOG_BACKUP_COUNT,
                                        encoding=LOG_ENCODING)

    #フォーマッタの作成
    formatter=Formatter("%(asctime)s %(levelname)-8s %(message)s")
    handler.setFormatter(formatter)

    #ロガーの作成
    logger = getLogger(__name__)
    logger.addHandler(handler)
    logger.setLevel(INFO)

    # SMTPHandlerの設定：アラートメールの発信
    smtp_handler = SMTPHandler(
        mailhost = (params["SMTP_SERVER"], params["SMTP_PORT"]), # SMTPサーバーのアドレスとポート
        fromaddr = params["FROM_ADDRESS"], # 送信元メールアドレス
        toaddrs = params["TO_ADDRESSES"], # 送信先メールアドレス
        subject = params["SUBJECT"], # メールタイトル
        secure=()
    )
    smtp_handler.setLevel(ERROR) # SMTPHandlerのログレベルをERRORに設定
    smtp_handler.setFormatter(formatter)  # フォーマッタを設定
    logger.addHandler(smtp_handler) # SMTPHandlerをロガーに追加

    # StreamHandlerの設定：標準出力への表示
    stream_handler = StreamHandler(sys.stdout)
    stream_handler.setLevel(INFO)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    return logger

if __name__ == "__main__":
    #ログ設定
    logger = create_logger()

    #RPA実行
    try:
        main()
    except Exception as e:
        print(traceback.format_exc())
        logger.error(traceback.format_exc())