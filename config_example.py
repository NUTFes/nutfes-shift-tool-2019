### アンケート回答シート ###
# スプレッドシートID
ANSWER_SPREADSHEET_ID = ''
# 取得範囲
ANSWER_ENABLE_RANGE = 'A:M'  # Aから始める
# 名前の列番号
ANSWER_NAME_COLUMN_NUM = 1  # 0スタート
# 準備日の回答の列番号
ANSWER_FRI_COLUMN_NUM = 6
# 1日目の回答の列番号
ANSWER_SAT_COLUMN_NUM = 8
# 2日目の回答の列番号
ANSWER_SUM_COLUMN_NUM = 10
# 片付け日の回答の列番号
ANSWER_MON_COLUMN_NUM = 12


### シフトシート ###
# スプレッドシートID
SHIFT_SPREADSHEET_ID = ''
# シート名とシートIDの対応辞書
SHIFT_SHEETNAME2SHEETID = {
    '準備日': 0,
    '1日目晴': 0,
    '1日目雨': 0,
    '2日目晴': 0,
    '2日目雨': 0,
    '片付け日': 0,
}
# 名前の列の範囲
SHIFT_NAME_RANGE = 'C2:EU2'
# 時間帯と行番号の対応辞書を頑張って作る
TIME2ROW = {}
START_TIME = '6:00'
START_TIME_ROW = 3
END_TIME = '23:00'
END_TIME_ROW = 36
MINUTES_LIST = ['00', '30']
HOUR_LIST = list(range(int(START_TIME.split(':')[0]), int(END_TIME.split(':')[0])+1))
for i in range(END_TIME_ROW - START_TIME_ROW + 1):
    time = '{}:{}~{}:{}'.format(HOUR_LIST[i // 2],
                                MINUTES_LIST[i % 2],
                                HOUR_LIST[(i+1) // 2],
                                MINUTES_LIST[(i+1) % 2])
    TIME2ROW[time] = i + START_TIME_ROW


# 簡単な承認用パスワード
PASSWORD = ''
