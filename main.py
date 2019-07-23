# -*- coding: utf-8 -*-
import re
import mojimoji
import numpy as np
from collections import Counter
from googleapiclient.errors import HttpError
from config import *
from spreadsheet_api import read_values, batch_update_values, batch_update, \
                            make_merge_request, make_update_values_data


def clean_up(text):
    """テキストを半角に変換し，スペースを削除する"""
    text = mojimoji.zen_to_han(text, kana=False)  # 全角を半角に変換
    text = text.replace(' ', '')  # スペースを削除
    return text


def convert_to_alphabet(column_num):
    """列番号を列名に変換する(ex. 1→A, 2→B, 3→C)"""
    i = int((column_num - 1) / 26)
    j = int(column_num - (i * 26))
    alpha = ''
    for k in i, j:
        if k != 0:
            alpha += chr(k + 64)
    return alpha


def get_name2column():
    """名前と列の対応辞書"""
    names = read_values(SHIFT_SPREADSHEET_ID, SHIFT_NAME_RANGE)[0]
    name2column = {}
    for i, name in enumerate(names):
        name2column[clean_up(name)] = [i+2, convert_to_alphabet(i+3)]  # TODO: Change values accoding to your shiftsheet
    return name2column


NAME2COLUMN = get_name2column()


def get_row_numbers(ans):
    """1つの回答から対応する行番号リストを返す"""
    if not ans:
        # 欠席なし->何もしない
        return None
    elif ans in ['準備日全日参加できない', '1日目全日参加できない', '2日目全日参加できない', '片付日全日参加できない']:  # TODO: Change candidates accoding to your shiftsheet
        # 全時間帯の行番号を返す
        return range(min(TIME2ROW.values()), max(TIME2ROW.values())+1)
    else:
        # 回答された時間帯の行番号を返す
        try:
            return [TIME2ROW[time] for time in ans.replace(' ', '').split(',')]
        except:
            raise ValueError(f'不正な回答です: "{ans}"')


def make_requests(name, ans, sheet_name):
    """×を埋めるためのAPIリクエストを生成"""
    # 行番号を取得
    row_nums = get_row_numbers(ans)
    # 回答がなければリクエストは生成しなくていい
    if not row_nums:
        return None, None
    # 連続の時間帯をまとめる
    indices = [idx for idx, num in enumerate(row_nums) if num-1 != row_nums[idx-1] and idx > 0]
    # indicesで指定した位置でsplitする
    split_row_nums = np.split(np.array(row_nums), indices)

    sheet_id = SHIFT_SHEETNAME2SHEETID[sheet_name]
    start_column = NAME2COLUMN[name][0]
    end_column = start_column + 1
    column = NAME2COLUMN[name][1]

    merge_requests = []
    update_values_data = []
    for nums in split_row_nums:
        # セルの結合
        start_row = int(nums[0]) - 1
        end_row = int(nums[-1])
        merge_requests.extend(make_merge_request(sheet_id, start_row, end_row, start_column, end_column))
        # データの書き込み
        range_name = f'{sheet_name}!{column}{nums[0]}'
        values = [['×']]
        update_values_data.extend(make_update_values_data(range_name, values))
    return merge_requests, update_values_data


def register_absences():
    """シフトシートに欠席を登録する"""
    messages = ''

    # アンケートシート全体の値を取得
    answer_data = read_values(ANSWER_SPREADSHEET_ID, ANSWER_ENABLE_RANGE)[1:]
    name_list = [clean_up(line[ANSWER_NAME_COLUMN_NUM]) for line in answer_data]
    name_counter = Counter(name_list)
    all_merge_request = []
    all_update_value_data = []

    for line in answer_data:
        name = clean_up(line[ANSWER_NAME_COLUMN_NUM])
        fri_ans = line[ANSWER_FRI_COLUMN_NUM] if len(line) > ANSWER_FRI_COLUMN_NUM else None
        sat_ans = line[ANSWER_SAT_COLUMN_NUM] if len(line) > ANSWER_SAT_COLUMN_NUM else None
        sun_ans = line[ANSWER_SUM_COLUMN_NUM] if len(line) > ANSWER_SUM_COLUMN_NUM else None
        mon_ans = line[ANSWER_MON_COLUMN_NUM] if len(line) > ANSWER_MON_COLUMN_NUM else None
        ans_list = [fri_ans, fri_ans, sat_ans, sat_ans, sun_ans, sun_ans, mon_ans, mon_ans]

        if name not in NAME2COLUMN.keys():
            print(f'該当しない名前です: {name}')
            messages += f'該当しない名前です: {name}\n'
            continue

        # 名前が複数あったら後の回答を採用する
        if name_counter[name] > 1:
            name_counter[name] -= 1
            continue

        assert len(SHIFT_SHEETNAME2SHEETID) == len(ans_list)

        for sheetname, ans in zip(SHIFT_SHEETNAME2SHEETID.keys(), ans_list):
            merge_requests, update_values_data = make_requests(name, ans, sheetname)
            if not merge_requests or not update_values_data:
                continue
            all_merge_request.extend(merge_requests)
            all_update_value_data.extend(update_values_data)

    # 全リクエストを一斉送信
    try:
        batch_update(SHIFT_SPREADSHEET_ID, all_merge_request)
    except HttpError as e:
        # 既にセルが結合されている場合に出るエラーの出力
        error_request_index_pattern = '.*Invalid requests\[(\d+)\].mergeCells.*'
        result = re.match(error_request_index_pattern, str(e))
        error_request_index = int(result.group(1))
        error_request = all_merge_request[error_request_index]
        sheet_id = error_request['mergeCells']['range']['sheetId']
        sheet_name = [k for k, v in SHEETNAME2ID.items() if v == sheet_id][0]
        start_column = error_request['mergeCells']['range']['startColumnIndex']
        error_column_name = [k for k, v in NAME2COLUMN.items() if v[0] == start_column][0]
        raise ValueError("セルの結合に失敗しました. 以前のセル結合が残っている場合は解除してから実行してください. シート名: '{}', 該当指名: '{}'".format(
            sheet_name, error_column_name))

    batch_update_values(SHIFT_SPREADSHEET_ID, all_update_value_data)
    print('欠席を登録しました．')

    return messages



if __name__ == '__main__':
    messages = register_absences()
