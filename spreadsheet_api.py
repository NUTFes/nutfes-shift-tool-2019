# -*- coding: utf-8 -*-
import os
import pickle
import httplib2
import argparse
from apiclient import discovery
from oauth2client import client, tools
from oauth2client.file import Storage


SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'

if os.path.exists('flags.pkl'):
    with open('flags.pkl', 'rb') as f:
        flags = pickle.load(f)
else:
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
    with open('flags.pkl', 'wb') as f:
        pickle.dump(flags, f)


def get_credentials():
    """アカウント承認"""
    credential_dir = '.credentials'
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')
    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


credentials = get_credentials()
http = credentials.authorize(httplib2.Http())
discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)


def read_values(spreadsheet_id, range_name):
    """指定した範囲のセルの値を読み込み，2次元配列で返す"""
    # ref: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/get
    if ':' not in range_name:
        range_name = f'{range_name}:{range_name}'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    return values


def batch_update_values(spreadsheet_id, batch_data):
    """複数の範囲と値を一度に更新する"""
    # ref: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/batchUpdate
    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'valueInputOption': 'USER_ENTERED', 'data': batch_data}
    ).execute()


def batch_update(spreadsheet_id, batch_request):
    """複数の範囲と値を一度に更新する．更新する前に検証が行われ，要求が無効な場合はエラーを出す．"""
    # ref: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/batchUpdate
    result = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': batch_request}
    ).execute()


def make_update_values_data(range_name, values):
    """セルの値を更新するdataを作る"""
    data = [
        {
            'range': range_name,
            'majorDimension': 'ROWS',
            'values': values
        }
    ]
    return data


def make_merge_request(sheet_id, start_row, end_row, start_column, end_column):
    """セルの結合を行うrequestを作る"""
    ranges = {
        'sheetId': sheet_id,
        'startRowIndex': start_row,
        'endRowIndex': end_row,
        'startColumnIndex': start_column,
        'endColumnIndex': end_column
    }
    request = [
        {
            'mergeCells': {
                'range': ranges,
                'mergeType': 'MERGE_ALL'
            }
        },
        {
            'repeatCell': {
                'range': ranges,
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': {  # 背景色はグレー
                            'red': 0.9,
                            'green': 0.9,
                            'blue': 0.9
                        },
                        'horizontalAlignment': 'CENTER',
                        'verticalAlignment': 'MIDDLE'
                    }
                },
                'fields': '*'
            }
        }
    ]
    return request


def make_unmerge_request(sheet_id, start_row, end_row, start_column, end_column):
    """セルの結合を解除するrequestを作る"""
    ranges = {
        'sheetId': sheet_id,
        'startRowIndex': start_row,
        'endRowIndex': end_row,
        'startColumnIndex': start_column,
        'endColumnIndex': end_column
    }
    request = [
        {
            'repeatCell': {
                'range': ranges,
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 1.0,
                            'green': 1.0,
                            'blue': 1.0
                        },
                        'horizontalAlignment': 'LEFT',
                        'verticalAlignment': 'TOP'
                    }
                },
                'fields': '*'
            }
        },
        {
            'unmergeCells': {
                'range': ranges,
            }
        }
    ]
    return request
