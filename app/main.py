

"""
import 
"""
import os
import sys
import collections
import numpy as np
import gspread
from   oauth2client.service_account import ServiceAccountCredentials
import numpy as np
import Levenshtein
import pandas as pd
from gspread_dataframe import set_with_dataframe
import scheduler



# create a client to interact with the Google Drive API
link1 = 'https://spreadsheets.google.com/feeds'
link2 = 'https://www.googleapis.com/auth/drive'
scope = [link1, link2]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'client_secret.json', scope
)
client= gspread.authorize(creds)



def main():
    # 情報を持つファイル
    fn_mem      = "MMALab_membership_list2022"
    # 進捗報告をまとめた情報を持つファイル
    fn_rep      = "progress_report_2022_b"
    # アンケートの回答ファイル, 指定ファイルの情報を基に更新
    fn_reps     = [ 
        #"[情報数理研究室]ゼミ (4/13)（回答）"
        #"[情報数理研究室]ゼミ (10/5)（回答）"
        #"[情報数理研究室]ゼミ (10/12) （回答）"
        #"[情報数理研究室]ゼミ (10/19)（回答）"
        #"[情報数理研究室]ゼミ (10/26)（回答）"
        #"[情報数理研究室]ゼミ (11/2)（回答）"
        #"[情報数理研究室]ゼミ (11/16)（回答）"
        #"[情報数理研究室]ゼミ (11/25)（回答）"
        "[情報数理研究室]ゼミ (11/30)（回答）"
    ]
    rep_insertWS = [
        #"2022/04/13"
        #"2022/10/05"
        #"2022/10/12"
        #"2022/10/19"
        #"2022/10/26"
        #"2022/11/2"
        #"2022/11/16"
        "2022/11/30"
    ]
    scheduler.collect_progressreports(client, fn_mem, fn_rep, fn_reps, rep_insertWS)



main()
input()
