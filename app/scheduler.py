
"""
import 
"""
import os
import sys
import collections
import gspread
from   oauth2client.service_account import ServiceAccountCredentials
import numpy as np
import Levenshtein
import pandas as pd
from gspread_dataframe import set_with_dataframe


"""
名前の類似度を計算する関数
#   LDist    は、文字列間の距離を表す関数で、ここではLevenshtein距離を利用
#   LDistsv  は、文字と名前列を比較して距離をベクトルで返す。
#   argDistsvは、その添え字を返す。
"""
LDist       = lambda s1,s2: Levenshtein.distance(s1.lower(), s2.lower())
LDistsv     = lambda s, v : [min(LDist(s, __v) for __v in _v) for _v in v]
LDistsv2    = lambda s, v : [min(LDist(_s, _v) for _s in s) for _v in v]
argLDistsv  = lambda s, v : np.argmin(LDistsv(s, v))
argLDistsv2 = lambda s, v : np.argmin(LDistsv2(s, v))





"""

"""
def collect_progressreports(client, fn_mem, fn_rep, fn_reps, rep_insertWS):
    # open sheets
    f_mem  = client.open(fn_mem )
    f_rep  = client.open(fn_rep )
    f_reps = [client.open(f) for f in fn_reps]
    
    col_course, col_name = 1, 3 
    
    mem_course = f_mem.sheet1.col_values(col_course)
    mem_name   = [s.split(", ") for s in f_mem.sheet1.col_values(col_name)]
    
    rep_title  = [ws.title for ws in f_rep.worksheets()]
    
    for i in range(len(f_reps)):
        ws_name = rep_insertWS[i] # 挿入するＷＳの名前
        
        if ws_name in rep_title: 
            f_rep.del_worksheet(f_rep.worksheet(ws_name))
        
        f_rep.add_worksheet(title=ws_name, rows=100, cols=20)
        
        wso = f_reps[i].sheet1 # 書き込み禁止
        ws  = f_rep.worksheet(ws_name) # 書き込みOK
        
        wso_col_name = 4
        
        wso_values = wso.get_all_values()
        wso_name  = wso.col_values(wso_col_name) # name
        
        J = len(mem_name)
        
        cosine = np.array( [-1] + [min(LDistsv2(name, wso_name[1::])) for name in mem_name[1::]] )
        index  = np.array( [-1] + [argLDistsv2(name, wso_name[1::]) + 1 for name in mem_name[1::]] )
        
        count = collections.Counter(index)
        for key in count:
            if count[key] == 1: continue
            l = np.array(range(len(index)))[index == key]
            k = np.argmin(cosine[l])
            index[l[l != l[k]]] = -1
        
        lines = []
        headline = [
            "所属(Course)", 
            "記入名 [辞書登録名](Name)", 
            "出席(Attend)", 
            "進捗報告(Progress report)", 
            "その他 (other ex. absent reason etc.)"
        ]
        
        for ii, jj in enumerate(index):
            line = []
            course = mem_course[ii]
            if jj != -1: 
                l_values = wso_values[jj]
                line = line + l_values[4::]
                #name = "{0} [{1}]".format(l_values[3], mem_name[ii][0])
                name = "{0}".format(l_values[3])
            else:
                name = "[{0}]".format(mem_name[ii][0])
            
            line = [course, name] + line
            
            if ii == 0:
                line = headline
            lines += [line]
        
        
        index_c = list(set(range(1, len(wso_name))) - set(index)) # 補集合
        
        for ii, jj in enumerate(index_c):
            l_values = wso_values[jj]
            lines = lines + [l_values[2::]]
        
        
        n_inputrow  = len(wso_values)
        n_outputrow = len(lines)
        n_atenndees, n_absentees, n_unknown = 0, 0, 0
        for n in range(n_outputrow):
            if len(lines[n]) < 3:
                n_unknown    += 1
                continue
            n_atenndees += (lines[n][2] == "出席できる(can attend)")
            n_absentees += (lines[n][2] == "出席できない(can't attend)")
        
        
        line = "n_inputrow: {0}, n_outputrow: {1}, n_atenndees: {2}, n_absentees: {3} n_unknown: {4}".format(n_inputrow, n_outputrow, n_atenndees, n_absentees, n_unknown)
        
        lines = lines + [[line]]
        
        df_lines = pd.DataFrame(lines)
        
        set_with_dataframe(
            ws, 
            df_lines, 
            include_column_header=False
        )



