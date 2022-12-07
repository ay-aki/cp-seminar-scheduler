# cp-seminar-scheduler

## 使用技術

以下の技術を使用

 - Python
 - numpy, pandas, gspread, Levenshtein, gspread_dataframe

## 何を目的として作ったか？

ゼミのスケジュール管理のために作成したものです。
主な機能は以下の３つです。
 - アンケートフォームで入力した情報により生成する資料(A)をまとめ、メールアドレス、氏名などの個人データ全体の進捗報告を一目で確認できる資料(C)を生成する（これは司会が画面共有で利用する）。
 - 資料(C)の内容を用いて出欠席をまとめたシートを追加した資料(C')を生成する。
 - 

資料(A)

![画像2](https://user-images.githubusercontent.com/55615907/206204264-b1f88e25-0fe1-4333-96ff-46f9f17d5454.png)

資料(B)

![画像3](https://user-images.githubusercontent.com/55615907/206204273-26d9bca0-89a5-4116-977a-bcc5a05890a9.png)

資料(C)

![画像4](https://user-images.githubusercontent.com/55615907/206204275-19dba4ca-a593-46ea-901c-69571c4038cb.png)

資料(C')

![画像5](https://user-images.githubusercontent.com/55615907/206204278-d6a6eed5-4533-4009-9a08-2bddc904b520.png)


## 参考資料

[1] https://www.twilio.com/blog/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python-jp
[2] https://programmer-life.work/python/gspread-write-spreadsheet#1
[3] https://docs.gspread.org/en/latest/user-guide.html#selecting-a-worksheet
[4] https://docs.gspread.org/en/v5.6.1/
    https://docs.gspread.org/en/v5.7.0/
[5] https://github.com/burnash/gspread/

