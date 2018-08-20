" "
" "
"*********************************************"
"                                             "
"   PC END LOG GET                            "
"      - Time Card Hack!!                     "
"   Version 1.1    UPDATE 2018/08/12          "
"                                             "
"*********************************************"
" "
" "
" "
# System.Windows.Formsアセンブリを有効化
[void][System.Reflection.Assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=Neutral, PublicKeyToken=b77a5c561934e089")

# 抽出条件を入力させる
$start_date = Read-Host "開始日をyyyy-mm-dd形式で入力してください"
$end_date = Read-Host "終了日をyyyy-mm-dd形式で入力してください"

# 出力先のファイル名を指定させる
# OpenFileDialogクラスをインスタンス化し、必要な情報を設定
$dialog = New-Object System.Windows.Forms.SaveFileDialog 
$dialog.Filter = "CSVファイル(*.CSV)|*.CSV;"
$dialog.InitialDirectory = "C:\"
$dialog.Title = "ファイルを選択してください"
# ダイアログを表示
if($dialog.ShowDialog() -ne "OK"){
    "出力を中止しました。"
    exit 1
}



# LogParserを実行する
& 'C:\Program Files (x86)\Log Parser 2.2\LogParser.exe' -i:EVT -o:CSV -stats:OFF "SELECT TO_DATE(TimeWritten) AS DATE ,MAX(TimeWritten) AS TIME FROM 'System' WHERE ( TimeWritten BETWEEN TO_LOCALTIME(TIMESTAMP('$start_date 00:00:00', 'yyyy-MM-dd hh:mm:ss')) AND TO_LOCALTIME(TIMESTAMP('$end_date 00:00:00', 'yyyy-MM-dd hh:mm:ss')) )  AND ((EventID = 7002) or (EventID = 6006)) group by DATE" > $dialog.FileName
