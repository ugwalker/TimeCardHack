" "
" "
"*********************************************"
"                                             "
"   Entry And Exit LOG TO LIST                "
"      - Time Card Hack!!                     "
"   Version 1.1    UPDATE 2018/08/20          "
"                                             "
"*********************************************"
" "
" "
" "
# System.Windows.Formsアセンブリを有効化
[void][System.Reflection.Assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=Neutral, PublicKeyToken=b77a5c561934e089")

# スクリプトのパスを取得する。
$scriptFile = $MyInvocation.MyCommand.Path
$sciriptPath = Split-Path -Parent $scriptFile

# リスト化する入退室ログ（JSON）を指定させる。
$dialog_EntryExit = New-Object System.Windows.Forms.OpenFileDialog 
$dialog_EntryExit.Filter = "JSONファイル(*.JSON)|*.JSON;"
$dialog_EntryExit.InitialDirectory = $sciriptPath
$dialog_EntryExit.Title = "入退室ログを選択してください"

if($dialog_EntryExit.ShowDialog() -ne "OK"){
    "出力を中止しました。"
    exit 1
}

# 更新するリスト（Excel）を指定させる。
$dialog_Excel = New-Object System.Windows.Forms.OpenFileDialog 
$dialog_Excel.Filter = "Excelファイル(*.XLSX)|*.XLSX;"
$dialog_Excel.InitialDirectory = $sciriptPath
$dialog_Excel.Title = "更新するExcelを選択してください"

if($dialog_Excel.ShowDialog() -ne "OK"){
    "出力を中止しました。"
    exit 1
}


# ExcelにCSVを貼り付ける。
try{
        # Excelオブジェクト作成
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false

        # 既存のExcel
        $book = $excel.Workbooks.Open($dialog_Excel.FileName)

        # シート取り出し
        $sheet_entry = $excel.Worksheets.Item("Entry")
        $sheet_exit = $excel.Worksheets.Item("Exit")

        # シートのクリア
        $result = $sheet_entry.Cells.Clear
        $result = $sheet_exit.Cells.Clear
        

        # 入退出ログを読み込み、貼り付ける。
        $EntryExit_Data = Get-Content $dialog_EntryExit.FileName -Encoding UTF8 | ConvertFrom-Json
        
        $Line_entry = 0
        $Line_exit = 0
        $Line_json = 0
        while($Line_json -lt $EntryExit_Data.length){

            # 入室
            if( $EntryExit_Data[$Line_json].type -eq '0'){
                $sheet_entry.Cells.Item($Line_entry + 1 ,1) = $EntryExit_Data[$Line_json].recodedate.Substring(0, 10)
                $sheet_entry.Cells.Item($Line_entry + 1 ,2) = $EntryExit_Data[$Line_json].recodedate.Substring(11, 8)

                $Line_entry++
            }
            # 退室
            else{
                $sheet_exit.Cells.Item($Line_exit + 1 ,1) = $EntryExit_Data[$Line_json].recodedate.Substring(0, 10)
                $sheet_exit.Cells.Item($Line_exit + 1 ,2) = $EntryExit_Data[$Line_json].recodedate.Substring(11, 8)
   
                $Line_exit++
            }
   
            $Line_json++
        }

        # 上書き保存
        $book.Save()

        # 閉じる
        $excel.Quit()

} finally {
　　　　 # null破棄
        $excel,$book,$sheet | foreach{$_ = $null}
}

# 完了メッセージを表示する。
" 終わりました！今月もおつかれやまです。"
