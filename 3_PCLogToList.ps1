" "
" "
"*********************************************"
"                                             "
"   PC LOG TO LIST                            "
"      - Time Card Hack!!                     "
"   Version 1.0    UPDATE 2018/08/12          "
"                                             "
"*********************************************"
" "
" "
" "
# System.Windows.Formsアセンブリを有効化
[void][System.Reflection.Assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=Neutral, PublicKeyToken=b77a5c561934e089")

# リスト化する起動ログ（CSV）を指定させる。
$dialog_PC_START = New-Object System.Windows.Forms.OpenFileDialog 
$dialog_PC_START.Filter = "CSVファイル(*.CSV)|*.CSV;"
$dialog_PC_START.InitialDirectory = "C:\"
$dialog_PC_START.Title = "起動ログを選択してください"

if($dialog_PC_START.ShowDialog() -ne "OK"){
    "出力を中止しました。"
    exit 1
}

# リスト化する停止ログ（CSV）を指定させる。
$dialog_PC_END = New-Object System.Windows.Forms.OpenFileDialog 
$dialog_PC_END.Filter = "CSVファイル(*.CSV)|*.CSV;"
$dialog_PC_END.InitialDirectory = "C:\"
$dialog_PC_END.Title = "停止ログを選択してください"

if($dialog_PC_END.ShowDialog() -ne "OK"){
    "出力を中止しました。"
    exit 1
}


# 更新するリスト（Excel）を指定させる。
$dialog_Excel = New-Object System.Windows.Forms.OpenFileDialog 
$dialog_Excel.Filter = "Excelファイル(*.XLSX)|*.XLSX;"
$dialog_Excel.InitialDirectory = "C:\"
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



        # ****** 起動ログをExcelに貼り付ける ******
        # シート取り出し
        $sheet = $excel.Worksheets.Item("PC_STR")

        # シートのクリア
        $result = $sheet.Cells.Clear
        

        # 起動ログを読み込み、貼り付ける。
        $PC_START_Data = Import-Csv -Encoding UTF8 -Path $dialog_PC_START.FileName
        
        $Line = 0
        while($Line -lt $PC_START_Data.Length){

            $sheet.Cells.Item($Line + 1 ,1) = $PC_START_Data[$Line].DATE
            $sheet.Cells.Item($Line + 1 ,2) = $PC_START_Data[$Line].TIME
            
            $Line++
        }


        # ****** 停止ログをExcelに貼り付ける ******
        # シート取り出し
        $sheet = $excel.Worksheets.Item("PC_END")

        # シートのクリア
        $result = $sheet.Cells.Clear
        

        # 停止ログを読み込み、貼り付ける。
        $PC_END_Data = Import-Csv -Encoding UTF8 -Path $dialog_PC_END.FileName
        
        $Line = 0
        while($Line -lt $PC_END_Data.Length){

            $sheet.Cells.Item($Line + 1 ,1) = $PC_END_Data[$Line].DATE
            $sheet.Cells.Item($Line + 1 ,2) = $PC_END_Data[$Line].TIME
            
            $Line++
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
