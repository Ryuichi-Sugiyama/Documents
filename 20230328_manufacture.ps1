##################################################################### 2023/03/28 ###########
# Excelファイルはコピーして使用しました。ディレクトリは全ファイル同じとしています。
#   【パック】申請受付入力シート(4744件) .xlsx → Pack.xlsx
#   トレーサビリティ管理表（端末）.xlsx → Traceability.xlsx
############################################################################################

$in_xlsx_1 = ".\Pack.xlsx"                          # ファイル名、シート名指定
$in_xlsx_1 =  (Get-ChildItem $in_xlsx_1).FullName
$sheet_in_1_name = "Standard・バリュー"
$in_xlsx_2 = ".\Traceability.xlsx"
$in_xlsx_2 =  (Get-ChildItem $in_xlsx_2).FullName
$sheet_in_2_name = "トレサビリティ管理表"

$excel = New-Object -ComObject Excel.Application    # エクセル起動
$Excel.DisplayAlerts = $False                       # 上書き保存時に表示されるアラートなどを非表示
$excel.Visible = $True

$book_in_1 = $excel.Workbooks.Open($in_xlsx_1, 0, $true)       # 入力ファイル(パック)オープン
$sheet_in_1 = $book_in_1.Worksheets.Item($sheet_in_1_name)     #     ワークシート指定
$book_in_2 = $excel.Workbooks.Open($in_xlsx_2, 0, $true)       # 入力ファイル(トレーサビリティ)オープン
$sheet_in_2 = $book_in_2.Worksheets.Item($sheet_in_2_name)     #     ワークシート指定

$Num_1      = $sheet_in_1.Range("A11:A4754") | % {$_.Text}     # 項番(パック)
$SerialNo_1 = $sheet_in_1.Range("D11:D4754") | % {$_.Text}     # 製造番号(パック)
$SerialNo_2 = $sheet_in_2.Range("H4:H6173") | % {$_.Text}      # シリアル番号(トレサ)
$place_2 = $sheet_in_2.Range("B4:B6173") | % {$_.Text}         # 学校名(トレサ)
$t_type_2 = $sheet_in_2.Range("F4:F6173") | % {$_.Text}        # 端末種別(トレサ) "【2in1】"が対象

    # 出力ファイル(csv)：パックファイルの製造番号からソート
$out_csv = "result_manu.aaa"                            # ファイル名
$fp = Get-Location
$out_csv = $fp.Path + "\" + $out_csv                    # フルパスに変換
Write-Output "項番,製造番号(パック),シリアル番号(トレサ),学校名" | Out-File $out_csv    # ヘッダ出力

    # 両ファイルのシリアル番号が一致したらcsv出力
for ( $i=0; $i -lt $SerialNo_1.Count; $i++ ) {
    $k = 0
    $x = $SerialNo_1[$i]
    for ( $j=0; $j -lt $SerialNo_2.Count; $j++ ) {
        $y = $SerialNo_2[$j]
        if ( $x -eq $y ) {
            $z = $Num_1[$i] +","+ $x +","+ $y +"," + $place_2[$j]
            Write-Output $z | Out-File $out_csv -Append
            $k = 1
            break
        }
    }
    if ( $k -eq 0 ) {
        $z = $Num_1[$i] +","+ $x +",,"
        Write-Output $z | Out-File $out_csv -Append
    }
}

$book_in_1.Close($False)                     # Excel ファイルクローズ
$book_in_2.Close($False)

$excel.Quit()                                # Excel 終了
$excel = $Null
