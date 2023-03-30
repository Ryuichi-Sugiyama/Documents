###########################################################
#     Rangeで一括でやった方が速いかなぁ？
# $datarange = $sheet_temp.Range($sheet_temp.Cells(5, 1), $sheet_temp.Cells($temp_LastRow, 1))
# $sheet_temp.Range($sheet_temp.Cells(3, 1), $sheet_temp.Cells($temp_LastRow-2, 1)) = $datarange
#
# $datarange = $sheet_temp.Range($sheet_temp.Cells(5, $i1), $sheet_temp.Cells($temp_LastRow, $i1))
# $sheet_temp.Range($sheet_temp.Cells(3, 2), $sheet_temp.Cells($temp_LastRow-2, 2)) = $datarange
#
# $datarange = $sheet_temp.Range($sheet_temp.Cells(5, $i2), $sheet_temp.Cells($temp_LastRow, $i2))
# $sheet_temp.Range($sheet_temp.Cells(3, 4), $sheet_temp.Cells($temp_LastRow-2, 4)) = $datarange
#
###########################################################
# メッセージボックスで、取得した2つの内容を表示
# Add-Type -Assembly System.Windows.Forms
# [System.Windows.Forms.MessageBox]::Show("G7のテキストは $text1 です。`nG7の数式は $Formula1 です。", "結果") 
###########################################################

    # 入力ファイル：Excelファイル
$temp_xlsx = ".\temperature_data.xlsx"              # 気温データ
$temp_xlsx = (Get-ChildItem $temp_xlsx).FullName    #   フルパスに変換
$sheetName_temp = "temperature_data"                #   シート名
$wind_xlsx = "wind_data.xlsx"                       # 風速データ
$wind_xlsx = (Get-ChildItem $wind_xlsx).FullName    #   フルパスに変換
$sheetName_wind = "wind_data"                       #   シート名

    # 出力ファイル：ファイルが存在する場合は確認せずに削除
$out_excel = "shukei.xlsx"                          # 出力エクセルファイル
$out_excel = (Get-Location).Path + "\" + $out_excel #   フルパスに変換
if ( Test-Path $out_excel ) {                       #   ファイルが存在すれば削除
    Remove-Item $out_excel
}

    # エクセルの準備
$excel = New-Object -ComObject Excel.Application    # エクセル起動
$excel.Visible = $False                             # Excel画面非表示
$Excel.DisplayAlerts = $False                       # 上書き保存時に表示されるアラートなどを非表示
	# 入力ファイル
$book_temp = $excel.Workbooks.Open($temp_xlsx, 0, $true)  # 入力ファイルオープン
$sheet_temp = $book_temp.Worksheets.Item($sheetName_temp) #     ワークシート指定
$book_wind = $excel.Workbooks.Open($wind_xlsx, 0, $true)  # 入力ファイルオープン
$sheet_wind = $book_wind.Worksheets.Item($sheetName_wind) #     ワークシート指定
	# 出力ファイル
$book = $excel.Workbooks.Add()                      # 新規Excelファイルオープン
$sheet = $book.Worksheets.Item(1)                   # book内1枚目のsheet
$sheet.Name = "Connect"                             # シート名：Connect

    # エクセルファイル(temperature)読み込み
          # Range.SpecialCellsメソッドのXlCellTypeに使われたセル範囲内の
          # 最後のセルを意味する定数11（xlCellTypeLastCell）を指定して、
          # Cellsオブジェクトの最終行数を取得します。
          # https://learn.microsoft.com/ja-jp/office/vba/api/excel.range.specialcells
          # https://learn.microsoft.com/ja-jp/office/vba/api/excel.xlcelltype
$temp_LastRow = $sheet_temp.Cells.SpecialCells(11).Row

        # 大阪市の平均温度の列 : i1
For ( $i1=1; $i1 -le $temp_LastRow; $i1++ ) {
    $p1 = $sheet_temp.Cells.Item(1, $i1)
    $p2 = $sheet_temp.Cells.Item(2, $i1)
    $p3 = $sheet_temp.Cells.Item(4, $i1)
    if ( ($p1.Value() -eq "大阪") -and ($p2.Value() -eq "平均気温(℃)") -and ($p3.Value() -eq $Null) ) {
                                                            # セルが空の場合は $Null となる
        break
    }
}
        # 堺市の平均温度の列 : i2
For ( $i2=1; $i2 -le $temp_LastRow; $i2++ ) {
    $p1 = $sheet_temp.Cells.Item(1, $i2)
    $p2 = $sheet_temp.Cells.Item(2, $i2)
    $p3 = $sheet_temp.Cells.Item(4, $i2)
    if ( ($p1.Value() -eq "堺") -and ($p2.Value() -eq "平均気温(℃)") -and ($p3.Value() -eq $Null) ) {
        break
    }
}
    # エクセルファイル(wind)読み込み
$wind_LastRow = $sheet_wind.Cells.SpecialCells(11).Row
        # 大阪市の平均湿度の列 : i3
For ( $i3=1; $i3 -le $wind_LastRow; $i3++ ) {
    $p1 = $sheet_wind.Cells.Item(1, $i3)
    $p2 = $sheet_wind.Cells.Item(2, $i3)
    $p3 = $sheet_wind.Cells.Item(4, $i3)
    if ( ($p1.Value() -eq "大阪") -and ($p2.Value() -eq "平均湿度(％)") -and ($p3.Value() -eq $Null) ) {
        break
    }
}
        # 堺市の平均湿度の列 : i4
For ( $i4=1; $i4 -le $wind_LastRow; $i4++ ) {
    $p1 = $sheet_wind.Cells.Item(1, $i4)
    $p2 = $sheet_wind.Cells.Item(2, $i4)
    $p3 = $sheet_wind.Cells.Item(4, $i4)
    if ( ($p1.Value() -eq "堺") -and ($p2.Value() -eq "平均湿度(％)") -and ($p3.Value() -eq $Null) ) {
        break
    }
}

    # エクセルファイル書き込み
        # 見出し行の書き込み
$sheet.Cells.Item(2, 1) = "年月日"
$sheet.Cells.Item(1, 2) = "大阪"
$sheet.Cells.Item(2, 2) = "平均気温(℃)"
$sheet.Cells.Item(2, 3) = "平均湿度(％)"
$sheet.Cells.Item(1, 4) = "堺"
$sheet.Cells.Item(2, 4) = "平均気温(℃)"
$sheet.Cells.Item(2, 5) = "平均湿度(％)"
        # セルの結合、中央揃え
$sheet.range( "B1:C1" ).mergecells = 1
$sheet.range( "B1" ).HorizontalAlignment = -4108
$sheet.range( "D1:E1" ).mergecells = 1
$sheet.range( "D1" ).HorizontalAlignment = -4108
        # A列の書式設定→yyyy/mm/dd
$sheet.Columns("A").NumberFormatLocal = "yyyy/mm/dd"
        # 日付、温湿度データのコピー
For ( $j=5; $j -le $temp_LastRow; $j++ ) {
    Write-Output( $j )
    $dd = $sheet_temp.Cells.Item($j, 1)
    if ( $dd -eq $Null ) {
        break
    }
        # 年月日
    $sheet.Cells.Item($j-2, 1) = $dd
        # 大阪市、気温
    $sheet.Cells.Item($j-2, 2) = $sheet_temp.Cells.Item($j, $i1)
        # 堺市、気温
    $sheet.Cells.Item($j-2, 4) = $sheet_temp.Cells.Item($j, $i2)
        # 
    For ( $k=5; $k -le $wind_LastRow; $k++ ) {
        $p1 = $sheet_wind.Cells.Item($k, 1)
        if ( $dd.Value() -eq $p1.Value() ) {         # 日付が同じ行の湿度
                # 大阪市、湿度
            $sheet.Cells.Item($j-2, 3) = $sheet_wind.Cells.Item($k, $i3)
                # 堺市、湿度
            $sheet.Cells.Item($j-2, 5) = $sheet_wind.Cells.Item($k, $i4)
            break
        } elseif ( $p1.Value() -eq $Null ) {         # 同じ日付がない場合：空欄
            $sheet.Cells.Item($j-2, 3) = $Null
                # 堺市、湿度
            $sheet.Cells.Item($j-2, 5) = $Null
            break
        }
    }
}

        # 列幅の調整→A-E列を文字幅に合わせて自動調整
$Sheet.Columns("A:E").AutoFit() | Out-Null

    # エクセルファイルの保存とクローズ
$excel.DisplayAlerts = $FALSE
$book.SaveAs($out_excel)                             # ファイル保存
$book.Close($False)                                  # ファイルクローズ

$book_temp.Close($False)                             # temp ファイルクローズ
$book_wind.Close($False)                             # wind ファイルクローズ

$excel.Quit()
$excel = $Null
