[9:40] 杉山 隆一
    # 入力ファイル
$cmdb_csv = ".\資材管理.csv"                        # 構成管理
$vcenter_csv = "vCenter.csv"                        # vCenter
$excelFile = ".\配線表.xlsx"                        # 配線表
$sheetName = "配線表"                               # 配線表のシート名
$excelFile = (Get-ChildItem $excelFile).FullName    # フルパスに変換

    # 出力ファイル
$out_csv = ".\result-3.csv"                            # 中間csvファイル
Write-Output "CMDB.資材管理番号,配線表.iLO,CMDB.資材名,配線表.ホスト名,vCenter.名前,vCenter.クラスタ,CMDB.クラスタ名,配線表.タイプ,CMDB.資材ステータス,vCenter.状態,CMDB.状況,CMDB.備考" | Out-File $out_csv$out_excel = "result-3.xlsx"                        # 出力エクセルファイル
$fp = Get-Location
$out_excel = $fp.Path + "\" + $out_excel            # フルパスに変換

    # エクセルの準備
$excel = New-Object -ComObject Excel.Application    # エクセル起動
$excel.Visible = $false                             # Excel画面非表示
$Excel.DisplayAlerts = $False                       # 上書き保存時に表示されるアラートなどを非表示
$book = $excel.Workbooks.Open($excelFile, 0, $true) # ファイルオープン
$sheet = $book.Worksheets.Item($sheetName)          # ワークシート指定
    # csvファイル読込
$csv1 = Import-Csv -Encoding oem $cmdb_csv | Sort-Object -Property 資材管理番号        # 構成管理のデータ読込
$csv2 = Import-Csv -Encoding oem $vcenter_csv       # vCenterのデータ読込

    # csvファイルへの書き出し
ForEach( $a in $csv1 ) {                            # CMDBの資材管理番号をキーとして検索
  # CMDB.資材管理番号 : $a.資材管理番号
  # 配線表.iLO : $h_iLO
  # CMDB.資材名 : $a.資材名
  # 配線表.ホスト名 : $h_HostName
  # vCenter.名前 : $vc_HostName
  # vCenter.クラスタ : $vc_Cluster
  # CMDB.クラスタ名 : $a.クラスタ名
  # 配線表.タイプ : $h_Type
  # CMDB.資材ステータス : $a.資材ステータス
  # vCenter.状態 : $vc_Status
  # CMDB.状況
  # CMDB.備考

    if ( $a.資材管理番号.Substring(0,4) -ne "FS2J" ) { continue }

  # 配線表データ取得
    if ( $sheet.Cells.Find($a.資材管理番号) -eq $nul ) {
        $h_iLO = ""
        $h_HostName = ""
        $h_Type = ""
    } else {
        $h_iLO = $sheet.Cells.Find($a.資材管理番号).Value2
        $cc = $sheet.Cells.Find($a.資材管理番号).Column    # 列
        $rr = $sheet.Cells.Find($a.資材管理番号).Row    # 行
        $h_HostName = $sheet.Cells.Item($rr,$cc-2).Value2
        $h_Type = $sheet.Cells.Item($rr,$cc+2).Value2
    }

  # vCenter
    $vc_HostName = ""
    $vc_Cluster = ""
    $vc_Status = ""
    ForEach( $b in $csv2 ) {
        if ( $b.名前.Split(".")[0] -eq $a.資材名 ) {
            $vc_HostName = $b.名前.Split(".")[0]
            $vc_Cluster = $b.クラスタ
            $vc_Status = $b.状態
            break
        }
    }

  # CMDB.状況、備考：改行とコンマの削除
    $jokyo = $a.状況.Replace("`n","").Replace("`r","").Replace(".","。")
    $bikou = $a.備考.Replace("`n","").Replace("`r","").Replace(".","。")

  # csvファイル出力
    Write-Output ( $a.資材管理番号 +","+ $h_iLO +","+ $a.資材名 +","+ $h_HostName +","+ $vc_HostName +","+ $vc_Cluster +","+ $a.クラスタ名 +","+ $h_Type +","+ $a.資材ステータス +","+ $vc_Status +","+ $jokyo +","+  $bikou ) | Out-File $out_csv -Append
}

$book.Close($False)                                    # 配線表ファイルクローズ

#######################################################################
# 出力したcsvファイルをテキストファイルとして読み込み, Excelに出力
# 後続の列に各情報の整合性判定関数を追加
#######################################################################
    # 前段で作成したcsvファイルをテキストとして読込、列数の取得
$line_csv = Get-Content -Encoding oem $out_csv        # ファイルをテキストファイルとして読込
$clm = $line_csv[0].Split(",").Count                # csvファイルの列数
    # エクセルの準備
$excel = New-Object -ComObject Excel.Application    # エクセル起動
$book = $excel.Workbooks.Add()                        # 新規Excelファイルオープン
$sheet = $book.Worksheets.Item(1)                    # book内1枚目のsheet
$sheet.Name = "Hosts"                                # シート名：Hosts
    # エクセルシートへの書込
For ( $i=0; $i -lt $line_csv.Count; $i++ ) {
    $cell_data = $line_csv[$i] -split ","            # 1行読込みコンマ区切りで分割
    For ($j=0; $j -lt $cell_data.Count; $j++) {
        $Sheet.Cells.Item($i+1, $j+1) = $cell_data[$j]    # [$i+1]行の各セルに順次書込み
    }
}
    # 後続の列への判定用関数追加
$Sheet.Cells.Item(1, $clm+1) = "資産管理`n番号`n整合性"    # 1行目：見出し行
$Sheet.Cells.Item(1, $clm+2) = "ホスト名`n整合性"
$Sheet.Cells.Item(1, $clm+3) = "クラスタ`n整合性"
$Sheet.Cells.Item(1, $clm+4) = "タイプ`n整合性"
For ( $i=2; $i -le $line_csv.Count; $i++ ) {
    $f = "=A" +$i+ "=B" +$i
    $Sheet.Cells.Item($i, $clm+1) = $f
    $f = "=IF(AND(C" +$i+ "=D" +$i+ ",D" +$i+ "=E" +$i+ "),TRUE,FALSE)"
    $Sheet.Cells.Item($i, $clm+2) = $f
    $f = "=F" +$i+ "=G" +$i
    $Sheet.Cells.Item($i, $clm+3) = $f
    $f = "=IF(AND(RIGHT(F" +$i+ ",2)=RIGHT(G" +$i+ ",2),RIGHT(G" +$i+ ",2)=LEFT(H" +$i+ ",2)),TRUE,FALSE)"
    $Sheet.Cells.Item($i, $clm+4) = $f
}
    # 見易いようにフィルタとか行の高さとか列幅とか
$col = $Sheet.UsedRange.Columns.Count                # フィルターをセット
$Sheet.Range($Sheet.Cells(1, 1), $Sheet.Cells(1, $col)).AutoFilter() | Out-Null
$Sheet.Range("A1:B1").RowHeight = 56.25              # 1行目の高さ
$Sheet.Range("A1:B1").ColumnWidth = 11.88            # A,B列の幅
$Sheet.Range("C1:E1").ColumnWidth = 23.75            # C,D,E列の幅
$Sheet.Range("F1:G1").ColumnWidth = 19.25            # F,G列の幅
$Sheet.Range("L1:L1").ColumnWidth = 17.5             # L列の幅
$Sheet.Range($Sheet.Cells(1, $clm+1), $Sheet.Cells(1, $clm+4)).interior.color = 0x00ffff
    # エクセルファイルの保存とクローズ
$excel.DisplayAlerts = $FALSE
$book.SaveAs($out_excel)                             # ファイル保存
$book.Close($False)                                  # ファイルクローズ

$excel.Quit()
$excel = $Null

