###########################################################
#     Range�ňꊇ�ł���������������Ȃ��H
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
# ���b�Z�[�W�{�b�N�X�ŁA�擾����2�̓��e��\��
# Add-Type -Assembly System.Windows.Forms
# [System.Windows.Forms.MessageBox]::Show("G7�̃e�L�X�g�� $text1 �ł��B`nG7�̐����� $Formula1 �ł��B", "����") 
###########################################################

    # ���̓t�@�C���FExcel�t�@�C��
$temp_xlsx = ".\temperature_data.xlsx"              # �C���f�[�^
$temp_xlsx = (Get-ChildItem $temp_xlsx).FullName    #   �t���p�X�ɕϊ�
$sheetName_temp = "temperature_data"                #   �V�[�g��
$wind_xlsx = "wind_data.xlsx"                       # �����f�[�^
$wind_xlsx = (Get-ChildItem $wind_xlsx).FullName    #   �t���p�X�ɕϊ�
$sheetName_wind = "wind_data"                       #   �V�[�g��

    # �o�̓t�@�C���F�t�@�C�������݂���ꍇ�͊m�F�����ɍ폜
$out_excel = "shukei.xlsx"                          # �o�̓G�N�Z���t�@�C��
$out_excel = (Get-Location).Path + "\" + $out_excel #   �t���p�X�ɕϊ�
if ( Test-Path $out_excel ) {                       #   �t�@�C�������݂���΍폜
    Remove-Item $out_excel
}

    # �G�N�Z���̏���
$excel = New-Object -ComObject Excel.Application    # �G�N�Z���N��
$excel.Visible = $False                             # Excel��ʔ�\��
$Excel.DisplayAlerts = $False                       # �㏑���ۑ����ɕ\�������A���[�g�Ȃǂ��\��
	# ���̓t�@�C��
$book_temp = $excel.Workbooks.Open($temp_xlsx, 0, $true)  # ���̓t�@�C���I�[�v��
$sheet_temp = $book_temp.Worksheets.Item($sheetName_temp) #     ���[�N�V�[�g�w��
$book_wind = $excel.Workbooks.Open($wind_xlsx, 0, $true)  # ���̓t�@�C���I�[�v��
$sheet_wind = $book_wind.Worksheets.Item($sheetName_wind) #     ���[�N�V�[�g�w��
	# �o�̓t�@�C��
$book = $excel.Workbooks.Add()                      # �V�KExcel�t�@�C���I�[�v��
$sheet = $book.Worksheets.Item(1)                   # book��1���ڂ�sheet
$sheet.Name = "Connect"                             # �V�[�g���FConnect

    # �G�N�Z���t�@�C��(temperature)�ǂݍ���
          # Range.SpecialCells���\�b�h��XlCellType�Ɏg��ꂽ�Z���͈͓���
          # �Ō�̃Z�����Ӗ�����萔11�ixlCellTypeLastCell�j���w�肵�āA
          # Cells�I�u�W�F�N�g�̍ŏI�s�����擾���܂��B
          # https://learn.microsoft.com/ja-jp/office/vba/api/excel.range.specialcells
          # https://learn.microsoft.com/ja-jp/office/vba/api/excel.xlcelltype
$temp_LastRow = $sheet_temp.Cells.SpecialCells(11).Row

        # ���s�̕��ω��x�̗� : i1
For ( $i1=1; $i1 -le $temp_LastRow; $i1++ ) {
    $p1 = $sheet_temp.Cells.Item(1, $i1)
    $p2 = $sheet_temp.Cells.Item(2, $i1)
    $p3 = $sheet_temp.Cells.Item(4, $i1)
    if ( ($p1.Value() -eq "���") -and ($p2.Value() -eq "���ϋC��(��)") -and ($p3.Value() -eq $Null) ) {
                                                            # �Z������̏ꍇ�� $Null �ƂȂ�
        break
    }
}
        # ��s�̕��ω��x�̗� : i2
For ( $i2=1; $i2 -le $temp_LastRow; $i2++ ) {
    $p1 = $sheet_temp.Cells.Item(1, $i2)
    $p2 = $sheet_temp.Cells.Item(2, $i2)
    $p3 = $sheet_temp.Cells.Item(4, $i2)
    if ( ($p1.Value() -eq "��") -and ($p2.Value() -eq "���ϋC��(��)") -and ($p3.Value() -eq $Null) ) {
        break
    }
}
    # �G�N�Z���t�@�C��(wind)�ǂݍ���
$wind_LastRow = $sheet_wind.Cells.SpecialCells(11).Row
        # ���s�̕��ώ��x�̗� : i3
For ( $i3=1; $i3 -le $wind_LastRow; $i3++ ) {
    $p1 = $sheet_wind.Cells.Item(1, $i3)
    $p2 = $sheet_wind.Cells.Item(2, $i3)
    $p3 = $sheet_wind.Cells.Item(4, $i3)
    if ( ($p1.Value() -eq "���") -and ($p2.Value() -eq "���ώ��x(��)") -and ($p3.Value() -eq $Null) ) {
        break
    }
}
        # ��s�̕��ώ��x�̗� : i4
For ( $i4=1; $i4 -le $wind_LastRow; $i4++ ) {
    $p1 = $sheet_wind.Cells.Item(1, $i4)
    $p2 = $sheet_wind.Cells.Item(2, $i4)
    $p3 = $sheet_wind.Cells.Item(4, $i4)
    if ( ($p1.Value() -eq "��") -and ($p2.Value() -eq "���ώ��x(��)") -and ($p3.Value() -eq $Null) ) {
        break
    }
}

    # �G�N�Z���t�@�C����������
        # ���o���s�̏�������
$sheet.Cells.Item(2, 1) = "�N����"
$sheet.Cells.Item(1, 2) = "���"
$sheet.Cells.Item(2, 2) = "���ϋC��(��)"
$sheet.Cells.Item(2, 3) = "���ώ��x(��)"
$sheet.Cells.Item(1, 4) = "��"
$sheet.Cells.Item(2, 4) = "���ϋC��(��)"
$sheet.Cells.Item(2, 5) = "���ώ��x(��)"
        # �Z���̌����A��������
$sheet.range( "B1:C1" ).mergecells = 1
$sheet.range( "B1" ).HorizontalAlignment = -4108
$sheet.range( "D1:E1" ).mergecells = 1
$sheet.range( "D1" ).HorizontalAlignment = -4108
        # A��̏����ݒ聨yyyy/mm/dd
$sheet.Columns("A").NumberFormatLocal = "yyyy/mm/dd"
        # ���t�A�����x�f�[�^�̃R�s�[
For ( $j=5; $j -le $temp_LastRow; $j++ ) {
    Write-Output( $j )
    $dd = $sheet_temp.Cells.Item($j, 1)
    if ( $dd -eq $Null ) {
        break
    }
        # �N����
    $sheet.Cells.Item($j-2, 1) = $dd
        # ���s�A�C��
    $sheet.Cells.Item($j-2, 2) = $sheet_temp.Cells.Item($j, $i1)
        # ��s�A�C��
    $sheet.Cells.Item($j-2, 4) = $sheet_temp.Cells.Item($j, $i2)
        # 
    For ( $k=5; $k -le $wind_LastRow; $k++ ) {
        $p1 = $sheet_wind.Cells.Item($k, 1)
        if ( $dd.Value() -eq $p1.Value() ) {         # ���t�������s�̎��x
                # ���s�A���x
            $sheet.Cells.Item($j-2, 3) = $sheet_wind.Cells.Item($k, $i3)
                # ��s�A���x
            $sheet.Cells.Item($j-2, 5) = $sheet_wind.Cells.Item($k, $i4)
            break
        } elseif ( $p1.Value() -eq $Null ) {         # �������t���Ȃ��ꍇ�F��
            $sheet.Cells.Item($j-2, 3) = $Null
                # ��s�A���x
            $sheet.Cells.Item($j-2, 5) = $Null
            break
        }
    }
}

        # �񕝂̒�����A-E��𕶎����ɍ��킹�Ď�������
$Sheet.Columns("A:E").AutoFit() | Out-Null

    # �G�N�Z���t�@�C���̕ۑ��ƃN���[�Y
$excel.DisplayAlerts = $FALSE
$book.SaveAs($out_excel)                             # �t�@�C���ۑ�
$book.Close($False)                                  # �t�@�C���N���[�Y

$book_temp.Close($False)                             # temp �t�@�C���N���[�Y
$book_wind.Close($False)                             # wind �t�@�C���N���[�Y

$excel.Quit()
$excel = $Null
