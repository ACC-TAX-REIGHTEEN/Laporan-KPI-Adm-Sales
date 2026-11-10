Option Explicit

Sub HitungDataPerBulanIRC()
    Dim wb As Workbook
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long
    Dim dataArr As Variant
    Dim dict As Object
    Dim i As Long
    Dim val As Variant
    Dim dt As Date
    Dim isValidDate As Boolean
    Dim keyPeriode As String
    Dim invalidCount As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set wsSource = wb.Sheets("IRC")
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "Sheet 'IRC' tidak ditemukan!", vbCritical, "Error"
        GoTo CleanUp
    End If
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Data pada sheet IRC kosong atau hanya berisi header.", vbExclamation, "Informasi"
        GoTo CleanUp
    End If
    
    dataArr = wsSource.Range("A2:A" & lastRow).Value2
    
    Set dict = CreateObject("Scripting.Dictionary")
    invalidCount = 0
    
    For i = 1 To UBound(dataArr, 1)
        val = dataArr(i, 1)
        isValidDate = False
        
        If Not IsError(val) And Not IsEmpty(val) Then
            If IsNumeric(val) Then
                If CDbl(val) > 20000 And CDbl(val) < 80000 Then
                    dt = CDate(Int(CDbl(val)))
                    isValidDate = True
                End If
            Else
                Dim strVal As String
                strVal = Trim(CStr(val))
                
                strVal = Replace(strVal, "Mei", "May", , , vbTextCompare)
                strVal = Replace(strVal, "Agu", "Aug", , , vbTextCompare)
                strVal = Replace(strVal, "Okt", "Oct", , , vbTextCompare)
                strVal = Replace(strVal, "Nop", "Nov", , , vbTextCompare)
                strVal = Replace(strVal, "Des", "Dec", , , vbTextCompare)
                
                If IsDate(strVal) Then
                    dt = CDate(strVal)
                    isValidDate = True
                End If
            End If
        End If
        
        If isValidDate Then
            keyPeriode = Format(dt, "yyyy-mm")
            If dict.Exists(keyPeriode) Then
                dict(keyPeriode) = dict(keyPeriode) + 1
            Else
                dict.Add keyPeriode, 1
            End If
        Else
            invalidCount = invalidCount + 1
        End If
    Next i
    
    On Error Resume Next
    Set wsDest = wb.Sheets("Rekap_IRC")
    On Error GoTo 0
    
    If wsDest Is Nothing Then
        Set wsDest = wb.Sheets.Add(After:=wsSource)
        wsDest.Name = "Rekap_IRC"
    Else
        wsDest.Cells.Clear
    End If
    
    With wsDest
        .Range("A1:E1").Value = Array("No", "Tahun", "Bulan", "Periode", "Jumlah Baris")
        With .Range("A1:E1")
            .Font.Bold = True
            .Font.Color = vbWhite
            .Interior.Color = RGB(31, 78, 120)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 24
        End With
    End With
    
    Dim keysArr() As Variant
    keysArr = dict.keys
    
    Dim j As Long, temp As Variant
    For i = LBound(keysArr) To UBound(keysArr) - 1
        For j = i + 1 To UBound(keysArr)
            If keysArr(i) > keysArr(j) Then
                temp = keysArr(i)
                keysArr(i) = keysArr(j)
                keysArr(j) = temp
            End If
        Next j
    Next i
    
    Dim namaBulan As Variant
    namaBulan = Array("", "Januari", "Februari", "Maret", "April", "Mei", "Juni", _
                      "Juli", "Agustus", "September", "Oktober", "November", "Desember")
    
    Dim rowIdx As Long: rowIdx = 2
    Dim thn As String, blnAngka As Integer
    
    For i = LBound(keysArr) To UBound(keysArr)
        keyPeriode = CStr(keysArr(i))
        thn = Left(keyPeriode, 4)
        blnAngka = CInt(Right(keyPeriode, 2))
        
        wsDest.Cells(rowIdx, 1).Value = rowIdx - 1
        wsDest.Cells(rowIdx, 2).Value = thn
        wsDest.Cells(rowIdx, 3).Value = namaBulan(blnAngka)
        wsDest.Cells(rowIdx, 4).Value = keyPeriode
        wsDest.Cells(rowIdx, 5).Value = dict(keyPeriode)
        rowIdx = rowIdx + 1
    Next i
    
    If invalidCount > 0 Then
        wsDest.Cells(rowIdx, 1).Value = rowIdx - 1
        wsDest.Cells(rowIdx, 2).Value = "-"
        wsDest.Cells(rowIdx, 3).Value = "Tanpa Tanggal / Format Tidak Dikenali"
        wsDest.Cells(rowIdx, 4).Value = "-"
        wsDest.Cells(rowIdx, 5).Value = invalidCount
        rowIdx = rowIdx + 1
    End If
    
    wsDest.Cells(rowIdx, 1).Value = "Total Keseluruhan"
    wsDest.Range(wsDest.Cells(rowIdx, 1), wsDest.Cells(rowIdx, 4)).Merge
    wsDest.Cells(rowIdx, 5).Formula = "=SUM(E2:E" & (rowIdx - 1) & ")"
    
    With wsDest.Range(wsDest.Cells(rowIdx, 1), wsDest.Cells(rowIdx, 5))
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With
    
    Dim tblRange As Range
    Set tblRange = wsDest.Range("A1:E" & rowIdx)
    
    tblRange.Borders.LineStyle = xlContinuous
    tblRange.Borders.Weight = xlThin
    
    wsDest.Range("A2:A" & rowIdx).HorizontalAlignment = xlCenter
    wsDest.Range("B2:B" & rowIdx).HorizontalAlignment = xlCenter
    wsDest.Range("D2:D" & rowIdx).HorizontalAlignment = xlCenter
    wsDest.Range("E2:E" & rowIdx).NumberFormat = "#,##0"
    
    wsDest.Columns("A:E").AutoFit
    wsDest.Activate

    MsgBox "Proses rekapitulasi data sheet IRC selesai!", vbInformation, "Sukses"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
