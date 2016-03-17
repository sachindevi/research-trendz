'''Description: This macro will list the distinct word from the text and count the number of words also yearwise.

'--------------------------------------------------------------------------------------------------------------------------------
'Following function is the main function to count the word year wise
'--------------------------------------------------------------------------------------------------------------------------------
Sub Word_Count_YearWise()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.StatusBar = True
    
    Dim shtSource As Worksheet
    Dim shtResult As Worksheet
    Dim shtRemove As Worksheet

    Dim tRowinSrc As Long, tRowResult As Long, tRowRemove As Long
    Dim tColResult As Long, goTh As Long, i As Long
    Dim strYear As Integer
    Dim strTitle As String, splitTitle() As String, colL As String
    
    Dim rngYear As Range, rngWord As Range, rngRemove As Range
    Dim resYear As Variant, resWord As Variant, resRemove As Variant
    
    Set shtSource = Worksheets("InputFile")
    Set shtResult = Worksheets("Word Count")
    Set shtRemove = Worksheets("Word to Remove")
    
    shtResult.Select
    
    Application.StatusBar = "Please Wait: Script is running....!!"
    
    tRowinSrc = shtSource.Range("E" & Rows.Count).End(xlUp).Row
    tRowResult = shtResult.Range("A" & Rows.Count).End(xlUp).Row
    tRowRemove = shtRemove.Range("A" & Rows.Count).End(xlUp).Row
    Set rngRemove = shtRemove.Range("A1:A" & tRowRemove)
    
    shtResult.Range("A2:AAB" & tRowResult).ClearContents                           'Clear prevous results of "word count" sheets
    shtResult.Range("C1:AAB1").ClearContents
    
    fileSource = ActiveWorkbook.FullName     '''''get the active workbook name for query source
    
    ConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fileSource & ";Extended Properties=Excel 8.0;Persist Security Info=False"
    Set oCn = New ADODB.Connection
    oCn.ConnectionString = ConnString

    oCn.Open
'----------------------------------------------------------------------------------------------------------------------------------------
'///////////////For News Search Dropdown logic start from here////////////////-----------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------------------------------
    '''For OC Status Dropdown
    SqlDrop = "Select DISTINCT[PubDate(Year)] from [" & shtSource.Name & "$A1:F" & tRowinSrc & "] ORDER BY [PubDate(Year)] ASC"
    
    Set oRS = New ADODB.Recordset
    oRS.Source = SqlDrop
    oRS.ActiveConnection = oCn

    oRS.Open
    
    rowD = 3
    While Not oRS.EOF
        If Not IsNull(oRS.Fields.Item(0)) Then
            shtResult.Cells(1, rowD) = oRS.Fields.Item(0)
            rowD = rowD + 1
        End If
        oRS.MoveNext
    Wend
Err:
    If oRS.State <> adStateClosed Then
        oRS.Close
    End If
    tColResult = shtResult.Cells(1, Columns.Count).End(xlToLeft).Column
    colL = Left(Cells(1, tColResult).Address(False, False), 1 - (tColResult > 26))
    Set rngYear = shtResult.Range("A1:" & colL & 1)
    Dim curWord As String
    For goTh = 2 To tRowinSrc
        Application.StatusBar = "Please Wait: Script is running....!! Row#[" & goTh & "] out of Rows#[" & tRowinSrc & "] is in progress..!!"
        
        strYear = shtSource.Range("C" & goTh)
        
        strTitle = Replace(shtSource.Range("E" & goTh), Chr(10), " ")
        splitTitle = Split(strTitle, " ")
        
        For i = 0 To UBound(splitTitle) - 1
            resRemove = Application.Match(splitTitle(i), rngRemove, 0)
            If IsError(resRemove) Then
                tRowResult = shtResult.Range("A" & Rows.Count).End(xlUp).Row
                Set rngWord = shtResult.Range("A1:A" & tRowResult)
                curWord = removeSpecial(Trim(splitTitle(i)))
                resWord = Application.Match(curWord, rngWord, 0)
                If IsError(resWord) Then
                    'shtResult.Range("A" & tRowResult + 1).NumberFormat = "@"
                    shtResult.Range("A" & tRowResult + 1) = curWord
                    shtResult.Range("B" & tRowResult + 1) = 1
                    resYear = Application.Match(strYear, rngYear, 0)
                    If IsError(resYear) Then
                    Else
                        shtResult.Cells(tRowResult + 1, resYear) = shtResult.Cells(tRowResult + 1, resYear) + 1
                    End If
                Else
                    shtResult.Range("B" & resWord) = shtResult.Range("B" & resWord) + 1
                    '''For Year wise count
                    resYear = Application.Match(strYear, rngYear, 0)
                    If IsError(resYear) Then
                    Else
                        shtResult.Cells(resWord, resYear) = shtResult.Cells(resWord, resYear) + 1
                    End If
                End If
            End If
        Next i
    Next goTh
    Dim colLetter As String
    tRowResult = shtResult.Range("A" & Rows.Count).End(xlUp).Row
    tColResult = shtResult.Cells(1, Columns.Count).End(xlToLeft).Column
    colLetter = Left(Cells(1, tColResult).Address(False, False), 1 - (tColResult > 26))
    shtResult.Range("C2:" & colLetter & tRowResult).Replace What:="", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    '''Sort the data
    shtResult.Sort.SortFields.Clear
    shtResult.Sort.SortFields.Add Key:=Range( _
        "B2:B" & tRowResult), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With shtResult.Sort
        .SetRange Range("A1:" & colLetter & tRowResult)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    '''End of Sort Logic
    shtResult.Range("A1") = "[Words]"
    shtResult.Range("B1") = "[Count of Word]"
    MsgBox "Words are counted...!!", vbInformation, "Success"
    ''''''End of finding Unique value for the year
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
End Sub

'--------------------------------------------------------------------------------------------------------------------------------
'Following function will Identify the occurance of the word
'--------------------------------------------------------------------------------------------------------------------------------
Sub Identify_Occurance()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.StatusBar = True
    Dim shtResult As Worksheet
    
    Dim tRowResult As Long, goT As Long
    Dim tCol As Integer, i As Integer
    
    Set shtResult = Worksheets("Word Count")
    
    tRowResult = shtResult.Range("A" & Rows.Count).End(xlUp).Row
    tCol = shtResult.Cells(1, Columns.Count).End(xlToLeft).Column
    For goT = 2 To tRowResult
        Application.StatusBar = "Please Wait: Script is running....!! Row#[" & goT & "] out of Rows#[" & tRowResult & "] is in progress..!!"
        For i = 3 To tCol
            If shtResult.Cells(goT, i) > 0 Then
                shtResult.Cells(goT, i).Interior.Color = vbGreen
                Exit For
            End If
        Next i
    Next goT
    
    Application.StatusBar = "Please Wait: Script is running....!!"
    
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'--------------------------------------------------------------------------------------------------------------------------------
'Following function will is to used select the input file to process (read the word)
'--------------------------------------------------------------------------------------------------------------------------------
Sub Select_Input_File()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    MasterShtName = ActiveWorkbook.Name
    
    openShtName = Application.GetOpenFilename("All Excel files (*.xls*), *.xls*", , _
                                "Please choose file")
    On Error GoTo 10:
    If Not openShtName Then
      Exit Sub
    End If
10:
    If Sheets.Count > 1 Then
        While (Sheets.Count <> 5)
            Sheets(6).Delete
        Wend
    End If
    
    If openShtName <> "FALSE" Then
        Workbooks.Open Filename:=openShtName
        
        openShtName = Split(openShtName, "\")
        openShtName = openShtName(UBound(openShtName))
        ordoroShtName = openShtName                          'stores the ordoroShtName
        Windows(openShtName).Activate
        totSheetsinOpenSht = Sheets.Count
        For i = 1 To 1 'totSheetsinOpenSht
            Sheets(i).Select
            Sheets(i).Copy After:=Workbooks(MasterShtName).Sheets(5)
            Windows(openShtName).Activate
        Next i
        ActiveWorkbook.Close
        Sheets(6).Name = "InputFile"
        Windows(MasterShtName).Activate
        Worksheets("Control Panel").Select
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'--------------------------------------------------------------------------------------------------------------------------------
'Following function will save the "Word Count" tab of the Excel file
'--------------------------------------------------------------------------------------------------------------------------------
Sub Excel_Save_Word_Count_Tab()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim csvFilName As String
    '''''
    csvFilName = Application.Application.GetSaveAsFilename(fileFilter:="Excel Files, *.xlsx")
    
    Sheets("Word Count").Select
    Sheets("Word Count").Copy
    
    If Len(csvFilName) > 5 Then
        ActiveWorkbook.SaveAs Filename:=csvFilName, FileFormat _
            :=xlOpenXMLWorkbook, CreateBackup:=False
    End If
    ActiveWorkbook.Save
    ActiveWindow.Close
    Sheets("Control Panel").Select
    
    MsgBox "Template file saved successfully...!!", vbInformation
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'--------------------------------------------------------------------------------------------------------------------------------
'Following function will remove the special characters from the given string
'--------------------------------------------------------------------------------------------------------------------------------
Function removeSpecial(sInput As String) As String
    Dim sSpecialChars As String
    Dim i As Long
    sSpecialChars = "’.,-[];\/:*?""<>|+'()±”"
    For i = 1 To Len(sSpecialChars)
        'sInput = Replace$(Right(sInput, 1), Mid$(sSpecialChars, i, 1), "")
        If Right(sInput, 1) = Mid$(sSpecialChars, i, 1) Then
            sInput = Left(sInput, Len(sInput) - 1)
        End If
    Next
    removeSpecial = sInput
End Function


