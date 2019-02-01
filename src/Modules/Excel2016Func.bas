Attribute VB_Name = "Excel2016Func"
Option Explicit

' [関数リファレンス | 経理・会計事務所向けエクセルスピードアップ講座](https://www.excelspeedup.com/category/kansuu/)

' [ifs関数の使い方とExcel2013以前の古いエクセルで使う方法](https://www.excelspeedup.com/ifs/)
Function IFS(ParamArray par())
  Dim i As Integer

  IFS = CVErr(xlErrNA)
  If UBound(par) Mod 2 = 0 Then
    Exit Function
  End If
  
  For i = LBound(par) To UBound(par) - 1 Step 2
    If par(i) Then
      IFS = par(i + 1)
      Exit Function
    End If
  Next
  
End Function

' [switch関数の使い方とExcel2013以前の古いエクセルで使う方法](https://www.excelspeedup.com/switch/)
Function SWITCH(ParamArray par())
  Dim i As Integer

  For i = LBound(par) + 1 To UBound(par) - 1 Step 2
    If par(LBound(par)) = par(i) Then
      SWITCH = par(i + 1)
      Exit Function
    End If
  Next

  If i = UBound(par) Then
    SWITCH = par(i)
  Else
    SWITCH = CVErr(xlErrNA)
  End If
End Function

' [concat関数の使い方とExcel2013以前の古いエクセルで使う方法](https://www.excelspeedup.com/concat2/)
Function CONCAT(ParamArray parA())
    Dim min_i As Long: min_i = LBound(parA)
    Dim max_i As Long: max_i = UBound(parA)
    Dim par As Variant
    Dim tRA As Variant
    Dim row As Long
    Dim column As Long

    CONCAT = ""
    
    If max_i - min_i + 1 < 1 Then Exit Function
    
    For Each par In parA
        If TypeName(par) = "Range" Then
            tRA = par.Value2
            If IsArray(tRA) Then
                For row = LBound(tRA, 1) To UBound(tRA, 1)
                    For column = LBound(tRA, 2) To UBound(tRA, 2)
                        CONCAT = CONCAT & CStr(tRA(row, column))
                    Next
                Next
            Else
                CONCAT = CONCAT & CStr(tRA)
            End If
        ElseIf IsArray(par) Then
            For Each tRA In par
                CONCAT = CONCAT & CStr(tRA)
            Next
        Else
            CONCAT = CONCAT & CStr(par)
        End If
    Next
End Function

' [TEXTJOIN関数の使い方とExcel2013以前の古いエクセルで使う方法](https://www.excelspeedup.com/TEXTJOIN2/)
Function TEXTJOIN(Delim, Ignore As Boolean, ParamArray parA())
    Dim min_i As Long: min_i = LBound(parA)
    Dim max_i As Long: max_i = UBound(parA)
    Dim par As Variant
    Dim tRA As Variant
    Dim row As Long
    Dim column As Long

    TEXTJOIN = ""
    
    If max_i - min_i + 1 < 1 Then Exit Function
    
    For Each par In parA
        If TypeName(par) = "Range" Then
            tRA = par.Value2
            If IsArray(tRA) Then
                For row = LBound(tRA, 1) To UBound(tRA, 1)
                    For column = LBound(tRA, 2) To UBound(tRA, 2)
                        If tRA(row, column) <> "" Or Ignore = False Then
                            TEXTJOIN = TEXTJOIN & Delim & CStr(tRA(row, column))
                        End If
                    Next
                Next
            Else
                If tRA <> "" Or Ignore = False Then
                    TEXTJOIN = TEXTJOIN & Delim & CStr(tRA)
                End If
            End If
        ElseIf IsArray(par) Then
            For Each tRA In par
                If tRA <> "" Or Ignore = False Then
                    TEXTJOIN = TEXTJOIN & Delim & CStr(tRA)
                End If
            Next
        Else
            If par <> "" Or Ignore = False Then
                TEXTJOIN = TEXTJOIN & Delim & CStr(par)
            End If
        End If
    Next
  
    TEXTJOIN = Mid(TEXTJOIN, Len(Delim) + 1)
End Function

' [ユーザー定義関数：MAXIFS・MINIFS（Excel 2013以前向け）](https://gist.github.com/furyutei/ca02a52e564535e051f1d96eba390e8d#file-maxifs)
Function MAXIFS(max_range As Range, ParamArray criteria_list())
    Dim max_range_value_array As Variant
    Dim max_range_width As Integer
    Dim max_range_height As Integer
    Dim row_index As Integer
    Dim column_index As Integer
    
    Dim criteria_range_array() As Range
    Dim criteria_range_value_array() As Variant
    Dim criteria_condition_array As Variant
    Dim criteria_number As Integer
    Dim criteria_index As Integer
    
    Dim is_valid As Boolean
    Dim max_value As Variant
    
    MAXIFS = CVErr(xlErrValue)
    
    criteria_number = UBound(criteria_list) - LBound(criteria_list) + 1
    
    If criteria_number Mod 2 <> 0 Then
        'On Error Resume Next
        'Err.Raise Number:=450 ' 引数の数が一致していません。または不正なプロパティを指定しています。
        'MsgBox CStr(Err.Number) & " : " & Err.Description
        'Err.Clear
        'On Error GoTo 0
        Exit Function
    End If
    
    criteria_number = criteria_number / 2
    ReDim criteria_range_array(criteria_number)
    ReDim criteria_range_value_array(criteria_number)
    ReDim criteria_condition_array(criteria_number)
    
    max_range_value_array = max_range
    
    max_range_height = UBound(max_range_value_array)
    max_range_width = UBound(max_range_value_array, 2)
    
    For criteria_index = 1 To criteria_number
        Set criteria_range_array(criteria_index) = criteria_list((criteria_index - 1) * 2)
        criteria_range_value_array(criteria_index) = criteria_list((criteria_index - 1) * 2)
        criteria_condition_array(criteria_index) = criteria_list((criteria_index - 1) * 2 + 1)
        
        If (UBound(criteria_range_value_array(criteria_index)) <> max_range_height) Or _
            (UBound(criteria_range_value_array(criteria_index), 2) <> max_range_width) _
        Then
            Exit Function
        End If
    Next criteria_index
    
    max_value = Empty
    
    For row_index = 1 To max_range_height
        For column_index = 1 To max_range_width
            is_valid = True
            
            For criteria_index = 1 To criteria_number
                ' TODO: 条件が式の場合に正しく動作しない→作り込みが困難なため、COUNTIF()を利用
                'If criteria_range_value_array(criteria_index)(row_index, column_index) <> criteria_condition_array(criteria_index) Then
                '    is_valid = False
                '    Exit For
                'End If
                
                If Application.WorksheetFunction.CountIf( _
                        criteria_range_array(criteria_index).Offset(row_index - 1, column_index - 1).Cells(1, 1), _
                        criteria_condition_array(criteria_index) _
                    ) = 0 _
                Then
                    is_valid = False
                    Exit For
                End If
            Next criteria_index
            
            If is_valid = True Then
                If max_value = Empty Then
                    max_value = max_range_value_array(row_index, column_index)
                Else
                    max_value = Application.WorksheetFunction.Max(max_value, max_range_value_array(row_index, column_index))
                End If
            End If
        Next column_index
    Next row_index
    
    If max_value <> Empty Then
        MAXIFS = max_value
    Else
        MAXIFS = 0
    End If
End Function

' [ユーザー定義関数：MAXIFS・MINIFS（Excel 2013以前向け）](https://gist.github.com/furyutei/ca02a52e564535e051f1d96eba390e8d#file-minifs)
Function MINIFS(min_range As Range, ParamArray criteria_list())
    Dim min_range_value_array As Variant
    Dim min_range_width As Integer
    Dim min_range_height As Integer
    Dim row_index As Integer
    Dim column_index As Integer
    
    Dim criteria_range_array() As Range
    Dim criteria_range_value_array() As Variant
    Dim criteria_condition_array As Variant
    Dim criteria_number As Integer
    Dim criteria_index As Integer
    
    Dim is_valid As Boolean
    Dim min_value As Variant
    
    MINIFS = CVErr(xlErrValue)
    
    criteria_number = UBound(criteria_list) - LBound(criteria_list) + 1
    
    If criteria_number Mod 2 <> 0 Then
        'On Error Resume Next
        'Err.Raise Number:=450 ' 引数の数が一致していません。または不正なプロパティを指定しています。
        'MsgBox CStr(Err.Number) & " : " & Err.Description
        'Err.Clear
        'On Error GoTo 0
        Exit Function
    End If
    
    criteria_number = criteria_number / 2
    ReDim criteria_range_array(criteria_number)
    ReDim criteria_range_value_array(criteria_number)
    ReDim criteria_condition_array(criteria_number)
    
    min_range_value_array = min_range
    
    min_range_height = UBound(min_range_value_array)
    min_range_width = UBound(min_range_value_array, 2)
    
    For criteria_index = 1 To criteria_number
        Set criteria_range_array(criteria_index) = criteria_list((criteria_index - 1) * 2)
        criteria_range_value_array(criteria_index) = criteria_list((criteria_index - 1) * 2)
        criteria_condition_array(criteria_index) = criteria_list((criteria_index - 1) * 2 + 1)
        
        If (UBound(criteria_range_value_array(criteria_index)) <> min_range_height) Or _
            (UBound(criteria_range_value_array(criteria_index), 2) <> min_range_width) _
        Then
            Exit Function
        End If
    Next criteria_index
    
    min_value = Empty
    
    For row_index = 1 To min_range_height
        For column_index = 1 To min_range_width
            is_valid = True
            
            For criteria_index = 1 To criteria_number
                ' TODO: 条件が式の場合に正しく動作しない→作り込みが困難なため、COUNTIF()を利用
                'If criteria_range_value_array(criteria_index)(row_index, column_index) <> criteria_condition_array(criteria_index) Then
                '    is_valid = False
                '    Exit For
                'End If
                
                If Application.WorksheetFunction.CountIf( _
                        criteria_range_array(criteria_index).Offset(row_index - 1, column_index - 1).Cells(1, 1), _
                        criteria_condition_array(criteria_index) _
                    ) = 0 _
                Then
                    is_valid = False
                    Exit For
                End If
            Next criteria_index
            
            If is_valid = True Then
                If min_value = Empty Then
                    min_value = min_range_value_array(row_index, column_index)
                Else
                    min_value = Application.WorksheetFunction.Min(min_value, min_range_value_array(row_index, column_index))
                End If
            End If
        Next column_index
    Next row_index
    
    If min_value <> Empty Then
        MINIFS = min_value
    Else
        MINIFS = 0
    End If
End Function


