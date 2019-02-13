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
Function CONCAT(ParamArray para()) As String
    CONCAT = TextJoinCommon("", True, para)
End Function

' [TEXTJOIN関数の使い方とExcel2013以前の古いエクセルで使う方法](https://www.excelspeedup.com/TEXTJOIN2/)
Function TEXTJOIN(Delim, Ignore As Boolean, ParamArray para()) As String
    TEXTJOIN = TextJoinCommon(Delim, Ignore, para)
End Function

Private Function TextJoinCommon(Delim, Ignore As Boolean, ByVal para As Variant) As String
    Dim min_i As Long: min_i = LBound(para)
    Dim max_i As Long: max_i = UBound(para)
    Dim par As Variant
    Dim value As Variant
    Dim value_counter As Long
    Dim delim_array As Variant: delim_array = ToArray(Delim)
    Dim delim_number As Long: delim_number = UBound(delim_array) - LBound(delim_array) + 1
    Dim delim_str As String
    
    TextJoinCommon = ""
    
    If max_i - min_i + 1 < 1 Then Exit Function
    
    value_counter = 0
    delim_str = ""

    For Each par In para
        For Each value In ToArray(par)
            value = CStr(value)
            If value <> "" Or Ignore = False Then
                delim_str = CStr(delim_array(value_counter Mod delim_number))
                TextJoinCommon = TextJoinCommon & CStr(value) & delim_str
                value_counter = value_counter + 1
            End If
        Next
    Next
    
    TextJoinCommon = Mid(TextJoinCommon, 1, Len(TextJoinCommon) - Len(delim_str))
End Function

Private Function ToArray(ByVal par As Variant) As Variant()
    Dim values As Variant
    Dim row As Long
    Dim column As Long
    Dim index As Long
    Dim text_array() As Variant
    
    If TypeName(par) = "Range" Then
        values = par.Value2
        If IsArray(values) Then
            ReDim text_array(0 To (UBound(values, 1) - LBound(values, 1) + 1) * (UBound(values, 2) - LBound(values, 2) + 1) - 1)
            index = 0
            For row = LBound(values, 1) To UBound(values, 1)
                For column = LBound(values, 2) To UBound(values, 2)
                    text_array(index) = values(row, column)
                    index = index + 1
                Next
            Next
        Else
            text_array = Array(values)
        End If
    ElseIf IsArray(par) Then
        ReDim Preserve par(0 To UBound(par) - LBound(par))
        text_array = par
    Else
        text_array = Array(par)
    End If
    
    ToArray = text_array
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


