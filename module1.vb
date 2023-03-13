Function GETCUSTOMDESCRIPTION(myDesc As String, myOption As Integer)
  
  Dim arrValues() As String
  arrValues = Split(myDesc, " ")
  
  Dim arrQty() As String
  Dim arrType() As String
  Dim arrWeight() As String
 
  Dim counter As Integer
  counter = 0
  
  Dim regex1 As Object
  Dim regex2 As Object
  Dim regex3 As Object
  Set regex1 = New RegExp
  Set regex2 = New RegExp
  Set regex3 = New RegExp
  regex1.Pattern = "^\d+(?=[a-zA-Z]+=)" 'e.g. 8OV=1.16CTW get 8
  regex2.Pattern = "[a-zA-Z]+(?==)" 'e.g. 8OV=1.16CTW get OV
  regex3.Pattern = "\d?\.?\d+(?=CT)" 'e.g. 8OV=1.16CTW get 1.16
  
  Dim componentQty As Object
  Dim componentType As Object
  Dim componentWeight As Object
  
  For Each Item In arrValues
    If (InStr(Item, "=") > 0) Then 'Items with equal values are components
        ReDim Preserve arrQty(counter)
        ReDim Preserve arrType(counter)
        ReDim Preserve arrWeight(counter)
        
        Set componentQty = regex1.Execute(Item)
        Set componentType = regex2.Execute(Item)
        Set componentWeight = regex3.Execute(Item)
        
        If componentQty.Count() > 0 Then
           arrQty(counter) = componentQty(0).Value
        Else
           arrQty(counter) = "1"
        End If
        
        If componentType.Count() Then
            If arrQty(counter) > 1 Then
              arrType(counter) = GetColumnValueByValue("COMPONENTS", componentType(0).Value, 3) 'if Plural
            Else
              arrType(counter) = GetColumnValueByValue("COMPONENTS", componentType(0).Value, 2)
            End If
        Else
           arrType(counter) = ""
        End If
        
        If componentWeight.Count() Then
            If Left(componentWeight(0).Value, 1) = "." Then
              arrWeight(counter) = "0" & componentWeight(0).Value
            Else
              arrWeight(counter) = componentWeight(0).Value
            End If
        Else
           arrWeight(counter) = ""
        End If
        
        counter = counter + 1
    End If
  Next Item
  
  Dim i As Integer
  Dim strDesc As String
  
  If myOption = 0 Then ' return stone quantity string
    For i = 0 To UBound(arrQty)
      If strDesc = "" Then
        strDesc = strDesc & arrQty(i) & " " & arrType(i)
      Else
        strDesc = strDesc & ", " & arrQty(i) & " " & arrType(i)
      End If
    Next i
  ElseIf myOption = 1 Then ' return stone weight string
    For i = 0 To UBound(arrQty)
      If strDesc = "" Then
        strDesc = strDesc & arrWeight(i) & "CT " & arrType(i)
      Else
        strDesc = strDesc & ", " & arrWeight(i) & "CT " & arrType(i)
      End If
    Next i
  End If
  
  GETCUSTOMDESCRIPTION = LCase(strDesc)
End Function

Function GetColumnValueByValue(sheetName As String, lookupValue As Variant, returnColumn As Long) As Variant
    Dim lookupSheet As Worksheet
    Dim columnValues As Range
    Dim lookupIndex As Long
    Dim lastRow As Long
    
    Set lookupSheet = ThisWorkbook.Worksheets(sheetName)
    
    ' Get the last row of data in column A
    lastRow = lookupSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set the range to include all columns and rows with data
    Set columnValues = lookupSheet.Range("A1").Resize(lastRow)
    
    ' Find the index of the lookup value in the column of values
    lookupIndex = Application.Match(lookupValue, columnValues, 0)
    
    If IsNumeric(lookupIndex) Then
        ' Return the value in the specified row and column of the column of values
        GetColumnValueByValue = lookupSheet.Range("A1").Offset(lookupIndex - 1, returnColumn - 1).Value
    Else
        ' Return an error message if the lookup value is not found
        GetColumnValueByValue = "Value not found"
    End If
End Function
