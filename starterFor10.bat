'Nicked the pos in array code from here https://stackoverflow.com/a/50382597/6166127

Public Function posInArray(ByVal itemSearched As Variant, ByVal aArray As Variant) As Long
Dim pos As Long, item As Variant

posInArray = -1
If IsArray(aArray) Then
    If Not IsEmpty(aArray) Then
        pos = 1
        For Each item In aArray
            If itemSearched = item Then
                posInArray = pos
                Exit Function
            End If
            pos = pos + 1
        Next item
        posInArray = 0
    End If
End If

End Function
    
Sub PrintArray(Data As Variant, Cl As Range)
    'I need to set the lbound to fix this
    Cl.Resize(1 + UBound(Data, 1), UBound(Data, 2)) = Data
End Sub


Sub biadjmat()
    
    Dim r As Range
    Dim target, source As Variant
    Dim grants, pis As Object
    
    target = ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").ListColumns(1).DataBodyRange.value
    source = ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").ListColumns(3).DataBodyRange.value
    
    
    Set grants = CreateObject("Scripting.Dictionary")
    Set pis = CreateObject("Scripting.Dictionary")
    
    
    Dim i As Long
    For i = LBound(target) To UBound(target)
        If Not grants.Exists(target(i, 1)) Then
            grants.Add target(i, 1), 1
        End If
        If Not pis.Exists(source(i, 1)) Then
            pis.Add source(i, 1), 1
        End If
    Next i
    
    gk = grants.keys
    pk = pis.keys
    
    Dim ban() As Integer
    ReDim ban(1 To grants.Count, 1 To pis.Count)
    
    
    Dim j As Long
        For j = LBound(target) To UBound(target)
            s = posInArray(target(j, 1), gk)
            t = posInArray(source(j, 1), pk)
            ban(s, t) = 1
        Next j
        
     oneModeProj = Application.MMult(Application.Transpose(ban), ban)
     
     PrintArray oneModeProj, ActiveWorkbook.Worksheets("Sheet2").[b2]
     
     'WIP
'    Dim TempG(), TempP() As Variant
'
'
'    ReDim TempG(0 To UBound(gk), 0 To 1)
'    ReDim TempP(0 To UBound(pk), 0 To 1)
'
'    For i = 0 To UBound(gk)
'        TempG(i, 0) = gk(i)
'    Next i
'    gk = TempG
'
'
'
'
'    For i = 0 To UBound(pk)
'        TempP(i, 0) = pk(i)
'    Next i
'    pk = Application.Transpose(TempP)
'
'
'    PrintArray gk, ActiveWorkbook.Worksheets("Sheet2").[a2]
'    PrintArray pk, ActiveWorkbook.Worksheets("Sheet2").[b1]
'    PrintArray ban, ActiveWorkbook.Worksheets("Sheet2").[b2]
End Sub
