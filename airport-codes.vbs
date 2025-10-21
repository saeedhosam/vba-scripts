Private Sub Worksheet_Change(ByVal Target As Range)
    Dim code As String, city As String
    Dim wsAirports As Worksheet
    Dim f As Range
    Dim match As Object
    Dim matches As Object
    Dim newValue As String
    If Target.CountLarge > 1 Then Exit Sub
    If IsEmpty(Target.Value) Then Exit Sub
    Set wsAirports = Sheets("airports")
    newValue = Target.Value
    With CreateObject("VBScript.RegExp")
        .Pattern = "\b[A-Z]{3}\b"
        .Global = True
        If .Test(Target.Value) Then
            Set matches = .Execute(Target.Value)
            For Each match In matches
                code = match.Value
                Set f = wsAirports.Columns(1).Find(What:=code, LookIn:=xlValues, LookAt:=xlWhole)
                If Not f Is Nothing Then
                    city = f.Offset(0, 1).Value
                    newValue = Replace(newValue, code, city)
                End If
            Next match
            If newValue <> Target.Value Then
                Application.EnableEvents = False
                Target.Value = newValue
                Application.EnableEvents = True
            End If
        End If
    End With
End Sub
