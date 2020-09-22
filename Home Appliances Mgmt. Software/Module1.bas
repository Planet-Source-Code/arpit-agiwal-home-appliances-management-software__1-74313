Attribute VB_Name = "Module1"
Public accounttype As String, username As String, logintime, emprecclick As String, suprecclick As String, prorecclick As String, salesclick As String, billno As String, salescustname As String, salescontactno As String, salesaddress As String, salesdate As String, purchasesclick, pbillno As String, psupid As String, psupname  As String, psupcontactno As String, psupemailid As String, psupcity As String, pdate As String

Public Sub UnloadAllForms(Optional FormToIgnore As String = "")
Dim f As Form
    For Each f In Forms
        If f.Name <> FormToIgnore Then
        Unload f
        Set f = Nothing
        End If
    Next f
End Sub

Public Function onlyalpha(k1 As Integer)
If (k1 >= 65 And k1 <= 90) Or (k1 >= 97 And k1 <= 122) Or (k1 = 32 Or k1 = 8) Then
Else
k1 = 0
End If
End Function

Public Function onlyemail(k As Integer)
If (k >= 65 And k <= 90) Or (k >= 97 And k <= 122) Or (k >= 48 And k <= 57) Or (k = 32 Or k = 8 Or k = 64 Or k = 46) Then
Else
k = 0
End If
End Function

Public Function onlynumeric(k As Integer)
If (k >= 48 And k <= 57) Or (k = 8) Then
Else
k = 0
End If
End Function
Public Function onlyalphanum(k As Integer)
If (k >= 65 And k <= 90) Or (k >= 97 And k <= 122) Or (k = 32 Or k = 8 Or k = 127) Or (k >= 48 And k <= 57) Then
Else
k = 0
End If
End Function

Public Function onlyaddress(k As Integer)
If (k >= 65 And k <= 90) Or (k >= 97 And k <= 122) Or (k = 32 Or k = 8 Or k = 44 Or k = 47) Or (k >= 48 And k <= 57) Then
Else
k = 0
End If
End Function

