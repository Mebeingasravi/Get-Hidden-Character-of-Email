Dim TrimedName As String

Sub Button1_Click()

Dim FullName As String
Dim valss() As String
Dim i As Boolean
Dim Email As String
Dim TempK As String


For Each cell In Range("A11:A101")
    'MsgBox cell.Value
    
    FullName = cell.Value
    
    valss = getSplitName(FullName)
    
    Email = getEmailAddress(cell.row, 3)  'MsgBox Email
    
    For Each k In valss
        TempK = k
        
     'check name and email 5th character
     i = CheckName(TempK, Email)
     
     If i = True Then
         'MsgBox cell.row & TrimedName & getTrimedEmail(Email)
         Cells(cell.row, 12) = TrimedName & getTrimedEmail(Email)
    End If
    Next
Next

'MsgBox "Check the list"


'FullName = Cells(4, 1)
'MsgBox FullName

'valss = getSplitName(FullName)

'Email = getEmailAddress()

'For Each k In valss
  '  TempK = k
    'MsgBox TempK
    
    'check name and email 5th character
 '   i = CheckName(TempK, Email)
    
 '   If i = True Then
  '      Cells(5, 12) = TrimedName & getTrimedEmail("****tri.prasanna123@gmail.com")
   '     MsgBox TrimedName & getTrimedEmail("****tri.prasanna123@gmail.com")
   ' End If
'Next

End Sub

Public Function getSplitName(FullName As String) As String()

Dim SliptNames() As String

SplitNames = Split(FullName, " ")

getSplitName = SplitNames

End Function

Public Function getEmailAddress(row As Integer, clm As Integer) As String

getEmailAddress = Cells(row, clm)

End Function

Public Function CheckName(Name As String, Email As String) As Variant

Dim i As Boolean
Dim fourthofName As String
Dim fourthofEmail As String

i = False
fourthofName = Mid(Name, 5, 1)
fourthofEmail = Mid(Email, 5, 1)

If fourthofName = fourthofEmail Then
    i = True
    TrimedName = Left(Name, 4)
Else
    i = False
End If

CheckName = i

End Function

Public Function getTrimedEmail(Email As String) As String
getTrimedEmail = Right(Email, Len(Email) - 4)
End Function
