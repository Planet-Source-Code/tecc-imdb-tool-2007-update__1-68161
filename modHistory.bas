Attribute VB_Name = "modHistory"
Public History() As MOVIE_DATA_IMDB


'history module

'Load History Sub

Public Sub LoadHistory()
On Error GoTo ERRR:
Dim Buf2 As String

Dim HS1() As String
Dim HS2() As String

ReDim History(0)


Buf2 = Space$(FileLen(IIf(Right(App.Path, 1) = "\", App.Path & "history.dat", App.Path & "\history.dat")))



Open IIf(Right(App.Path, 1) = "\", App.Path & "history.dat", App.Path & "\history.dat") For Binary Access Read As #1
    Get #1, , Buf2
Close #1

HS1 = Split(Buf2, "»®©§" & vbCrLf & vbCrLf)
'"«©©®"
'"»®©§"

If UBound(HS1) >= 0 Then
    For i = 0 To UBound(HS1)
        HS2 = Split(HS1(i), "«©©®")
        If UBound(HS2) > 0 Then
            'add another history entry
            ReDim Preserve History(UBound(History) + 1)
                With History(UBound(History))
                    .Country = HS2(0)
                    .CoverURL = HS2(1)
                    .Language = HS2(2)
                    .mDate = HS2(3)
                    .mGenre = HS2(4)
                    .mSypnosys = HS2(5)
                    .mTitle = HS2(6)
                    .Offline = True
                    .Runtime = HS2(7)
                    .userRating = HS2(8)
                    .ttID = HS2(9)
                    .v4_Tagline = HS2(10)
                    .v4_MpaaRating = HS2(11)
                End With
        End If
    Next
End If

Exit Sub
ERRR:
If Err.Number = 53 Then
    'no history found
    frmMain.sb1.Panels(2).Text = "Warning: No history file was found!"
End If
End Sub

Public Sub SaveHistory()
Dim OUT__STR As String
For i = 0 To UBound(History)
    With History(i)
        If Trim(.mTitle) <> "" Then
            OUT__STR = OUT__STR & .Country & "«©©®" & _
                .CoverURL & "«©©®" & _
                .Language & "«©©®" & _
                .mDate & "«©©®" & _
                .mGenre & "«©©®" & _
                .mSypnosys & "«©©®" & _
                .mTitle & "«©©®" & _
                .Runtime & "«©©®" & _
                .userRating & "«©©®" & _
                .ttID & "«©©®" & _
                .v4_Tagline & "«©©®" & _
                .v4_MpaaRating & "»®©§" & vbCrLf & vbCrLf
        End If
    End With
Next

Open IIf(Right(App.Path, 1) = "\", App.Path & "history.dat", App.Path & "\history.dat") For Output As #1
    Print #1, OUT__STR
Close #1
End Sub

Public Sub Add_History_Item(sItem As MOVIE_DATA_IMDB)
Dim uTN As Long
Dim SSSSS() As String
'check for open entry
For i = 0 To UBound(History)
    With History(i)
        If Trim(.mTitle) = "" Then
            'use this number
            uTN = i
            GoTo FOUND_FREE:
        End If
    End With
Next

uTN = UBound(History) + 1
ReDim Preserve History(uTN)

'append to history
FOUND_FREE:

SSSSS = Split(sItem.CoverURL, "/")

With History(uTN)
    If UBound(SSSSS) < 0 Then
    .CoverURL = ""
    Else
    .CoverURL = SSSSS(UBound(SSSSS))
    End If
    .Offline = True
    .Country = sItem.Country
    .Language = sItem.Language
    .mDate = sItem.mDate
    .mGenre = sItem.mGenre
    .mSypnosys = sItem.mSypnosys
    .mTitle = sItem.mTitle
    .Runtime = sItem.Runtime
    .userRating = sItem.userRating
    .ttID = sItem.ttID
    .v4_Tagline = sItem.v4_Tagline
    .v4_MpaaRating = sItem.v4_MpaaRating
End With
End Sub

