VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const LOCALE_USER_DEFAULT = &H400
Const LOCALE_SENGCOUNTRY = &H1002 ' «”„ «·œÊ·… »«·≈‰Ã·Ì“Ì
Const LOCALE_SENGLANGUAGE = &H1001 ' «”„ «··€… »«·≈‰Ã·Ì“Ì
Const LOCALE_SNATIVELANGNAME = &H4 ' «”„ «··€… »«··€… «·Êÿ‰Ì…
Const LOCALE_SNATIVECTRYNAME = &H8 ' «”„ «·œÊ·… »«··€… «·Êÿ‰Ì…

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Function GetInfo(ByVal lInfo As Long) As String
    Dim Buffer As String, Ret As String
    
    Buffer = String$(256, 0)
    Ret = GetLocaleInfo(LOCALE_USER_DEFAULT, lInfo, Buffer, Len(Buffer))
    If Ret > 0 Then
        GetInfo = Left$(Buffer, Ret - 1)
    Else
        GetInfo = ""
    End If
End Function

Private Sub Form_Load()
    MsgBox "√‰  „ﬁÌ„ ›Ì " & GetInfo(LOCALE_SNATIVECTRYNAME) & " (" & GetInfo(LOCALE_SENGCOUNTRY) & ")," & vbCrLf & "Ê ·€ ﬂ ÂÌ " & GetInfo(LOCALE_SNATIVELANGNAME) & " (" & GetInfo(LOCALE_SENGLANGUAGE) & ").", vbInformation Or vbMsgBoxRight Or vbMsgBoxRtlReading

End Sub
