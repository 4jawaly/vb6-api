VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Mobily.ws"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6060
      Top             =   4980
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox SSTab1 
      Height          =   7815
      Left            =   120
      ScaleHeight     =   7755
      ScaleWidth      =   7185
      TabIndex        =   2
      Top             =   120
      Width           =   7245
      Begin VB.TextBox TxtActivecode2 
         Height          =   375
         Left            =   1680
         TabIndex        =   21
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox TxtActivationCode 
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CommandButton CmdRegisterSender 
         Caption         =   "«—”«· ﬂÊœ «· ›⁄Ì· "
         Height          =   615
         Left            =   1920
         TabIndex        =   15
         Top             =   3960
         Width           =   2175
      End
      Begin VB.CommandButton CmdAddSender 
         Caption         =   "ÕÃ“ «”„ «·„—”· ⁄·Ï «·„Êﬁ⁄"
         Height          =   615
         Left            =   2160
         TabIndex        =   14
         Top             =   7680
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton CmdCheckSender 
         Caption         =   "«÷€ÿ Â‰« ·«÷«›Â «”„ «·„—”·"
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox TxtSenderName 
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "«—”‹‹‹«· «·‹‹—”‹‹‹‹«·‹‹Â"
         Height          =   495
         Left            =   -74400
         TabIndex        =   6
         Top             =   6960
         Width           =   2175
      End
      Begin VB.TextBox txtNumbers 
         Height          =   1935
         Left            =   -74520
         TabIndex        =   5
         Text            =   "966555664326,966564101705"
         Top             =   4920
         Width           =   5655
      End
      Begin VB.TextBox txtMessage 
         Height          =   1935
         Left            =   -74520
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "mobily.frx":0000
         Top             =   2280
         Width           =   5655
      End
      Begin VB.TextBox txtSender 
         Height          =   375
         Left            =   -71280
         TabIndex        =   3
         Text            =   "4jawaly.net"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   " :ﬂÊœ «· ›⁄Ì· «·„—”· ⁄·Ï «·ÃÊ«·"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4200
         TabIndex        =   22
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "›Ì Õ«· ÿ·» «”„ „—”· —ﬁ„ ÃÊ«· Ì „ «—”«· ﬂÊœ ·· Õﬁﬁ „‰ „·ﬂÌ ﬂ ··—ﬁ„ - Ì „ «œŒ«· ﬂÊœ «· Õﬁﬁ Â‰« "
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   20
         Top             =   2280
         Width           =   6735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "›Ì Õ«·… ⁄œ„ ÊÃÊœ «”„ «·„—”· Ì „ ÕÃ“ «”„ «·„—”· ⁄·Ï «·„Êﬁ⁄"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   6960
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„—”· 11 "" Õ—› «‰Ã·Ì“Ì "" «Ê "" —ﬁ„ "" ‘«„·Â «·„”«›«  Ê·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ «”„ «·„—”· »√Õ—› ⁄—»ÌÂ"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   480
         TabIndex        =   18
         Top             =   1440
         Width           =   6135
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   " :—ﬁ„ «·ÃÊ«· «·Œ«’ »ﬂ"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4080
         TabIndex        =   17
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "«œŒ· «”„ «·„‹‹—”‹‹·"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblBalance 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -74640
         TabIndex        =   10
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "‰‹‹‹’ «·‹—”‹‹«·‹Â"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -70080
         TabIndex        =   9
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "«”‹‹„ «·„‹‹—”‹‹·"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -69840
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   " :«·«—ﬁ‹‹‹«„ «·„—«œ «·«—”‹‹‹‹«· «·ÌÂ‹‹‹«"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   -71160
         TabIndex        =   7
         Top             =   4560
         Width           =   2415
      End
   End
   Begin VB.PictureBox Inet13 
      Height          =   480
      Left            =   1440
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   23
      Top             =   8400
      Width           =   1200
   End
   Begin VB.ListBox List1 
      Height          =   2205
      ItemData        =   "mobily.frx":001C
      Left            =   2640
      List            =   "mobily.frx":01D6
      TabIndex        =   1
      Top             =   8400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   2205
      ItemData        =   "mobily.frx":0392
      Left            =   4560
      List            =   "mobily.frx":054C
      TabIndex        =   0
      Top             =   8400
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''###################'''''''''''''''''''''''
'''''''''''''''#                 #'''''''''''''''''''''''
'''''''''''''''# www.4jawaly.net #'''''''''''''''''''''''
'''''''''''''''#                 #'''''''''''''''''''''''
'''''''''''''''###################'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'===================================================================
Const UserName = "1111" ' Enter Your User Name Here
Const Password = "1111" ' Enter Your Password Here
'===================================================================


Private Sub CmdAddSender_Click()
On Error Resume Next
Dim s As String
Dim result As String

s = "http://www.4jawaly.net/apiSjawaly/addSender.php?username=" & UserName & "&password=" & Password & "&Sendername=" & TxtSenderName.Text

result = Inet1.OpenURL(s)

Select Case result
Case "101": MsgBox ("«·»Ì«‰«  ‰«ﬁ’…")
Case "102": MsgBox ("«”„ «·„” Œœ„ €Ì— ’ÕÌÕ")
Case "103": MsgBox ("ﬂ·„… «·„—Ê— €Ì— ’ÕÌÕ…")
Case "104": MsgBox ("·« ÌÊÃœ —’Ìœ ›Ï «·Õ”«»")
Case "105": MsgBox ("«·«—”«· „€·ﬁ")
Case "106": MsgBox ("«·Õ”«» €Ì— „›⁄·")
Case "107": MsgBox ("«·Õ”«» „ÊﬁÊ›")
Case "108": MsgBox ("€Ì— „›⁄· ÃÊ«·")
Case "109": MsgBox ("€Ì— „›⁄· »—Ìœ «·ﬂ —Ê‰Ï")
Case "110": MsgBox (" „ «÷«›… «”„ «·„—”· »‰Ã«Õ Ê”Ê› Ì „  ›⁄Ì·… ›Ï «ﬁ—» Êﬁ ")
Case "111": MsgBox ("„‰ ›÷·ﬂ «œŒ· «”„ „—”· ’ÕÌÕ")
Case "112": MsgBox ("«·”‰œ— ‰Ì„ ·«»œ Ê«‰ ÌﬂÊ‰ 15 —ﬁ„ «Ê 11 Õ—› Ê—ﬁ„  - _. ‘«„·« «·„”«›« ")
Case "113": MsgBox ("«”„ «·„—”· „‰ «·«”„«¡ «·„ÕÃÊ»… «Ê „‘«»… ·«”„ „‰ «·«”„«¡ «·„ÕÃÊ»…")
Case "114": MsgBox ("«”„ «·„—”· „ÊÃÊœ „‰ ﬁ»·")
Case "115": MsgBox ("·Ì” ·œÌﬂ —’Ìœ ·ÿ·» «”„ „—”· —ﬁ„Ï")
Case "116": MsgBox ("›‘· ›Ï «·«—”«· Õ«Ê· „—… «Œ—Ï «Ê « ’· »«·œ⁄„ «·›‰Ï")
Case "117": MsgBox (" „ «÷«›… «”„ «·„—”· »‰Ã«Õ Ê „ «—”«· ﬂÊœ  ›⁄Ì· «·Ï «·—ﬁ„")
    End Select

TxtActivationCode.Text = Split(result, "#")(1)

End Sub

Private Sub CmdCheckSender_Click()
On Error Resume Next
Dim s As String
Dim result As String
s = "http://www.4jawaly.net/apiSjawaly/addSender.php?username=" & UserName & "&password=" & Password & "&Sendername=" & TxtSenderName.Text
result = Inet1.OpenURL(s)

Select Case result
    Case "101": MsgBox ("«·»Ì«‰«  ‰«ﬁ’…")
    Case "102": MsgBox ("«”„ «·„” Œœ„ €Ì— ’ÕÌÕ")
    Case "103": MsgBox ("ﬂ·„… «·„—Ê— €Ì— ’ÕÌÕ…")
    Case "104": MsgBox ("·« ÌÊÃœ —’Ìœ ›Ï «·Õ”«»")
    Case "105": MsgBox ("«·«—”«· „€·ﬁ")
    Case "106": MsgBox ("«·Õ”«» €Ì— „›⁄·")
    Case "107": MsgBox ("«·Õ”«» „ÊﬁÊ›")
    Case "108": MsgBox ("€Ì— „›⁄· ÃÊ«·")
    Case "109": MsgBox ("€Ì— „›⁄· »—Ìœ «·ﬂ —Ê‰Ï")
    Case "110": MsgBox (" „ «÷«›… «”„ «·„—”· »‰Ã«Õ Ê”Ê› Ì „  ›⁄Ì·… ›Ï «ﬁ—» Êﬁ ")
    Case "111": MsgBox ("„‰ ›÷·ﬂ «œŒ· «”„ „—”· ’ÕÌÕ")
    Case "112": MsgBox ("«·”‰œ— ‰Ì„ ·«»œ Ê«‰ ÌﬂÊ‰ 15 —ﬁ„ «Ê 11 Õ—› Ê—ﬁ„  - _. ‘«„·« «·„”«›« ")
    Case "113": MsgBox ("«”„ «·„—”· „‰ «·«”„«¡ «·„ÕÃÊ»… «Ê „‘«»… ·«”„ „‰ «·«”„«¡ «·„ÕÃÊ»…")
    Case "114": MsgBox ("«”„ «·„—”· „ÊÃÊœ „‰ ﬁ»·")
    Case "115": MsgBox ("·Ì” ·œÌﬂ —’Ìœ ·ÿ·» «”„ „—”· —ﬁ„Ï")
    Case "116": MsgBox ("›‘· ›Ï «·«—”«· Õ«Ê· „—… «Œ—Ï «Ê « ’· »«·œ⁄„ «·›‰Ï")
    Case "117": MsgBox (" „ «÷«›… «”„ «·„—”· »‰Ã«Õ Ê „ «—”«· ﬂÊœ  ›⁄Ì· «·Ï «·—ﬁ„")
    Case Else: MsgBox (result)
End Select

End Sub

Private Sub CmdRegisterSender_Click()
On Error Resume Next
Dim s As String
Dim result As String


s = "http://www.4jawaly.net/apiSjawaly/ActiveSende.php?username=" & UserName & "&password=" & Password & "&Snderid=" & TxtActivationCode.Text & "&Activecode=" & TxtActivecode2.Text

result = Inet1.OpenURL(s)

Select Case result

    Case "101": MsgBox ("«·»Ì«‰«  ‰«ﬁ’…")
    Case "102": MsgBox ("«”„ «·„” Œœ„ €Ì— ’ÕÌÕ")
    Case "103": MsgBox ("ﬂ·„… «·„—Ê— €Ì— ’ÕÌÕ…")
    Case "104": MsgBox ("·« ÌÊÃœ —’Ìœ ›Ï «·Õ”«»")
    Case "105": MsgBox ("«·«—”«· „€·ﬁ")
    Case "106": MsgBox ("«·Õ”«» €Ì— „›⁄·")
    Case "107": MsgBox ("«·Õ”«» „ÊﬁÊ›")
    Case "108": MsgBox ("€Ì— „›⁄· ÃÊ«·")
    Case "109": MsgBox ("€Ì— „›⁄· »—Ìœ «·ﬂ —Ê‰Ï")
    Case "110": MsgBox (" „  ›⁄Ì· «”„ «·„—”· »‰Ã«Õ")
    Case "111": MsgBox ("«·»Ì«‰«  ‰«ﬁ’…")
    Case "112": MsgBox (" ‘Ì— ”Ã·« ‰« «‰ «”„ «·„—”·  „  ›⁄Ì·… „‰ ﬁ»·")
    Case "113": MsgBox ("›‘· ›Ï «· ›⁄Ì·")
    Case "114": MsgBox ("›‘· ›Ï «· ÕœÌÀ Ê „ Õ–› «”„ «·„—”· · ŒÿÏ ⁄œœ „—«  «·”„«Õ")
    Case "115": MsgBox (" ‘Ì— ”Ã·« ‰« «‰… ·« ÌÊÃœ «”„ „—”· „ÿ«»ﬁ ··»ÕÀ")

End Select



End Sub

Private Sub Command1_Click()
    sendMessage
    updateBalance
End Sub

Private Sub sendMessage()
Dim t As String

't = send(UserName, URLEncode(Password), ConvertToUnicode(ConvertString(txtMessage.Text)), txtSender.Text, txtNumbers.Text)
t = send(UserName, URLEncode(Password), ConvertToUnicode(txtMessage.Text), txtSender.Text, txtNumbers.Text)
ShowResult (t)

End Sub

Private Function send(UserName As String, Password As String, msg As String, sender As String, numbers As String) As String
On Error Resume Next
Dim s As String

s = "http://www.4jawaly.net/api/sendsms.php?username=" & UserName & "&password=" & Password & "&message=" & msg & "&numbers=" & numbers & "&sender=" & sender & "&unicode=U"
send = Inet1.OpenURL(s)
End Function
Function GetBalance(UserName As String, Password As String) As String
On Error Resume Next
Dim s As String

s = "http://www.4jawaly.net/api/getbalance.php?username=" & UserName & "&password=" & Password

GetBalance = Inet1.OpenURL(s)
End Function

Private Sub ShowResult(val As String)

Select Case val
 
    Case "100": MsgBox (" „ «·«—”«· »‰Ã«Õ") 'sent
    Case "101": MsgBox ("«·»Ì«‰«  ‰«ﬁ’… ") '
    Case "102": MsgBox ("«”„ «·„” Œœ„ €Ì— ’ÕÌÕ ")
    Case "103": MsgBox ("ﬂ·„Â «·„—Ê— €Ì— ’ÕÌÕÂ")
    Case "104": MsgBox ("Œÿ«¡ »ﬁÊ«⁄œ «·»Ì«‰« ")
    Case "105": MsgBox ("«·—’Ìœ ·« Ìﬂ›Ì")
    Case "106": MsgBox ("«”„ «·„—”· €Ì— „›⁄·")
    Case "107": MsgBox ("«”„ «·„—”· „ÕŸÊ—")
    Case "108": MsgBox ("·„ Ì „ Ê÷⁄ «—ﬁ«„ «Ê «·«—ﬁ«„ €Ì— ’ÕÌÕÂ")
    Case "109": MsgBox ("·« Ì„ﬂ‰ «·«—”«· ·«ﬂÀ— „‰ 8 „ﬁ«ÿ⁄")
    Case "110": MsgBox ("Œÿ«¡ ›Ì «·«—”«·")
    Case "111": MsgBox ("«·«—”«· „€·ﬁ „‰ «œ«—Â «·„Êﬁ⁄ Õ«Ê· „—Â ›Ì Êﬁ  ·«Õﬁ")
    Case "113": MsgBox ("«·Õ”«» «·Œ«’ »ﬂ €Ì— „›⁄·")
    Case "114": MsgBox ("«·Õ”«» „ÊﬁÊ› ... —«Ã⁄ „œÌ— «·Õ”«»")
    Case "114": MsgBox ("«·Õ”«» „ÊﬁÊ› ... —«Ã⁄ „œÌ— «·Õ”«»")
    Case "115": MsgBox ("ÌÃ»  ›⁄Ì· —ﬁ„ «·ÃÊ«· «·Œ«’ »ﬂ »«·„Êﬁ⁄")
    Case "116": MsgBox ("ÌÃ»  ›⁄Ì· «·»—Ìœ «·«·ﬂ —Ê‰Ì «·Œ«’ »ﬂ »«·„Êﬁ⁄")
    Case "117": MsgBox ("·ﬁœ  „  «·⁄„·Ì… »‰Ã«Õ")
    Case "118": MsgBox ("·ﬁœ  „  «·⁄„·Ì… »‰Ã«Õ")
    Case "119": MsgBox ("·ﬁœ  „  «·⁄„·Ì… »‰Ã«Õ")
    
    Case "1010": MsgBox ("Œÿ«¡ ›Ì «· ‘›Ì— ·‰’ «·—”«·Â ")
    Case "1011": MsgBox ("·„ » „ Ê÷⁄ «”„ „” Œœ„ «Ê «”„ «·„” Œœ„ ›«—€")
    Case "1012": MsgBox ("·„ Ì „ Ê÷⁄ ﬂ·„Â „—Ê—")
    Case "1013": MsgBox ("‰’ «·—”«·Â ›«—€")
    Case "1014": MsgBox ("—ﬁ„ «·„” ﬁ»· ›«—€")
    Case "1015": MsgBox ("«”„ «·„—”· ›«—€")

    Case Else: MsgBox (val)
End Select

End Sub

Private Function isArabic(val As String) As Boolean

Dim I As Integer
Dim str As String
str = "œÃÕŒÂ⁄€›ﬁÀ’÷ÿﬂ„‰ «·»Ì”‘Ÿ“Ê…Ï·«—ƒ¡∆≈·≈√·√¬·¬"

For I = 0 To Len(val)
    If InStr(0, str, Mid(val, I, 1), vbTextCompare) <> 0 Then
        isArabic = True
    End If
Next I

isArabic = False
           
End Function

Function ConvertString(s As String) As String
Dim Arr() As String
Dim I As Integer
Arr = Split(s, vbNewLine)
Dim st As String
For I = 0 To UBound(Arr)
st = st & Arr(I) & "'"
Next
ConvertString = st
End Function

Function ConvertToUnicode(st As String) As String
          Dim chrArray(0 To 149) As String
          Dim unicodeArray(0 To 149) As String

10        chrArray(0) = "°"
20        unicodeArray(0) = "060D"
30        chrArray(1) = "∫"
40        unicodeArray(1) = "061B"
50        chrArray(2) = "ø"
60        unicodeArray(2) = "061F"
70        chrArray(3) = "¡"
80        unicodeArray(3) = "0621"
90        chrArray(4) = "¬"
100       unicodeArray(4) = "0622"
110       chrArray(5) = "√"
120       unicodeArray(5) = "0623"
130       chrArray(6) = "ƒ"
140       unicodeArray(6) = "0624"
150       chrArray(7) = "≈"
160       unicodeArray(7) = "0625"
170       chrArray(8) = "∆"
180       unicodeArray(8) = "0626"
190       chrArray(9) = "«"
200       unicodeArray(9) = "0627"
210       chrArray(10) = "»"
220       unicodeArray(10) = "0628"
230       chrArray(11) = "…"
240       unicodeArray(11) = "0629"
250       chrArray(12) = " "
260       unicodeArray(12) = "062A"
270       chrArray(13) = "À"
280       unicodeArray(13) = "062B"
290       chrArray(14) = "Ã"
300       unicodeArray(14) = "062C"
310       chrArray(15) = "Õ"
320       unicodeArray(15) = "062D"
330       chrArray(16) = "Œ"
340       unicodeArray(16) = "062E"
350       chrArray(17) = "œ"
360       unicodeArray(17) = "062F"
370       chrArray(18) = "–"
380       unicodeArray(18) = "0630"
390       chrArray(19) = "—"
400       unicodeArray(19) = "0631"
410       chrArray(20) = "“"
420       unicodeArray(20) = "0632"
430       chrArray(21) = "”"
440       unicodeArray(21) = "0633"
450       chrArray(22) = "‘"
460       unicodeArray(22) = "0634"
470       chrArray(23) = "’"
480       unicodeArray(23) = "0635"
490       chrArray(24) = "÷"
500       unicodeArray(24) = "0636"
510       chrArray(25) = "ÿ"
520       unicodeArray(25) = "0637"
530       chrArray(26) = "Ÿ"
540       unicodeArray(26) = "0638"
550       chrArray(27) = "⁄"
560       unicodeArray(27) = "0639"
570       chrArray(28) = "€"
580       unicodeArray(28) = "063A"
590       chrArray(29) = "›"
600       unicodeArray(29) = "0641"
610       chrArray(30) = "ﬁ"
620       unicodeArray(30) = "0642"
630       chrArray(31) = "ﬂ"
640       unicodeArray(31) = "0643"
650       chrArray(32) = "·"
660       unicodeArray(32) = "0644"
670       chrArray(33) = "„"
680       unicodeArray(33) = "0645"
690       chrArray(34) = "‰"
700       unicodeArray(34) = "0646"
710       chrArray(35) = "Â"
720       unicodeArray(35) = "0647"
730       chrArray(36) = "Ê"
740       unicodeArray(36) = "0648"
750       chrArray(37) = "Ï"
760       unicodeArray(37) = "0649"
770       chrArray(38) = "Ì"
780       unicodeArray(38) = "064A"
790       chrArray(39) = "‹"
800       unicodeArray(39) = "0640"
810       chrArray(40) = ""
820       unicodeArray(40) = "064B"
830       chrArray(41) = "Ò"
840       unicodeArray(41) = "064C"
850       chrArray(42) = "Ú"
860       unicodeArray(42) = "064D"
870       chrArray(43) = "Û"
880       unicodeArray(43) = "064E"
890       chrArray(44) = "ı"
900       unicodeArray(44) = "064F"
910       chrArray(45) = "ˆ"
920       unicodeArray(45) = "0650"
930       chrArray(46) = "¯"
940       unicodeArray(46) = "0651"
950       chrArray(47) = "˙"
960       unicodeArray(47) = "0652"
970       chrArray(48) = "!"
980       unicodeArray(48) = "0021"
990       chrArray(49) = """"
1000      unicodeArray(49) = "0022"
1010      chrArray(50) = "#"
1020      unicodeArray(50) = "0023"
1030      chrArray(51) = "$"
1040      unicodeArray(51) = "0024"
1050      chrArray(52) = "%"
1060      unicodeArray(52) = "0025"
1070      chrArray(53) = "&"
1080      unicodeArray(53) = "0026"
1090      chrArray(54) = "'"
1100      unicodeArray(54) = "0027"
1110      chrArray(55) = "("
1120      unicodeArray(55) = "0028"
1130      chrArray(56) = ")"
1140      unicodeArray(56) = "0029"
1150      chrArray(57) = "*"
1160      unicodeArray(57) = "002A"
1170      chrArray(58) = "+"
1180      unicodeArray(58) = "002B"
1190      chrArray(59) = ","
1200      unicodeArray(59) = "002C"
1210      chrArray(60) = "-"
1220      unicodeArray(60) = "002D"
1230      chrArray(61) = "."
1240      unicodeArray(61) = "002E"
1250      chrArray(62) = "/"
1260      unicodeArray(62) = "002F"
1270      chrArray(63) = "0"
1280      unicodeArray(63) = "0030"
1290      chrArray(64) = "1"
1300      unicodeArray(64) = "0031"
1310      chrArray(65) = "2"
1320      unicodeArray(65) = "0032"
1330      chrArray(66) = "3"
1340      unicodeArray(66) = "0033"
1350      chrArray(67) = "4"
1360      unicodeArray(67) = "0034"
1370      chrArray(68) = "5"
1380      unicodeArray(68) = "0035"
1390      chrArray(69) = "6"
1400      unicodeArray(69) = "0036"
1410      chrArray(70) = "7"
1420      unicodeArray(70) = "0037"
1430      chrArray(71) = "8"
1440      unicodeArray(71) = "0038"
1450      chrArray(72) = "9"
1460      unicodeArray(72) = "0039"
1470      chrArray(73) = ":"
1480      unicodeArray(73) = "003A"
1490      chrArray(74) = ""
1500      unicodeArray(74) = "003B"
1510      chrArray(75) = "<"
1520      unicodeArray(75) = "003C"
1530      chrArray(76) = "="
1540      unicodeArray(76) = "003D"
1550      chrArray(77) = ">"
1560      unicodeArray(77) = "003E"
1570      chrArray(78) = "?"
1580      unicodeArray(78) = "003F"
1590      chrArray(79) = "@"
1600      unicodeArray(79) = "0040"
1610      chrArray(80) = "A"
1620      unicodeArray(80) = "0041"
1630      chrArray(81) = "B"
1640      unicodeArray(81) = "0042"
1650      chrArray(82) = "C"
1660      unicodeArray(82) = "0043"
1670      chrArray(83) = "D"
1680      unicodeArray(83) = "0044"
1690      chrArray(84) = "E"
1700      unicodeArray(84) = "0045"
1710      chrArray(85) = "F"
1720      unicodeArray(85) = "0046"
1730      chrArray(86) = "G"
1740      unicodeArray(86) = "0047"
1750      chrArray(87) = "H"
1760      unicodeArray(87) = "0048"
1770      chrArray(88) = "I"
1780      unicodeArray(88) = "0049"
1790      chrArray(89) = "J"
1800      unicodeArray(89) = "004A"
1810      chrArray(90) = "K"
1820      unicodeArray(90) = "004B"
1830      chrArray(91) = "L"
1840      unicodeArray(91) = "004C"
1850      chrArray(92) = "M"
1860      unicodeArray(92) = "004D"
1870      chrArray(93) = "N"
1880      unicodeArray(93) = "004E"
1890      chrArray(94) = "O"
1900      unicodeArray(94) = "004F"
1910      chrArray(95) = "P"
1920      unicodeArray(95) = "0050"
1930      chrArray(96) = "Q"
1940      unicodeArray(96) = "0051"
1950      chrArray(97) = "R"
1960      unicodeArray(97) = "0052"
1970      chrArray(98) = "S"
1980      unicodeArray(98) = "0053"
1990      chrArray(99) = "T"
2000      unicodeArray(99) = "0054"
2010      chrArray(100) = "U"
2020      unicodeArray(100) = "0055"
2030      chrArray(101) = "V"
2040      unicodeArray(101) = "0056"
2050      chrArray(102) = "W"
2060      unicodeArray(102) = "0057"
2070      chrArray(103) = "X"
2080      unicodeArray(103) = "0058"
2090      chrArray(104) = "Y"
2100      unicodeArray(104) = "0059"
2110      chrArray(105) = "Z"
2120      unicodeArray(105) = "005A"
2130      chrArray(106) = "[" '"("
2140      unicodeArray(106) = "005B"
2150      chrArray(107) = Trim("\ ")
2160      unicodeArray(107) = "005C"
2170      chrArray(108) = "]" '")"
2180      unicodeArray(108) = "005D"
2190      chrArray(109) = "^"
2200      unicodeArray(109) = "005E"
2210      chrArray(110) = "_"
2220      unicodeArray(110) = "005F"
2230      chrArray(111) = "`"
2240      unicodeArray(111) = "0060"
2250      chrArray(112) = "a"
2260      unicodeArray(112) = "0061"
2270      chrArray(113) = "b"
2280      unicodeArray(113) = "0062"
2290      chrArray(114) = "c"
2300      unicodeArray(114) = "0063"
2310      chrArray(115) = "d"
2320      unicodeArray(115) = "0064"
2330      chrArray(116) = "e"
2340      unicodeArray(116) = "0065"
2350      chrArray(117) = "f"
2360      unicodeArray(117) = "0066"
2370      chrArray(118) = "g"
2380      unicodeArray(118) = "0067"
2390      chrArray(119) = "h"
2400      unicodeArray(119) = "0068"
2410      chrArray(120) = "i"
2420      unicodeArray(120) = "0069"
2430      chrArray(121) = "j"
2440      unicodeArray(121) = "006A"
2450      chrArray(122) = "k"
2460      unicodeArray(122) = "006B"
2470      chrArray(123) = "l"
2480      unicodeArray(123) = "006C"
2490      chrArray(124) = "m"
2500      unicodeArray(124) = "006D"
2510      chrArray(125) = "n"
2520      unicodeArray(125) = "006E"
2530      chrArray(126) = "o"
2540      unicodeArray(126) = "006F"
2550      chrArray(127) = "p"
2560      unicodeArray(127) = "0070"
2570      chrArray(128) = "q"
2580      unicodeArray(128) = "0071"
2590      chrArray(129) = "r"
2600      unicodeArray(129) = "0072"
2610      chrArray(130) = "s"
2620      unicodeArray(130) = "0073"
2630      chrArray(131) = "t"
2640      unicodeArray(131) = "0074"
2650      chrArray(132) = "u"
2660      unicodeArray(132) = "0075"
2670      chrArray(133) = "v"
2680      unicodeArray(133) = "0076"
2690      chrArray(134) = "w"
2700      unicodeArray(134) = "0077"
2710      chrArray(135) = "x"
2720      unicodeArray(135) = "0078"
2730      chrArray(136) = "y"
2740      unicodeArray(136) = "0079"
2750      chrArray(137) = "z"
2760      unicodeArray(137) = "007A"
2770      chrArray(138) = "{"
2780      unicodeArray(138) = "007B"
2790      chrArray(139) = "|"
2800      unicodeArray(139) = "007C"
2810      chrArray(140) = "}"
2820      unicodeArray(140) = "007D"
2830      chrArray(141) = "~"
2840      unicodeArray(141) = "007E"
2850      chrArray(142) = "©"
2860      unicodeArray(142) = "00A9"
2870      chrArray(143) = "Æ"
2880      unicodeArray(143) = "00AE"
2890      chrArray(144) = "˜"
2900      unicodeArray(144) = "00F7"
2910      chrArray(145) = "◊"
2920      unicodeArray(145) = "00F7"
2930      chrArray(146) = "ß"
2940      unicodeArray(146) = "00A7"
2950      chrArray(147) = " "
2960      unicodeArray(147) = "0020"
2970      chrArray(148) = "\n"
2980      unicodeArray(148) = "000D"
2990      chrArray(149) = "\r"
3000      unicodeArray(149) = "000A"

          
          
          Dim strResult As String, I As Integer, C As Integer
3010          strResult = ""
3020          For I = 1 To Len(st)
3030              For C = 0 To 149
3040                  If (chrArray(C) = Mid(st, I, 1)) Then
3050                      strResult = strResult & unicodeArray(C)
3060                  End If
3070              Next C
3080          Next I
3090          ConvertToUnicode = strResult

End Function
Function ToUnicode(ch As String) As String
Dim I As Integer

If ch = "'" Then
ToUnicode = "000D"
Exit Function
End If

For I = 0 To List1.ListCount - 1

If ch = List1.List(I) Then
ToUnicode = List2.List(I)
Exit Function
End If

Next
End Function
Private Sub Form_Load()
Dim I As Integer

List1.AddItem "'"
List1.AddItem " "

For I = 0 To List2.ListCount - 1
List2.List(I) = Fourdigit(List2.List(I))
Next

List2.AddItem "000D"
List2.AddItem "0020"

updateBalance
End Sub
Private Sub updateBalance()
Dim b As String
b = GetBalance(UserName, Password)

Select Case b

 Case "100": MsgBox ("«·»Ì«‰«  ‰«ﬁ’Â")
 Case "102": MsgBox ("«”„ «·„” Œœ„ €Ì— ’ÕÌÕ")
 Case "103": MsgBox ("ﬂ·„Â «·„—Ê— €Ì— ’ÕÌÕÂ…")
 Case "104": MsgBox ("·« ÌÊÃœ —’Ìœ »«·Õ”«»")
 Case "111": MsgBox ("«·„Êﬁ⁄ „€·ﬁ ")
 Case "113": MsgBox ("«·Õ”«» „ÊﬁÊ› —«Ã⁄ „œÌ— Õ”«»ﬂ")
 Case "114": MsgBox ("«·ÃÊ«· €Ì— „›⁄· »«·„Êﬁ⁄")
 Case "116": MsgBox ("«·»—Ìœ «·«·ﬂ —Ê‰Ì €Ì— „›⁄· »«·„Êﬁ⁄")
        
        Exit Sub
End Select

lblBalance.Caption = "Your Balance is: " & b
End Sub
Function Fourdigit(ch As String) As String

If Len(ch) = 1 Then
    Fourdigit = "000" & ch
    Exit Function
End If

If Len(ch) = 2 Then
    Fourdigit = "00" & ch
    Exit Function
End If

If Len(ch) = 3 Then
    Fourdigit = "0" & ch
    Exit Function
End If

If Len(ch) = 4 Then
    Fourdigit = ch
    Exit Function
End If

End Function


Function URLEncode(ByVal str As String) As String
    Dim intLen As Integer
    Dim x As Integer
    Dim curChar As Long
    Dim newStr As String

    intLen = Len(str)
    newStr = ""
    For x = 1 To intLen
        curChar = Asc(Mid$(str, x, 1))
          
        If (curChar < 48 Or curChar > 57) And _
              (curChar < 65 Or curChar > 90) And _
              (curChar < 97 Or curChar > 122) Then
                            newStr = newStr & "%" & Hex(curChar)
        Else
            newStr = newStr & Chr(curChar)
        End If
    Next x
              
    URLEncode = newStr
End Function


