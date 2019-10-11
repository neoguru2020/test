VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CBRate"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   4080
      Top             =   120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strRateCB As String
Dim dtDate As Date


Private Sub Timer1_Timer()

    If dtDate <> Date Then
        If fnGetALLRateCBDBSQL = True Then
            dtDate = Date
            strRateCB = ""
            Text1.text = Text1.text & "За " & dtDate & " курсы обновлены." & vbCrLf
        Else
        
        End If
    End If

End Sub


Function fnSetRateCBDBSQL(ByVal CharCode As String, ByVal dat As Date) As Boolean
        
        On Error GoTo fnSetRateCBDBSQL_Err
        
        Dim strRate As String
100     strRate = fnFind(strRateCB, "<VALUE>", "</VALUE>", InStr(1, strRateCB, "<CHARCODE>" & CharCode & "</CHARCODE>"))
102     If strRate = "ERROR" Then
104         fnSetRateCBDBSQL = False
            Exit Function
        End If
        
106     If fnRateCBDBSQL(CharCode, dat) = strRate Then
108         fnSetRateCBDBSQL = True
            Exit Function
        End If
        
        Dim strSQL As String
110     If fnRateCBDBSQL(CharCode, dat) = 0 Then
112         strSQL = "INSERT INTO [dbo].[Курс]([дата],[валюта],[курс]) VALUES ('" & dat & "','" & UCase(CharCode) & "'," & Replace$(strRate, ",", ".") & ")"
        Else
114         strSQL = "UPDATE [dbo].[Курс] SET [курс] = " & Replace$(strRate, ",", ".") & " WHERE [дата] = '" & dat & "' AND [валюта] = '" & UCase(CharCode) & "'"
        End If
    
        Dim strConnectionString As String
116     strConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Auto Translate=True;Persist Security Info=False;Initial Catalog=test;Data Source=UNIT1\SQLEXPRESS"
        Dim Con As New ADODB.Connection
118     Con.ConnectionString = strConnectionString
120     Con.Open
122     Con.Execute strSQL, , adExecuteNoRecords
124     Con.Close
126     Set Con = Nothing

128     fnSetRateCBDBSQL = True
        Exit Function

fnSetRateCBDBSQL_Err:
130     fnSetRateCBDBSQL = False

End Function


Function fnGetALLRateCBDBSQL() As Boolean
        
        On Error GoTo fnGetALLRateCBDBSQL_Err

100     If fnGetRateCB(Date) = False Then
102         fnGetALLRateCBDBSQL = False
            Exit Function
        End If
        
        Dim strConnectionString As String
        Dim strSQL As String
104     strConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Auto Translate=True;Persist Security Info=False;Initial Catalog=test;Data Source=UNIT1\SQLEXPRESS"
106     strSQL = "SELECT [CharCode] FROM [dbo].[Валюты]"
        Dim Con As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
108     Con.ConnectionString = strConnectionString
110     Con.Open
112     With cmd
114         .ActiveConnection = Con
116         .CommandText = strSQL
118         .CommandType = adCmdText
        End With
120     With rs
122         .CursorType = adOpenStatic
124         .CursorLocation = adUseClient
126         .LockType = adLockOptimistic
128         .Open cmd
        End With
130     If Not (rs.BOF And rs.EOF) Then
132         Do While Not rs.EOF
134             If fnSetRateCBDBSQL(rs![CharCode], Date) = False Then
                    fnGetALLRateCBDBSQL = False
                    Exit Function
                End If
136         rs.MoveNext
            Loop
138         fnGetALLRateCBDBSQL = True
        Else
140         fnGetALLRateCBDBSQL = False
        End If
142     rs.Close
144     Con.Close
146     Set Con = Nothing
148     Set cmd = Nothing
150     Set rs = Nothing

        Exit Function

fnGetALLRateCBDBSQL_Err:
        fnGetALLRateCBDBSQL = False
        
End Function


Function fnGetRateCB(ByVal dat As Date) As Boolean

        On Error GoTo fnGetRateCB_Err

        Dim request$
        Dim response$
        Dim objHTTP As Object
100     request = "http://www.cbr.ru/scripts/XML_daily.asp?date_req=" & Format(dat, "DD.MM.YYYY")
102     Set objHTTP = CreateObject("MSXML2.XMLHTTP")
104     objHTTP.Open "GET", request, False
106     objHTTP.Send
108     response = UCase(objHTTP.ResponseText)
110     strRateCB = response
        fnGetRateCB = True
        Exit Function

fnGetRateCB_Err:
        fnGetRateCB = False
        
End Function


Function fnFind(ByVal text As String, ByVal String1 As String, ByVal String2 As String, Optional ByVal Start As Integer = 1) As String
        
        On Error GoTo fnFind_Err

        Dim nid1 As Integer, nid2 As Integer
100     nid1 = InStr(Start, text, String1) + Len(String1)
102     nid2 = InStr(nid1, text, String2)
104     If InStr(Start, text, String1) = 0 Or InStr(nid1, text, String2) = 0 Then
106         fnFind = "ERROR"
        End If
108     fnFind = Mid(text, nid1, nid2 - nid1)

        Exit Function

fnFind_Err:
        fnFind = "ERROR"

End Function
    
    
' Sample: MsgBox fnRateCBDBSQL("usd", Date), , "RateCBDBSQL"
' Sample: MsgBox fnRateCBDBSQL("eur", Date), , "RateCBDBSQL"

Function fnRateCBDBSQL(ByVal CharCode As String, ByVal dat As Date) As Currency
        
        On Error GoTo fnRateCBDBSQL_Err
    
        Dim strConnectionString As String
        Dim strSQL As String
100     strConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Auto Translate=True;Persist Security Info=False;Initial Catalog=test;Data Source=UNIT1\SQLEXPRESS"
102     strSQL = "SELECT [курс] FROM [dbo].[Курс] WHERE [валюта] = '" & UCase(CharCode) & "' AND [дата] = '" & dat & "'"
        Dim Con As New ADODB.Connection
        Dim cmd As New ADODB.Command
        Dim rs As New ADODB.Recordset
104     Con.ConnectionString = strConnectionString
106     Con.Open
108     With cmd
110         .ActiveConnection = Con
112         .CommandText = strSQL
114         .CommandType = adCmdText
        End With
116     With rs
118         .CursorType = adOpenStatic
120         .CursorLocation = adUseClient
122         .LockType = adLockOptimistic
124         .Open cmd
        End With
126     If Not (rs.BOF And rs.EOF) Then
128         fnRateCBDBSQL = CCur(rs![курс])
        Else
130         fnRateCBDBSQL = 0
        End If
132     rs.Close
134     Con.Close
136     Set Con = Nothing
138     Set cmd = Nothing
140     Set rs = Nothing

        Exit Function

fnRateCBDBSQL_Err:
        fnRateCBDBSQL = 0
        
End Function

