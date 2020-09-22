VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMainWin 
   Caption         =   "Download Internet Content"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboURL 
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Text            =   "http://www.abacusinternet.com/images/keyboard.jpg"
      Top             =   1320
      Width           =   5895
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   7800
      TabIndex        =   11
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdGetURL 
      Caption         =   "Get Page"
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   2400
      Width           =   855
   End
   Begin VB.Frame frmProxy 
      Caption         =   "Proxy Settings"
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   3375
      Begin VB.OptionButton optProxy 
         Caption         =   "Use Proxy"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optNoProxy 
         Caption         =   "No Proxy"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox txtProxy 
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Text            =   "proxy.myproxy.com"
      Top             =   1800
      Width           =   5775
   End
   Begin MSWinsockLib.Winsock WS_HTTP 
      Left            =   9840
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtPort 
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Text            =   "80"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtStatus 
      Height          =   975
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   9855
   End
   Begin VB.TextBox txtResponse 
      Height          =   4935
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3000
      Width           =   10095
   End
   Begin VB.Label lblProxy 
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblURL 
      Caption         =   "URL:"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   255
      Left            =   7560
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
End
Attribute VB_Name = "frmMainWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ---------------------------------------------------------
' WEB DOWNLOAD USING WINSOCK CONTROL
' Download web pages (http://www.whatever.com/default.htm)
' and graphics (binaries) using only the Winsock Control.
' Binaries are saved to hard drive in your application path.
' Works with proxy servers too!
'
' Please send comments to:
'       Kevin McGahey
'       kmcgahey@abacusinternet.com
' ---------------------------------------------------------

Private mstrURL As String
Private mstrResponseDocument As String
Private mblnIsProxyUsed As Boolean
Private mblnIsPicture As Boolean
Private mblnIsHeader As Boolean
Private mstrReturnHeader As String
Private mstrRequestHeader As String
Private mstrLocalFile As String


Private Sub cmdClear_Click()
    txtResponse.Text = ""
    txtStatus.Text = ""
End Sub

Private Sub cmdClose_Click()
    WS_HTTP.Close
    Unload Me
End Sub

Private Sub cmdGetURL_Click()
    Dim strPureURL As String
    Dim strServerAddress As String
    Dim strServerHostIP As String
    Dim strDocumentURI As String
    Dim lngStartPos As Long
    Dim lngServerPort As Long
    
    Dim strRequestTemplate As String
     
     mstrURL = cboURL.Text
             
    If (optProxy.Value = True) Then
        mblnIsProxyUsed = True
    End If

    
    If UCase(Left(mstrURL, 7)) <> "HTTP://" Then
        MsgBox "Please enter url With http://", vbCritical + vbOK
        Exit Sub
    End If
    
    ' Note: This section of code (header) is based on code posted
    ' by Tair Abdurman on http://www.planetsourcecode.com
    ' - Thanks for the proxy help Tair
    mstrRequestHeader = ""
    strRequestTemplate = "GET _$-$_$- HTTP/1.0" & Chr(13) & Chr(10) & _
    "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-comet, */*" & Chr(13) & Chr(10) & _
    "Accept-Language: en" & Chr(13) & Chr(10) & _
    "Accept-Encoding: gzip , deflate" & Chr(13) & Chr(10) & _
    "Cache-Control: no-cache" & Chr(13) & Chr(10) & _
    "Proxy-Connection: Keep-Alive" & Chr(13) & Chr(10) & _
    "User-Agent: SSM Agent 1.0" & Chr(13) & Chr(10) & _
    "Host: @$@@$@" & Chr(13) & Chr(10)
    
    ' Remove "http://"
    strPureURL = Right(mstrURL, Len(mstrURL) - 7)
    lngStartPos = InStr(1, strPureURL, "/")
    
    If lngStartPos < 1 Then
        strServerAddress = strPureURL
        strDocumentURI = "/"
    Else
        strServerAddress = Left(strPureURL, lngStartPos - 1)
        strDocumentURI = Right(strPureURL, Len(strPureURL) - lngStartPos + 1)
        mstrLocalFile = App.Path & "\" & Right(strPureURL, Len(strPureURL) - InStrRev(strPureURL, "/"))
    End If
            
    If strServerAddress = "" Or strDocumentURI = "" Then
        MsgBox "Unable To detect target page!", vbCritical + vbOK
        Exit Sub
    End If
            
    If mblnIsProxyUsed Then
        strServerHostIP = txtProxy.Text
        mstrRequestHeader = strRequestTemplate
        mstrRequestHeader = Replace(mstrRequestHeader, "_$-$_$-", mstrURL)
        lngServerPort = 80
    Else
        strServerHostIP = strServerAddress
        lngServerPort = 80
        mstrRequestHeader = strRequestTemplate
        mstrRequestHeader = Replace(mstrRequestHeader, "_$-$_$-", strDocumentURI)
    End If
            
    mstrRequestHeader = Replace(mstrRequestHeader, "@$@@$@", strServerAddress)
    mstrRequestHeader = mstrRequestHeader & Chr(13) & Chr(10)
    txtStatus.Text = "Connecting To server ..." & vbCrLf
    txtStatus.Refresh
    
    ' Are we retreiving a picture
    If (UCase(Right(mstrURL, 3)) = "GIF" Or _
        UCase(Right(mstrURL, 3)) = "JPG") Then
        mblnIsPicture = True
        If (FileExists(mstrLocalFile)) Then Kill mstrLocalFile
        
        ' Open mstrLocalFile For Binary As #1
        Open mstrLocalFile For Binary Access Write As #1
        mblnIsHeader = True
    Else
        mblnIsHeader = False
        mblnIsPicture = False
    End If
           
    WS_HTTP.Connect strServerHostIP, lngServerPort
End Sub

Private Sub Form_Load()
    cboURL.AddItem "http://www.abacusinternet.com/images/keyboard.jpg"
    cboURL.AddItem "http://www.microsoft.com/default.htm"
    cboURL.AddItem "http://www.yahoo.com"
End Sub

Private Sub WS_HTTP_Close()
    WS_HTTP.Close
    txtStatus.Text = txtStatus.Text & "Transaction completed ..." & vbCrLf
    txtStatus.Refresh
    If (mblnIsPicture) Then
        Close #1
        MsgBox "File saved to:" & vbCrLf & mstrLocalFile, vbInformation, "Download Complete!"
    Else
        MsgBox "Web Page Download Complete.", vbInformation, "Web Page Retrieved!"
    End If
End Sub


Private Sub WS_HTTP_Connect()
    WS_HTTP.SendData mstrRequestHeader
    txtStatus.Text = txtStatus.Text & "Connected, try To obtain page ..." & vbCrLf
    txtStatus.Refresh
    frmMainWin.txtResponse.Text = ""
    frmMainWin.txtResponse.Refresh
End Sub


Private Sub WS_HTTP_DataArrival(ByVal bytesTotal As Long)
    Dim strTemp As String
    Dim lngBytes As Long
    Dim blnFoundHeadEndByte As Boolean
    Dim b() As Byte
    Dim b2() As Byte
    Dim aryMyArray As Variant
    Dim i As Long
    Dim j As Long
    Dim strChr As String
    
    ' -------------------------------------
    ' Check Header
    ' Some Servers will send the header then
    ' then the graphic, others will send the
    ' header along with a chunk of the
    ' graphic in the first pass
    ' --------------------------------------
    
    If (mblnIsPicture) Then
        WS_HTTP.PeekData strTemp, vbString
        If (mblnIsHeader) Then
            WS_HTTP.GetData b(), vbByte, 350
            mstrReturnHeader = StrConv(b(), vbUnicode)
            
            ' Loop through the first chunk of
            ' data and remove the header
            ' p.s. this part of the code required
            ' me to use one of my 2 free calls to Microsoft
            ' Developers Support so I hope you appreciate it!
            For i = LBound(b) + 3 To UBound(b)
                If (blnFoundHeadEndByte) Then
                    b2(j) = b()(i)
                    j = j + 1
                Else
                    If b()(i - 3) = 13 And b()(i - 2) = 10 And _
                       b()(i - 1) = 13 And b()(i) = 10 Then
                        ReDim b2(UBound(b) - i)
                        blnFoundHeadEndByte = True
                        j = 0
                    End If
                End If
                    
            Next i
            
            If (UBound(b2) > 0) Then Put #1, , b2()
            mblnIsHeader = False
        Else
            WS_HTTP.GetData b(), vbByte
            Put #1, , b()
        End If
    Else
        WS_HTTP.GetData strTemp, vbString
    End If
      
    txtResponse.Text = txtResponse.Text & strTemp
    txtStatus.Text = txtStatus.Text & "Data from server, continue ..." & vbCrLf

End Sub


Private Sub WS_HTTP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WS_HTTP.Close
    txtStatus.Text = txtStatus.Text & "Errors occured ..." & vbCrLf
    txtStatus.Text = txtStatus.Text & "Number: " & Number & "  Description: " & Description & vbCrLf

    txtStatus.Refresh
    'Me.Hide
    'Unload Me
End Sub


'-----------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
'
Public Function FileExists(ByVal strPathName As String) As Boolean
    Dim intFileNum As Integer

    On Error Resume Next

    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = (Err.Number = 0)

    Close intFileNum
    Err.Clear
End Function

