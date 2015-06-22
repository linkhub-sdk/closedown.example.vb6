VERSION 5.00
Begin VB.Form closedownExam 
   Caption         =   "휴폐업조회 API SDK Example for VB6"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   8130
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton btnUnitCost 
      Caption         =   "조회단가 확인"
      Height          =   400
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "과금 API"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   6855
      Begin VB.CommandButton btnGetBalance 
         Caption         =   "잔여포인트 조회"
         Height          =   400
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "휴폐업조회 API"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   6975
      Begin VB.CommandButton btnCheckCorpNums 
         Caption         =   "대량조회"
         Height          =   375
         Left            =   5400
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton btnCheckCorpNum 
         Caption         =   "조회"
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtCorpNum 
         Height          =   270
         Left            =   1320
         TabIndex        =   2
         Text            =   "4108600477"
         Top             =   440
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "사업자번호 : "
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "closedownExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'링크아이디
Private Const linkID = "TESTER"
'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private ClosedownChecker As New Closedown
'휴폐업조회 - 단건
Private Sub btnCheckCorpNum_Click()
    Dim Corpstate As Corpstate
    
    Set Corpstate = ClosedownChecker.CheckCorpNum(txtCorpNum.Text)
    
    If Corpstate Is Nothing Then
        MsgBox ("[" + CStr(ClosedownChecker.LastErrCode) + "]" + ClosedownChecker.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "corpNum : " + Corpstate.corpNum + vbCrLf
    tmp = tmp + "state : " + Corpstate.state + vbCrLf
    tmp = tmp + "type : " + Corpstate.ctype + vbCrLf
    tmp = tmp + "stateDate(휴폐업일자) : " + Corpstate.stateDate + vbCrLf
    tmp = tmp + "checkDate(국세청 확인일자) : " + Corpstate.checkDate + vbCrLf + vbCrLf
    
    tmp = tmp + "* state (휴폐업상태) : null-알수없음, 0-등록되지 않은 사업자번호, 1-사업중, 2-폐업, 3-휴업" + vbCrLf
    tmp = tmp + "* type (사업 유형) : null-알수없음, 1-일반과세자, 2-면세과세자, 3-간이과세자, 4-비영리법인, 국가기관"
        
    MsgBox tmp
End Sub
'휴폐업조회 - 대량(최대 1000건)
Private Sub btnCheckCorpNums_Click()
    Dim resultList As Collection
    
    Dim CorpNumList As New Collection
    
    '사업자번호 배열, 최대 1000건
    CorpNumList.Add "1234567890"
    CorpNumList.Add "4108600477"
    CorpNumList.Add "4352343543"
    
    Set resultList = ClosedownChecker.CheckCorpNums(CorpNumList)
    
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(ClosedownChecker.LastErrCode) + "] " + ClosedownChecker.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    Dim stateInfo As Corpstate
    
    tmp = "* state (휴폐업상태) : null-알수없음, 0-등록되지 않은 사업자번호, 1-사업중, 2-폐업, 3-휴업" + vbCrLf
    tmp = tmp + "* type (사업 유형) : null-알수없음, 1-일반과세자, 2-면세과세자, 3-간이과세자, 4-비영리법인, 국가기관" + vbCrLf + vbCrLf
    
    For Each stateInfo In resultList
        tmp = tmp + "corpNum : " + stateInfo.corpNum + vbCrLf
        tmp = tmp + "state : " + stateInfo.state + vbCrLf
        tmp = tmp + "type : " + stateInfo.ctype + vbCrLf
        tmp = tmp + "stateDate(휴폐업일자) : " + stateInfo.stateDate + vbCrLf
        tmp = tmp + "checkDate(국세청 확인일자) : " + stateInfo.checkDate + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp
End Sub
'잔여포인트 조회
Private Sub btnGetBalance_Click(index As Integer)
    Dim balance As Double
    
    balance = ClosedownChecker.GetBalance()
    
    If balance < 0 Then
        MsgBox ("[" + CStr(ClosedownChecker.LastErrCode) + "] " + ClosedownChecker.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
End Sub
'조회단가 확인
Private Sub btnUnitCost_Click(index As Integer)
    Dim unitCost As Double
    
    unitCost = ClosedownChecker.GetUnitCost()
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(ClosedownChecker.LastErrCode) + "] " + ClosedownChecker.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "조회단가 : " + CStr(unitCost)
End Sub

Private Sub txtCorpNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call btnCheckCorpNum_Click
    End If
End Sub

Private Sub Form_Load()
    '휴폐업조회 서비스모듈 초기화
    ClosedownChecker.Initialize linkID, SecretKey
    
End Sub
