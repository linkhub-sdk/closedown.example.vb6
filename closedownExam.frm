VERSION 5.00
Begin VB.Form closedownExam 
   Caption         =   "�������ȸ API SDK Example for VB6"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   8130
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton btnUnitCost 
      Caption         =   "��ȸ�ܰ� Ȯ��"
      Height          =   400
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "���� API"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "�ܿ�����Ʈ ��ȸ"
         Height          =   400
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�������ȸ API"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "�뷮��ȸ"
         Height          =   375
         Left            =   5400
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton btnCheckCorpNum 
         Caption         =   "��ȸ"
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
         Caption         =   "����ڹ�ȣ : "
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

'��ũ���̵�
Private Const linkID = "TESTER"
'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private ClosedownChecker As New Closedown
'�������ȸ - �ܰ�
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
    tmp = tmp + "stateDate(���������) : " + Corpstate.stateDate + vbCrLf
    tmp = tmp + "checkDate(����û Ȯ������) : " + Corpstate.checkDate + vbCrLf + vbCrLf
    
    tmp = tmp + "* state (���������) : null-�˼�����, 0-��ϵ��� ���� ����ڹ�ȣ, 1-�����, 2-���, 3-�޾�" + vbCrLf
    tmp = tmp + "* type (��� ����) : null-�˼�����, 1-�Ϲݰ�����, 2-�鼼������, 3-���̰�����, 4-�񿵸�����, �������"
        
    MsgBox tmp
End Sub
'�������ȸ - �뷮(�ִ� 1000��)
Private Sub btnCheckCorpNums_Click()
    Dim resultList As Collection
    
    Dim CorpNumList As New Collection
    
    '����ڹ�ȣ �迭, �ִ� 1000��
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
    
    tmp = "* state (���������) : null-�˼�����, 0-��ϵ��� ���� ����ڹ�ȣ, 1-�����, 2-���, 3-�޾�" + vbCrLf
    tmp = tmp + "* type (��� ����) : null-�˼�����, 1-�Ϲݰ�����, 2-�鼼������, 3-���̰�����, 4-�񿵸�����, �������" + vbCrLf + vbCrLf
    
    For Each stateInfo In resultList
        tmp = tmp + "corpNum : " + stateInfo.corpNum + vbCrLf
        tmp = tmp + "state : " + stateInfo.state + vbCrLf
        tmp = tmp + "type : " + stateInfo.ctype + vbCrLf
        tmp = tmp + "stateDate(���������) : " + stateInfo.stateDate + vbCrLf
        tmp = tmp + "checkDate(����û Ȯ������) : " + stateInfo.checkDate + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp
End Sub
'�ܿ�����Ʈ ��ȸ
Private Sub btnGetBalance_Click(index As Integer)
    Dim balance As Double
    
    balance = ClosedownChecker.GetBalance()
    
    If balance < 0 Then
        MsgBox ("[" + CStr(ClosedownChecker.LastErrCode) + "] " + ClosedownChecker.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub
'��ȸ�ܰ� Ȯ��
Private Sub btnUnitCost_Click(index As Integer)
    Dim unitCost As Double
    
    unitCost = ClosedownChecker.GetUnitCost()
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(ClosedownChecker.LastErrCode) + "] " + ClosedownChecker.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��ȸ�ܰ� : " + CStr(unitCost)
End Sub

Private Sub txtCorpNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call btnCheckCorpNum_Click
    End If
End Sub

Private Sub Form_Load()
    '�������ȸ ���񽺸�� �ʱ�ȭ
    ClosedownChecker.Initialize linkID, SecretKey
    
End Sub
