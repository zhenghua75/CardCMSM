VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "面包工坊卡读写测试"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6945
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "消费"
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "积分兑换"
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "充值"
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "发卡"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "读卡"
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "错误"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   3840
      Width           =   5055
   End
   Begin VB.Label Label3 
      Caption         =   "积分："
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "余额："
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "卡号："
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() '读卡
    Set m1 = New M1Card
    Dim strCardNo As String
    Dim dCharge As Double
    Dim dIg As Long
    Dim ret As String
    
    ret = m1.ReadCard(strCardNo, dCharge, dIg)
    Label4.Caption = ret
End Sub

Private Sub Command2_Click() '发卡
    Set m1 = New M1Card
    Dim strCardNo As String
    Dim dCharge As Double
    Dim dIg As Long
    Dim ret As String
    If Len(Text1.Text) = 0 Or Len(Text2.Text) = 0 Or Len(Text3.Text) = 0 Then
       MsgBox "卡号、余额、积分不可为空"
       Exit Sub
    End If
    strCardNo = Text1.Text
    dCharge = CDbl(Text2.Text)
    dIg = CLng(Text3.Text)
    ret = m1.PutCard(strCardNo, dCharge, dIg)
    Label4.Caption = ret
End Sub

Private Sub Command3_Click() '充值
    Set m1 = New M1Card
    Dim strCardNo As String
    Dim dCharge As Double
    Dim dChargeBak As Double

    Dim ret As String
    
    If Len(Text2.Text) = 0 Then
       MsgBox "余额不可为空"
       Exit Sub
    End If
    strCardNo = Text1.Text
    dCharge = CDbl(Text2.Text)
    dCharge = dChargeBak
    
    dIg = CLng(Text3.Text)
    ret = m1.WriteCharge(dCharge, dChargeBak)
    Label4.Caption = ret
End Sub

Private Sub Command4_Click()
    Set m1 = New M1Card
    Dim dIg As Long
    Dim ret As String
    If Len(Text3.Text) = 0 Then
       MsgBox "积分不可为空"
       Exit Sub
    End If
    dIg = CLng(Text3.Text)
    ret = m1.WriteIg(dIg)
    Label4.Caption = ret
End Sub

Private Sub Command5_Click()
    Set m1 = New M1Card
    Dim dCharge As Double
    Dim dChargeBak As Double
    Dim dIg As Long
    Dim ret As String
    If Len(Text2.Text) = 0 Or Len(Text3.Text) = 0 Then
       MsgBox "余额、积分不可为空"
       Exit Sub
    End If
    dCharge = CDbl(Text2.Text)
    dChargeBak = dCharge
    dIg = CLng(Text3.Text)
    ret = m1.WriteCard(dCharge, dChargeBak, dIg)
    Label4.Caption = ret
End Sub
