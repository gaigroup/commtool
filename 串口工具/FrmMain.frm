VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ڹ���"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   10935
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame3 
      Caption         =   "Hex����(��һ���ո�ָ�)"
      Height          =   1335
      Left            =   6600
      TabIndex        =   17
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton Command4 
         Caption         =   "���� ʮ����������"
         Height          =   855
         Left            =   3360
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   975
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��������"
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   10695
      Begin VB.CommandButton Command3 
         Caption         =   "����"
         Height          =   855
         Left            =   9840
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Top             =   240
         Width           =   9615
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   10080
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6375
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��"
         Height          =   495
         Left            =   4440
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�ر�"
         Height          =   495
         Left            =   5400
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "У��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "���ں�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "������:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "����״̬:δ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "����COM������"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   10695
      Begin VB.TextBox Text1 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   240
         Width           =   9615
      End
      Begin VB.CommandButton Command11 
         Caption         =   "���"
         Height          =   855
         Left            =   9840
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   2400
         Width           =   9615
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error GoTo ErrCommand1_Click
    With MSComm1
        .CommPort = Combo1.Text '�򿪵Ĵ��ں�
        If Combo3.Text = "None" Then
            .Settings = Combo2.Text & ",N,8,1" '�������ò�����+��У��+8λ����+1λ��ֹ
        ElseIf Combo3.Text = "Odd" Then
            .Settings = Combo2.Text & ",O,8,1" '�������ò�����+��У��+8λ����+1λ��ֹ
        ElseIf Combo3.Text = "Even" Then
            .Settings = Combo2.Text & ",E,8,1" '�������ò�����+żУ��+8λ����+1λ��ֹ
        ElseIf Combo3.Text = "Mark" Then
            .Settings = Combo2.Text & ",M,8,1" '�������ò�����+1У��+8λ����+1λ��ֹ
        ElseIf Combo3.Text = "Space" Then
            .Settings = Combo2.Text & ",S,8,1" '�������ò�����+0У��+8λ����+1λ��ֹ
        End If
        .InputMode = comInputModeBinary '����ģʽ
        .NullDiscard = False
        .DTREnable = False 'true -- ���˿ڱ���ʱ Data Terminal Ready ������Ϊ�ߵ�ƽ
        .EOFEnable = False
        .RTSEnable = False
        .InBufferCount = 0 '������뻺��
        .OutBufferCount = 0 '����������
        .SThreshold = 0 '����OnComm�¼�������С�ַ���
        .RThreshold = 1 '����OnComm�¼�������С�ַ���
        .InBufferSize = 512 '�趨���뻺���С
        .OutBufferSize = 512 '�趨��������С
        .PortOpen = True '�򿪴���
    End With
    OpenCom = True
    Label3.Caption = "����״̬:COM" & Combo1.Text & "�Ѵ�"
    MsgBox "COM" & Combo1.Text & "�򿪳ɹ�!"
    Exit Sub
ErrCommand1_Click:
    MsgBox Err.Description, vbInformation, "ϵͳ��Ϣ"
End Sub

Private Sub Command11_Click()
    ComBuf = ""
    Text1.Text = ""
    Label8.Caption = ""
End Sub

Private Sub Command2_Click()
On Error GoTo ErrCommand2_Click
    With MSComm1
        .PortOpen = False '�رմ���
    End With
    OpenCom = False
    Label3.Caption = "����״̬:δ��"
    MsgBox "�رճɹ�!"
    Exit Sub
ErrCommand2_Click:
    MsgBox Err.Description, vbInformation, "ϵͳ��Ϣ"
End Sub

Private Sub Command3_Click()
    Dim Buffer As String
    Dim HexData() As Byte
    Dim Tmp As String
    Dim Loopi As Integer
    If OpenCom = False Then
        MsgBox "������δ��!", vbInformation, "ϵͳ��Ϣ"
        Exit Sub
    End If
    MSComm1.Output = Text2.Text '����ָ��
    Do Until MSComm1.OutBufferCount = 0 '�ȴ����ͽ���
        DoEvents
    Loop
End Sub

Private Sub Command4_Click()
    Dim Tmp As Integer
    Dim TheStr As String
    Dim i As Integer
    Dim j As Integer
    Dim TmpByte(1 To 1) As Byte
    If OpenCom = False Then
        MsgBox "������δ��!", vbInformation, "ϵͳ��Ϣ"
        Exit Sub
    End If
    TheStr = Trim(Text3.Text)
    i = 1
    Do Until InStr(i, TheStr, " ") = 0
        j = InStr(i, TheStr, " ")
        Tmp = GetHex(Mid(TheStr, i, j - i))
        If Tmp = -1 Then
            MsgBox "���ݲ���ȷ:" & Mid(TheStr, i, j - i), vbInformation
            Exit Sub
        Else
            TmpByte(1) = Tmp
            MSComm1.Output = TmpByte '����
            Do Until MSComm1.OutBufferCount = 0 '�ȴ����ͽ���
                DoEvents
            Loop
        End If
        i = j + 1
    Loop
    Tmp = GetHex(Trim(Mid(TheStr, i)))
    If Tmp = -1 Then
        MsgBox "���ݲ���ȷ:" & Trim(Mid(TheStr, i)), vbInformation
        Exit Sub
    Else
        TmpByte(1) = Tmp
        MSComm1.Output = TmpByte '����
        Do Until MSComm1.OutBufferCount = 0 '�ȴ����ͽ���
            DoEvents
        Loop
    End If
End Sub

Private Sub Form_Load()
    With Combo1
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
        .AddItem "9"
        .AddItem "10"
        .AddItem "11"
        .AddItem "12"
        .ListIndex = 0
    End With
    With Combo2
        .Clear
        .AddItem "110"
        .AddItem "300"
        .AddItem "600"
        .AddItem "1200"
        .AddItem "2400"
        .AddItem "4800"
        .AddItem "9600"
        .AddItem "14400"
        .AddItem "19200"
        .AddItem "38400"
        .AddItem "56000"
        .AddItem "57600"
        .AddItem "115200"
        .ListIndex = 12
    End With
    With Combo3
        .Clear
        .AddItem "None"
        .AddItem "Odd"
        .AddItem "Even"
        .AddItem "Mark"
        .AddItem "Space"
        .ListIndex = 0
    End With
    OpenCom = False
    ComBuf = "" '��ʼ�����ܻ����ַ���
End Sub

Private Sub MSComm1_OnComm()
On Error GoTo ErrMSComm1_OnComm
    Dim RecData() As Byte
    Dim Buffer As Variant
    Dim TempD As Double
    Dim TempI As Integer
    Dim Loopi As Integer
    If MSComm1.CommEvent = comEvReceive Then
        ComBuf = ComBuf & MSComm1.Input
        RecData = ComBuf
        Text1.Text = ""
        Label8.Caption = ""
        For Loopi = 0 To UBound(RecData)
            Text1.Text = Text1.Text & Chr(RecData(Loopi))
            Label8.Caption = Label8.Caption & Hex(RecData(Loopi)) & " "
        Next
    Else
        'MsgBox MSComm1.CommEvent
        
    End If
    Exit Sub
ErrMSComm1_OnComm:
    MsgBox Err.Description, vbInformation
End Sub

Private Function GetHex(Num As String) As Integer
    Dim Tmp1 As Integer
    Dim Tmp2 As Integer
    If Len(Num) <> 2 Then
        GetHex = -1
        Exit Function
    End If
    Tmp1 = GetHexNum(Asc(Mid(Num, 1)))
    Tmp2 = GetHexNum(Asc(Mid(Num, 2)))
    If Tmp1 = -1 Or Tmp2 = -1 Then
        GetHex = -1
        Exit Function
    End If
    GetHex = Tmp1 * 16 + Tmp2
End Function

Private Function GetHexNum(Num As Integer) As Integer
    Dim Tmp As Integer
    Tmp = Num
    If Tmp >= 48 And Tmp <= 57 Then
        Tmp = Tmp - 48
    ElseIf Tmp >= 65 And Tmp <= 70 Then
        Tmp = Tmp - 65 + 10
    ElseIf Tmp >= 97 And Tmp <= 102 Then
        Tmp = Tmp - 97 + 10
    Else
        GetHexNum = -1
        Exit Function
    End If
    GetHexNum = Tmp
End Function
