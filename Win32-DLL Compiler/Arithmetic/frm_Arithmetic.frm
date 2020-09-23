VERSION 5.00
Begin VB.Form frm_Arithmetic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arithmetic"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame fra_Container 
      Caption         =   "Result"
      Height          =   690
      Index           =   2
      Left            =   90
      TabIndex        =   9
      Top             =   900
      Width           =   4515
      Begin VB.TextBox txt_Num 
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   270
         Width           =   4290
      End
   End
   Begin VB.CommandButton cmd_Operator 
      Caption         =   "&About"
      Height          =   420
      Index           =   4
      Left            =   1530
      TabIndex        =   8
      Top             =   1710
      Width           =   1635
   End
   Begin VB.CommandButton cmd_Operator 
      Caption         =   "\"
      Height          =   420
      Index           =   3
      Left            =   2745
      TabIndex        =   7
      Top             =   225
      Width           =   420
   End
   Begin VB.CommandButton cmd_Operator 
      Caption         =   "*"
      Height          =   420
      Index           =   2
      Left            =   2340
      TabIndex        =   6
      Top             =   225
      Width           =   420
   End
   Begin VB.CommandButton cmd_Operator 
      Caption         =   "-"
      Height          =   420
      Index           =   1
      Left            =   1935
      TabIndex        =   5
      Top             =   225
      Width           =   420
   End
   Begin VB.CommandButton cmd_Operator 
      Caption         =   "+"
      Height          =   420
      Index           =   0
      Left            =   1530
      TabIndex        =   4
      Top             =   225
      Width           =   420
   End
   Begin VB.Frame fra_Container 
      Caption         =   "Number2"
      Height          =   690
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   90
      Width           =   1365
      Begin VB.TextBox txt_Num 
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame fra_Container 
      Caption         =   "Number1"
      Height          =   690
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1365
      Begin VB.TextBox txt_Num 
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frm_Arithmetic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***Arithmetic, for testing DLL_Arithmetic.dll
'***made by KFGG, China.P.R
'***12/12/2004
Option Explicit
Private Declare Function Plus Lib "DLL_Arithmetic.dll" (ByVal a As Long, ByVal b As Long) As Long
Private Declare Function Minus Lib "DLL_Arithmetic.dll" (ByVal a As Long, ByVal b As Long) As Long
Private Declare Function Multiply Lib "DLL_Arithmetic.dll" (ByVal a As Long, ByVal b As Long) As Long
Private Declare Function Divide Lib "DLL_Arithmetic.dll" (ByVal a As Long, ByVal b As Long) As Long
Private Declare Sub About Lib "DLL_Arithmetic.dll" ()

Private Sub cmd_Operator_Click(Index As Integer)
  Select Case Index
  
  Case 0
    Me.txt_Num(2).Text = Trim(Str(Plus(Val(Me.txt_Num(0).Text), Val(Me.txt_Num(1).Text))))
  Case 1
    Me.txt_Num(2).Text = Trim(Str(Minus(Val(Me.txt_Num(0).Text), Val(Me.txt_Num(1).Text))))
  Case 2
    Me.txt_Num(2).Text = Trim(Str(Multiply(Val(Me.txt_Num(0).Text), Val(Me.txt_Num(1).Text))))
  Case 3
    Me.txt_Num(2).Text = Trim(Str(Divide(Val(Me.txt_Num(0).Text), Val(Me.txt_Num(1).Text))))
  Case 4
    About
  End Select
  
End Sub
