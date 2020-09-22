VERSION 5.00
Object = "*\ACustomSelectControl.vbp"
Begin VB.Form frmTest 
   Caption         =   "SelectBox Test"
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin CustomSelectControl.SelectBox SelectBox9 
      Height          =   225
      Left            =   2940
      TabIndex        =   9
      Top             =   750
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      BoxBorderColor  =   8454143
      BoxBackgroundColor=   255
      BoxStyle        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectMarkColor =   16777215
   End
   Begin CustomSelectControl.SelectBox SelectBox8 
      Height          =   225
      Left            =   2940
      TabIndex        =   8
      Top             =   450
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      BoxBorderColor  =   16744576
      BoxBackgroundColor=   16761024
      BoxStyle        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectMarkColor =   128
   End
   Begin CustomSelectControl.SelectBox SelectBox7 
      Height          =   225
      Left            =   2940
      TabIndex        =   7
      Top             =   150
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      BoxBackgroundColor=   16777215
      BoxStyle        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CustomSelectControl.SelectBox SelectBox6 
      Height          =   225
      Left            =   1500
      TabIndex        =   6
      Top             =   750
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      BoxBorderColor  =   65535
      BoxBackgroundColor=   255
      BoxStyle        =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectMarkColor =   16777215
   End
   Begin CustomSelectControl.SelectBox SelectBox5 
      Height          =   225
      Left            =   1500
      TabIndex        =   5
      Top             =   450
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      BoxBorderColor  =   16744576
      BoxBackgroundColor=   16761024
      BoxStyle        =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectMarkColor =   128
   End
   Begin CustomSelectControl.SelectBox SelectBox4 
      Height          =   225
      Left            =   1500
      TabIndex        =   4
      Top             =   150
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      BoxBackgroundColor=   16777215
      BoxStyle        =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CustomSelectControl.SelectBox SelectBox3 
      Height          =   225
      Left            =   150
      TabIndex        =   3
      Top             =   750
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      BoxBorderColor  =   65535
      BoxBackgroundColor=   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectMarkColor =   16777215
   End
   Begin CustomSelectControl.SelectBox SelectBox2 
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   450
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      BoxBorderColor  =   16744576
      BoxBackgroundColor=   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SelectMarkColor =   128
   End
   Begin CustomSelectControl.SelectBox SelectBox1 
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   397
      BoxBackgroundColor=   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Unload Me"
      Height          =   360
      Left            =   1575
      TabIndex        =   0
      Top             =   1125
      Width           =   1110
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdExit_Click()
     Unload Me
End Sub

Private Sub SelectBox5_Click()
     MsgBox "You clicked SelectBox5.  It's value is " + CStr(SelectBox5.SelectValue), _
          vbOKOnly + vbInformation + vbApplicationModal, "Message"
End Sub
