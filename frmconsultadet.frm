VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmconsultadet 
   BackColor       =   &H00CC9966&
   Caption         =   "Consulta Det"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15480
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   15480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Abrir Aplicativo DET"
      Height          =   255
      Left            =   6240
      TabIndex        =   13
      Top             =   5520
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   7440
      TabIndex        =   6
      Top             =   5040
      WhatsThisHelpID =   20
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   5040
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Limpiar"
      Height          =   300
      Left            =   6240
      TabIndex        =   5
      Top             =   5040
      WhatsThisHelpID =   20
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5400
      TabIndex        =   7
      Text            =   "descripcion"
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   5520
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   11280
      TabIndex        =   1
      Top             =   5040
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   11280
      MaxLength       =   14
      TabIndex        =   4
      Top             =   5520
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4455
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   17418
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   17418
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00CC9966&
      Caption         =   "Lassy Radio"
      Height          =   255
      Left            =   14400
      TabIndex        =   12
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CC9966&
      Caption         =   "Digite el apellido 1"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00CC9966&
      Caption         =   "Digite los nombres"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00CC9966&
      Caption         =   "Digite el apellido 2"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9000
      TabIndex        =   9
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00CC9966&
      Caption         =   "Digite el nit"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9000
      TabIndex        =   8
      Top             =   5520
      Width           =   2175
   End
End
Attribute VB_Name = "frmconsultadet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub Command1_Click()
    Text1.Text = ""
    Text3.Text = ""
    Text5.Text = ""
    Text7.Text = ""

    
End Sub

Private Sub Command2_Click()
Dim archivo As String
Dim ok As String

ok = MsgBox("nit a buscar " & DataGrid1.Columns(0) & " en el det", vbInformation, "Lassy")


CommonDialog1.Filter = "archivos Exe(*.exe"
CommonDialog1.ShowOpen
archivo = CommonDialog1.FileName
Shell (archivo)


End Sub
Private Sub Form_Activate()

    DataGrid1.Columns(0).Width = 2000
    DataGrid1.Columns(1).Width = 3000
    DataGrid1.Columns(2).Width = 3000
    DataGrid1.Columns(3).Width = 4000
        Text1.SetFocus

End Sub

Private Sub Form_Load()
 Set db = New Connection

CommonDialog1.Filter = "Base de datos de access 97 (*.mdb|*.mdb"
CommonDialog1.ShowOpen
pri = CommonDialog1.FileName

 If db.State = adStateOpen Then
        db.Close
    End If
    db.CursorLocation = adUseClient

db.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" + pri + ""

Set consulta = New Recordset
consulta.Open "select datGeneral.nit, datGeneral.apellido1, datGeneral.apellido2, datGeneral.nombres from datGeneral ", db, adOpenStatic, adLockOptimistic
Set DataGrid1.DataSource = consulta

End Sub

Private Sub text1_Change()
    If Text2.Text = "descripcion" Then
        If Trim(Text1.Text) = "" Then
            consulta.Filter = adFilterNone
            consulta.Requery
            Form_Activate
        Else
            consulta.Filter = "apellido1 Like '%" + Text1.Text + "%'"
        End If
    End If
End Sub






Private Sub text5_Change()

    If Text2.Text = "descripcion" Then
        If Trim(Text5.Text) = "" Then
            consulta.Filter = adFilterNone
            consulta.Requery
            Form_Activate
        Else
            consulta.Filter = "apellido2 Like '%" + Text5.Text + "%'"
            
        End If
    End If
    
End Sub



Private Sub Text3_Change()

    If Text2.Text = "descripcion" Then
        If Trim(Text3.Text) = "" Then
            consulta.Filter = adFilterNone
            consulta.Requery
            Form_Activate
        Else
            consulta.Filter = "nombres Like '%" + Text3.Text + "%'"
        End If
    End If
End Sub

Private Sub Text7_Change()
If Text2.Text = "descripcion" Then
        If Trim(Text7.Text) = "" Then
            consulta.Filter = adFilterNone
            consulta.Requery
            Form_Activate
        Else
            consulta.Filter = "nit like '%" + Text7.Text + "%'"
        End If
    End If
End Sub



