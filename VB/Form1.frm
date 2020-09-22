VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4350
   ClientLeft      =   5625
   ClientTop       =   2265
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   9990
   Begin VB.CommandButton cmdExecSP 
      Caption         =   "Execute SP"
      Height          =   420
      Left            =   7845
      TabIndex        =   2
      Top             =   570
      Width           =   2085
   End
   Begin VB.CommandButton cmdExecSQL 
      Caption         =   "Execute SQL"
      Height          =   420
      Left            =   7845
      TabIndex        =   1
      Top             =   90
      Width           =   2085
   End
   Begin VB.TextBox Text1 
      Height          =   4125
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   90
      Width           =   7725
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoConn As New ADODB.Connection

Private Sub cmdExecSP_Click()
    ExecuteStoredProcedure
End Sub

Private Sub cmdExecSQL_Click()
    ExecuteSQLStatement
End Sub

Private Sub Form_Load()
    adoConn.ConnectionString = "provider=LCPI.IBProvider;data source=localhost:" & _
        "C:\Program Files\Borland\InterBase\examples\Database\Employee.gdb;ctype=" & _
        "win1251;user id=SYSDBA;password=masterkey"

    adoConn.Open
End Sub


Private Sub ExecuteSQLStatement()
    Dim rst As New Recordset

    Text1 = ""
    
    rst.Source = "SELECT CUSTOMER.CONTACT_FIRST, " & _
                "CUSTOMER.CONTACT_LAST, CUSTOMER.COUNTRY " & _
                "FROM CUSTOMER"
                
    rst.ActiveConnection = adoConn
    
    adoConn.BeginTrans
        
    rst.Open
        If rst.RecordCount Then
            Do Until rst.EOF
                Text1 = Text1 & rst.Fields(0).Value & ", " & rst.Fields(1).Value & ", " & rst.Fields(2).Value & vbCrLf
                rst.MoveNext
            Loop
        End If
    adoConn.CommitTrans
End Sub


Private Sub ExecuteStoredProcedure()
    Dim rst As New Recordset
    Dim cmd As New ADODB.Command
    
    Text1 = ""
    
    With cmd
        .ActiveConnection = adoConn
        .CommandText = "Select * FROM DEPT_BUDGET (100)"
    End With

    adoConn.BeginTrans
        Set rst = cmd.Execute
        If rst.RecordCount Then
            Do Until rst.EOF
                Text1 = Text1 & rst.Fields(0).Value
                rst.MoveNext
            Loop
        End If
    adoConn.CommitTrans
End Sub

