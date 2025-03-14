VERSION 5.00
Begin VB.Form frmInicio 
   Caption         =   "Form1"
   ClientHeight    =   2025
   ClientLeft      =   5130
   ClientTop       =   4665
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   90
      Top             =   1500
   End
   Begin VB.Label lblTitulo 
      Caption         =   "Seja Bem-Vindo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   0
      Top             =   780
      Width           =   18135
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ArquivosConfig(3) As String
Dim Conn As Object ' Objeto de conexão com o banco

Sub CarregarArquivos()
    ArquivosConfig(0) = App.Path & "/arquivo1.ini"
    ArquivosConfig(1) = App.Path & "/arquivo2.ini"
    ArquivosConfig(2) = App.Path & "/arquivo3.ini"
    ArquivosConfig(3) = "DRIVER={PostgreSQL Unicode};Server=localhost;Port=5432;Database=vbdb;Uid=postgres;Pwd=1234;"
End Sub

Sub VerificarArquivos(Num)
    Dim Resultado As String
    
    ' Verifica se é um caminho de arquivo (0 a 2) ou se é a string de conexão (3)
    If Num < 3 Then
        Resultado = Dir(ArquivosConfig(Num))
        If Len(Resultado) > 0 Then
            lblTitulo.Caption = lblTitulo.Caption & " Arquivo carregado!"
        Else
            MsgBox "O arquivo " & ArquivosConfig(Num) & " não foi carregado corretamente.", vbCritical, "ERRO AO CARREGAR OS ARQUIVOS INICIAIS"
            Timer1.Enabled = False
            Unload Me
        End If
    Else
        ' Se for a string de conexão, testa a conexão com o banco
        If Not TestarConexaoPostgreSQL(ArquivosConfig(3)) Then
            MsgBox "Falha ao conectar ao banco de dados!", vbCritical, "ERRO DE CONEXÃO"
            Timer1.Enabled = False
            Unload Me
        Else
            lblTitulo.Caption = "Banco de dados conectado com sucesso!"
        End If
    End If
End Sub

Function TestarConexaoPostgreSQL(StringConexao As String) As Boolean
    On Error Resume Next ' Evita que erros quebrem o código
    Set Conn = CreateObject("ADODB.Connection")
    
    Conn.Open "Provider=MSDASQL;Driver={PostgreSQL Unicode};" & StringConexao

    ' Verifica se conectou corretamente
    If Conn.State = 1 Then
        TestarConexaoPostgreSQL = True
        Conn.Close
    Else
        TestarConexaoPostgreSQL = False
    End If
    
    Set Conn = Nothing
End Function

Sub ConsultaDiretorio(Tempo As Integer)

    Select Case Tempo
        Case 1:
            lblTitulo.Caption = "Carregando os arquivos."
            Call VerificarArquivos(0)
        Case 2:
            lblTitulo.Caption = "Carregando os arquivos.."
            Call VerificarArquivos(1)
        Case 3:
            lblTitulo.Caption = "Carregando os arquivos..."
            Call VerificarArquivos(2)
        Case 4:
            lblTitulo.Caption = "Conectando ao banco de dados..."
            Call VerificarArquivos(3)
        Case Else
            Timer1.Enabled = False
            Unload Me
            frmLogin.Show
    End Select
End Sub

Private Sub Form_Load()
    Call CarregarArquivos
    Timer1.Enabled = True ' Ativa o Timer para iniciar o processo
End Sub

Private Sub Timer1_Timer()
    Static Tempo As Integer
    Tempo = Tempo + 1
    Call ConsultaDiretorio(Tempo)
End Sub

