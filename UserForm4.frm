VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Eventus - FICHAS"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Dim wbde As Worksheet
Set wbde = Worksheets("BDE")
Dim bdlin As Integer

If MsgBox("Deseja salvar este evento?", vbYesNo) = vbYes Then

bdlin = 11 + wbde.Cells(7, 2)



wbde.Cells(bdlin, 1) = Me.TextBox9.Value   'POREM ESTAVA SEM VALOR (NAO E PARA DIGITAR)
wbde.Cells(bdlin, 2) = Me.TextBox1.Value   ' num emp
wbde.Cells(bdlin, 3) = Me.TextBox8.Value   ' nome emp
wbde.Cells(bdlin, 4) = Me.TextBox7.Value   ' quem inclui
wbde.Cells(bdlin, 5) = Me.TextBox11.Value   ' data inclui
wbde.Cells(bdlin, 6) = Me.ComboBox2.Value   ' origem do evento
wbde.Cells(bdlin, 7) = Me.ComboBox3.Value   ' destino_grupo
wbde.Cells(bdlin, 8) = Me.ComboBox1.Value   ' destino_pessoa
wbde.Cells(bdlin, 9) = Me.TextBox5.Value   ' evento
wbde.Cells(bdlin, 10) = Me.TextBox6.Value   ' observ
wbde.Cells(bdlin, 11) = Me.TextBox9.Value  ' resolvido/etc
wbde.Cells(bdlin, 12) = Me.TextBox12.Value  ' data programada
wbde.Cells(bdlin, 13) = Me.TextBox10.Value  ' solução
wbde.Cells(bdlin, 14) = Me.TextBox2.Value  ' data fim (quando resolveu)
wbde.Cells(bdlin + 1, 1) = "fim"





End If


End Sub

Private Sub CommandButton13_Click()
Unload Me
UserForm3.Show

End Sub

Private Sub CommandButton3_Click()

MsgBox ("Origem do Evento" + Chr(13) + Chr(13) + "Telefone, E-mail se a solicitação foi por estas vias" + Chr(13) + Chr(13) + "Interno, se for de uma pessoa para outra" + Chr(13) + Chr(13) + "Direto, quando o contato é direto com o cliente (veio aqui ou fomos lá)" + Chr(13) + Chr(13) + "Conferencia, quando aparece um erro ou pendencia em um processo que necessita de outro," + Chr(13) + "Ou pura e simples conferência")
End Sub
Private Sub commandbutton5_click()

MsgBox ("DATA FIM é a data programada para terminar a tarefa/evento" + Chr(13) + Chr(13) + "O campo é opcional")
End Sub

Private Sub ScrollBar1_Change()
Dim wbde As Worksheet
Set wbde = Worksheets("BDE")

Dim pont As Integer

pont = 11 + ScrollBar1.Value
'MsgBox (pont)

'MsgBox (wbde.Cells(pont, 1))

TextBox18.Value = wbde.Cells(pont, 1)
TextBox13.Value = wbde.Cells(pont, 2)
TextBox17.Value = wbde.Cells(pont, 3)
TextBox7.Value = wbde.Cells(pont, 4)
TextBox11.Value = wbde.Cells(pont, 5)
ComboBox5.Text = wbde.Cells(pont, 6) '*****
ComboBox6.Text = wbde.Cells(pont, 7) '*****
ComboBox4.Text = wbde.Cells(pont, 8) '*****
TextBox15.Value = wbde.Cells(pont, 9)
TextBox16.Value = wbde.Cells(pont, 10)
'TextBox19.Value = wbde.Cells(pont, 11)  este são as opções
TextBox20.Value = wbde.Cells(pont, 12)
TextBox19.Value = wbde.Cells(pont, 13)
TextBox14.Value = wbde.Cells(pont, 14)

End Sub

Private Sub TextBox1_Change()

Dim wbanco As Worksheet
Set wbanco = Worksheets("BD")
wbanco.Cells(2, 2) = TextBox1.Value

End Sub

Private Sub TextBox5_Change()
Dim wbanco As Worksheet
Set wbanco = Worksheets("BD")
TextBox8.Text = wbanco.Cells(2, 3)

End Sub

Private Sub UserForm_Initialize()

Dim Waux As Worksheet
Set Waux = Worksheets("AUX")

ComboBox1.AddItem "------"
ComboBox1.AddItem "Maximo"
ComboBox1.AddItem "Rita"
ComboBox1.AddItem "Maria"
ComboBox1.AddItem "Leonardo"
ComboBox1.AddItem "Leocir"
ComboBox1.AddItem "Leandro"
ComboBox1.AddItem "Jackson"
ComboBox1.AddItem "Estagiario"

ComboBox1.Text = ComboBox1.List(0)
'Call constroi_combo_user

'************** COMBO 3 = GRUPO DE TRABALHO

ComboBox3.AddItem "------"
ComboBox3.AddItem "CONTABIL"
ComboBox3.AddItem "FISCAL"
ComboBox3.AddItem "PESSOAL"
ComboBox3.AddItem "REPARTIÇÕES"
ComboBox3.AddItem "DECLARAÇÕES"
ComboBox3.AddItem "INFORMATICA"
ComboBox3.AddItem "OUTROS"
ComboBox3.AddItem "COBRANÇA**"

ComboBox3.Text = ComboBox3.List(0)

'************** COMBO 2 = Origem do evento

ComboBox2.AddItem "------"
ComboBox2.AddItem "Telefone"
ComboBox2.AddItem "E-mail"
ComboBox2.AddItem "Direto"
ComboBox2.AddItem "Interno"
ComboBox2.AddItem "Conferencia"
ComboBox2.AddItem "Pesquisa"

ComboBox2.Text = ComboBox2.List(0)

TextBox7.Value = Waux.Cells(4, 3)

End Sub


