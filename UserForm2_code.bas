Private Sub UserForm_Initialize()

Dim Waux As Worksheet
Set Waux = Worksheets("AUX")
ev_combos
ev_zeratela
ev_liberatela
Me.ScrollBar1.Max = Val(Waux.Cells(7, 12))

If Waux.Cells(10, 12) <> Empty Then

   pont = Val(Waux.Cells(10, 13))
'Aqui dá tipos incompativeis qdo nao por LCTO pois a M10 do 'aux' fica com #N/D
   ev_lebanco (pont)
   ev_travatela
   Me.ScrollBar1.Value = Val(Waux.Cells(10, 13)) - 11
   'Me.ScrollBar1.RowSource = Waux.Cells(10, 12)
   
   Waux.Cells(10, 12) = Empty
   
End If

End Sub
Private Sub ev_zeratela()

Dim Waux As Worksheet
Set Waux = Worksheets("AUX")

With Me
TextBox4.Locked = True
TextBox5.Locked = True               'data
TextBox16.Locked = True              'hora
TextBox2.SetFocus                    'inicia em empresa

TextBox1.Value = Waux.Cells(2, 3)       'numero lancamento
TextBox2 = Empty
TextBox3 = Empty
TextBox4.Value = Waux.Cells(4, 3)       'usuario que inclui
TextBox5.Text = Date
TextBox16.Text = Time
TextBox9 = Empty
TextBox10 = Empty
TextBox12 = Empty
TextBox13 = Empty
TextBox14 = Empty
TextBox15 = Empty
ComboBox1.Text = ComboBox1.List(0)
ComboBox2.Text = ComboBox2.List(0)
ComboBox3.Text = ComboBox3.List(0)

End With

End Sub

Private Sub ev_grava(tipo As Integer)
'
' tipo = 1 --> inclusao de novo (bdlin vem da funcao max na planilha)
' tipo = 2 --> alteracao (tem que achar a linha a ser alterada)

Dim wbde As Worksheet
Set wbde = Worksheets("BDE")
Dim bdlin As Integer

bdlin = 11 + wbde.Cells(7, 2)       'CELLS(7,2)É A LOCALIZ DA PALAVRA 'FIM'

wbde.Cells(bdlin, 1) = Me.TextBox1.Value   'N.LCTO (NAO E PARA DIGITAR)
wbde.Cells(bdlin, 2) = Me.TextBox2.Value   ' num emp
wbde.Cells(bdlin, 3) = Me.TextBox3.Value   ' nome emp
wbde.Cells(bdlin, 4) = Me.TextBox4.Value   ' quem inclui
wbde.Cells(bdlin, 5) = Val(Me.TextBox5.Value)   ' data inclui
wbde.Cells(bdlin, 6) = Me.ComboBox1.Value   ' origem do evento
wbde.Cells(bdlin, 7) = Me.ComboBox2.Value   ' destino_grupo
wbde.Cells(bdlin, 8) = Me.ComboBox3.Value   ' destino_pessoa
wbde.Cells(bdlin, 9) = Me.TextBox9.Value   ' evento
wbde.Cells(bdlin, 10) = Me.TextBox10.Value   ' observ
If Me.OptionButton1.Value Then
   wbde.Cells(bdlin, 11) = 0
Else
   wbde.Cells(bdlin, 11) = 1
End If
wbde.Cells(bdlin, 12) = Val(Me.TextBox12.Value)  ' vencimento
wbde.Cells(bdlin, 13) = Me.TextBox13.Value  ' solução
wbde.Cells(bdlin, 14) = Val(Me.TextBox14.Value)  ' data fim (quando resolveu)
wbde.Cells(bdlin, 15) = Me.TextBox15.Value
wbde.Cells(bdlin, 16) = Me.TextBox16.Value
If Me.CheckBox2 = True Then
    wbde.Cells(bdlin, 17) = 1
Else
    wbde.Cells(bdlin, 17) = 0
End If

wbde.Cells(bdlin + 1, 1) = "fim"

End Sub

Private Sub CommandButton1_Click()

If MsgBox("Deseja salvar este evento?", vbYesNo) = vbYes Then
ev_grava (1)
End If

End Sub
Private Sub CommandButton3_Click()
'muda de consulta para inclusao

With Me
.CommandButton3.Visible = False
.CommandButton4.Visible = True
.ScrollBar1.Visible = False
.CheckBox1.Visible = True

ev_liberatela
ev_zeratela

End With

End Sub
Private Sub CheckBox1_Click()
' abre os campos para Editar data e hora

If Me.CheckBox1.Value Then
   Me.TextBox5.Locked = False
   Me.TextBox16.Locked = False
   Me.TextBox16.BackColor = &H80000005
   Me.TextBox5.BackColor = &H80000005
Else
   Me.TextBox5.Locked = True
   Me.TextBox16.Locked = True
   Me.TextBox5.Text = Date
   Me.TextBox16.Text = Time
   Me.TextBox16.BackColor = &H80000000
   Me.TextBox5.BackColor = &H80000000
End If

End Sub
Private Sub Checkbox2_click()
'apenas OCULTAR
End Sub

Private Sub CheckBox3_Click()
' Resolvido agora

If CheckBox3 Then
   Me.TextBox12.Value = Date
   Me.TextBox14.Value = Date
   Me.TextBox15.Value = Time
   Me.OptionButton1.Value = False
   Me.OptionButton2.Value = True
Else
   Me.TextBox12.Value = Empty
   Me.TextBox14.Value = Empty
   Me.TextBox15.Value = Empty
   Me.OptionButton1.Value = True
   Me.OptionButton2.Value = False
End If

End Sub

Private Sub CommandButton4_Click()
'muda de inclusao para consulta

With Me
.CommandButton4.Visible = False
.CommandButton3.Visible = True
'.ScrollBar1.Visible = True
'.CheckBox1.Visible = False


ev_travatela

End With

End Sub
Private Sub CommandButton5_Click()

If MsgBox("Deseja alterar este evento?", vbYesNo) = vbYes Then
Me.CommandButton5.Visible = False
Me.CommandButton4.Visible = True
Me.ScrollBar1.Visible = False
Me.CheckBox1.Visible = True

ev_liberatela

End If

End Sub

Private Sub CommandButton6_Click()
' TROCA DE TELA
Unload Me
UserForm3.Show

End Sub

Private Sub TextBox2_afterupdate()
    Dim wbanco As Worksheet
    Set wbanco = Worksheets("BD")
    
    
    TextBox2.Value = Val(TextBox2.Value)
    
    wbanco.Cells(2, 2) = TextBox2.Value
    TextBox3.Text = wbanco.Cells(2, 3)

    
    
' tem que mudar a hora que se faz esta consulta (colocar o nome da empresa)
' passar para outro evento ou fazer diferente
'atualmente esta em textbox9

End Sub
Private Sub TextBox5_afterupdate()

dDate = Me.TextBox5.Value
If Len(dDate) = 4 Then Me.TextBox5.Value = Format(Mid(dDate, 1, 2) + "/" + Mid(dDate, 3, 2), "dd/mm/yyyy")
If Len(dDate) = 5 Then Me.TextBox5.Value = Format(dDate, "dd/mm/yyyy")
If Len(dDate) = 8 Then Me.TextBox5.Value = Format(dDate, "dd/mm/yyyy")

End Sub

Private Sub TextBox12_afterupdate()

dDate = Me.TextBox12.Value
If Len(dDate) = 4 Then Me.TextBox12.Value = Format(Mid(dDate, 1, 2) + "/" + Mid(dDate, 3, 2), "dd/mm/yyyy")
If Len(dDate) = 5 Then Me.TextBox12.Value = Format(dDate, "dd/mm/yyyy")
If Len(dDate) = 8 Then Me.TextBox12.Value = Format(Mid(dDate, 1, 2) + "/" + Mid(dDate, 3, 2) + "/" + Mid(dDate, 5, 4), "dd/mm/yyyy")

'If Me.TextBox12.ValuE < Me.TextBox5.Value Then
'   MsgBox ("Estou alterando para proximo ano")
'   Me.TextBox12.Value = Format(Mid(dDate, 1, 2) + "/" + Mid(dDate, 3, 2) + "/" + Str(Val(Mid(dDate, 5, 4)) + 1), "dd/mm/yyyy")
'End If


End Sub


Private Sub ScrollBar1_Change()
' é a barra de rolagem da consulta

Dim wbde As Worksheet
Set wbde = Worksheets("BDE")

Dim pont As Integer

pont = 11 + ScrollBar1.Value

If pont <= 11 Or wbde.Cells(pont, 1) = "fim" Then Exit Sub

Label11.Caption = "Registro " + CStr(Me.ScrollBar1.Value) + " de " + CStr(Me.ScrollBar1.Max)

'fixar a empresa
'If CheckBox3 Then
'   If TextBox4 <> CheckBox3.Caption Then
'    pont = pont + 1
'    ScrollBar1.value = ....+1  mas nao deu
'End If
'End If
'ScrollBar1.Value = Val(ScrollBar1.Value) + 1
'CheckBox3.Caption = VarType(ScrollBar1.Value)

ev_lebanco (pont)


End Sub
Private Sub ev_lebanco(pont As Integer)
Dim wbde As Worksheet
Set wbde = Worksheets("BDE")

TextBox3.Enabled = True
TextBox5.Enabled = True
TextBox16.Enabled = True
TextBox4.Enabled = True

TextBox1.Value = wbde.Cells(pont, 1)
TextBox2.Value = wbde.Cells(pont, 2)
TextBox3.Value = wbde.Cells(pont, 3)
TextBox4.Value = wbde.Cells(pont, 4)
TextBox5.Value = wbde.Cells(pont, 5)
ComboBox1.Text = wbde.Cells(pont, 6) '*****
ComboBox2.Text = wbde.Cells(pont, 7) '*****
ComboBox3.Text = wbde.Cells(pont, 8) '*****
TextBox9.Value = wbde.Cells(pont, 9)
TextBox10.Value = wbde.Cells(pont, 10)

If wbde.Cells(pont, 11) = 0 Then
   Me.OptionButton1.Value = True
   Me.OptionButton2.Value = False
Else
   Me.OptionButton1.Value = False
   Me.OptionButton2.Value = True
End If
   
ev_travapend
   

TextBox12.Value = wbde.Cells(pont, 12)
TextBox13.Value = wbde.Cells(pont, 13)
TextBox14.Value = wbde.Cells(pont, 14)
TextBox15.Value = wbde.Cells(pont, 15)
TextBox16.Value = Format(wbde.Cells(pont, 16), "hh:mm")


If wbde.Cells(pont, 17) = 1 Then Me.CheckBox2 = True
If wbde.Cells(pont, 17) = 0 Then Me.CheckBox2 = False

End Sub

Private Sub ev_travatela()
With Me
.ScrollBar1.Visible = True
.CheckBox1.Visible = False
.TextBox1.BackColor = &H80000000
.TextBox2.BackColor = &H80000000
.TextBox4.BackColor = &H80000000
.TextBox5.BackColor = &H80000000
.ComboBox1.BackColor = &H80000000
.ComboBox2.BackColor = &H80000000
.ComboBox3.BackColor = &H80000000
.TextBox9.BackColor = &H80000000
.TextBox10.BackColor = &H80000000
'.TextBox11.BackColor = &H80000000
.TextBox12.BackColor = &H80000000
.TextBox13.BackColor = &H80000000
.TextBox14.BackColor = &H80000000
.TextBox15.BackColor = &H80000000
.TextBox16.BackColor = &H80000000
.TextBox1.Locked = True
.TextBox2.Locked = True
.TextBox3.Locked = True
.TextBox4.Locked = True
.TextBox5.Locked = True
.ComboBox1.Locked = True
.ComboBox2.Locked = True
.ComboBox3.Locked = True
.TextBox9.Locked = True
.TextBox10.Locked = True
'.TextBox11.Locked = True
.TextBox12.Locked = True
.TextBox13.Locked = True
.TextBox14.Locked = True
.TextBox15.Locked = True
.TextBox16.Locked = True
.OptionButton1.Locked = True
.OptionButton2.Locked = True
.CheckBox3.Visible = False

End With

End Sub

Private Sub ev_liberatela()
With Me

.TextBox1.BackColor = &H80000005
.TextBox2.BackColor = &H80000005
.TextBox4.BackColor = &H80000005
.TextBox5.BackColor = &H80000005
.ComboBox1.BackColor = &H80000005
.ComboBox2.BackColor = &H80000005
.ComboBox3.BackColor = &H80000005
.TextBox9.BackColor = &H80000005
.TextBox10.BackColor = &H80000005
'.TextBox11.BackColor = &H80000005
.TextBox12.BackColor = &H80000005
.TextBox13.BackColor = &H80000005
.TextBox14.BackColor = &H80000005
.TextBox15.BackColor = &H80000005
.TextBox16.BackColor = &H80000005

.TextBox1.Locked = False
.TextBox2.Locked = False
.TextBox3.Locked = False
.TextBox4.Locked = False
.TextBox5.Locked = False
.ComboBox1.Locked = False
.ComboBox2.Locked = False
.ComboBox3.Locked = False
.TextBox9.Locked = False
.TextBox10.Locked = False
'.TextBox11.Locked = False
.TextBox12.Locked = False
.TextBox13.Locked = False
.TextBox14.Locked = False
.TextBox15.Locked = False
.TextBox16.Locked = False
.OptionButton1.Locked = False
.OptionButton2.Locked = False
.CheckBox3.Visible = True

End With

End Sub
Private Sub ev_travapend()
   TextBox13.Locked = True
   TextBox14.Locked = True
   TextBox15.Locked = True
   TextBox13.BackColor = &H80000000
   TextBox14.BackColor = &H80000000
   TextBox15.BackColor = &H80000000

End Sub
Private Sub ev_liberapend()
   TextBox13.Locked = False
   TextBox14.Locked = False
   TextBox15.Locked = False
   TextBox13.BackColor = &H80000005
   TextBox14.BackColor = &H80000005
   TextBox15.BackColor = &H80000005

End Sub

Private Sub OptionButton1_Click()
If OptionButton1 Then
   ev_travapend
End If

End Sub
Private Sub OptionButton2_Click()
If OptionButton2 Then
   ev_liberapend
End If
End Sub


Public Sub ev_combos()
ComboBox3.AddItem "------"
ComboBox3.AddItem "Maximo"
ComboBox3.AddItem "Rita"
ComboBox3.AddItem "Maria"
ComboBox3.AddItem "Leonardo"
ComboBox3.AddItem "Leocir"
ComboBox3.AddItem "Leandro"
ComboBox3.AddItem "Jackson"
ComboBox3.AddItem "Estagiario"

ComboBox3.Text = ComboBox3.List(0)

'************** COMBO 2  =  GRUPO DE TRABALHO

ComboBox2.AddItem "------"
ComboBox2.AddItem "CONTABIL"
ComboBox2.AddItem "FISCAL"
ComboBox2.AddItem "PESSOAL"
ComboBox2.AddItem "REPARTIÇÕES"
ComboBox2.AddItem "DECLARAÇÕES"
ComboBox2.AddItem "INFORMATICA"
ComboBox2.AddItem "OUTROS"

'Dim Waux As Worksheet
'Set Waux = Worksheets("AUX")

If Worksheets("AUX").Cells(4, 4) = "permite" Then
   ComboBox2.AddItem "COBRANÇA**"
Else
   ComboBox2.AddItem "**********"
End If

ComboBox2.Text = ComboBox2.List(0)

'************** COMBO 1 = Origem do evento

ComboBox1.AddItem "------"
ComboBox1.AddItem "Telefone"
ComboBox1.AddItem "E-mail"
ComboBox1.AddItem "Direto"
ComboBox1.AddItem "Interno"
ComboBox1.AddItem "Conferencia"
ComboBox1.AddItem "Pesquisa"

ComboBox1.Text = ComboBox1.List(0)
'------------------------------------------

End Sub
