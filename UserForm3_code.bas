Private Sub CommandButton1_Click()

Dim Waux As Worksheet
Set Waux = Worksheets("AUX")

v_queminclui = ComboBox1.Value
Waux.Cells(4, 3) = v_queminclui

UserForm2.Label17.Caption = "Usuário: " + v_queminclui

If Waux.Cells(4, 4) = "permite" Then
   UserForm2.CheckBox2.Visible = True
 Else
   UserForm2.CheckBox2.Visible = False
End If

Unload Me
UserForm2.Show

End Sub

Private Sub UserForm_Initialize()

Dim Waux As Worksheet
Set Waux = Worksheets("AUX")

ComboBox1.AddItem "------"
ComboBox1.AddItem "MAX"
ComboBox1.AddItem "RIT"
ComboBox1.AddItem "MAR"
ComboBox1.AddItem "LMO"
ComboBox1.AddItem "LVM"
ComboBox1.AddItem "LEM"
ComboBox1.AddItem "JKS"
ComboBox1.AddItem "ETG"

ComboBox1.Text = ComboBox1.List(0)
ComboBox1.Text = Waux.Cells(4, 3)


End Sub



