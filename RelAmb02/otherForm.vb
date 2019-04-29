
Public Class otherForm
	Inherits Form

	Private otherInstr As New instructionBox(horizontalDist:=0.65, verticalDist:=0.25)
	Private WithEvents otherBox1 As New TextBox
	Private WithEvents otherBox2 As New TextBox
	Private otherPanel1 As New labelledBox(Me.otherBox1, "Name des 1. signifikanten anderen :", boxWidth:=180)
	Private otherPanel2 As New labelledBox(Me.otherBox2, "Name des 2. signifikanten anderen :", boxWidth:=180)
	Private WithEvents contButton As New continueButton
	Private otherPos As New List(Of String)
	Private otherNeg As New List(Of String)

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Me.WindowState = FormWindowState.Maximized
		Me.FormBorderStyle = FormBorderStyle.None
		Me.BackColor = Color.White

		Me.Controls.Add(Me.otherInstr)
		objCenter(Me.otherInstr, 0.25)
		Me.otherInstr.Rtf = My.Resources.ResourceManager.GetString("_1_other" & mainForm.firstValence)

		Me.Controls.Add(Me.contButton)
		objCenter(Me.contButton, 0.8)

		Me.Controls.Add(Me.otherPanel1)
		objCenter(Me.otherPanel1, 0.5, 0.45)
		Me.Controls.Add(Me.otherPanel2)
		objCenter(Me.otherPanel2, 0.6, 0.45)

		Me.otherBox1.MaxLength = 14
		Me.otherBox2.MaxLength = 14

		Me.otherBox1.Select()

		If debugMode Then
			Me.otherBox1.Text = "Adam"
			Me.otherBox2.Text = "Chloe"
		End If

	End Sub

	Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click

		If Me.otherBox1.Text = "" OrElse Me.otherBox2.Text = "" OrElse
			Not IsName(Me.otherBox1.Text) OrElse Not IsName(Me.otherBox2.Text) OrElse
				Me.otherBox1.Text.Length < 2 OrElse Me.otherBox2.Text.Length < 2 OrElse
			Me.otherPos.Contains(Me.otherBox1.Text) OrElse Me.otherNeg.Contains(Me.otherBox1.Text) OrElse
				Me.otherPos.Contains(Me.otherBox2.Text) OrElse Me.otherNeg.Contains(Me.otherBox2.Text) OrElse
				Me.otherBox1.Text = Me.otherBox2.Text Then
			'Me.otherBox1.Text.Length > 14 OrElse Me.otherBox2.Text.Length > 14 OrElse
			' this should be dealth with by the 'MaxLength' property

			MsgBox("Bitte geben Sie gültige Namen ein!" & vbCrLf &
					"[Namen sollten zwischen 2 und 14 Buchstaben lang sein," & vbCrLf &
					"nur Buchstaben enthalten (Umlaute und Bindestriche sind erlaubt)," & vbCrLf & "und sollten nicht mit einem der anderen 3 Namen identisch sein.]", MsgBoxStyle.Critical, Title:="Fehler!")
			Exit Sub


		ElseIf Me.otherPos.Count = 0 AndAlso Me.otherNeg.Count = 0 Then
			Select Case mainForm.firstValence
				Case "Pos"
					Me.otherPos.AddRange({StrConv(Me.otherBox1.Text, vbProperCase), StrConv(Me.otherBox2.Text, vbProperCase)})
					Me.otherInstr.Rtf = My.Resources.ResourceManager.GetString("_1_otherNeg")
				Case "Neg"
					Me.otherNeg.AddRange({StrConv(Me.otherBox1.Text, vbProperCase), StrConv(Me.otherBox2.Text, vbProperCase)})
					Me.otherInstr.Rtf = My.Resources.ResourceManager.GetString("_1_otherPos")
			End Select

			Me.otherBox1.Text = ""
			Me.otherBox2.Text = ""

			If debugMode Then
				Me.otherBox1.Text = "Xavier"
				Me.otherBox2.Text = "Zoe"
			End If

			Me.otherBox1.Select()

		ElseIf Me.otherPos.Count = 2 OrElse Me.otherNeg.Count = 2 Then
			Select Case mainForm.firstValence
				Case "Pos"
					mainForm.otherNeg.AddRange({StrConv(Me.otherBox1.Text, vbProperCase), StrConv(Me.otherBox2.Text, vbProperCase)})
					mainForm.otherPos = Me.otherPos
				Case "Neg"
					mainForm.otherPos.AddRange({StrConv(Me.otherBox1.Text, vbProperCase), StrConv(Me.otherBox2.Text, vbProperCase)})
					mainForm.otherNeg = Me.otherNeg
			End Select

			dataFrame("otherPos1") = mainForm.otherPos(0).ToString
			dataFrame("otherPos2") = mainForm.otherPos(1).ToString
			dataFrame("otherNeg1") = mainForm.otherNeg(0).ToString
			dataFrame("otherNeg2") = mainForm.otherNeg(1).ToString

			Me.Close()
		End If
	End Sub

	Private Sub pressEnter(sender As Object, e As KeyEventArgs) Handles otherBox1.KeyDown, otherBox2.KeyDown
		If e.KeyCode = Keys.Enter Then
			If 0 Then
			ElseIf Me.otherBox1.Text = "" Then
				Me.otherBox1.Select()
			ElseIf Me.otherBox2.Text = "" Then
				Me.otherBox2.Select()
			Else Me.contButton.PerformClick()
			End If
		End If
	End Sub

	Private Sub suppressNonAlpha(sender As Object, e As KeyPressEventArgs) Handles otherBox1.KeyPress, otherBox2.KeyPress
		If e.KeyChar <> ControlChars.Back AndAlso Not IsName(Me.otherBox1.Text & e.KeyChar) AndAlso Not IsName(Me.otherBox2.Text & e.KeyChar) Then
			e.Handled = True
		End If
	End Sub


End Class