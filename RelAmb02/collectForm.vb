
Public Class collectForm
	Inherits Form

	Private collectInstr As New instructionBox(horizontalDist:=0.65, verticalDist:=0.25)
	Private WithEvents collectBox1 As New TextBox
	Private WithEvents collectBox2 As New TextBox
	Private collectPanel1 As New labelledBox(Me.collectBox1, boxWidth:=180)
	Private collectPanel2 As New labelledBox(Me.collectBox2, boxWidth:=180)
	Private WithEvents contButton As New continueButton

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Me.WindowState = FormWindowState.Maximized
		Me.FormBorderStyle = FormBorderStyle.None
		Me.BackColor = Color.White

		Me.Controls.Add(Me.collectInstr)
		objCenter(Me.collectInstr, 0.25)

		Me.Controls.Add(Me.contButton)
		objCenter(Me.contButton, 0.8)

		Me.Controls.Add(Me.collectPanel1)
		objCenter(Me.collectPanel1, 0.5, 0.45)
		Me.Controls.Add(Me.collectPanel2)
		objCenter(Me.collectPanel2, 0.6, 0.45)

		Me.collectBox1.MaxLength = 14
		Me.collectBox2.MaxLength = 14

		Me.setQuestion()

	End Sub

	Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click

		For Each list In collectFrame.Values
			If list.Contains(Me.collectBox1.Text) OrElse list.Contains(Me.collectBox2.Text) Then
				MsgBox("Bitte geben Sie keine identischen Namen ein!", MsgBoxStyle.Critical, Title:="Fehler!")
				Exit Sub
			End If
		Next

		If Me.collectBox1.Text = "" OrElse Me.collectBox2.Text = "" OrElse
				Me.collectBox1.Text.ToLower = Me.collectBox2.Text.ToLower OrElse
			Not isNoun(Me.collectBox1.Text) OrElse Not isNoun(Me.collectBox2.Text) OrElse
				Me.collectBox1.Text.Length < 2 OrElse Me.collectBox2.Text.Length < 2 Then
			MsgBox("Bitte geben Sie gültige Namen ein!" & vbCrLf &
					"[Namen sollten zwischen 2 und 14 Buchstaben lang sein," & vbCrLf &
					"nur Buchstaben enthalten (Umlaute und Bindestriche sind erlaubt)," & vbCrLf & "und sollten nicht mit einem der anderen Namen identisch sein.]", MsgBoxStyle.Critical, Title:="Fehler!")
			Exit Sub

		ElseIf mainForm.currentValence = mainForm.firstValence Then

			collectFrame("collect" & mainForm.currentBlock & mainForm.currentValence) = New List(Of String) From {StrConv(Me.collectBox1.Text, vbProperCase), StrConv(Me.collectBox2.Text, vbProperCase)}
			mainForm.currentValence = mainForm.secondValence

			Me.setQuestion()

		ElseIf mainForm.currentValence = mainForm.secondValence Then

			collectFrame("collect" & mainForm.currentBlock & mainForm.currentValence) = New List(Of String) From {StrConv(Me.collectBox1.Text, vbProperCase), StrConv(Me.collectBox2.Text, vbProperCase)}
			mainForm.currentValence = mainForm.firstValence

			dataFrame("collect" & mainForm.currentBlock & "Pos1") = collectFrame("collect" & mainForm.currentBlock & "Pos")(0).ToString
			dataFrame("collect" & mainForm.currentBlock & "Pos2") = collectFrame("collect" & mainForm.currentBlock & "Pos")(1).ToString
			dataFrame("collect" & mainForm.currentBlock & "Neg1") = collectFrame("collect" & mainForm.currentBlock & "Neg")(0).ToString
			dataFrame("collect" & mainForm.currentBlock & "Neg2") = collectFrame("collect" & mainForm.currentBlock & "Neg")(1).ToString

			Me.Close()
		End If
	End Sub

	Private Sub setQuestion()

		Dim CBlock As New String("")
		Dim CVal As New String("")

		Select Case mainForm.currentBlock
			Case "Others"
				CBlock = "anderen"
			Case "Objects"
				CBlock = "Objekts"
		End Select
		Select Case mainForm.currentValence
			Case "Pos"
				CVal = "positiven"
			Case "Neg"
				CVal = "negativen"
		End Select

		Me.collectPanel1.reLabel(collectBox1, "Name des 1. " & CVal & " signifikanten " & CBlock & ":")
		Me.collectPanel2.reLabel(collectBox2, "Name des 2. " & CVal & " signifikanten " & CBlock & ":")

		Me.collectBox1.Text = ""
		Me.collectBox2.Text = ""
		Me.collectBox1.Select()

		Me.collectInstr.Rtf = My.Resources.ResourceManager.GetString("_1_collect" & mainForm.currentBlock & mainForm.currentValence)

	End Sub

	Private Sub suppressNonAlpha(sender As Object, e As KeyPressEventArgs) Handles collectBox1.KeyPress, collectBox2.KeyPress
		If e.KeyChar <> ControlChars.Back AndAlso Not isNoun(Me.collectBox1.Text & e.KeyChar) AndAlso Not isNoun(Me.collectBox2.Text & e.KeyChar) AndAlso Not e.KeyChar = Keys.OemSemicolon.ToString Then
			e.Handled = True
		End If
	End Sub


End Class