Public Class subjectForm
	Inherits Form

	Private WithEvents contButton As New continueButton(Txt:="Bestätigen")
	Private WithEvents textBox As New TextBox
	Private WithEvents subjPanel As New labelledBox(textBox, "VPNr:")
	Private condN As String
	Private subjN As String

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
		Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height

		Me.WindowState = FormWindowState.Normal
		Me.FormBorderStyle = FormBorderStyle.None

		Me.Size = New Point(250, 200)
		Me.BackColor = Color.Gray
		Me.Top = (3 * screenHeight / 4) - (Me.Height / 2)
		Me.Left = (screenWidth / 2) - (Me.Width / 2)

		Me.Controls.Add(Me.contButton)
		objCenter(Me.contButton)

		Me.Controls.Add(Me.subjPanel)
		objCenter(Me.subjPanel, 0.4)

		Me.textBox.MaxLength = 3
		Me.textBox.Select()

	End Sub

	Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click

		If Me.textBox.Text = "999" Then
			debugMode = True
			'Me.textBox.Text = "499"
		ElseIf Me.textBox.Text = "" OrElse Val(Me.textBox.Text) < 1 OrElse Val(Me.textBox.Text) > 500 Then
		'Val("these are letters") returns 0
		MsgBox("Bitte geben Sie eine korrekte VPNr ein!", MsgBoxStyle.Critical, Title:="Fehler!")
			Exit Sub

		Else

			Me.subjN = Me.textBox.Text
			Me.condN = setCond(Me.subjN)

			Select Case Val(Me.condN(0))
				Case 0
					mainForm.keyAss = "Apos"
				Case 1
					mainForm.keyAss = "Aneg"
			End Select

			Select Case Val(Me.condN(1))
				Case 0
					mainForm.firstValence = "Pos"
					mainForm.secondValence = "Neg"
					mainForm.currentValence = "Pos"
				Case 1
					mainForm.firstValence = "Neg"
					mainForm.secondValence = "Pos"
					mainForm.currentValence = "Neg"
			End Select

			Select Case Val(Me.condN(2))
				Case 0
					mainForm.firstBlock = "Others"
					mainForm.secondBlock = "Objects"
					mainForm.currentBlock = "Others"
				Case 1
					mainForm.firstBlock = "Objects"
					mainForm.secondBlock = "Others"
					mainForm.currentBlock = "Objects"
			End Select

			dataFrame("Subject") = Me.subjN
			dataFrame("Condition") = Me.condN
			dataFrame("Key") = mainForm.keyAss
			dataFrame("firstBlock") = mainForm.firstBlock
			dataFrame("firstValence") = mainForm.firstValence

			Me.Close()
		End If
	End Sub

	Private Sub enterConfirm(sender As Object, e As KeyEventArgs) Handles textBox.KeyDown
		If e.KeyCode = Keys.Enter Then
			Me.contButton.PerformClick()
		End If
	End Sub

	Private Sub suppressNonNumeric(sender As Object, e As KeyPressEventArgs) Handles textBox.KeyPress
		If e.KeyChar <> ControlChars.Back AndAlso Not IsNumeric(e.KeyChar) Then
			e.Handled = True
		End If
	End Sub


End Class