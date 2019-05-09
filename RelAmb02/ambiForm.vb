Public Class ambiForm
	Inherits Form

	Private WithEvents contButton As New continueButton(Txt:="Bestätigen")
	Private WithEvents textBox As New TextBox

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Dim screenWidth As Integer = Screen.PrimaryScreen.Bounds.Width
		Dim screenHeight As Integer = Screen.PrimaryScreen.Bounds.Height

		Me.WindowState = FormWindowState.Normal
		Me.FormBorderStyle = FormBorderStyle.None

		Me.Size = New Point(screenWidth * 0.7, screenHeight * 0.6)
		Me.Top = (screenHeight * 0.6) - (Me.Height * 0.5)
		Me.Left = (screenWidth * 0.5) - (Me.Width * 0.5)

		Me.Controls.Add(Me.contButton)
		objCenter(Me.contButton, 0.9)

		Me.textBox.Multiline = True
		Me.textBox.Height = Me.Height * 0.70
		Me.textBox.Width = Me.Width * 0.9
		Me.textBox.Font = sansSerif20
		Me.Controls.Add(Me.textBox)
		objCenter(Me.textBox, 0.4)
		Me.textBox.Select()

	End Sub

	Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click
		dataFrame("ambiSuggestions") = Me.textBox.Text.Replace(vbCr, " ").Replace(vbLf, " ")
		Me.Close()
	End Sub

	Private Sub suppressColon(sender As Object, e As KeyPressEventArgs) Handles textBox.KeyPress
		e.Handled = e.KeyChar = ","
	End Sub

End Class