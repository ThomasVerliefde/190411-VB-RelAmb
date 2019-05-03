
Public Class collectForm
	Inherits Form

	Private collectInstr As New instructionBox(horizontalDist:=0.65, verticalDist:=0.25)
	Private WithEvents collectBox1 As New TextBox
	Private WithEvents collectBox2 As New TextBox
	Private collectPanel1 As New labelledBox(Me.collectBox1, boxWidth:=180)
	Private collectPanel2 As New labelledBox(Me.collectBox2, boxWidth:=180)
	Private WithEvents contButton As New continueButton
	'Private collectPos As New List(Of String)
	'Private collectNeg As New List(Of String)

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Me.WindowState = FormWindowState.Maximized
		Me.FormBorderStyle = FormBorderStyle.None
		Me.BackColor = Color.White

		Me.Controls.Add(Me.collectInstr)
		objCenter(Me.collectInstr, 0.25)
		Me.collectInstr.Rtf = My.Resources.ResourceManager.GetString("_1_collect" & mainForm.currentBlock & mainForm.currentValence)

		Me.Controls.Add(Me.contButton)
		objCenter(Me.contButton, 0.8)

		Me.Controls.Add(Me.collectPanel1)
		objCenter(Me.collectPanel1, 0.5, 0.45)
		Me.Controls.Add(Me.collectPanel2)
		objCenter(Me.collectPanel2, 0.6, 0.45)

		'PLACEHOLDERS FOR LABELS, SHOULD BE CHANGED FOR FINAL
		Select Case mainForm.currentBlock
			Case "Others"
				Me.collectPanel1.reLabel(collectBox1, "Name des 1. signifikanten anderen :")
				Me.collectPanel2.reLabel(collectBox2, "Name des 2. signifikanten anderen :")

				'If debugMode Then
				'	Me.otherBox1.Text = "Adam"
				'	Me.otherBox2.Text = "Chloe"
				'End If

			Case "Objects"
				Me.collectPanel1.reLabel(collectBox1, "Name des 1. signifikanten objekten :")
				Me.collectPanel2.reLabel(collectBox2, "Name des 2. signifikanten objekten :")

				'If debugMode Then
				'	Me.otherBox1.Text = "Piano"
				'	Me.otherBox2.Text = "Teddy"
				'End If

		End Select

		Me.collectBox1.MaxLength = 14
		Me.collectBox2.MaxLength = 14

		Me.collectBox1.Select()

	End Sub

	Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click


		'I need to change detection for other primes

		For Each list In collectFrame.Values
			If list.Contains(Me.collectBox1.Text) OrElse list.Contains(Me.collectBox2.Text) Then
				MsgBox("Bitte geben Sie keine identischen Namen ein!", MsgBoxStyle.Critical, Title:="Fehler!")
				Exit Sub
			End If
		Next

		If Me.collectBox1.Text = "" OrElse Me.collectBox2.Text = "" OrElse
			Not isNoun(Me.collectBox1.Text) OrElse Not isNoun(Me.collectBox2.Text) OrElse
				Me.collectBox1.Text.Length < 2 OrElse Me.collectBox2.Text.Length < 2 Then
			MsgBox("Bitte geben Sie gültige Namen ein!" & vbCrLf &
					"[Namen sollten zwischen 2 und 14 Buchstaben lang sein," & vbCrLf &
					"nur Buchstaben enthalten (Umlaute und Bindestriche sind erlaubt)," & vbCrLf & "und sollten nicht mit einem der anderen Namen identisch sein.]", MsgBoxStyle.Critical, Title:="Fehler!")
			Exit Sub

		ElseIf mainForm.currentValence = mainForm.firstValence Then

			collectFrame("collect" & mainForm.currentBlock & mainForm.currentValence) = New List(Of String) From {StrConv(Me.collectBox1.Text, vbProperCase), StrConv(Me.collectBox2.Text, vbProperCase)}

			Me.collectInstr.Rtf = My.Resources.ResourceManager.GetString("_1_collect" & mainForm.currentBlock & mainForm.currentValence)

			'Select Case mainForm.firstValence
			'	Case "Pos"
			'		Me.collectPos.AddRange({StrConv(Me.collectBox1.Text, vbProperCase), StrConv(Me.collectBox2.Text, vbProperCase)})
			'		Me.collectInstr.Rtf = My.Resources.ResourceManager.GetString("_1_otherNeg")
			'	Case "Neg"
			'		Me.collectNeg.AddRange({StrConv(Me.collectBox1.Text, vbProperCase), StrConv(Me.collectBox2.Text, vbProperCase)})
			'		Me.collectInstr.Rtf = My.Resources.ResourceManager.GetString("_1_otherPos")
			'End Select

			Me.collectBox1.Text = ""
			Me.collectBox2.Text = ""

			'If debugMode Then
			'	Me.otherBox1.Text = "Xavier"
			'	Me.otherBox2.Text = "Zoe"
			'End If

			mainForm.currentValence = mainForm.secondValence
			Me.collectBox1.Select()

		ElseIf mainForm.currentValence = mainForm.secondValence Then

			collectFrame("collect" & mainForm.currentBlock & mainForm.currentValence) = New List(Of String) From {StrConv(Me.collectBox1.Text, vbProperCase), StrConv(Me.collectBox2.Text, vbProperCase)}

			'Select Case mainForm.firstValence
			'	Case "Pos"
			'		mainForm.collectNeg.AddRange({StrConv(Me.collectBox1.Text, vbProperCase), StrConv(Me.collectBox2.Text, vbProperCase)})
			'		mainForm.collectPos = Me.collectPos
			'	Case "Neg"
			'		mainForm.collectPos.AddRange({StrConv(Me.collectBox1.Text, vbProperCase), StrConv(Me.collectBox2.Text, vbProperCase)})
			'		mainForm.collectNeg = Me.collectNeg
			'End Select

			dataFrame("collect" & mainForm.currentBlock & "Pos1") = collectFrame("collect" & mainForm.currentBlock & "Pos")(0).ToString
			dataFrame("collect" & mainForm.currentBlock & "Pos2") = collectFrame("collect" & mainForm.currentBlock & "Pos")(1).ToString
			dataFrame("collect" & mainForm.currentBlock & "Neg1") = collectFrame("collect" & mainForm.currentBlock & "Neg")(0).ToString
			dataFrame("collect" & mainForm.currentBlock & "Neg2") = collectFrame("collect" & mainForm.currentBlock & "Neg")(1).ToString

			mainForm.currentValence = mainForm.firstValence

			Me.Close()
		End If
	End Sub

	'Private Sub pressEnter(sender As Object, e As KeyEventArgs) Handles collectBox1.KeyDown, collectBox2.KeyDown
	'	If e.KeyCode = Keys.Enter Then
	'		If 0 Then
	'		ElseIf Me.collectBox1.Text = "" Then
	'			Me.collectBox1.Select()
	'		ElseIf Me.collectBox2.Text = "" Then
	'			Me.collectBox2.Select()
	'		Else Me.contButton.PerformClick()
	'		End If
	'	End If
	'End Sub

	Private Sub suppressNonAlpha(sender As Object, e As KeyPressEventArgs) Handles collectBox1.KeyPress, collectBox2.KeyPress
		If e.KeyChar <> ControlChars.Back AndAlso Not isNoun(Me.collectBox1.Text & e.KeyChar) AndAlso Not isNoun(Me.collectBox2.Text & e.KeyChar) Then
			e.Handled = True
		End If
	End Sub


End Class