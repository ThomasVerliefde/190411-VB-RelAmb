
Public Class explicitForm
	Inherits Form

	Private WithEvents trackB1 As New TrackBar
	Private labB1 As New labelledTrackbar(Me.trackB1)

	Private WithEvents trackB2 As New TrackBar
	Private labB2 As New labelledTrackbar(Me.trackB2)

	Private WithEvents trackB3 As New TrackBar
	Private labB3 As New labelledTrackbar(Me.trackB3)

	Private labIntro As New Label
	Private labName As New Label

	Private WithEvents relText As New TextBox
	Private relBox As New labelledBox(Me.relText, boxWidth:=600, fieldHeight:=350)

	Private WithEvents numText As New TextBox
	Private numBox As New labelledBox(Me.numText, "Wie viele Leute kennen Sie mit diesem Vornamen? ")

	Private WithEvents contButton As New continueButton
	Private correctRel As Boolean
	Private correctNum As Boolean

	Private collectKeys As New List(Of String)({"Pos1", "Pos2", "Neg1", "Neg2"})
	Private collectCount As Integer
	Private questionCount As Integer

	Private collectKey As String
	Private collectName As String

	Private tempFrame As New SortedDictionary(Of String, String)

	Private B1 As Boolean
	Private B2 As Boolean
	Private B3 As Boolean

	Private amountQuestions As Integer = 2

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Me.WindowState = FormWindowState.Maximized
		Me.FormBorderStyle = FormBorderStyle.None
		Me.BackColor = Color.White

		mainForm.currentBlock = mainForm.firstBlock

		Me.trackB1.Name = "B1"
		Me.trackB2.Name = "B2"
		Me.trackB3.Name = "B3"

		shuffleList(Me.collectKeys)

		With Me.labIntro
			.Text = "Bitte beantworten Sie die folgenden Fragen zu:"
			.Font = sansSerif22
			.Width = TextRenderer.MeasureText(.Text, sansSerif22).Width
			.Height = TextRenderer.MeasureText(.Text, sansSerif22).Height
		End With

		Me.Controls.Add(Me.labIntro)
		objCenter(Me.labIntro, verticalDist:=0.1, setLeft:=100)

		Me.Controls.Add(Me.labName)
		objCenter(Me.labName, verticalDist:=0.095)

		Me.Controls.Add(Me.labB1)
		Me.Controls.Add(Me.labB2)
		Me.Controls.Add(Me.labB3)
		Me.Controls.Add(Me.numBox)
		Me.Controls.Add(Me.relBox)
		objCenter(Me.labB1, verticalDist:=0.3)
		objCenter(Me.labB2, verticalDist:=0.5)
		objCenter(Me.labB3, verticalDist:=0.7)
		objCenter(Me.numBox, 0.35, setLeft:=100)
		objCenter(Me.relBox, 0.55, setLeft:=100)

		Me.relText.Multiline = True

		Me.Controls.Add(Me.contButton)
		objCenter(Me.contButton)

		Me.setQuestion()

	End Sub

	Private Sub contButton_Click(sender As Object, e As EventArgs) Handles contButton.Click

		If Me.questionCount Mod Me.amountQuestions = 0 Then
			Me.tempFrame(mainForm.currentBlock & "_" & Me.collectKey & "_Rel") = Me.relText.Text.ToString.Replace(vbCr, " ").Replace(vbLf, " ")
			Me.tempFrame(mainForm.currentBlock & "_" & Me.collectKey & "_Num") = Me.numText.Text.ToString
			Me.tempFrame(mainForm.currentBlock & "_" & Me.collectKey & "_Dir_Pos") = Me.trackB1.Value.ToString
			Me.tempFrame(mainForm.currentBlock & "_" & Me.collectKey & "_Dir_Neg") = Me.trackB2.Value.ToString
			Me.tempFrame(mainForm.currentBlock & "_" & Me.collectKey & "_Dir_Amb") = Me.trackB3.Value.ToString
		End If

		If Me.collectCount >= Me.collectKeys.Count Then

			If mainForm.currentBlock = mainForm.firstBlock Then
				mainForm.currentBlock = mainForm.secondBlock
				Me.collectCount = 0
				shuffleList(Me.collectKeys)
			Else
				For Each item In Me.tempFrame
					dataFrame.Add(item.Key, item.Value)
				Next
				Me.Close()
			End If

		Else

			Me.setQuestion()

		End If

	End Sub

	Private Sub setQuestion()

		Me.collectKey = Me.collectKeys(Me.collectCount)
		Me.collectName = dataFrame("collect" & mainForm.currentBlock & Me.collectKey)

		With Me.labName
			.Text = Me.collectName
			.Font = sansSerif25B
			.Width = TextRenderer.MeasureText(.Text, sansSerif25B).Width
			.Height = TextRenderer.MeasureText(.Text, sansSerif25B).Height
		End With

		'Resetting the checks whether the questions have received attention
		Me.B1 = False
		Me.B2 = False
		Me.B3 = False
		Me.correctRel = False
		Me.correctNum = mainForm.currentBlock = "Objects" 'This needs to be True if we are not showing the numBox, in order to make contButton enabled
		Me.contButton.Enabled = debugMode 'False if debugMode is off, True if debugMode is on

		Select Case Me.questionCount Mod Me.amountQuestions

			Case 0

				Me.labB1.Visible = False
				Me.labB2.Visible = Me.labB1.Visible
				Me.labB3.Visible = Me.labB1.Visible
				Me.relBox.Visible = True
				Me.numBox.Visible = mainForm.currentBlock = "Others"

				Me.relText.ResetText()
				Me.numText.ResetText()

				If mainForm.currentBlock = "Others" Then
					Me.relBox.reLabel(Me.relText, "Was ist deine soziale Beziehung zu dieser Person?")
					Me.numBox.Select()
				ElseIf mainForm.currentBlock = "Objects" Then
					Me.relBox.reLabel(Me.relText, "Was ist dieses Objekt genau und warum ist es für Sie von Bedeutung?")
					Me.relText.Select()
					Me.numBox.Text = "NA"
				End If

				If debugMode Then
					Me.relText.Text = "DEBUG"
					Me.numText.Text = 1
				End If

			Case 1 ' Explicit positive, negative, and ambivalent

				Me.labB1.Visible = True
				Me.labB2.Visible = Me.labB1.Visible
				Me.labB3.Visible = Me.labB1.Visible
				Me.relBox.Visible = False
				Me.numBox.Visible = Me.relBox.Visible

				Me.labB1.reInit("Wie positiv finden Sie " & Me.collectName & "?" & vbCrLf &
								" Konzentrieren Sie sich für Ihr Urteil bitte nur auf die positiven Aspekte" & vbCrLf &
								"und ignorieren Sie mögliche negative Aspekte.", Me.trackB1, 0.5, "neutral", "sehr positiv", minVal:=0, maxVal:=100, freqVal:=50, defVal:=0)
				Me.labB2.reInit("Wie negativ finden Sie " & Me.collectName & "?" & vbCrLf &
								" Konzentrieren Sie sich für Ihr Urteil bitte nur auf die negativen Aspekte" & vbCrLf &
								"und ignorieren Sie mögliche positive Aspekte.", Me.trackB2, 0.5, "neutral", "sehr negativ", minVal:=0, maxVal:=100, freqVal:=50, defVal:=0)
				Me.labB3.reInit("Wie hin- und hergerissen fühlen Sie sich angesichts " & Me.collectName, Me.trackB3,, "überhaupt nicht", "sehr", minVal:=0, maxVal:=100, freqVal:=50, defVal:=0)

				Me.collectCount += 1

		End Select

		Me.questionCount += 1

	End Sub


	Private Sub enableTrack(sender As Object, e As EventArgs) Handles trackB1.MouseDown, trackB2.MouseDown, trackB3.MouseDown

		Select Case DirectCast(sender, TrackBar).Name
			Case "B1"
				Me.B1 = True

			Case "B2"
				Me.B2 = True

			Case "B3"
				Me.B3 = True

		End Select

		If Me.B1 AndAlso Me.B2 AndAlso Me.B3 Then
			Me.contButton.Enabled = True
		End If

	End Sub

	Private Sub enableNum(sender As Object, e As EventArgs) Handles numText.TextChanged
		If Val(Me.numText.Text) < 1 OrElse Val(Me.numText.Text) > 25 Then
			Me.correctNum = False
		ElseIf Val(Me.numText.Text) > 0 Then
			Me.correctNum = True
		End If
		Me.contButton.Enabled = Me.correctNum AndAlso Me.correctRel
	End Sub

	Private Sub enableRel(sender As Object, e As EventArgs) Handles relText.TextChanged
		If Me.relText.Text.Length > 1 Then
			Me.correctRel = True
		ElseIf Me.relText.Text.Length <= 1 Then
			Me.correctRel = False
		End If
		Me.contButton.Enabled = Me.correctNum AndAlso Me.correctRel
	End Sub

	Private Sub suppressNonNumeric(sender As Object, e As KeyPressEventArgs) Handles numText.KeyPress
		e.Handled = e.KeyChar <> ControlChars.Back AndAlso Not IsNumeric(e.KeyChar)
	End Sub

	Private Sub suppressColon(sender As Object, e As KeyPressEventArgs) Handles relText.KeyPress
		e.Handled = e.KeyChar = ","
	End Sub

End Class