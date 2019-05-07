Imports NodaTime

Public Class mainForm
	Inherits Form

	Public instructionCount As New Integer 'main counter, to control the flow of the experiment; increases after each continuebutton press in this form

	Friend WithEvents contButton As New continueButton 'main button, to advance the flow of the experiment
	Private WithEvents buttonTimer As New Timer() 'Timer to disable the continuebutton after immediately pressing it, to prevent double-click registration
	Private ReadOnly instrText As New instructionBox 'main method of displaying the instructions to participants; disabled & readonly

	'All NodaTime.Duration variables, to check the actual duration (difference in sequential starting points) of each part
	Private timeCollect01 As Duration
	Private timePractice01 As Duration
	Private timeExperiment01 As Duration
	Private timeCollect02 As Duration
	Private timePractice02 As Duration
	Private timeExperiment02 As Duration
	Private timeExplicit As Duration
	Private timeDemographics As Duration
	Private timeAmbi As Duration
	Private timeTotal As Duration

	'Variable necessary for grabbing the correct instruction sheet, depending on conditions, see subjectForm.vb as well
	Friend keyAss As String
	Friend firstValence As String
	Friend secondValence As String
	Friend currentValence As String
	Friend firstBlock As String
	Friend secondBlock As String
	Friend currentBlock As String

	'Pools of primes for both practice and experiment blocks; see mainModule for the lists of trials (perhaps should move those here or vice versa..)
	Public practicePrimes As List(Of List(Of String))
	Public experimentPrimes As List(Of List(Of String))

	Private Sub formLoad(sender As Object, e As EventArgs) Handles MyBase.Load

		Me.WindowState = FormWindowState.Maximized
		Me.FormBorderStyle = FormBorderStyle.None
		Me.BackColor = Color.White

		subjectForm.ShowDialog()

		Me.Controls.Add(Me.instrText)
		objCenter(Me.instrText, 0.42)
		Me.instrText.Rtf = My.Resources.ResourceManager.GetString("_0_mainInstr")

		Me.Controls.Add(Me.contButton)
		objCenter(Me.contButton)

		Me.buttonTimer.Interval = 1500

	End Sub

	Public Sub loadNext(sender As Object, e As EventArgs) Handles contButton.Click

		delayContinue()

		Select Case Me.instructionCount
			Case 0 'Start of the First Block
				timeFrame("startT") = time.GetCurrentInstant()
				Me.instrText.Rtf = My.Resources.ResourceManager.GetString("_1_collect" & Me.currentBlock)

			Case 1 'Collecting Names & Making Primes for the First Block
				timeFrame("collect" & Me.currentBlock & "T") = time.GetCurrentInstant()
				collectForm.ShowDialog()
				subjectForm.Dispose()

				Me.instrText.Rtf = My.Resources.ResourceManager.GetString("_2_practice" & Me.keyAss)
				collectForm.Dispose()

				' Creating All Practice & Experiment Trials for the Current Block

				' Practice Trials

				Dim collectPractice As New List(Of String)(My.Resources.ResourceManager.GetString("practice" & Me.currentBlock).Split(" "))
				collectPractice = compareList(collectFrame("collect" & Me.currentBlock & "Neg"), collectPractice)
				collectPractice = compareList(collectFrame("collect" & Me.currentBlock & "Pos"), collectPractice)
				shuffleList(collectPractice)
				Dim practicePrime_Pos As New List(Of String)(My.Resources.practicePrime_Pos.Split(" "))
				shuffleList(practicePrime_Pos)
				Dim practicePrime_Neg As New List(Of String)(My.Resources.practicePrime_Neg.Split(" "))
				shuffleList(practicePrime_Neg)
				Dim practicePrime_Str As New List(Of String)(My.Resources.practicePrime_Str.Split(" "))
				shuffleList(practicePrime_Str)

				'practicePrimes is a List of List of String with 5 lists: PositiveNounPrimes, NegativeNounPrimes, PositiveOthers, NegativeOthers, MatchingLetterStrings
				Me.practicePrimes = New List(Of List(Of String))({
								New List(Of String)({practicePrime_Pos(0)}),
								New List(Of String)({practicePrime_Neg(0)}),
								New List(Of String)({collectPractice(0)}),
								New List(Of String)({collectPractice(1)}),
								New List(Of String)({practicePrime_Str(0)})
								})

				practiceTrials = createTrials(
										Me.practicePrimes,
										New List(Of List(Of String))({
											 New List(Of String)(My.Resources.practiceTarget_Pos.Split(" ")),
											 New List(Of String)(My.Resources.practiceTarget_Neg.Split(" "))
											 }),
										timesPrimes:=2 'How often each prime is paired with a target from each category
										)
				' Results in 20 Trials (Can shorten to 10 by setting timesPrimes to 1)

				shuffleList(practiceTrials)

				Dim savePrimes As New List(Of String)
				Me.practicePrimes.ForEach(Sub(x) savePrimes.Add(String.Join(" ", x)))
				dataFrame("practicePrimes" & Me.currentBlock) = String.Join(" ", savePrimes)

				Dim saveTrials As New List(Of String)
				practiceTrials.ForEach(Sub(x) saveTrials.Add(String.Join(" ", x)))
				dataFrame("practiceTrials" & Me.currentBlock) = String.Join("-", saveTrials)

				' Experiment Trials

				Me.experimentPrimes = createPrimes(
					collectFrame("collect" & Me.currentBlock & "Pos"),
					collectFrame("collect" & Me.currentBlock & "Neg"),
					New List(Of String)(My.Resources.experimentPrime_Pos.Split(" ")),
					New List(Of String)(My.Resources.experimentPrime_Neg.Split(" ")),
					New List(Of String)(My.Resources.experimentPrime_Str.Split(" "))
				)

				experimentTrials = createTrials(
					Me.experimentPrimes,
					New List(Of List(Of String))({
						New List(Of String)(My.Resources.experimentTarget_Pos.Split(" ")),
						New List(Of String)(My.Resources.experimentTarget_Neg.Split(" "))
					}),
					timesPrimes:=4 'How often each prime is paired with a target from each category
				)
				' Results in 96 Trials (12 [2+2+2+2+4 Primes] x 2 [Targets] x 4 [timesPrimes]) 

				shuffleList(experimentTrials)

				savePrimes.Clear()
				Me.experimentPrimes.ForEach(Sub(x) savePrimes.Add(String.Join(" ", x)))
				dataFrame("experimentPrimes" & Me.currentBlock) = String.Join(" ", savePrimes)

				saveTrials.Clear()
				experimentTrials.ForEach(Sub(x) saveTrials.Add(String.Join(" ", x)))
				dataFrame("experimentTrials" & Me.currentBlock) = String.Join("-", saveTrials)

				If debugMode Then
					Console.WriteLine("------" & Me.currentBlock & "------")
					Console.WriteLine("- practicePrimes -")
					For Each c In Me.practicePrimes
						For Each d In c
							Console.Write(" * " + d)
						Next
						Console.WriteLine("")
					Next
					Console.WriteLine("")

					Console.WriteLine("- practiceTrials -")

					Dim amount As Integer
					For Each c In practiceTrials
						For Each d In c
							Console.Write(" * " + d)
						Next
						amount += c.Count
						Console.WriteLine("")
					Next
					Console.WriteLine("Amount of Trials: " & amount)

					Console.WriteLine("- experimentPrimes -")
					For Each c In Me.experimentPrimes
						For Each d In c
							Console.Write(" * " + d)
						Next
						Console.WriteLine("")
					Next
					Console.WriteLine("")

					Console.WriteLine("- experimentTrials -")

					amount = 0
					For Each c In experimentTrials
						For Each d In c
							Console.Write(" * " + d)
						Next
						amount += c.Count
						Console.WriteLine("")
					Next
					Console.WriteLine("Amount of Trials: " & amount)
				End If

			Case 2 'Practice Trials for the First Block
				'Me.practiceT = time.GetCurrentInstant()
				timeFrame("practice" & Me.currentBlock & "T") = time.GetCurrentInstant()
				practiceForm.ShowDialog()
				Me.instrText.Rtf = My.Resources.ResourceManager.GetString("_3_experiment" & Me.keyAss)
				practiceForm.Dispose()

			Case 3 'Experiment Proper of the First Block
				'Me.experimentT = time.GetCurrentInstant()
				timeFrame("experiment" & Me.currentBlock & "T") = time.GetCurrentInstant()
				experimentForm.ShowDialog()

				' Starting the second Block, or continues to Explicit
				Select Case Me.currentBlock
					Case Me.firstBlock
						Me.instrText.Rtf = My.Resources.ResourceManager.GetString("_3_breakInstr")
						Me.instructionCount = -1
						Me.currentBlock = Me.secondBlock
					Case Me.secondBlock
						Me.instrText.Rtf = My.Resources.ResourceManager.GetString("_4_explicitInstr")
				End Select
				experimentForm.Dispose()

			Case 4 'Explicit/Direct Measurements of Ambivalence
				'Me.explicitT = time.GetCurrentInstant()
				timeFrame("explicitT") = time.GetCurrentInstant()
				explicitForm.ShowDialog()
				Me.instrText.Rtf = My.Resources.ResourceManager.GetString("_5_demoInstr")
				explicitForm.Dispose()

			Case 5 'Demographic Information
				'Me.demographicsT = time.GetCurrentInstant()
				timeFrame("demographicsT") = time.GetCurrentInstant()
				demographicsForm.ShowDialog()
				Me.instrText.Rtf = My.Resources.ResourceManager.GetString("_6_ambiInstr")
				demographicsForm.Dispose()

				'Ambivalent Word Collection
				' This has no separate case, as it is implemented below the _6_ambiInstr text

				timeFrame("ambiT") = time.GetCurrentInstant()
				ambiForm.ShowDialog()

				'End
				Me.instrText.Rtf = My.Resources.ResourceManager.GetString("_7_endInstr")

				'Case 6 'Ambivalent Word Collection
				'Me.ambiT = time.GetCurrentInstant()
				'Me.instrText.Rtf = My.Resources.ResourceManager.GetString("_7_endInstr")
				Me.instrText.Font = sansSerif40
				ambiForm.Dispose()

				Me.contButton.Text = "Speichern"
				'Me.endT = time.GetCurrentInstant()
				timeFrame("endT") = time.GetCurrentInstant()

				Me.timeCollect01 = timeFrame("practice" & Me.firstBlock & "T") - timeFrame("collect" & Me.firstBlock & "T")
				Me.timePractice01 = timeFrame("experiment" & Me.firstBlock & "T") - timeFrame("practice" & Me.firstBlock & "T")
				Me.timeExperiment01 = timeFrame("collect" & Me.secondBlock & "T") - timeFrame("experiment" & Me.firstBlock & "T")
				Me.timeCollect02 = timeFrame("practice" & Me.secondBlock & "T") - timeFrame("collect" & Me.secondBlock & "T")
				Me.timePractice02 = timeFrame("experiment" & Me.secondBlock & "T") - timeFrame("practice" & Me.secondBlock & "T")
				Me.timeExperiment02 = timeFrame("explicitT") - timeFrame("experiment" & Me.secondBlock & "T")
				Me.timeExplicit = timeFrame("demographicsT") - timeFrame("explicitT")
				Me.timeDemographics = timeFrame("ambiT") - timeFrame("demographicsT")
				Me.timeAmbi = timeFrame("endT") - timeFrame("ambiT")
				Me.timeTotal = timeFrame("endT") - timeFrame("startT")

				dataFrame("timeCollect01") = Me.timeCollect01.TotalMinutes.ToString.Replace(",", ".")
				dataFrame("timePractice01") = Me.timePractice01.TotalMinutes.ToString.Replace(",", ".")
				dataFrame("timeExperiment01") = Me.timeExperiment01.TotalMinutes.ToString.Replace(",", ".")
				dataFrame("timeCollect02") = Me.timeCollect02.TotalMinutes.ToString.Replace(",", ".")
				dataFrame("timePractice02") = Me.timePractice02.TotalMinutes.ToString.Replace(",", ".")
				dataFrame("timeExperiment02") = Me.timeExperiment02.TotalMinutes.ToString.Replace(",", ".")
				dataFrame("timeExplicit") = Me.timeExplicit.TotalMinutes.ToString.Replace(",", ".")
				dataFrame("timeDemographics") = Me.timeDemographics.TotalMinutes.ToString.Replace(",", ".")
				dataFrame("timeAmbi") = Me.timeAmbi.TotalMinutes.ToString.Replace(",", ".")
				dataFrame("timeTotal") = Me.timeTotal.TotalMinutes.ToString.Replace(",", ".")
				dataFrame("hostName") = Net.Dns.GetHostName()

				IO.Directory.CreateDirectory("Data")
				saveCSV(dataFrame, "Data\RelAmb02_" & dataFrame("Subject") & "_" & Net.Dns.GetHostName & ".csv")

			Case Else
				Me.Close()
		End Select
		Me.instructionCount += 1
		delayContinue()
	End Sub

	Private Sub delayContinue()

		Me.contButton.Enabled = False
		Me.buttonTimer.Start()

	End Sub

	Private Sub buttonEnabler(sender As Object, e As EventArgs) Handles buttonTimer.Tick

		Me.buttonTimer.Stop()
		Me.contButton.Enabled = True

	End Sub

End Class
