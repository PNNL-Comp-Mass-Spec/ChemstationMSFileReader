Option Strict On

' Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2012

' Tests the ChemstationMSFileReader DLL

Module modMain

	Public Const PROGRAM_DATE As String = "August 7, 2014"

	Public Function Main() As Integer

		TestChemstationReader()
		Return 0

	End Function

	Private Sub TestChemstationReader()

		Dim strMessage As String
		Dim intScanNumber As Integer

		Try

			Dim lstFilesToTest = New List(Of String)
			lstFilesToTest.Add("GCData\3hr_10cGy_10r.D\Data.ms")
			lstFilesToTest.Add("GCData\FSFA_Diffusion_Pre_nondev_010.D\Data.ms")
			lstFilesToTest.Add("GCData\Heliumtest30.D\Data.ms")

			For Each sDatafilePath As String In lstFilesToTest

				Using oReader As ChemstationMSFileReader.clsChemstationDataMSFileReader = New ChemstationMSFileReader.clsChemstationDataMSFileReader(sDatafilePath)

					Console.WriteLine()
					Console.WriteLine(oReader.Header.DatasetName)
					Console.WriteLine(oReader.Header.Description)
					Console.WriteLine(oReader.Header.FileTypeName)
					Console.WriteLine(oReader.Header.AcqDate)

					strMessage = "SpectrumDescription" & ControlChars.Tab & "Minutes" & ControlChars.Tab & "BPI" & ControlChars.Tab & "TIC" & ControlChars.Tab & "TotalSignalRawFromIndex"

					Dim lstSpecInfo As List(Of String) = New List(Of String)
					lstSpecInfo.Add(strMessage)

					Dim intModValue As Integer
					intModValue = CInt(Math.Ceiling(oReader.Header.SpectraCount / 10))

					For intSpectrumIndex As Integer = 0 To oReader.Header.SpectraCount - 1

						intScanNumber = intSpectrumIndex + 1

						Dim oSpectrum As ChemstationMSFileReader.clsSpectralRecord = Nothing
						Dim intTotalSignalRawFromIndex As Integer = 0

						oReader.GetSpectrum(intSpectrumIndex, oSpectrum, intTotalSignalRawFromIndex)

						strMessage = "SpectrumIndex " & intSpectrumIndex & " at " & oSpectrum.RetentionTimeMinutes.ToString("0.00") & " minutes; base peak " & oSpectrum.BasePeakMZ & " m/z with intensity " & oSpectrum.BasePeakAbundance & "; TIC = " & oSpectrum.TIC.ToString("0") & ControlChars.Tab & oSpectrum.RetentionTimeMinutes.ToString("0.00") & ControlChars.Tab & oSpectrum.BasePeakAbundance & ControlChars.Tab & oSpectrum.TIC & ControlChars.Tab & intTotalSignalRawFromIndex
						lstSpecInfo.Add(strMessage)
						If intSpectrumIndex Mod intModValue = 0 Then
							Console.WriteLine(strMessage)
						End If

						Dim lstMZs As List(Of Single) = oSpectrum.Mzs
						Dim lstIntensities As List(Of Int32) = oSpectrum.Intensities

					Next

					Dim strOutFilePath As String
					strOutFilePath = IO.Path.Combine("Results_Debug_" & oReader.Header.DatasetName & ".txt")

					Using swOutFile As IO.StreamWriter = New IO.StreamWriter(New IO.FileStream(strOutFilePath, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.Read))


						swOutFile.WriteLine(oReader.Header.DatasetName)
						swOutFile.WriteLine(oReader.Header.Description)
						swOutFile.WriteLine(oReader.Header.FileTypeName)
						swOutFile.WriteLine(oReader.Header.AcqDate)
						swOutFile.WriteLine(oReader.Header.MiscInfo)
						swOutFile.WriteLine(oReader.Header.OperatorName)
						swOutFile.WriteLine(oReader.Header.InstrumentModel)

						For Each strEntry As String In lstSpecInfo
							swOutFile.WriteLine(strEntry)
						Next

					End Using

				End Using

			Next

		Catch ex As Exception
			Console.WriteLine("Error at scan " & intScanNumber & ": " & ControlChars.NewLine & ex.Message)
			Console.WriteLine(ex.StackTrace)
		End Try

	End Sub

End Module
