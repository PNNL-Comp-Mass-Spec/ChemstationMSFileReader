Option Strict On

' Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2012

' Tests the ChemstationMSFileReader DLL

Module modMain

	Public Const PROGRAM_DATE As String = "March 27, 2012"

	Public Function Main() As Integer

		TestChemstationReader()
		Return 0

	End Function

	Private Sub TestChemstationReader()

		Dim ioSourceFile As System.IO.FileInfo

		Dim lstFilesToTest As System.Collections.Generic.List(Of String) = New System.Collections.Generic.List(Of String)
		Dim strMessage As String

		Try

			lstFilesToTest.Add("GCData\3hr_10cGy_10r.D\Data.ms")
			lstFilesToTest.Add("GCData\FSFA_Diffusion_Pre_nondev_010.D\Data.ms")
			lstFilesToTest.Add("GCData\Heliumtest30.D\Data.ms")

			For Each sDatafilePath As String In lstFilesToTest

				ioSourceFile = New System.IO.FileInfo(sDatafilePath)

				Using oReader As ChemstationMSFileReader.clsChemstationDataMSFileReader = New ChemstationMSFileReader.clsChemstationDataMSFileReader(sDatafilePath)

					Console.WriteLine()
					Console.WriteLine(oReader.Header.DatasetName)
					Console.WriteLine(oReader.Header.Description)
					Console.WriteLine(oReader.Header.FileTypeName)
					Console.WriteLine(oReader.Header.AcqDate)

					strMessage = "SpectrumDescription" & ControlChars.Tab & "Minutes" & ControlChars.Tab & "BPI" & ControlChars.Tab & "TIC"

					Dim lstSpecInfo As System.Collections.Generic.List(Of String) = New System.Collections.Generic.List(Of String)
					lstSpecInfo.Add(strMessage)

					Dim intModValue As Integer
					intModValue = CInt(Math.Ceiling(oReader.Header.SpectraCount / 10))

					For intSpectrumIndex As Integer = 0 To oReader.Header.SpectraCount - 1
						Dim oSpectrum As ChemstationMSFileReader.clsChemstationDataMSFileReader.clsSpectralRecord = Nothing

						oReader.GetSpectrum(intSpectrumIndex, oSpectrum)

						strMessage = "SpectrumIndex " & intSpectrumIndex & " at " & oSpectrum.RetentionTimeMinutes.ToString("0.00") & " minutes; base peak " & oSpectrum.BasePeakMZ & " m/z with intensity " & oSpectrum.BasePeakAbundance & "; TIC = " & oSpectrum.TIC.ToString("0") & ControlChars.Tab & oSpectrum.RetentionTimeMinutes.ToString("0.00") & ControlChars.Tab & oSpectrum.BasePeakAbundance & ControlChars.Tab & oSpectrum.TIC
						lstSpecInfo.Add(strMessage)
						If intSpectrumIndex Mod intModValue = 0 Then
							Console.WriteLine(strMessage)
						End If

						Dim lstMZs As System.Collections.Generic.List(Of Single)
						Dim lstIntensities As System.Collections.Generic.List(Of Int32)

						lstMZs = oSpectrum.Mzs
						lstIntensities = oSpectrum.Intensities

					Next

					Dim strOutFilePath As String
					strOutFilePath = System.IO.Path.Combine("Results_Debug_" & oReader.Header.DatasetName & ".txt")

					Using swOutFile As System.IO.StreamWriter = New System.IO.StreamWriter(New System.IO.FileStream(strOutFilePath, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.Read))


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
			Console.WriteLine("Error: " & ControlChars.NewLine & ex.Message)
			Console.WriteLine(ex.StackTrace)
		End Try

	End Sub

End Module
