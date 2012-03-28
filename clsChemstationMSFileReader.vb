Option Strict On

' This class reads Chemstation Data.MS files
' These are binary files that use a file format originally developed by HP in the 1980s
' Agilent still uses this file format for GC/MS data
'
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2012
'
' Last modified March 27, 2012

Public Class clsChemstationDataMSFileReader
	Implements IDisposable

#Region "Structures"

	Protected Structure udtNormalizationRecordType
		Public Mass As Single
		Public Slope As Single
		Public Intercept As Single
	End Structure

	Protected Structure udtIndexEntryType
		Public OffsetBytes As Int32					' Stored as integer (in words); converted to Byte offset by this class
		Public RetentionTimeMsec As Int32			' Stored in milliseconds
		Public TotalSignalRaw As Int32				' This Total Signal value is representative of TIC, but it has been scaled down by some sort of polynomial transformation
		Public ReadOnly Property RetentionTimeMinutes() As Single
			Get
				Return RetentionTimeMsec / 60000.0!
			End Get
		End Property
	End Structure

#End Region

#Region "Member Variables"

	Public Header As clsDataHeader

	Protected mFileStream As System.IO.FileStream
	Protected mNormalizationRecordList As System.Collections.Generic.List(Of udtNormalizationRecordType)

	' Index entries (aka Directory Records)
	Protected mIndexList As System.Collections.Generic.List(Of udtIndexEntryType)

#End Region

	''' <summary>
	''' Open the specified data file and read the data headers
	''' </summary>
	''' <param name="sDatafilePath">Path to the file to read</param>
	''' <remarks></remarks>
	Public Sub New(ByVal sDatafilePath As String)

		mNormalizationRecordList = New System.Collections.Generic.List(Of udtNormalizationRecordType)
		mIndexList = New System.Collections.Generic.List(Of udtIndexEntryType)

		' Read the headers from the data file
		ReadHeaders(sDatafilePath)
	End Sub

	Public Function GetSpectrum(ByVal intSpectrumIndex As Integer, ByRef oSpectrum As clsSpectralRecord) As Boolean
		Return GetSpectrum(intSpectrumIndex, oSpectrum, 0)
	End Function

	''' <summary>
	'''  Returns the mass spectrum at the specified index
	''' </summary>
	''' <param name="intSpectrumIndex">0-based spectrum index</param>
	''' <param name="oSpectrum">Spectrum object (output)</param>
	''' <param name="intTotalSignalRawFromIndex">TIC value as reported by the Index; this value has been scaled down by some sort of polynomial transformation</param>
	''' <returns>True if success, false if an error</returns>
	Public Function GetSpectrum(ByVal intSpectrumIndex As Integer, ByRef oSpectrum As clsSpectralRecord, ByRef intTotalSignalRawFromIndex As Integer) As Boolean

		Dim sngRetentionTimeMinutes As Single

		sngRetentionTimeMinutes = 0
		intTotalSignalRawFromIndex = 0

		If intSpectrumIndex < 0 OrElse intSpectrumIndex >= mIndexList.Count Then
			' Index out of range
			oSpectrum = New clsSpectralRecord()
			Return False
		Else
			Dim intByteOffset As Integer

			intByteOffset = mIndexList(intSpectrumIndex).OffsetBytes
			sngRetentionTimeMinutes = mIndexList(intSpectrumIndex).RetentionTimeMinutes()

			' These Total Signal values are not accurate
			' Use the TIC value reported by oSpectrum instead
			intTotalSignalRawFromIndex = mIndexList(intSpectrumIndex).TotalSignalRaw

			oSpectrum = New clsSpectralRecord(mFileStream, intByteOffset)

			If oSpectrum.RetentionTimeMinutes <> sngRetentionTimeMinutes Then
				Console.WriteLine("  ... retention time mismatch; this is unexpected: " & oSpectrum.RetentionTimeMinutes & " vs. " & sngRetentionTimeMinutes)
			End If
		End If

		Return False
	End Function

	''' <summary>
	''' Open the data file and read the header sections from the data file
	''' </summary>
	''' <param name="sDataFilePath">Path to the file to read</param>
	''' <returns>True if success, false if an error</returns>
	''' <remarks>The file handle will remain open until this class is disposed of</remarks>
	Protected Function ReadHeaders(ByVal sDataFilePath As String) As Boolean

		Dim blnSuccess As Boolean = False

		If Not System.IO.File.Exists(sDataFilePath) Then
			Throw New System.IO.FileNotFoundException("Data file not found", sDataFilePath)
		End If

		Try
			' Open the data file
			mFileStream = New System.IO.FileStream(sDataFilePath, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)

		Catch ex As Exception
			Throw New Exception("Error opening data file: " & ex.Message, ex)
			Return False
		End Try

		Try
			' Read the header data
			' Note that clsDataHeader will throw an exception if it occurs
			Header = New clsDataHeader(mFileStream)
		Catch ex As Exception
			Throw ex
			Return False
		End Try

		If Header Is Nothing OrElse Not Header.Valid Then
			Return False
		End If

		Try
			' Read the Normalization records
			blnSuccess = ReadNormalizationRecords(mFileStream)
		Catch ex As Exception
			Throw ex
			Return False
		End Try

		If Not blnSuccess Then Return False

		Try
			' Read the Index entries
			blnSuccess = ReadIndexRecords(mFileStream)
		Catch ex As Exception
			Throw ex
			Return False
		End Try

		Return blnSuccess

	End Function


	''' <summary>
	''' Reads the index entries (aka directory records)
	''' </summary>
	''' <param name="fsDatafile"></param>
	''' <returns>True if success, false if an error</returns>
	Protected Function ReadIndexRecords(ByRef fsDatafile As System.IO.FileStream) As Boolean

		Dim bc As New clsByteConverter()

		'Const LOG_BASE As Integer = 8
		'Dim intTotalSignalMin As Integer = Integer.MaxValue
		'Dim intTotalSignalMax As Integer = 0

		Try
			' Move the filestream to the correct byte offset
			fsDatafile.Seek(Header.DirectoryOffset, IO.SeekOrigin.Begin)

			For intIndex As Integer = 0 To Header.SpectraCount - 1
				Dim udtEntry As udtIndexEntryType

				udtEntry.OffsetBytes = bc.WordOffsetToBytes(bc.ReadInt32SwapBytes(fsDatafile) - 1)
				udtEntry.RetentionTimeMsec = bc.ReadInt32SwapBytes(fsDatafile)
				udtEntry.TotalSignalRaw = bc.ReadInt32SwapBytes(fsDatafile)

				'If udtEntry.TotalSignalRaw > intTotalSignalMax Then intTotalSignalMax = udtEntry.TotalSignalRaw
				'If udtEntry.TotalSignalRaw < intTotalSignalMin Then intTotalSignalMin = udtEntry.TotalSignalRaw

				mIndexList.Add(udtEntry)

			Next

			'If mIndexList.Count > 0 Then
			'	' Now compute TotalSignalScaled for each entry in mIndexList
			'	Dim dblScaledSignal As Double
			'	Dim dblTotalSignalMinLog As Double = Math.Log(intTotalSignalMin, LOG_BASE)
			'	Dim dblTotalSignalMaxLog As Double = Math.Log(intTotalSignalMax, LOG_BASE)

			'	For intIndex As Integer = 0 To mIndexList.Count - 1
			'		Dim udtEntry As udtIndexEntryType

			'		udtEntry = mIndexList(intIndex)

			'		' Compute the log of the number
			'		dblScaledSignal = Math.Log(udtEntry.TotalSignalRaw, LOG_BASE)

			'		' Scale to a value between 0 and 1
			'		dblScaledSignal = (dblScaledSignal - dblTotalSignalMinLog) / (dblTotalSignalMaxLog - dblTotalSignalMinLog)

			'		' Scale to the range Header.SignalMinimum to Header.SignalMaximum
			'		udtEntry.TotalSignalScaled = dblScaledSignal * (Header.SignalMaximum - Header.SignalMinimum) + Header.SignalMinimum

			'		mIndexList(intIndex) = udtEntry
			'	Next
			'End If


		Catch ex As Exception
			Throw New Exception("Error reading index records: " & ex.Message, ex)
			Return False
		End Try

		Return True

	End Function

	''' <summary>
	''' Reads the 10 normalization records from the data file
	''' </summary>
	''' <param name="fsDatafile"></param>
	''' <returns>True if success, false if an error</returns>
	Protected Function ReadNormalizationRecords(ByRef fsDatafile As System.IO.FileStream) As Boolean
		Dim bc As New clsByteConverter()

		Try
			If Header.NormalizationRecordsOffset > 0 Then

				' Move the filestream to the correct byte offset
				fsDatafile.Seek(Header.NormalizationRecordsOffset, IO.SeekOrigin.Begin)

				For intIndex As Integer = 0 To 9
					Dim udtRecord As udtNormalizationRecordType

					udtRecord.Mass = bc.ReadSingleSwapBytes(fsDatafile)
					udtRecord.Slope = bc.ReadSingleSwapBytes(fsDatafile)
					udtRecord.Intercept = bc.ReadSingleSwapBytes(fsDatafile)

					mNormalizationRecordList.Add(udtRecord)

				Next

			End If

		Catch ex As Exception
			Throw New Exception("Error reading normalization records: " & ex.Message, ex)
			Return False
		End Try

		Return True

	End Function

#Region "IDisposable Support"
	Private disposedValue As Boolean ' To detect redundant calls

	' IDisposable
	Protected Overridable Sub Dispose(disposing As Boolean)
		If Not Me.disposedValue Then
			If disposing Then
				If Not mFileStream Is Nothing Then
					mFileStream.Close()
				End If
			End If

			If Not mNormalizationRecordList Is Nothing Then
				mNormalizationRecordList.Clear()
			End If

			If Not mIndexList Is Nothing Then
				mIndexList.Clear()
			End If
		End If
		Me.disposedValue = True
	End Sub

	Public Sub Dispose() Implements IDisposable.Dispose
		' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
		Dispose(True)
		GC.SuppressFinalize(Me)
	End Sub
#End Region

End Class
