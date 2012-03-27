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
		Public TotalSignalRaw As Int32
		Public TotalSignalScaled As Double			' Total Signal re-scaled to the range defined by Header.SignalMinimum to Header.SignalMaximum (via a log transformation)
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

	''' <summary>
	'''  Returns the mass spectrum at the specified index
	''' </summary>
	''' <param name="intSpectrumIndex">0-based spectrum index</param>
	''' <param name="oSpectrum">Spectrum object (output)</param>
	''' <returns>True if success, false if an error</returns>
	Public Function GetSpectrum(ByVal intSpectrumIndex As Integer, ByRef oSpectrum As clsSpectralRecord) As Boolean

		Dim sngRetentionTimeMinutes As Single
		Dim intTotalSignalRaw As Int32
		Dim dblTotalSignalScaled As Double

		sngRetentionTimeMinutes = 0
		intTotalSignalRaw = 0
		dblTotalSignalScaled = 0

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
			intTotalSignalRaw = mIndexList(intSpectrumIndex).TotalSignalRaw
			dblTotalSignalScaled = mIndexList(intSpectrumIndex).TotalSignalScaled

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
		Const LOG_BASE As Integer = 8

		Dim bc As New clsByteConverter()

		Dim intTotalSignalMin As Integer = Integer.MaxValue
		Dim intTotalSignalMax As Integer = 0

		Try
			' Move the filestream to the correct byte offset
			fsDatafile.Seek(Header.DirectoryOffset, IO.SeekOrigin.Begin)

			For intIndex As Integer = 0 To Header.SpectraCount - 1
				Dim udtEntry As udtIndexEntryType

				udtEntry.OffsetBytes = bc.WordOffsetToBytes(bc.ReadInt32SwapBytes(fsDatafile) - 1)
				udtEntry.RetentionTimeMsec = bc.ReadInt32SwapBytes(fsDatafile)
				udtEntry.TotalSignalRaw = bc.ReadInt32SwapBytes(fsDatafile)

				If udtEntry.TotalSignalRaw > intTotalSignalMax Then intTotalSignalMax = udtEntry.TotalSignalRaw
				If udtEntry.TotalSignalRaw < intTotalSignalMin Then intTotalSignalMin = udtEntry.TotalSignalRaw

				mIndexList.Add(udtEntry)

			Next

			If mIndexList.Count > 0 Then
				' Now compute TotalSignalScaled for each entry in mIndexList
				Dim dblScaledSignal As Double
				Dim dblTotalSignalMinLog As Double = Math.Log(intTotalSignalMin, LOG_BASE)
				Dim dblTotalSignalMaxLog As Double = Math.Log(intTotalSignalMax, LOG_BASE)

				For intIndex As Integer = 0 To mIndexList.Count - 1
					Dim udtEntry As udtIndexEntryType

					udtEntry = mIndexList(intIndex)

					' Compute the log of the number
					dblScaledSignal = Math.Log(udtEntry.TotalSignalRaw, LOG_BASE)

					' Scale to a value between 0 and 1
					dblScaledSignal = (dblScaledSignal - dblTotalSignalMinLog) / (dblTotalSignalMaxLog - dblTotalSignalMinLog)

					' Scale to the range Header.SignalMinimum to Header.SignalMaximum
					udtEntry.TotalSignalScaled = dblScaledSignal * (Header.SignalMaximum - Header.SignalMinimum) + Header.SignalMinimum

					mIndexList(intIndex) = udtEntry
				Next
			End If


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

#Region "Data Header Class"

	Public Class clsDataHeader

		' The data header in the Data.MS file is composed of the following entries
		'
		Protected mFileNumber As String = String.Empty		 '  3-byte string
		Protected mFileString As String = String.Empty		 ' 19 bytes
		Protected mDataName As String = String.Empty		 ' 61 bytes
		Protected mMiscInfo As String = String.Empty		 ' 61 bytes
		Protected mOperatorName As String = String.Empty	 ' 29 bytes
		Protected mAcqDate As String = String.Empty			 ' 29 bytes
		Protected mInstrumentModel As String = String.Empty	 '  9 bytes
		Protected mInlet As String = String.Empty			 '  9 bytes
		Protected mMethodFile As String = String.Empty		 ' 19 bytes
		Protected mFileType As Int32
		Protected mSeqIndex As Int16
		Protected mALSBottle As Int16
		Protected mReplicate As Int16
		Protected mDirectoryEntryType As Int16
		Protected mDirectoryOffset As Int32					 ' Stored as integer (in words); converted to Byte offset by this class
		Protected mDataOffset As Int32						 ' Stored as integer (in words); converted to Byte offset by this class
		Protected mRunTableOffset As Int32					 ' Stored as integer (in words), unused; converted to Byte offset by this class
		Protected mNormalizationRecordsOffset As Int32		 ' Stored as integer (in words); converted to Byte offset by this class
		Protected mExtraRecords As Int16
		Protected mDataRecordCount As Int32
		Protected mRetentionTimeMsecStart As Int32
		Protected mRetentionTimeMsecEnd As Int32
		Protected mSignalMaximum As Int32
		Protected mSignalMinimum As Int32

		Protected mValid As Boolean
		Protected mDatasetName As String = String.Empty

#Region "Properties"

		Public ReadOnly Property AcqDate() As DateTime
			Get
				Dim dtDate As System.DateTime
				If DateTime.TryParse(mAcqDate, dtDate) Then
					Return dtDate
				Else
					Return System.DateTime.MinValue
				End If
			End Get
		End Property

		Public ReadOnly Property AcqDateText() As String
			Get
				Return mAcqDate
			End Get
		End Property

		Public ReadOnly Property ALSBottle() As Int16
			Get
				Return mALSBottle
			End Get
		End Property

		Public ReadOnly Property DataOffset() As Int32
			Get
				Return mDataOffset
			End Get
		End Property

		Public ReadOnly Property DatasetName() As String
			Get
				Return mDatasetName
			End Get
		End Property

		Public ReadOnly Property Description() As String
			Get
				Return mDataName
			End Get
		End Property

		Public ReadOnly Property DirectoryEntryType() As Int16
			Get
				Return mDirectoryEntryType
			End Get
		End Property

		Public ReadOnly Property DirectoryOffset() As Int32
			Get
				Return mDirectoryOffset
			End Get
		End Property

		Public ReadOnly Property FileNumber() As String
			Get
				Return mFileNumber
			End Get
		End Property

		Public ReadOnly Property FileType() As Int32
			Get
				Return mFileType
			End Get
		End Property

		Public ReadOnly Property FileTypeName() As String
			Get
				Return mFileString
			End Get
		End Property

		Public ReadOnly Property Inlet() As String
			Get
				Return mInlet
			End Get
		End Property

		Public ReadOnly Property InstrumentModel() As String
			Get
				Return mInstrumentModel
			End Get
		End Property

		Public ReadOnly Property MethodFile() As String
			Get
				Return mMethodFile
			End Get
		End Property

		Public ReadOnly Property MiscInfo() As String
			Get
				Return mMiscInfo
			End Get
		End Property

		Public ReadOnly Property NormalizationRecordsOffset() As Int32
			Get
				Return mNormalizationRecordsOffset
			End Get
		End Property

		Public ReadOnly Property OperatorName() As String
			Get
				Return mOperatorName
			End Get
		End Property

		Public ReadOnly Property Replicate() As Int16
			Get
				Return mReplicate
			End Get
		End Property

		Public ReadOnly Property RetentionTimeMsecEnd() As Int32
			Get
				Return mRetentionTimeMsecEnd
			End Get
		End Property

		Public ReadOnly Property RetentionTimeMsecStart() As Int32
			Get
				Return mRetentionTimeMsecStart
			End Get
		End Property

		Public ReadOnly Property RetentionTimeMinutesStart() As Single
			Get
				Return mRetentionTimeMsecStart / 60000.0!
			End Get
		End Property

		Public ReadOnly Property RetentionTimeMinutesEnd As Single
			Get
				Return mRetentionTimeMsecEnd / 60000.0!
			End Get
		End Property

		Public ReadOnly Property RunTableOffset() As Int32
			Get
				Return mRunTableOffset
			End Get
		End Property

		Public ReadOnly Property SeqIndex() As Int16
			Get
				Return mSeqIndex
			End Get
		End Property

		Public ReadOnly Property SignalMaximum() As Int32
			Get
				Return mSignalMaximum
			End Get
		End Property

		Public ReadOnly Property SignalMinimum() As Int32
			Get
				Return mSignalMinimum
			End Get
		End Property

		Public ReadOnly Property SpectraCount() As Int32
			Get
				Return mDataRecordCount
			End Get
		End Property

		Public ReadOnly Property Valid As Boolean
			Get
				Return mValid
			End Get
		End Property

#End Region

		''' <summary>
		''' Read header from the specified file
		''' </summary>
		''' <param name="fsDatafile"></param>
		''' <remarks></remarks>
		Public Sub New(ByRef fsDatafile As System.IO.FileStream)
			ReadFromFile(fsDatafile)
		End Sub

		''' <summary>
		''' Read header from the specified file
		''' </summary>
		''' <param name="fs">Input file stream</param>
		''' <returns>True if success, false if an error</returns>
		Protected Function ReadFromFile(ByRef fs As System.IO.FileStream) As Boolean

			Dim bc As New clsByteConverter()

			Try

				mDatasetName = System.IO.Path.GetFileName(fs.Name)
				If mDatasetName.ToLower() = "data.ms" Then
					' Use the folder name as the dataset name
					Dim fiFile As System.IO.FileInfo
					fiFile = New System.IO.FileInfo(fs.Name)
					mDatasetName = fiFile.Directory.Name
					If mDatasetName.ToUpper().EndsWith(".D") Then
						mDatasetName = mDatasetName.Substring(0, mDatasetName.Length - 2)
					End If
				End If

				' Must skip the first byte to properly read the data
				fs.Seek(1, IO.SeekOrigin.Begin)

				mFileNumber = bc.ReadString(fs, 3, True)			'  3-byte string

				mFileString = bc.ReadString(fs, 19, True)			' 19 bytes

				mDataName = bc.ReadString(fs, 61, True).Trim()		' 61 bytes

				mMiscInfo = bc.ReadString(fs, 61, True).Trim()		' 61 bytes

				mOperatorName = bc.ReadString(fs, 29, True)			' 29 bytes

				mAcqDate = bc.ReadString(fs, 29, True)				' 29 bytes

				mInstrumentModel = bc.ReadString(fs, 9, True)		'  9 bytes

				mInlet = bc.ReadString(fs, 9, True)					'  9 bytes

				mMethodFile = bc.ReadString(fs, 19, False).Trim()	' 19 bytes

				mFileType = bc.ReadInt32SwapBytes(fs)
				mSeqIndex = bc.ReadInt16SwapBytes(fs)
				mALSBottle = bc.ReadInt16SwapBytes(fs)
				mReplicate = bc.ReadInt16SwapBytes(fs)
				mDirectoryEntryType = bc.ReadInt16SwapBytes(fs)

				mDirectoryOffset = bc.WordOffsetToBytes(bc.ReadInt32SwapBytes(fs) - 1)

				mDataOffset = bc.WordOffsetToBytes(bc.ReadInt32SwapBytes(fs) - 1)

				mRunTableOffset = bc.WordOffsetToBytes(bc.ReadInt32SwapBytes(fs) - 1)

				mNormalizationRecordsOffset = bc.WordOffsetToBytes(bc.ReadInt32SwapBytes(fs) - 1)

				mExtraRecords = bc.ReadInt16SwapBytes(fs)
				mDataRecordCount = bc.ReadInt32SwapBytes(fs)			' Number of spectra

				mRetentionTimeMsecStart = bc.ReadInt32SwapBytes(fs)
				mRetentionTimeMsecEnd = bc.ReadInt32SwapBytes(fs)

				mSignalMaximum = bc.ReadInt32SwapBytes(fs)
				mSignalMinimum = bc.ReadInt32SwapBytes(fs)

			Catch ex As Exception
				Throw New Exception("Error reading header: " & ex.Message, ex)
				mValid = False
			End Try

			mValid = True

			Return True

		End Function

	End Class
#End Region

#Region "Byte Converter Class"

	Protected Class clsByteConverter

		''' <summary>
		''' Reads an Int16 from the filestream
		''' Swaps bytes prior to converting
		''' </summary>
		''' <param name="fs">FileStream object</param>
		''' <returns>The value, as an Int16 number</returns>
		Public Function ReadInt16SwapBytes(ByRef fs As System.IO.FileStream) As Int16

			Dim byteArray As Byte()
			ReDim byteArray(1)
			byteArray(1) = CByte(fs.ReadByte())
			byteArray(0) = CByte(fs.ReadByte())

			Return BitConverter.ToInt16(byteArray, 0)

		End Function

		''' <summary>
		''' Reads an UInt16 from the filestream
		''' Swaps bytes prior to converting
		''' </summary>
		''' <param name="fs">FileStream object</param>
		''' <returns>The value, as an Int32 number</returns>
		Public Function ReadUInt16SwapBytes(ByRef fs As System.IO.FileStream) As Int32

			Dim byteArray As Byte()
			ReDim byteArray(3)
			byteArray(3) = CByte(fs.ReadByte())
			byteArray(2) = CByte(fs.ReadByte())
			byteArray(1) = CByte(fs.ReadByte())
			byteArray(0) = CByte(fs.ReadByte())

			Return BitConverter.ToUInt16(byteArray, 0)

		End Function

		''' <summary>
		''' Reads an Int32 from the filestream
		''' Swaps bytes prior to converting
		''' </summary>
		''' <param name="fs">FileStream object</param>
		''' <returns>The value, as an Int32 number</returns>
		Public Function ReadInt32SwapBytes(ByRef fs As System.IO.FileStream) As Int32

			Dim byteArray As Byte()
			ReDim byteArray(3)
			byteArray(3) = CByte(fs.ReadByte())
			byteArray(2) = CByte(fs.ReadByte())
			byteArray(1) = CByte(fs.ReadByte())
			byteArray(0) = CByte(fs.ReadByte())

			Return BitConverter.ToInt32(byteArray, 0)

		End Function

		''' <summary>
		''' Reads a 4-byte single (real) from the filestream
		''' Swaps bytes prior to converting
		''' </summary>
		''' <param name="fs">FileStream object</param>
		''' <returns>The value, as a single-precision number</returns>
		Public Function ReadSingleSwapBytes(ByRef fs As System.IO.FileStream) As Single

			Dim byteArray As Byte()
			ReDim byteArray(3)
			byteArray(3) = CByte(fs.ReadByte())
			byteArray(2) = CByte(fs.ReadByte())
			byteArray(1) = CByte(fs.ReadByte())
			byteArray(0) = CByte(fs.ReadByte())

			Return BitConverter.ToSingle(byteArray, 0)

		End Function

		''' <summary>
		''' Reads a fixed-length string from the filestream
		''' Optionally advances the reader one byte after reading the string
		''' </summary>
		''' <param name="fs">FileStream object</param>
		''' <param name="iStringLength">String length</param>
		''' <param name="bAdvanceExtraByte">If true, then advances the read an extra byte after reading the string</param>
		''' <returns>The string read</returns>
		Public Function ReadString(ByRef fs As System.IO.FileStream, ByVal iStringLength As Integer, ByVal bAdvanceExtraByte As Boolean) As String
			Dim byteArray() As Byte

			ReDim byteArray(iStringLength - 1)
			fs.Read(byteArray, 0, iStringLength)

			If bAdvanceExtraByte Then
				fs.ReadByte()
			End If

			' Remove entries from the end of byteArray that are null
			Dim intIndexEnd As Integer = byteArray.Length - 1
			For intIndex As Integer = byteArray.Length - 1 To 0 Step -1
				If byteArray(intIndex) = 0 Then
					intIndexEnd -= 1
				Else
					Exit For
				End If
			Next

			If intIndexEnd > -1 Then
				Dim sText As String = System.Text.Encoding.ASCII.GetString(byteArray, 0, intIndexEnd + 1)
				Return sText
			Else
				Return String.Empty
			End If

		End Function

		''' <summary>
		''' Convert byte-offset stored in words into bytes
		''' </summary>
		''' <param name="iOffsetWords"></param>
		''' <returns>The byte offset, in bytes</returns>		
		Public Function WordOffsetToBytes(ByVal iOffsetWords As Integer) As Integer
			If iOffsetWords > 0 Then
				Return (iOffsetWords * 2)
			Else
				Return 0
			End If
		End Function
	End Class
#End Region

#Region "Spectral Record Class"

	Public Class clsSpectralRecord

		' Each spectral record is composed of the following entries, followed by a list of mass and abundance values
		'
		Protected mNumberOfWords As Int16
		Protected mRetentionTimeMsec As Int32
		Protected mNumberOfWordsLess3 As Int16
		Protected mDataType As Int16
		Protected mStatusWord As Int16
		Protected mNumberOfPeaks As Int16
		Protected mBasePeak20x As Int32			' Stores Mass * 20.  Stored as UInt16, but using Int32 to avoid a "Not CLS-Compliant" warning
		Protected mBasePeakAbundance As Int32

		Protected mMzMin As Single
		Protected mMzMax As Single
		Protected mTIC As Double

		Protected mValid As Boolean

		' The data is stored in the Data.ms file as Mass and abundance pairs
		' Mass is represented by UInt16 and stores the mass value times 20
		' Intensity is represented by a packed Int16 value

		Protected mMzs As System.Collections.Generic.List(Of Single)
		Protected mIntensites As System.Collections.Generic.List(Of Int32)

#Region "Properties"

		Public ReadOnly Property BasePeakAbundance As Int32
			Get
				Return mBasePeakAbundance
			End Get
		End Property

		Public ReadOnly Property BasePeakMZ As Single
			Get
				Return mBasePeak20x / 20.0!
			End Get
		End Property

		Public ReadOnly Property Count As Int16
			Get
				Return mNumberOfPeaks
			End Get
		End Property

		Public ReadOnly Property DataType As Int16
			Get
				Return mDataType
			End Get
		End Property

		Public ReadOnly Property Intensities() As System.Collections.Generic.List(Of Int32)
			Get
				Return mIntensites
			End Get
		End Property

		Public ReadOnly Property Mzs() As System.Collections.Generic.List(Of Single)
			Get
				Return mMzs
			End Get
		End Property

		Public ReadOnly Property RetentionTimeMsec As Int32
			Get
				Return mRetentionTimeMsec
			End Get
		End Property

		Public ReadOnly Property RetentionTimeMinutes As Single
			Get
				Return mRetentionTimeMsec / 60000.0!
			End Get
		End Property

		Public ReadOnly Property StatusWord As Int16
			Get
				Return mStatusWord
			End Get
		End Property

		Public ReadOnly Property TIC As Double
			Get
				Return mTIC
			End Get
		End Property

		Public ReadOnly Property Valid As Boolean
			Get
				Return mValid
			End Get
		End Property

#End Region

		''' <summary>
		''' Instantiate a new spectrum object
		''' </summary>
		''' <remarks></remarks>
		Public Sub New()
			Me.Clear()
		End Sub

		''' <summary>
		''' Populate a spectrum object with the data at the specified byte offset
		''' </summary>
		Public Sub New(ByRef fsDatafile As System.IO.FileStream, ByVal intByteOffsetStart As Integer)
			ReadFromFile(fsDatafile, intByteOffsetStart)
		End Sub

		''' <summary>
		''' Initialize the variables and data structures
		''' </summary>
		''' <remarks></remarks>
		Protected Sub Clear()

			If mMzs Is Nothing Then
				mMzs = New System.Collections.Generic.List(Of Single)
			Else
				mMzs.Clear()
			End If

			If mIntensites Is Nothing Then
				mIntensites = New System.Collections.Generic.List(Of Int32)
			Else
				mIntensites.Clear()
			End If

			mNumberOfPeaks = 0
			mMzMin = 0
			mMzMax = 0
			mTIC = 0
		End Sub

		''' <summary>
		''' Read the spectrum at the specified byte offset
		''' </summary>
		''' <param name="fsDatafile"></param>
		''' <param name="intByteOffsetStart"></param>
		''' <remarks></remarks>
		Protected Sub ReadFromFile(ByRef fsDatafile As System.IO.FileStream, ByVal intByteOffsetStart As Integer)

			Dim bc As New clsByteConverter()

			Dim intMass20x() As Int16
			Dim intAbundance() As Int32

			Dim intPackedAbundance As Int16

			Try
				Me.Clear()

				' Move the filestream to the correct byte offset
				fsDatafile.Seek(intByteOffsetStart, IO.SeekOrigin.Begin)

				mNumberOfWords = bc.ReadInt16SwapBytes(fsDatafile)				' Total bytes stored in this spectral record
				mRetentionTimeMsec = bc.ReadInt32SwapBytes(fsDatafile)
				mNumberOfWordsLess3 = bc.ReadInt16SwapBytes(fsDatafile)

				mDataType = bc.ReadInt16SwapBytes(fsDatafile)
				mStatusWord = bc.ReadInt16SwapBytes(fsDatafile)
				mNumberOfPeaks = bc.ReadInt16SwapBytes(fsDatafile)
				mBasePeak20x = bc.ReadUInt16SwapBytes(fsDatafile)		   ' Note: Stored as UInt16; stores Mass * 20

				mBasePeakAbundance = UnpackAbundance(bc.ReadInt16SwapBytes(fsDatafile))		' Stored as a packed Int16; we unpack using UnpackAbundance

				ReDim intMass20x(mNumberOfPeaks - 1)
				ReDim intAbundance(mNumberOfPeaks - 1)

				For intIndex As Integer = 0 To mNumberOfPeaks - 1
					intMass20x(intIndex) = bc.ReadInt16SwapBytes(fsDatafile)
					intPackedAbundance = bc.ReadInt16SwapBytes(fsDatafile)					' Stored as a packed Int16; we unpack using UnpackAbundance

					intAbundance(intIndex) = UnpackAbundance(intPackedAbundance)
					mTIC += intAbundance(intIndex)
				Next

				' Data is typically sorted by descending abundance
				' Re-sort by ascending m/z, then store in the generic list objects

				Array.Sort(intMass20x, intAbundance)

				For intIndex As Integer = 0 To mNumberOfPeaks - 1
					mMzs.Add(intMass20x(intIndex) / 20.0!)
					mIntensites.Add(intAbundance(intIndex))
				Next

				mMzMin = mMzs(0)
				mMzMax = mMzs(mMzs.Count - 1)

			Catch ex As Exception
				Throw New Exception("Error reading spectrum: " & ex.Message, ex)
				mValid = False
			End Try

			mValid = True
		End Sub

		''' <summary>
		''' Unpack abundance stored as 4-bit scale with 12 bit mantissa
		''' </summary>
		''' <param name="intAbundancePacked">Packed abundance</param>
		''' <returns>Unpacked abundance</returns>
		Private Function UnpackAbundance(ByVal intAbundancePacked As Int16) As Int32

			Dim intAbundanceScale As Byte
			Dim intAbundanceMantissa As Int32

			Dim intScaleMask As UInt16 = 61440			' 1111 0000 0000 0000
			Dim intMantissaMask As UInt16 = 4095		' 0000 1111 1111 1111

			' Abundance is packed in powers of 8
			' The first 4 bits of intAbundancePacked represent the Scale and will be 0, 1, 2, or 3 (x1, x8, x64, or x512)
			' The remaining 12 bits of intAbundancePacked are the Mantissa, ranging from 0 to 16383

			' Extract the first 4 bits by applying bitmask intScaleMask, then bit shifting 12 bits to the right
			intAbundanceScale = CByte((intAbundancePacked And intScaleMask) >> 12)

			' Extract the mantissa by applying bitmask intMantissaMask; no need to bit shift
			intAbundanceMantissa = intAbundancePacked And intMantissaMask

			' Scale the abundance by powers of 8 (if intAbundanceScale > 0)
			For intIndex As Integer = 1 To intAbundanceScale
				intAbundanceMantissa *= 8
			Next

			Return intAbundanceMantissa

			' The following code shows how we can use a BitArray object to reverse the bits
			' This was explored to confirm that the bit order in intAbundancePacked is correct

			'Dim bits As BitArray
			'Dim bitsReversed As BitArray

			'bits = New BitArray(BitConverter.GetBytes(intAbundancePacked))
			'bitsReversed = New BitArray(bits.Length)

			'For intBitIndex As Integer = 0 To bits.Length - 1
			'	bitsReversed(intBitIndex) = bits(bits.Length - intBitIndex - 1)
			'Next

			'Dim newBytes() As Byte
			'ReDim newBytes(CInt(bitsReversed.Length / 8))
			'bitsReversed.CopyTo(newBytes, 0)

			'Dim intAbundancePackedAlt As Int16
			'Dim intAbundanceMantissaAlt As Int32
			'intAbundancePackedAlt = BitConverter.ToInt16(newBytes, 0)

			'Dim intMantissaMaskAlt As UInt16 = 65520	' 1111 1111 1111 0000
			'intAbundanceMantissaAlt = (intAbundancePackedAlt And intMantissaMaskAlt) >> 4

			'' Scale the abundance by powers of 8
			'For intIndex As Integer = 1 To intAbundanceScale
			'	intAbundanceMantissaAlt *= 8
			'Next


		End Function
	End Class
#End Region

End Class
