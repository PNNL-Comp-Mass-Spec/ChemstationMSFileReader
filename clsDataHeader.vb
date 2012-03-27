Option Strict On

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