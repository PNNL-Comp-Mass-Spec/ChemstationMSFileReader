Option Strict On

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

		Dim intMass20x() As Int32		' Stored as a UInt16 number
		Dim intAbundance() As Int32

		Dim intIndexCurrent As Integer

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

			mBasePeakAbundance = ReadPackedAbundance(fsDatafile)						' Stored as a packed Int16; we unpack using ReadPackedAbundance

			ReDim intMass20x(mNumberOfPeaks - 1)
			ReDim intAbundance(mNumberOfPeaks - 1)

			For intIndex As Integer = 0 To mNumberOfPeaks - 1
				intIndexCurrent = intIndex

				intMass20x(intIndex) = bc.ReadUInt16SwapBytes(fsDatafile)
				intAbundance(intIndex) = ReadPackedAbundance(fsDatafile)				' Stored as a packed Int16; we unpack using ReadPackedAbundance

				mTIC += intAbundance(intIndex)
			Next

			' Data is typically sorted by descending abundance
			' Re-sort by ascending m/z, then store in the generic list objects

			Array.Sort(intMass20x, intAbundance)

			For intIndex As Integer = 0 To mNumberOfPeaks - 1
				mMzs.Add(intMass20x(intIndex) / 20.0!)
				mIntensites.Add(intAbundance(intIndex))
			Next

			If mNumberOfPeaks > 0 Then
				mMzMin = mMzs(0)
				mMzMax = mMzs(mMzs.Count - 1)
			Else
				mMzMin = 0
				mMzMax = 0
			End If

		Catch ex As Exception
			Throw New Exception("Error reading spectrum, index=" & intIndexCurrent & ": " & ex.Message, ex)
			mValid = False
		End Try

		mValid = True
	End Sub

	''' <summary>
	''' Read packed abundance stored as 2-bit scale with 14 bit mantissa
	''' </summary>
	''' <param name="fs">FileStream object</param>
	''' <returns>Unpacked abundance</returns>
	Private Function ReadPackedAbundance(ByRef fs As System.IO.FileStream) As Int32

		Dim byteArray As Byte()
		Dim byteArrayRev As Byte()

		ReDim byteArray(1)
		ReDim byteArrayRev(1)
		byteArray(0) = CByte(fs.ReadByte())
		byteArray(1) = CByte(fs.ReadByte())

		byteArrayRev(0) = byteArray(1)
		byteArrayRev(1) = byteArray(0)

		Dim intAbundanceScale As Integer
		Dim intAbundanceMantissa As Int32

		Try
			' The abundance scale is stored in the first 2 bits of the first byte in byteArray()
			' Apply a bitmask of 0000 0011 to extract the value
			intAbundanceScale = byteArray(0) And 3

			Dim intAbundancePacked As UInt16
			intAbundancePacked = BitConverter.ToUInt16(byteArrayRev, 0)

			' Shift off the first 2 bits then obtain the bytes
			Dim bytesMantissa() As Byte
			bytesMantissa = BitConverter.GetBytes(intAbundancePacked >> 2)

			' Convert the newly obtained bytes back to a UInt16 number
			intAbundanceMantissa = BitConverter.ToUInt16(bytesMantissa, 0)

			' Scale the abundance by powers of 8 (if intAbundanceScale > 0)
			For intIndex As Integer = 1 To intAbundanceScale
				intAbundanceMantissa *= 8
			Next

			Return intAbundanceMantissa

		Catch ex As Exception
			Return 0
		End Try

		Return 0

	End Function

End Class
