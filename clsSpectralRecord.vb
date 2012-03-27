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
