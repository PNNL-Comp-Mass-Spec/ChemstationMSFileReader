Option Strict On

Friend Class clsByteConverter

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
