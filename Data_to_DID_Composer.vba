Option Explicit

Sub DIDcompMacro()

    Dim HeadersRange As Range
    Dim NameRange As Range
    Dim SizeRange As Range
    Dim ByteRange As Range
    Dim BitRange As Range
    Dim LengthRange As Range
    Dim ByteStart As Integer
    Dim bitOff As Integer
    Dim length As Integer
    Dim i As Integer
    Dim Cell As Range
    Dim bits As Integer
    Dim param As Integer

    Set HeadersRange = Range("HeaderDIDcomp", Range("HeaderDIDcomp").End(xlToRight).Address)
    HeadersRange.Select

    'Set NameRange = Range(HeadersRange.Find("Name (optional)").Address, HeadersRange.Find("Name (optional)").End(xlDown))
    Set SizeRange = Range(HeadersRange.Find("Size").Address, HeadersRange.Find("Size").End(xlDown))
    Set ByteRange = Range(HeadersRange.Find("Byte Start").Address)
    Set BitRange = Range(HeadersRange.Find("Bit Offset").Address)
    Set LengthRange = Range(HeadersRange.Find("Length").Address)

    bits = 0
    i = 1

    For Each Cell In SizeRange.Cells

        If i = 1 Then
            'i = 1 is the header itself, i.e. "Name". Better to do like this using Range and each cell in range. Can find a better solution
        ElseIf i = 2 Then
            ByteRange.Cells(i, 1) = 1
            BitRange.Cells(i, 1) = 0
            bits = SizeRange.Cells(i, 1)
        Else
            ByteRange.Cells(i, 1) = bits \ 8 + 1
            BitRange.Cells(i, 1) = bits Mod 8
            bits = bits + SizeRange.Cells(i, 1)

        End If

        ByteRange.Cells(i, 1).Interior.Color = RGB(255, 255, 0)
        ByteRange.Cells(i, 1).HorizontalAlignment = xlCenter
        BitRange.Cells(i, 1).Interior.Color = RGB(255, 255, 0)
        BitRange.Cells(i, 1).HorizontalAlignment = xlCenter

        i = i + 1

    Next Cell

        param = 2
        length = bits \ 8
        If bits Mod 8 Then
            length = length + 1
        End If

        For param = 2 To (i - 1)
            LengthRange.Cells(param, 1) = length
            LengthRange.Cells(param, 1).Interior.Color = RGB(255, 255, 0)
            LengthRange.Cells(param, 1).HorizontalAlignment = xlCenter
        Next param

End Sub
