'
' ===============================================
' Code 128(B|C) Generator
'
' Copyright 2020 Adam Fiser | Wanex.cz
' All rights are reserved
' ===============================================

Public Function ContentStringGenerator(content As String) As String
    ' Supports B and C charsets only; values 00-94, 99,101, 103-105 for B, 00-101, 103-105 for C
    
    Dim WeightSum As Single
    Const XmmTopt As Single = 0.351
    Const YmmTopt As Single = 0.351
    Const XCompRatio As Single = 0.9


    Const Tbar_Symbol As String * 2 = "11"
    Dim CurBar As Integer
    Dim i, j, k, CharIndex, SymbolIndex As Integer
    Dim tstr2 As String * 2
    Dim tstr1 As String * 1
    Dim ContentString As String ' bars sequence
    Const Asw As String * 1 = "A" ' alpha switch
    Const Dsw As String * 1 = "D" 'digital switch
    Const Arrdim As Byte = 30


    Dim Sw, PrevSw As String * 1  ' switch
    Dim BlockIndex, BlockCount, DBlockMod2, DBlockLen As Byte


    Dim BlockLen(Arrdim) As Byte
    Dim BlockSw(Arrdim) As String * 1


    Dim SymbolValue(0 To 106) As Integer ' values
    Dim SymbolString(0 To 106) As String * 11 'bits sequence
    Dim SymbolCharB(0 To 106) As String * 1  'Chars in B set
    Dim SymbolCharC(0 To 106) As String * 2  'Chars in B set


    For i = 0 To 106 ' values
        SymbolValue(i) = i
    Next i


    ' Symbols in charset B
    For i = 0 To 94
        SymbolCharB(i) = Chr(i + 32)
    Next i


    ' Symbols in charset C
    SymbolCharC(0) = "00"
    SymbolCharC(1) = "01"
    SymbolCharC(2) = "02"
    SymbolCharC(3) = "03"
    SymbolCharC(4) = "04"
    SymbolCharC(5) = "05"
    SymbolCharC(6) = "06"
    SymbolCharC(7) = "07"
    SymbolCharC(8) = "08"
    SymbolCharC(9) = "09"
    For i = 10 To 99
        SymbolCharC(i) = CStr(i)
    Next i


    ' bit sequences
    SymbolString(0) = "11011001100"
    SymbolString(1) = "11001101100"
    SymbolString(2) = "11001100110"
    SymbolString(3) = "10010011000"
    SymbolString(4) = "10010001100"
    SymbolString(5) = "10001001100"
    SymbolString(6) = "10011001000"
    SymbolString(7) = "10011000100"
    SymbolString(8) = "10001100100"
    SymbolString(9) = "11001001000"
    SymbolString(10) = "11001000100"
    SymbolString(11) = "11000100100"
    SymbolString(12) = "10110011100"
    SymbolString(13) = "10011011100"
    SymbolString(14) = "10011001110"
    SymbolString(15) = "10111001100"
    SymbolString(16) = "10011101100"
    SymbolString(17) = "10011100110"
    SymbolString(18) = "11001110010"
    SymbolString(19) = "11001011100"
    SymbolString(20) = "11001001110"
    SymbolString(21) = "11011100100"
    SymbolString(22) = "11001110100"
    SymbolString(23) = "11101101110"
    SymbolString(24) = "11101001100"
    SymbolString(25) = "11100101100"
    SymbolString(26) = "11100100110"
    SymbolString(27) = "11101100100"
    SymbolString(28) = "11100110100"
    SymbolString(29) = "11100110010"
    SymbolString(30) = "11011011000"
    SymbolString(31) = "11011000110"
    SymbolString(32) = "11000110110"
    SymbolString(33) = "10100011000"
    SymbolString(34) = "10001011000"
    SymbolString(35) = "10001000110"
    SymbolString(36) = "10110001000"
    SymbolString(37) = "10001101000"
    SymbolString(38) = "10001100010"
    SymbolString(39) = "11010001000"
    SymbolString(40) = "11000101000"
    SymbolString(41) = "11000100010"
    SymbolString(42) = "10110111000"
    SymbolString(43) = "10110001110"
    SymbolString(44) = "10001101110"
    SymbolString(45) = "10111011000"
    SymbolString(46) = "10111000110"
    SymbolString(47) = "10001110110"
    SymbolString(48) = "11101110110"
    SymbolString(49) = "11010001110"
    SymbolString(50) = "11000101110"
    SymbolString(51) = "11011101000"
    SymbolString(52) = "11011100010"
    SymbolString(53) = "11011101110"
    SymbolString(54) = "11101011000"
    SymbolString(55) = "11101000110"
    SymbolString(56) = "11100010110"
    SymbolString(57) = "11101101000"
    SymbolString(58) = "11101100010"
    SymbolString(59) = "11100011010"
    SymbolString(60) = "11101111010"
    SymbolString(61) = "11001000010"
    SymbolString(62) = "11110001010"
    SymbolString(63) = "10100110000"
    SymbolString(64) = "10100001100"
    SymbolString(65) = "10010110000"
    SymbolString(66) = "10010000110"
    SymbolString(67) = "10000101100"
    SymbolString(68) = "10000100110"
    SymbolString(69) = "10110010000"
    SymbolString(70) = "10110000100"
    SymbolString(71) = "10011010000"
    SymbolString(72) = "10011000010"
    SymbolString(73) = "10000110100"
    SymbolString(74) = "10000110010"
    SymbolString(75) = "11000010010"
    SymbolString(76) = "11001010000"
    SymbolString(77) = "11110111010"
    SymbolString(78) = "11000010100"
    SymbolString(79) = "10001111010"
    SymbolString(80) = "10100111100"
    SymbolString(81) = "10010111100"
    SymbolString(82) = "10010011110"
    SymbolString(83) = "10111100100"
    SymbolString(84) = "10011110100"
    SymbolString(85) = "10011110010"
    SymbolString(86) = "11110100100"
    SymbolString(87) = "11110010100"
    SymbolString(88) = "11110010010"
    SymbolString(89) = "11011011110"
    SymbolString(90) = "11011110110"
    SymbolString(91) = "11110110110"
    SymbolString(92) = "10101111000"
    SymbolString(93) = "10100011110"
    SymbolString(94) = "10001011110"
    SymbolString(95) = "10111101000"
    SymbolString(96) = "10111100010"
    SymbolString(97) = "11110101000"
    SymbolString(98) = "11110100010"
    SymbolString(99) = "10111011110"
    SymbolString(100) = "10111101110"
    SymbolString(101) = "11101011110"
    SymbolString(102) = "11110101110"
    SymbolString(103) = "11010000100"
    SymbolString(104) = "11010010000"
    SymbolString(105) = "11010011100"
    SymbolString(106) = "11000111010"


    X = X / XmmTopt 'mm to pt
    Y = Y / YmmTopt 'mm to pt
    Height = Height / YmmTopt 'mm to pt


    If IsNumeric(content) = True And Len(content) Mod 2 = 0 Then 'numeric, mode C
       WeightSum = SymbolValue(105) ' start-c
       ContentString = ContentString + SymbolString(105)
       i = 0 ' symbol count
       For j = 1 To Len(content) Step 2
          tstr2 = Mid(content, j, 2)
          i = i + 1
          k = 0
          Do While tstr2 <> SymbolCharC(k)
             k = k + 1
          Loop
          WeightSum = WeightSum + i * SymbolValue(k)
          ContentString = ContentString + SymbolString(k)
       Next j
       ContentString = ContentString + SymbolString(SymbolValue(WeightSum Mod 103))
       ContentString = ContentString + SymbolString(106)
       ContentString = ContentString + Tbar_Symbol

    Else ' alpha-numeric

       ' first digit
       Select Case IsNumeric(Mid(content, 1, 1))
       Case Is = True 'digit
          Sw = Dsw
       Case Is = False 'alpha
          Sw = Asw
       End Select
       BlockCount = 1
       BlockSw(BlockCount) = Sw
       BlockIndex = 1
       BlockLen(BlockCount) = 1 'block length



       i = 2 ' symbol index

       Do While i <= Len(content)
          Select Case IsNumeric(Mid(content, i, 1))
          Case Is = True 'digit
             Sw = Dsw
          Case Is = False 'alpha
             Sw = Asw
          End Select

          If Sw = BlockSw(BlockCount) Then
             BlockLen(BlockCount) = BlockLen(BlockCount) + 1
          Else
             BlockCount = BlockCount + 1
             BlockSw(BlockCount) = Sw
             BlockLen(BlockCount) = 1
             BlockIndex = BlockIndex + 1


          End If

          i = i + 1
       Loop



       'encoding
       CharIndex = 1 'index of Content character
       SymbolIndex = 0

       For BlockIndex = 1 To BlockCount ' encoding by blocks


          If BlockSw(BlockIndex) = Dsw And BlockLen(BlockIndex) >= 4 Then ' switch to C
             Select Case BlockIndex
             Case Is = 1
                WeightSum = SymbolValue(105) ' Start-C
                ContentString = ContentString + SymbolString(105)
             Case Else
                SymbolIndex = SymbolIndex + 1
                WeightSum = WeightSum + SymbolIndex * SymbolValue(99) 'switch c
                ContentString = ContentString + SymbolString(99)
             End Select
             PrevSw = Dsw

             ' encoding even amount of chars in a D block
             DBlockMod2 = BlockLen(BlockIndex) Mod 2
             If DBlockMod2 <> 0 Then 'even chars always to encode
                DBlockLen = BlockLen(BlockIndex) - DBlockMod2
             Else
                DBlockLen = BlockLen(BlockIndex)
             End If

             For j = 1 To DBlockLen / 2 Step 1
                tstr2 = Mid(content, CharIndex, 2)
                CharIndex = CharIndex + 2
                SymbolIndex = SymbolIndex + 1
                k = 0
                Do While tstr2 <> SymbolCharC(k)
                   k = k + 1
                Loop
                WeightSum = WeightSum + SymbolIndex * SymbolValue(k)
                ContentString = ContentString + SymbolString(k)
             Next j

             If DBlockMod2 <> 0 Then ' switch to B, encode 1 char
                PrevSw = Asw
                SymbolIndex = SymbolIndex + 1
                WeightSum = WeightSum + SymbolIndex * SymbolValue(100) 'switch b
                ContentString = ContentString + SymbolString(100)

                'CharIndex = CharIndex + 1
                SymbolIndex = SymbolIndex + 1
                tstr1 = Mid(content, CharIndex, 1)
                k = 0
                Do While tstr1 <> SymbolCharB(k)
                   k = k + 1
                Loop
                WeightSum = WeightSum + SymbolIndex * SymbolValue(k)
                ContentString = ContentString + SymbolString(k)
                CharIndex = CharIndex + 1
             End If


          Else 'alpha in B mode
             Select Case BlockIndex
             Case Is = 1
             '   PrevSw = Asw
                WeightSum = SymbolValue(104) ' start-b
                ContentString = ContentString + SymbolString(104)
             Case Else
                If PrevSw <> Asw Then
                   SymbolIndex = SymbolIndex + 1
                   WeightSum = WeightSum + SymbolIndex * SymbolValue(100) 'switch b
                   ContentString = ContentString + SymbolString(100)

                End If
             End Select
             PrevSw = Asw

             For j = CharIndex To CharIndex + BlockLen(BlockIndex) - 1 Step 1
                tstr1 = Mid(content, j, 1)
                SymbolIndex = SymbolIndex + 1
                k = 0
                Do While tstr1 <> SymbolCharB(k)
                   k = k + 1
                Loop
                WeightSum = WeightSum + SymbolIndex * SymbolValue(k)
                ContentString = ContentString + SymbolString(k)
             Next j
             CharIndex = j


          End If
       Next BlockIndex
       ContentString = ContentString + SymbolString(SymbolValue(WeightSum Mod 103))
       ContentString = ContentString + SymbolString(106)
       ContentString = ContentString + Tbar_Symbol

    End If
    
    ContentStringGenerator = ContentString
End Function



'Option Explicit

'you call this, it's the main worker here. content is the barcode data. r is cell address where you want your bar code to appear
Sub mainBarCoder(content As String, ByVal r As Range, Optional ByVal barHeight As Integer = 20, _
                            Optional fontSize As Integer = 7, _
                            Optional sideMargin As Integer = 10)
    Dim data As String
    data = ContentStringGenerator(content)
    
    Dim i As Integer
    Dim k As Integer
    Dim shapeArr() As String 'to store all the names of the new shapes
    Dim sh As Worksheet
    Set sh = r.Worksheet
    
    k = Len(data)
    
    ReDim shapeArr(0 To k + 1)
    
    'creating the white background
    shapeArr(0) = sh.Shapes.AddShape(msoShapeRectangle, _
        r.Left, r.Top, k + (2 * sideMargin), barHeight + fontSize + 10).Name
    
    With sh.Shapes.Range(shapeArr(0))
        .Fill.Visible = msoCTrue
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoFalse
    End With
    
    'actual barcode maker
    Dim l As Integer
    Dim startPos As Integer
    For i = 1 To k
        startPos = i
        l = 1
        Do While i <> k And Mid(data, i, 1) = Mid(data, i + 1, 1) 'checking for continus block
            l = l + 1
            i = i + 1
        Loop
        
        If CInt(Mid(data, i, 1)) Then
            shapeArr(i) = barDrawer(r, startPos, CInt(Mid(data, i, 1)), l, barHeight, sideMargin)
        End If
    Next i
    Dim grp As Variant
    
    shapeArr(UBound(shapeArr)) = textBoxDrawer(content, r, barHeight, k, fontSize)
    
    'grouping all shapes into one unit
    Set grp = ActiveSheet.Shapes.Range(shapeArr).Group
    
    'This renders barcode shapes into a vector image
    grp.Copy
    r.Select
    sh.PasteSpecial Format:="Picture (Enhanced Metafile)", Link:=False, DisplayAsIcon:=False
    grp.Delete
    
End Sub

'This draws the barcode shapes
Private Function barDrawer(r As Range, X As Integer, ch As Integer, blockLength As Integer, _
                            barHeight As Integer, sideMargin As Integer) As String
    Dim sh As Shape
    Set sh = ActiveSheet.Shapes.AddShape(msoShapeRectangle, r.Left + sideMargin + X, r.Top + 3, blockLength, barHeight)
    sh.Fill.Visible = msoCTrue
    sh.Fill.ForeColor.RGB = RGB(0, 0, 0)
    sh.Line.Visible = msoFalse
    sh.Placement = xlMove
    barDrawer = sh.Name
End Function

'this draws text box under the barcode
Private Function textBoxDrawer(content As String, _
                                r As Range, _
                                highFromTop As Integer, _
                                length As Integer, _
                Optional ByVal fontSize As Integer = 7) As String
                
    Dim textBox As Shape
    Set textBox = r.Worksheet.Shapes.AddShape(msoShapeRectangle, _
        r.Left + 10, r.Top + highFromTop + 3, length, fontSize + 10)
        
    With textBox
        With .TextFrame2
            .TextRange.Font.Size = fontSize
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .MarginBottom = 0
            .AutoSize = msoAutoSizeShapeToFitText
            .MarginTop = 0
            
            With .TextRange
                
                With .Font
                    .NameComplexScript = "Arial"
                    .NameFarEast = "Arial"
                    .Name = "Arial"
                    .BaselineOffset = 0
                    .Spacing = 1
                
                    With .Fill
                        .Visible = msoTrue
                        .ForeColor.RGB = RGB(0, 0, 0)
                        .ForeColor.TintAndShade = 0
                        .ForeColor.Brightness = 0
                        .Transparency = 0
                        .Solid
                    End With
                End With
                .Characters.Text = content

            End With
        End With
        
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
    textBoxDrawer = textBox.Name
End Function

'Remove before generated barcodes
Private Sub DeletePic(ByVal r As Range)
    Dim xPicRg As Range
    Dim xPic As Picture
    Dim xRg As Range
    
    Set xRg = r
    
    Application.ScreenUpdating = False
    'Set xRg = Range("A5:B8")
    For Each xPic In ActiveSheet.Pictures
        Set xPicRg = Range(xPic.TopLeftCell.Address & ":" & xPic.BottomRightCell.Address)
        If Not Intersect(xRg, xPicRg) Is Nothing Then xPic.Delete
    Next
    Application.ScreenUpdating = True
End Sub


'run
Sub generateResCodes()
      
    'FaNo
    Call DeletePic(Range("C40"))
    Call mainBarCoder(Range("T2"), Range("C40"), 40, 10)
    
    'OrderNo
    Call DeletePic(Range("C46"))
    Call mainBarCoder(Range("U17"), Range("C46"), 40, 10)
    
    'Amount
    Call DeletePic(Range("C52"))
    Call mainBarCoder(Range("T25"), Range("C52"), 40, 10)
End Sub

'Sub Worksheet_Change(ByVal Target As Range)
'   If Not Intersect(Target, Range("B1:B26")) Is Nothing Then
'        Call generateResCodes
'   End If
'End Sub





