Attribute VB_Name = "modDIB"
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbAlpha As Byte
End Type
 
Public Type BITMAPINFOHEADER
    bmSize As Long
    bmWidth As Long
    bmHeight As Long
    bmPlanes As Integer
    bmBitCount As Integer
    bmCompression As Long
    bmSizeImage As Long
    bmXPelsPerMeter As Long
    bmYPelsPerMeter As Long
    bmClrUsed As Long
    bmClrImportant As Long
End Type
 
Public Type BITMAPINFO
    bmHeader As BITMAPINFOHEADER
    bmColors(0 To 255) As RGBQUAD
End Type

'The GetObject API call gives us the bitmap variables we need for the other API calls
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long


'The magical API DIB function calls (they're long!)
Public Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dWidth As Long, ByVal dHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long, ByVal RasterOp As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'The array that will hold our pixel data
Public ImageData1() As Byte
Public ImageData2() As Byte

Public Sub DIBlend(DstPicture As PictureBox, SrcPicture1 As PictureBox, SrcPicture2 As PictureBox, ByVal dBlendPerc As Byte)
    'Coordinate variables
    Dim x As Long, y As Long
    Dim DEST_r As Long
    Dim DEST_g As Long
    Dim DEST_b As Long
    'picture 1 operators
    Dim P1op_r As Long
    Dim P1op_g As Long
    Dim P1op_b As Long
    'picture 2 operators
    Dim P2op_r As Long
    Dim P2op_g As Long
    Dim P2op_b As Long
    
    'Temporary width and height variables are faster than accessing the Scale properties over and over again
    Dim TempWidth As Long, TempHeight As Long
    
    'Get the pixel data into our ImageData array
    GetImageData SrcPicture1, ImageData1()
    GetImageData SrcPicture2, ImageData2()
    
    TempWidth = DstPicture.ScaleWidth - 1
    TempHeight = DstPicture.ScaleHeight - 1
    'run a loop through the picture to change every pixel
    For x = 0 To TempWidth
    For y = 0 To TempHeight
        
        'set operators
        P1op_r = ImageData1(2, x, y)
        P1op_g = ImageData1(1, x, y)
        P1op_b = ImageData1(0, x, y)
        
        P2op_r = ImageData2(2, x, y)
        P2op_g = ImageData2(1, x, y)
        P2op_b = ImageData2(0, x, y)
    
        DEST_r = P1op_r + (((P2op_r - P1op_r) / 255) * dBlendPerc)
        DEST_g = P1op_g + (((P2op_g - P1op_g) / 255) * dBlendPerc)
        DEST_b = P1op_b + (((P2op_b - P1op_b) / 255) * dBlendPerc)
        
        ImageData1(2, x, y) = DEST_r
        ImageData1(1, x, y) = DEST_g
        ImageData1(0, x, y) = DEST_b
    Next y
        'refresh the picture box every 25 lines (a nice progress bar effect if AutoRedraw is set)
        'If DstPicture.AutoRedraw = True And (x Mod 25) = 0 Then SetImageData DstPicture, ImageData()
    Next x
    'final picture refresh
    SetImageData DstPicture, ImageData1()
End Sub

Public Sub DIReflect(DstPicture As PictureBox, SrcPicture1 As PictureBox, SrcPicture2 As PictureBox, Optional REFL As Long = 50, Optional ATTN As Double = 1)
    'Coordinate variables
    Dim x As Long, y As Long
    Dim DEST_r As Long
    Dim DEST_g As Long
    Dim DEST_b As Long
    'picture 1 operators
    Dim P1op_r As Long
    Dim P1op_g As Long
    Dim P1op_b As Long
    'picture 2 operators
    Dim P2op_r As Long
    Dim P2op_g As Long
    Dim P2op_b As Long
    Dim dBlendPerc As Long
    dBlendPerc = 1
    'Temporary width and height variables are faster than accessing the Scale properties over and over again
    Dim TempWidth As Long, TempHeight As Long
    
    'Get the pixel data into our ImageData array
    GetImageData SrcPicture1, ImageData1()
    GetImageData SrcPicture2, ImageData2()
    
    TempWidth = DstPicture.ScaleWidth - 1
    TempHeight = DstPicture.ScaleHeight - 1
    'run a loop through the picture to change every pixel
    For x = 0 To TempWidth
    For y = 0 To TempHeight
        If ImageData1(2, x, y) = 255 And ImageData1(0, x, y) = 255 And ImageData1(1, x, y) = 0 Then
        DEST_r = ImageData2(2, x, y)
        DEST_g = ImageData2(1, x, y)
        DEST_b = ImageData2(0, x, y)
        Else
        'set operators
        P1op_r = ImageData1(2, x, y)
        P1op_g = ImageData1(1, x, y)
        P1op_b = ImageData1(0, x, y)
        
        P2op_r = ImageData2(2, x, y)
        P2op_g = ImageData2(1, x, y)
        P2op_b = ImageData2(0, x, y)
    
        DEST_r = P1op_r + (((P2op_r - P1op_r) / 255) * dBlendPerc)
        DEST_g = P1op_g + (((P2op_g - P1op_g) / 255) * dBlendPerc)
        DEST_b = P1op_b + (((P2op_b - P1op_b) / 255) * dBlendPerc)
        End If
        'apply flip
        ImageData1(2, x, y) = DEST_r
        ImageData1(1, x, y) = DEST_g
        ImageData1(0, x, y) = DEST_b
        
        ImageData2(2, x, y) = DEST_r
        ImageData2(1, x, y) = DEST_g
        ImageData2(0, x, y) = DEST_b
    
    dBlendPerc = Int(((y * ATTN) * 255) / TempHeight) + REFL
    If dBlendPerc > 255 Then dBlendPerc = 255
    Next y
        'refresh the picture box every 25 lines (a nice progress bar effect if AutoRedraw is set)
        'If DstPicture.AutoRedraw = True And (x Mod 25) = 0 Then SetImageData DstPicture, ImageData()
        
    Next x
    
    For x = 0 To TempWidth
    For y = 0 To TempHeight

        ImageData1(2, x, y) = ImageData2(2, x, (TempHeight - y))
        ImageData1(1, x, y) = ImageData2(1, x, (TempHeight - y))
        ImageData1(0, x, y) = ImageData2(0, x, (TempHeight - y))
    
    Next y
    Next x
    'final picture refresh
    SetImageDataREFL DstPicture, ImageData1()
End Sub


'Public Sub DrawDIBBrightness(DstPicture As PictureBox, SrcPicture As PictureBox, ByVal Brightness As Single)
'    'Coordinate variables
'    Dim x As Long, y As Long
'    'Build a look-up table for all possible brightness values
'    Dim bTable(0 To 255) As Long
'    Dim TempColor As Long
'    For x = 0 To 255
'        'Calculate the brightness for pixel value x
'        TempColor = Int(CSng(x) * Brightness)
'        'Make sure that the calculated value is between 0 and 255 (so we don't get an error)
'        ByteMe TempColor
'        'Place the corrected value into its array spot
'        bTable(x) = TempColor
'    Next x
'    'Get the pixel data into our ImageData array
'    GetImageData SrcPicture, ImageData1()
'    'Temporary width and height variables are faster than accessing the Scale properties over and over again
'    Dim TempWidth As Long, TempHeight As Long
'    TempWidth = DstPicture.ScaleWidth - 1
'    TempHeight = DstPicture.ScaleHeight - 1
'    'run a loop through the picture to change every pixel
'    For x = 0 To TempWidth
'    For y = 0 To TempHeight
'        'Use the values in the look-up table to quickly change the brightness values
'        'of each color.  The look-up table is much faster than doing the math
'        'over and over for each individual pixel.
'        ImageData(2, x, y) = bTable(ImageData(2, x, y))   'Change the red
'        ImageData(1, x, y) = bTable(ImageData(1, x, y))   'Change the green
'        ImageData(0, x, y) = bTable(ImageData(0, x, y))   'Change the blue
'    Next y
'        'refresh the picture box every 25 lines (a nice progress bar effect if AutoRedraw is set)
'        If DstPicture.AutoRedraw = True And (x Mod 25) = 0 Then SetImageData DstPicture, ImageData()
'    Next x
'    'final picture refresh
'    SetImageData DstPicture, ImageData()
'End Sub

Public Sub FlipMe(ByRef ImageData() As Byte)

End Sub

'Routine to get an image's pixel information into an array dimensioned (rgb, x, y)
Public Sub GetImageData(ByRef SrcPictureBox As PictureBox, ByRef ImageData() As Byte)
    'Declare us some variables of the necessary bitmap types
    Dim bm As BITMAP
    Dim bmi As BITMAPINFO
    'Now we fill up the bmi (Bitmap information variable) with all of the appropriate data
    bmi.bmHeader.bmSize = 40 'Size, in bytes, of the header (always 40)
    bmi.bmHeader.bmPlanes = 1 'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24 'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0 'Compression: standard/none or RLE
    'Calculate the size of the bitmap type (in bytes)
    Dim bmLen As Long
    bmLen = Len(bm)
    'Get the picture box information from SrcPictureBox and put it into our 'bm' variable
    GetObject SrcPictureBox.Image, bmLen, bm
    'Build a correctly sized array
    ReDim ImageData(0 To 2, 0 To bm.bmWidth - 1, 0 To bm.bmHeight - 1)
    'Finish building the 'bmi' variable we want to pass to the GetDIBits call (the same one we used above)
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    'Now that we've completely filled up the 'bmi' variable, we use GetDIBits to take the data from
    'SrcPictureBox and put it into the ImageData() array using the settings we specified in 'bmi'
    GetDIBits SrcPictureBox.hdc, SrcPictureBox.Image, 0, bm.bmHeight, ImageData(0, 0, 0), bmi, 0
End Sub

'Routine to set an image's pixel information from an array dimensioned (rgb, x, y)
Public Sub SetImageData(ByRef DstPictureBox As PictureBox, ByRef ImageData() As Byte)
    'Declare us some variables of the necessary bitmap types
    Dim bm As BITMAP
    Dim bmi As BITMAPINFO
    'Now we fill up the bmi (Bitmap information variable) with all of the appropriate data
    bmi.bmHeader.bmSize = 40 'Size, in bytes, of the header (always 40)
    bmi.bmHeader.bmPlanes = 1 'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24 'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0 'Compression: standard/none or RLE
    'Calculate the size of the bitmap type (in bytes)
    Dim bmLen As Long
    bmLen = Len(bm)
    'Get the picture box information from DstPictureBox and put it into our 'bm' variable
    GetObject DstPictureBox.Image, bmLen, bm
    'Now that we know the object's size, finish building the temporary header to pass to the StretchDIBits call
    '(continuing to use the 'bmi' we used above)
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    'Now that we've built the temporary header, we use StretchDIBits to take the data from the
    'ImageData() array and put it into SrcPictureBox using the settings specified in 'bmi' (the
    'StretchDIBits call should be on one continuous line)
    StretchDIBits DstPictureBox.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, bm.bmWidth, bm.bmHeight, ImageData(0, 0, 0), bmi, 0, vbSrcCopy
    'Since this doesn't automatically initialize AutoRedraw, we have to do it manually
    'Note: Always set AutoRedraw to true when using DIB sections; when AutoRedraw is false
    'you will get unpredictable results.
    If DstPictureBox.AutoRedraw = True Then
        DstPictureBox.Picture = DstPictureBox.Image
        DstPictureBox.Refresh
    End If
End Sub

Public Sub SetImageDataREFL(ByRef DstPictureBox As PictureBox, ByRef ImageData() As Byte)
    'Declare us some variables of the necessary bitmap types
    Dim bm As BITMAP
    Dim bmi As BITMAPINFO
    'Now we fill up the bmi (Bitmap information variable) with all of the appropriate data
    bmi.bmHeader.bmSize = 40 'Size, in bytes, of the header (always 40)
    bmi.bmHeader.bmPlanes = 1 'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24 'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0 'Compression: standard/none or RLE
    'Calculate the size of the bitmap type (in bytes)
    Dim bmLen As Long
    bmLen = Len(bm)
    'Get the picture box information from DstPictureBox and put it into our 'bm' variable
    GetObject DstPictureBox.Image, bmLen, bm
    'Now that we know the object's size, finish building the temporary header to pass to the StretchDIBits call
    '(continuing to use the 'bmi' we used above)
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    'Now that we've built the temporary header, we use StretchDIBits to take the data from the
    'ImageData() array and put it into SrcPictureBox using the settings specified in 'bmi' (the
    'StretchDIBits call should be on one continuous line)
    SetStretchBltMode DstPictureBox.hdc, 4
    StretchDIBits DstPictureBox.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, bm.bmWidth, bm.bmHeight, ImageData(0, 0, 0), bmi, 0, vbSrcCopy
    'Since this doesn't automatically initialize AutoRedraw, we have to do it manually
    'Note: Always set AutoRedraw to true when using DIB sections; when AutoRedraw is false
    'you will get unpredictable results.
    If DstPictureBox.AutoRedraw = True Then
        DstPictureBox.Picture = DstPictureBox.Image
        DstPictureBox.Refresh
    End If
End Sub

'Standardized routine for converting to absolute byte values
Public Sub ByteMe(ByRef TempVar As Long)
    If TempVar > 255 Then TempVar = 255: Exit Sub
    If TempVar < 0 Then TempVar = 0: Exit Sub
End Sub
