Attribute VB_Name = "mHLS"
Option Explicit

Public Sub RGBToHLS( _
      ByVal r As Long, ByVal g As Long, ByVal b As Long, _
      hue As Long, sat As Long, lum As Long _
   )
Dim Max As Single
Dim Min As Single
Dim delta As Single
Dim h As Single, s As Single, l As Single
Dim rR As Single, rG As Single, rB As Single

   rR = r / 255: rG = g / 255: rB = b / 255

'{Given: rgb each in [0,1].
' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
        Max = Maximum(rR, rG, rB)
        Min = Minimum(rR, rG, rB)
        l = (Max + Min) / 2    '{This is the lightness}
        '{Next calculate saturation}
        If Max = Min Then
            'begin {Acrhomatic case}
            s = 0
            h = 0
           'end {Acrhomatic case}
        Else
           'begin {Chromatic case}
                '{First calculate the saturation.}
           If l <= 0.5 Then
               s = (Max - Min) / (Max + Min)
           Else
               s = (Max - Min) / (2 - Max - Min)
            End If
            '{Next calculate the hue.}
            delta = Max - Min
           If rR = Max Then
                h = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
           ElseIf rG = Max Then
                h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
           ElseIf rB = Max Then
                h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
            End If
            'Debug.Print h
            'h = h * 60
           'If h < 0# Then
           '     h = h + 360            '{Make degrees be nonnegative}
           'End If
        'end {Chromatic Case}
      End If
      
      hue = (h + 1) * 255# / 6#
      sat = s * 255#
      lum = l * 255#
      
'end {RGB_to_HLS}
End Sub

Public Sub HLSToRGB( _
      ByVal hue As Long, ByVal sat As Long, ByVal lum As Long, _
      r As Long, g As Long, b As Long _
   )
Dim rR As Single, rG As Single, rB As Single
Dim hF As Single, sF As Single, lF As Single
Dim Min As Single, Max As Single

   hF = ((hue * 6#) / 255#) - 1#
   sF = sat / 255#
   lF = lum / 255#


   If sF = 0 Then
      ' Achromatic case:
      rR = lF: rG = lF: rB = lF
   Else
      ' Chromatic case:
      ' delta = Max-Min
      If lF <= 0.5 Then
         's = (Max - Min) / (Max + Min)
         ' Get Min value:
         Min = lF * (1 - sF)
      Else
         's = (Max - Min) / (2 - Max - Min)
         ' Get Min value:
         Min = lF - sF * (1 - lF)
      End If
      ' Get the Max value:
      Max = 2 * lF - Min
      
      ' Now depending on sector we can evaluate the h,l,s:
      If (hF < 1) Then
         rR = Max
         If (hF < 0) Then
            rG = Min
            rB = rG - hF * (Max - Min)
         Else
            rB = Min
            rG = hF * (Max - Min) + rB
         End If
      ElseIf (hF < 3) Then
         rG = Max
         If (hF < 2) Then
            rB = Min
            rR = rB - (hF - 2) * (Max - Min)
         Else
            rR = Min
            rB = (hF - 2) * (Max - Min) + rR
         End If
      Else
         rB = Max
         If (hF < 4) Then
            rR = Min
            rG = rR - (hF - 4) * (Max - Min)
         Else
            rG = Min
            rR = (hF - 4) * (Max - Min) + rG
         End If
         
      End If
            
   End If
   r = rR * 255: g = rG * 255: b = rB * 255
   
End Sub
Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
   If (rR > rG) Then
      If (rR > rB) Then
         Maximum = rR
      Else
         Maximum = rB
      End If
   Else
      If (rB > rG) Then
         Maximum = rB
      Else
         Maximum = rG
      End If
   End If
End Function
Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
   If (rR < rG) Then
      If (rR < rB) Then
         Minimum = rR
      Else
         Minimum = rB
      End If
   Else
      If (rB < rG) Then
         Minimum = rB
      Else
         Minimum = rG
      End If
   End If
End Function


