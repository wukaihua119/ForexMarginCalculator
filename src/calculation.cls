VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "�u�@��13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True




'計算浮動損益
Private Sub CommandButton1_Click()
Cells(9, 11).Value = Cells(6, 11).Value
End Sub

Private Sub CommandButton2_Click()

'計算保證金

    
'計算XXXUSD的保證金
  For i = 6 To 9
   If Cells(4, 19).Value = Cells(i, 3).Value Then
      Cells(7, 19).Value = Cells(5, 19).Value * 100000 / Cells(6, 19).Value * Cells(i, 4).Value
      Cells(8, 19).Value = Cells(7, 19).Value * Cells(31, 5).Value
      End If
  Next i
  
  
'計算USDXXX的保證金
For i = 10 To 12
   If Cells(4, 19).Value = Cells(i, 3).Value Then
      Cells(7, 19).Value = Cells(5, 19).Value * 100000 / Cells(6, 19).Value
      Cells(8, 19).Value = Cells(7, 19).Value * Cells(31, 5).Value
   End If
  Next i
  
'交叉貨幣保證金


'計算EURXXX交叉貨幣保證金

  For i = 13 To 17
    If Cells(4, 19).Value = Cells(i, 3).Value Then
       Cells(7, 19).Value = Cells(5, 19).Value * 100000 / Cells(6, 19).Value * Cells(6, 4).Value
       Cells(8, 19).Value = Cells(7, 19).Value * Cells(31, 5).Value
    End If
  Next i
  
  
'計算AUDXXX交叉貨幣保證金

  For i = 18 To 20
    If Cells(4, 19).Value = Cells(i, 3).Value Then
       Cells(7, 19).Value = Cells(5, 19).Value * 100000 / Cells(6, 19).Value * Cells(7, 4).Value
       Cells(8, 19).Value = Cells(7, 19).Value * Cells(31, 5).Value
    End If
  Next i
  
  
'計算GBPXXX交叉貨幣保證金

  For i = 21 To 24
    If Cells(4, 19).Value = Cells(i, 3).Value Then
       Cells(7, 19).Value = Cells(5, 19).Value * 100000 / Cells(6, 19).Value * Cells(8, 4).Value
       Cells(8, 19).Value = Cells(7, 19).Value * Cells(31, 5).Value
    End If
  Next i
  
  
'計算NZDXXX交叉貨幣保證金

  For i = 25 To 26
    If Cells(4, 19).Value = Cells(i, 3).Value Then
       Cells(7, 19).Value = Cells(5, 19).Value * 100000 / Cells(6, 19).Value * Cells(9, 4).Value
       Cells(8, 19).Value = Cells(7, 19).Value * Cells(31, 5).Value
    End If
  Next i
  
  
'計算CADJPY交叉貨幣保證金

  For i = 27 To 27
    If Cells(4, 19).Value = Cells(i, 3).Value Then
       Cells(7, 19).Value = Cells(5, 19).Value * 100000 / Cells(6, 19).Value / Cells(12, 5).Value
       Cells(8, 19).Value = Cells(7, 19).Value * Cells(31, 5).Value
    End If
  Next i
  

  

End Sub

'損益試算
Private Sub CommandButton3_Click()
If Cells(11, 19).Value = "買價" Then '買入算損益
    If Cells(13, 19).Value = "間接貨幣" Then  'XXXUSD
       Cells(17, 19).Value = (Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000
       Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
    

For i = 5 To 10  'XXXJPY
    If Cells(13, 19).Value = Worksheets("數據資料").Cells(i, 3).Value Then
       Cells(17, 19).Value = (Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000 / Cells(10, 5).Value
       Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
Next i

For i = 11 To 15  'XXXCAD
    If Cells(13, 19).Value = Worksheets("數據資料").Cells(i, 3).Value Then
       Cells(17, 19).Value = (Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000 / Cells(12, 5).Value
       Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
Next i

For i = 16 To 18  'XXXNZD
    If Cells(13, 19).Value = Worksheets("數據資料").Cells(i, 3).Value Then
       Cells(17, 19).Value = (Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000 * Cells(9, 4).Value
       Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
Next i

For i = 19 To 20 'XXXAUD
    If Cells(13, 19).Value = Worksheets("數據資料").Cells(i, 3).Value Then
       Cells(17, 19).Value = (Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000 * Cells(7, 4).Value
       Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
Next i

For i = 21 To 21 'EURGBP
     If Cells(13, 19).Value = Worksheets("數據資料").Cells(i, 3).Value Then
        Cells(17, 19).Value = (Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000 * Cells(8, 4).Value
        Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
Next i

For i = 22 To 22 'USDCNH
     If Cells(13, 19).Value = Worksheets("數據資料").Cells(i, 3).Value Then
        Cells(17, 19).Value = (Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000 / Cells(11, 5).Value
        Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
Next i

ElseIf Cells(11, 19).Value = "賣價" Then '賣出算損益
    If Cells(13, 19).Value = "間接貨幣" Then  'XXXUSD
       Cells(17, 19).Value = -(Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000
       Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
    

For i = 5 To 10  'XXXJPY
    If Cells(13, 19).Value = Worksheets("數據資料").Cells(i, 3).Value Then
       Cells(17, 19).Value = -(Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000 / Cells(10, 5).Value
       Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
Next i

For i = 11 To 15  'XXXCAD
    If Cells(13, 19).Value = Worksheets("數據資料").Cells(i, 3).Value Then
       Cells(17, 19).Value = -(Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000 / Cells(12, 5).Value
       Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
Next i

For i = 16 To 18  'XXXNZD
    If Cells(13, 19).Value = Worksheets("數據資料").Cells(i, 3).Value Then
       Cells(17, 19).Value = -(Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000 * Cells(9, 4).Value
       Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
Next i

For i = 19 To 20 'XXXAUD
    If Cells(13, 19).Value = Worksheets("數據資料").Cells(i, 3).Value Then
       Cells(17, 19).Value = -(Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000 * Cells(7, 4).Value
       Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
Next i

For i = 21 To 21 'EURGBP
     If Cells(13, 19).Value = Worksheets("數據資料").Cells(i, 3).Value Then
        Cells(17, 19).Value = -(Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000 * Cells(8, 4).Value
        Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
Next i

For i = 22 To 22 'USDCNH
     If Cells(13, 19).Value = Worksheets("數據資料").Cells(i, 3).Value Then
        Cells(17, 19).Value = -(Cells(16, 19).Value - Cells(15, 19).Value) * Cells(14, 19).Value * 100000 / Cells(11, 5).Value
        Cells(18, 19).Value = Cells(17, 19).Value * Cells(31, 5).Value
    End If
Next i
End If

End Sub

'計算隔夜利息
Private Sub CommandButton4_Click()

Dim Spread_JPYtoUSD, Spread_CNHtoUSD, Spread_CADtoUSD, Spread_NZDtoUSD, Spread_GBPtoUSD, Spread_AUDtoUSD As Long
    Spread_JPYtoUSD = Cells

If Cells(16, 11).Value = Cells(i, 3).Value Then
    If Cells(15, 11).Value = "買價" Then
        Cells(i, 7).Select
    ElseIf Cells(15, 11).Value = "賣價" Then
        Cells(i, 6).Value.Select

'計算總記息天數
    If Cells(17, 11).Value < TimeValue("05:00") And Cells(19, 11) >= TimeValue("05:00") Then
       Cells(22, 11).Value = Cells(21, 11).Value + 1 + Cells(20, 11).Value * 2
    ElseIf Cells(17, 11).Value < TimeValue("05:00") And Cells(19, 11) < TimeValue("05:00") Then
       Cells(22, 11).Value = Cells(21, 11).Value + Cells(20, 11).Value * 2
    ElseIf Cells(17, 11).Value >= TimeValue("05:00") And Cells(19, 11) >= TimeValue("05:00") Then
       Cells(22, 11).Value = Cells(21, 11).Value + Cells(20, 11).Value * 2
    ElseIf Cells(17, 11).Value >= TimeValue("05:00") And Cells(19, 11) < TimeValue("05:00") Then
       Cells(22, 11).Value = Cells(21, 11).Value - 1 + Cells(20, 11).Value * 2
    End If

'計算隔夜利息
For i = 6 To 27
    If Cells(4, 11).Value = Cells(4, 4).Value And Cells(5, 11).Value = Cells(i, 3).Value Then
       Cells(14, 11).Value = Cells(i, 6).Value * Cells(7, 11).Value * Cells(22, 11).Value
    ElseIf Cells(4, 11).Value = Cells(4, 5).Value And Cells(5, 11).Value = Cells(i, 3).Value Then
       Cells(14, 11).Value = Cells(i, 7).Value * Cells(7, 11).Value * Cells(22, 11).Value
    End If
Next i
End Sub

Private Sub CommandButton5_Click()

'計算手數

Dim CNHtoUSD, JPYtoUSD, CADtoUSD, EURtoUSD, AUDtoUSD, GBPtoUSD, NZDtoUSD, USDtoUSD As Long
 
 '將持有貨幣換算成美金
If Cells(4, 24).Text = "CNH" Then
 CNHtoUSD = Cells(5, 24).Value / Cells(11, 5).Value
ElseIf Cells(4, 24).Text = "JPY" Then
 JPYtoUSD = Cells(5, 24).Value / Cells(10, 5).Value
ElseIf Cells(4, 24).Value = "CAD" Then
 CADtoUSD = Cells(5, 24).Value / Cells(12, 5).Value
ElseIf Cells(4, 24).Value = "EUR" Then
 EURtoUSD = Cells(5, 24).Value * Cells(6, 4).Value
ElseIf Cells(4, 24).Value = "AUD" Then
 AUDtoUSD = Cells(5, 24).Value * Cells(7, 4).Value
ElseIf Cells(4, 24).Value = "GBP" Then
 GBPtoUSD = Cells(5, 24).Value * Cells(8, 4).Value
ElseIf Cells(4, 24).Value = "NZD" Then
 NZDtoUSD = Cells(5, 24).Value * Cells(9, 4).Value
ElseIf Cells(4, 24).Value = "USD" Then
 USDtoUSD = Cells(5, 24).Value
End If

    '算手數
'若持有貨幣是CNH
If Cells(4, 24).Text = "CNH" Then
    For i = 4 To 7    'AUDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = CNHtoUSD / (100000 / Cells(7, 24).Value * Cells(7, 4).Value)
      End If
    Next i
    
    For i = 8 To 8    'CADJPY
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = CNHtoUSD / (100000 / Cells(7, 24).Value / Cells(12, 5).Value)
      End If
    Next i
    
    For i = 9 To 14    'EURXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = CNHtoUSD / (100000 / Cells(7, 24).Value * Cells(6, 4).Value)
    End If
    Next i
    
    For i = 15 To 19    'GBPXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = CNHtoUSD / (100000 / Cells(7, 24).Value * Cells(8, 4).Value)
      End If
    Next i
    
    For i = 20 To 22    'NZDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = CNHtoUSD / (100000 / Cells(7, 24).Value * Cells(9, 4).Value)
      End If
    Next i
    
    For i = 23 To 25    'USDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = USDtoUSD / (100000 / Cells(7, 24).Value)
      End If
    Next i
End If

 '若持有貨幣是JPY
If Cells(4, 24).Text = "JPY" Then
    For i = 4 To 7    'AUDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = JPYtoUSD / (100000 / Cells(7, 24).Value * Cells(7, 4).Value)
      End If
    Next i
    
    For i = 8 To 8    'CADJPY
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = JPYtoUSD / (100000 / Cells(7, 24).Value / Cells(12, 5).Value)
      End If
    Next i
    
    For i = 9 To 14    'EURXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = JPYtoUSD / (100000 / Cells(7, 24).Value * Cells(6, 4).Value)
    End If
    Next i
    
    For i = 15 To 19    'GBPXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = JPYtoUSD / (100000 / Cells(7, 24).Value * Cells(8, 4).Value)
      End If
    Next i
    
    For i = 20 To 22    'NZDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = JPYtoUSD / (100000 / Cells(7, 24).Value * Cells(9, 4).Value)
      End If
    Next i
    
    For i = 23 To 25    'USDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = JPYtoUSD / (100000 / Cells(7, 24).Value)
      End If
    Next i
End If

 '若持有貨幣是CAD
If Cells(4, 24).Text = "CAD" Then
    For i = 4 To 7    'AUDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = CADtoUSD / (100000 / Cells(7, 24).Value * Cells(7, 4).Value)
      End If
    Next i
    
    For i = 8 To 8    'CADJPY
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = CADtoUSD / (100000 / Cells(7, 24).Value / Cells(12, 5).Value)
      End If
    Next i
    
    For i = 9 To 14    'EURXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = CADtoUSD / (100000 / Cells(7, 24).Value * Cells(6, 4).Value)
    End If
    Next i
    
    For i = 15 To 19    'GBPXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = CADtoUSD / (100000 / Cells(7, 24).Value * Cells(8, 4).Value)
      End If
    Next i
    
    For i = 20 To 22    'NZDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = CADtoUSD / (100000 / Cells(7, 24).Value * Cells(9, 4).Value)
      End If
    Next i
    
    For i = 23 To 25    'USDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = CADtoUSD / (100000 / Cells(7, 24).Value)
      End If
    Next i
End If

 '若持有貨幣是EUR
If Cells(4, 24).Text = "EUR" Then
    For i = 4 To 7    'AUDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = EURtoUSD / (100000 / Cells(7, 24).Value * Cells(7, 4).Value)
      End If
    Next i
    
    For i = 8 To 8    'CADJPY
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = EURtoUSD / (100000 / Cells(7, 24).Value / Cells(12, 5).Value)
      End If
    Next i
    
    For i = 9 To 14    'EURXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = EURtoUSD / (100000 / Cells(7, 24).Value * Cells(6, 4).Value)
    End If
    Next i
    
    For i = 15 To 19    'GBPXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = EURtoUSD / (100000 / Cells(7, 24).Value * Cells(8, 4).Value)
      End If
    Next i
    
    For i = 20 To 22    'NZDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = EURtoUSD / (100000 / Cells(7, 24).Value * Cells(9, 4).Value)
      End If
    Next i
    
    For i = 23 To 25    'USDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = EURtoUSD / (100000 / Cells(7, 24).Value)
      End If
    Next i
End If

 '若持有貨幣是AUD
If Cells(4, 24).Text = "AUD" Then
    For i = 4 To 7    'AUDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = AUDtoUSD / (100000 / Cells(7, 24).Value * Cells(7, 4).Value)
      End If
    Next i
    
    For i = 8 To 8    'CADJPY
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = AUDtoUSD / (100000 / Cells(7, 24).Value / Cells(12, 5).Value)
      End If
    Next i
    
    For i = 9 To 14    'EURXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = AUDtoUSD / (100000 / Cells(7, 24).Value * Cells(6, 4).Value)
    End If
    Next i
    
    For i = 15 To 19    'GBPXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = AUDtoUSD / (100000 / Cells(7, 24).Value * Cells(8, 4).Value)
      End If
    Next i
    
    For i = 20 To 22    'NZDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = AUDtoUSD / (100000 / Cells(7, 24).Value * Cells(9, 4).Value)
      End If
    Next i
    
    For i = 23 To 25    'USDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = AUDtoUSD / (100000 / Cells(7, 24).Value)
      End If
    Next i
End If

 '若持有貨幣是GBP
If Cells(4, 24).Text = "GBP" Then
    For i = 4 To 7    'AUDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = GBPtoUSD / (100000 / Cells(7, 24).Value * Cells(7, 4).Value)
      End If
    Next i
    
    For i = 8 To 8    'CADJPY
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = GBPtoUSD / (100000 / Cells(7, 24).Value / Cells(12, 5).Value)
      End If
    Next i
    
    For i = 9 To 14    'EURXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = GBPtoUSD / (100000 / Cells(7, 24).Value * Cells(6, 4).Value)
    End If
    Next i
    
    For i = 15 To 19    'GBPXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = GBPtoUSD / (100000 / Cells(7, 24).Value * Cells(8, 4).Value)
      End If
    Next i
    
    For i = 20 To 22    'NZDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = GBPtoUSD / (100000 / Cells(7, 24).Value * Cells(9, 4).Value)
      End If
    Next i
    
    For i = 23 To 25    'USDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = GBPtoUSD / (100000 / Cells(7, 24).Value)
      End If
    Next i
End If

 '若持有貨幣是NZD
If Cells(4, 24).Text = "NZD" Then
    For i = 4 To 7    'AUDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = NZDtoUSD / (100000 / Cells(7, 24).Value * Cells(7, 4).Value)
      End If
    Next i
    
    For i = 8 To 8    'CADJPY
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = NZDtoUSD / (100000 / Cells(7, 24).Value / Cells(12, 5).Value)
      End If
    Next i
    
    For i = 9 To 14    'EURXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = NZDtoUSD / (100000 / Cells(7, 24).Value * Cells(6, 4).Value)
    End If
    Next i
    
    For i = 15 To 19    'GBPXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = NZDtoUSD / (100000 / Cells(7, 24).Value * Cells(8, 4).Value)
      End If
    Next i
    
    For i = 20 To 22    'NZDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = NZDtoUSD / (100000 / Cells(7, 24).Value * Cells(9, 4).Value)
      End If
    Next i
    
    For i = 23 To 25    'USDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = NZDtoUSD / (100000 / Cells(7, 24).Value)
      End If
    Next i
End If

 '若持有貨幣是USD
If Cells(4, 24).Text = "USD" Then
 For i = 4 To 7    'AUDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = Cells(5, 24).Value / (100000 / Cells(7, 24).Value * Cells(7, 4).Value)
      End If
    Next i
    
    For i = 8 To 8    'CADJPY
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = Cells(5, 24).Value / (100000 / Cells(7, 24).Value / Cells(12, 5).Value)
      End If
    Next i
    
    For i = 9 To 14    'EURXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = Cells(5, 24).Value / (100000 / Cells(7, 24).Value * Cells(6, 4).Value)
    End If
    Next i
    
    For i = 15 To 19    'GBPXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = Cells(5, 24).Value / (100000 / Cells(7, 24).Value * Cells(8, 4).Value)
      End If
    Next i
    
    For i = 20 To 22    'NZDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = Cells(5, 24).Value / (100000 / Cells(7, 24).Value * Cells(9, 4).Value)
      End If
    Next i
    
    For i = 23 To 25    'USDXXX
      If Cells(6, 24).Value = Worksheets("數據資料").Cells(i, 5).Value Then
         Cells(9, 24).Value = Cells(5, 24).Value / (100000 / Cells(7, 24).Value)
      End If
    Next i
End If


End Sub



