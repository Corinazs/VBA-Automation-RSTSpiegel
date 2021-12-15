Sub PivotKontenAusblenden(gesellschaft As Worksheet)

  Dim bolVisible As Boolean
  Dim pvTab As PivotTable
  Dim pvField As PivotField, pvItem As PivotItem
  Dim index As Integer
  
  Set pvTab = gesellschaft.PivotTables(1)
  Set pvField = pvTab.PivotFields("Loc.Acc.")
  
  Dim arrValues As Variant
  
  Application.ScreenUpdating = False


  pvTab.RefreshTable
  pvField.ClearAllFilters
  
  For Each pvItem In pvField.VisibleItems
    bolVisible = True
   
   'Nur Konten, die mit 1 oder 2 anfangen werden in die Pivot gemapped
        If Not (VBA.left(pvItem.name, 1) = 1 Or VBA.left(pvItem.name, 1) = 2) Then
            bolVisible = False

        End If
        
        'folgene Konten müssen zusätzlich nicht gemapped werden
        If (VBA.left(pvItem.name, 10) = XXX Or VBA.left(pvItem.name, 10) = XXX Or VBA.left(pvItem.name, 10) = XXX) Then
            bolVisible = False

        End If

    
    If bolVisible = False Then pvItem.Visible = False
  Next pvItem
  Application.ScreenUpdating = True

End Sub
