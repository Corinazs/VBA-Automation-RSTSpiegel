Sub RSTSpiegel-Automatisierung()

Dim str_User As String
Dim str_Password As String
Dim strSpeicherort As String
Dim dateiname As String
Dim int_Jahr
Dim int_Monat
Dim ws As Workbook
Dim wsSheet As Worksheet
Dim gesellschaft As Worksheet

Dim export As Workbook
Dim exportSheet As Worksheet

Set ws = ThisWorkbook
Set wsSheet = ws.Sheets("Datengrundlage")


'SAP öffnen
Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", 4
Set WshShell = CreateObject("WScript.Shell")

Set SapGui = GetObject("SAPGUI")

Set Appl = SapGui.GetScriptingEngine

Set Connection = Appl.Openconnection("1. N20 - Production", True)

'SAP Anmeldedaten eingeben
login = MsgBox("Please log in to SAP, then confirm with 'OK'.", vbOKCancel, "SAP")

Set session = Connection.Children(0)

UserForm1.Show

int_Jahr = UserForm1.ComboBox2.Value
int_Monat = UserForm1.ComboBox1.Value
   

If int_Monat = "Januar" Then
    int_Monat = 1
ElseIf int_Monat = "Februar" Then
    int_Monat = 2
ElseIf int_Monat = "März" Then
    int_Monat = 3
ElseIf int_Monat = "April" Then
    int_Monat = 4
ElseIf int_Monat = "Mai" Then
    int_Monat = 5
ElseIf int_Monat = "Juni" Then
    int_Monat = 6
ElseIf int_Monat = "Juli" Then
    int_Monat = 7
ElseIf int_Monat = "August" Then
    int_Monat = 8
ElseIf int_Monat = "September" Then
    int_Monat = 9
ElseIf int_Monat = "Oktober" Then
    int_Monat = 10
ElseIf int_Monat = "November" Then
    int_Monat = 11
ElseIf int_Monat = "Dezember" Then
    int_Monat = 12
End If

int_Monat = CInt(int_Monat)
int_Jahr = CInt(int_Jahr)


If (int_Monat < 10) Then
    dateiname = "XXX_Rückstellungsspiegel_0" & int_Monat & "." & int_Jahr
    ActiveWorkbook.SaveAs "\\XXX & dateiname & ".xlsm"
                            
End If

If (int_Monat >= 10) Then
    dateiname = "XXX_Rückstellungsspiegel_" & int_Monat & "." & int_Jahr
        ActiveWorkbook.SaveAs "\\XXX & dateiname & ".xlsm"

End If


'Daten eingeben
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = 'SAP-Transaction'
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtSD_KTOPL-LOW").Text = 'generall settings'
session.findById("wnd[0]/usr/ctxtSD_KTOPL-LOW").caretPosition = 4
session.findById("wnd[0]/usr/btn%_SD_BUKRS_%_APP_%-VALU_PUSH").press

'Buchungskreise auswählen
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = XXX
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/radPRSSALL").Select

'Berichtsjahr und Monate eingeben
session.findById("wnd[0]/usr/ctxtB0JAHR").Text = int_Jahr
session.findById("wnd[0]/usr/txtB0MONA-LOW").Text = "1"
session.findById("wnd[0]/usr/txtB0MONA-HIGH").Text = int_Monat

'Abruf durchführen
session.findById("wnd[0]/usr/txtB0MONA-HIGH").SetFocus
session.findById("wnd[0]/usr/txtB0MONA-HIGH").caretPosition = 2
'MsgBox "SAP is working. This could last a few minutes. Please wait."

session.findById("wnd[0]/tbar[1]/btn[8]").press

'Als Excel XXL exportieren (ohne Zwischensummen)
session.findById("wnd[0]").maximize
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").Select
session.findById("wnd[1]/usr/cmbG_LISTBOX").SetFocus
session.findById("wnd[1]/usr/cmbG_LISTBOX").Key = "08"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press


dateiname = Year(Now) & Month(Now) & Day(Now) & "_Rückstellungen_SAPDatengrundlage_XXX"

'Datei auf Laufwerk speichern
ActiveWorkbook.SaveAs "\\XXX ".xlsx"

'exportierte Datei schließen
session.findById("wnd[1]/tbar[0]/btn[0]").press

'Datei öffnen
Set export = Workbooks.Open("XXX ".xlsx")
Set exportSheet = export.Sheets("Tabelle1")

'letzteZeile in wsSheet=Rückstellungsspiegel
letzteZeile = wsSheet.Cells(Rows.Count, 1).End(xlUp).Row

'wsSheet leeren
wsSheet.Range("A6:R" & letzteZeile).Delete

'letzteZeile in export
letzteZeileEx = exportSheet.Cells(Rows.Count, 1).End(xlUp).Row

'aus export Daten kopieren + in wsSheet einfügen
exportSheet.Range("A2:R" & letzteZeileEx).Copy _
    Destination:=wsSheet.Range("A6:R" & letzteZeileEx)
    
'ExportDatei schließen
export.Close

'------------------------------------------'

letzteZeile = wsSheet.Cells(Rows.Count, 1).End(xlUp).Row


Set gesellschaft = ws.Sheets("XXX")

'this needs to be distinguished due to different storage locations
'Pivot aktualisieren
 If (int_Monat < 10) Then
    gesellschaft.PivotTables("PivotTable1").ChangePivotCache ActiveWorkbook. _
    PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    "\XXX "\[XXX_Rückstellungsspiegel_0" & int_Monat & "." & int_Jahr & ".xlsm]Datengrundlage!R5C1:R" & letzteZeile & "C18" _
    , Version:=6)
   End If
   
If (int_Monat >= 10) Then
    gesellschaft.PivotTables("PivotTable1").ChangePivotCache ActiveWorkbook. _
    PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    "XXX\[XXX_Rückstellungsspiegel_" & int_Monat & "." & int_Jahr & ".xlsm]Datengrundlage!R5C1:R" & letzteZeile & "C18" _
    , Version:=6)
End If
   

'für US-GAAP Ansicht (alle 1er Konten)
Set gesellschaft = ws.Sheets("XXX")

'this needs to be distinguished due to different storage locations
 If (int_Monat < 10) Then
     gesellschaft.PivotTables("PivotTable1").ChangePivotCache ActiveWorkbook. _
    PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    "\XXX & "\[XXX_Rückstellungsspiegel_0" & int_Monat & "." & int_Jahr & ".xlsm]Datengrundlage!R5C1:R" & letzteZeile & "C18" _
    , Version:=6)
End If

If (int_Monat >= 10) Then
     gesellschaft.PivotTables("PivotTable1").ChangePivotCache ActiveWorkbook. _
    PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    "\XXX & "\[XXX_Rückstellungsspiegel_" & int_Monat & "." & int_Jahr & ".xlsm]Datengrundlage!R5C1:R" & letzteZeile & "C18" _
    , Version:=6)
End If
        
Call PivotKontenAusblenden(gesellschaft)
Call SheetsAktualieren(gesellschaft)
Call DatumÄndern(gesellschaft, int_Monat)


Final = MsgBox("Evaluation completed.", vbOKOnly, "Finished")

End Sub



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



Sub DatumÄndern(gesellschaft As Worksheet, int_Monat As Variant)

If (int_Monat = 1 Or int_Monat = 3 Or int_Monat = 5 Or int_Monat = 7 Or int_Monat = 8 Or int_Monat = 10 Or int_Monat = 12) Then
    gesellschaft.Cells(4, 17).Value = DateSerial(Year(Now), int_Monat, 31)
End If

If (int_Monat = 4 Or int_Monat = 6 Or int_Monat = 9 Or int_Monat = 11) Then
    gesellschaft.Cells(4, 17).Value = DateSerial(Year(Now), int_Monat, 30)
End If

If (int_Monat = 2) Then

    If Year(Now) Mod 4 = 0 And Year(Now) Mod 100 <> 0 Or Year(Now) Mod 400 = 0 Then
       gesellschaft.Cells(4, 17).Value = DateSerial(Year(Now), int_Monat, 29)
    Else
        gesellschaft.Cells(4, 17).Value = DateSerial(Year(Now), int_Monat, 28)
    End If
    
End If

End Sub


Sub SheetsAktualieren(gesellschaft As Worksheet)

letzteZeileSheets = gesellschaft.Cells(Rows.Count, 3).End(xlUp).Row
letzteZeilealt = gesellschaft.Cells(Rows.Count, 12).End(xlUp).Row
'Daten leeren
gesellschaft.Range("L5:T" & letzteZeilealt) = ""
gesellschaft.Range("W5:Y8") = ""

'calucation....
'The calculation was deleted due to sensitive data
    


Dim arr_speicher(50) As Integer
'finde die Zwischensummen
k = 0
For i = 1 To letzteZeileSheets + 1
        If (gesellschaft.Range("C" & i).MergeCells) Then 'Wenn in B verbundene Zellen = Zwischensumme
            zs = i 'merke in welcher Zeile Zwischensumme
            arr_speicher(k) = zs
            k = k + 1
        End If
Next

'alte Formatierung löschen
    With gesellschaft.Range("L5:S" & 500).Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    gesellschaft.Range("L5:S" & 500).Borders(xlDiagonalDown).LineStyle = xlNone
    gesellschaft.Range("L5:S" & 500).Borders(xlDiagonalUp).LineStyle = xlNone
    gesellschaft.Range("L5:S" & 500).Borders(xlEdgeLeft).LineStyle = xlNone
    gesellschaft.Range("L5:S" & 500).Borders(xlEdgeTop).LineStyle = xlNone
    gesellschaft.Range("L5:S" & 500).Borders(xlEdgeBottom).LineStyle = xlNone
    gesellschaft.Range("L5:S" & 500).Borders(xlEdgeRight).LineStyle = xlNone
    gesellschaft.Range("L5:S" & 500).Borders(xlInsideVertical).LineStyle = xlNone
    gesellschaft.Range("L5:S" & 500).Borders(xlInsideHorizontal).LineStyle = xlNone

'Währung und Format anpassen
 gesellschaft.Range("L5:S" & letzteZeileSheets + 2).NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
 k = 0
i = arr_speicher(k)
    While (arr_speicher(k) <> 0)
        gesellschaft.Range("B" & i & ":S" & i).Font.Bold = True
        With gesellschaft.Range("B" & i & ":S" & i).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent6
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        k = k + 1
        i = arr_speicher(k)
    Wend
'blau
    gesellschaft.Range("L" & letzteZeileSheets + 2 & ":S" & letzteZeileSheets + 2).Font.Bold = True
        With gesellschaft.Range("L" & letzteZeileSheets + 2 & ":S" & letzteZeileSheets + 2).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        
    gesellschaft.Range("L5:S" & letzteZeileSheets + 2).Borders(xlDiagonalDown).LineStyle = xlNone
    gesellschaft.Range("L5:S" & letzteZeileSheets + 2).Borders(xlDiagonalUp).LineStyle = xlNone
    With gesellschaft.Range("L5:S" & letzteZeileSheets + 2).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With gesellschaft.Range("L5:S" & letzteZeileSheets + 2).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With gesellschaft.Range("L5:S" & letzteZeileSheets + 2).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With gesellschaft.Range("L5:S" & letzteZeileSheets + 2).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With gesellschaft.Range("L5:S" & letzteZeileSheets + 2).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With gesellschaft.Range("L5:S" & letzteZeileSheets + 2).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    gesellschaft.Range("R2").Copy
    gesellschaft.Range("R5:R" & letzteZeileSheets + 2).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    

End Sub




