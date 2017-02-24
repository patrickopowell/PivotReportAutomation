Sub UpdatePivotData()

On Error GoTo ErrorHandler

Dim Data_sht As Worksheet
Dim Pivot_sht As Worksheet
Dim StartPoint As Range
Dim DataRange As Range
Dim PivotName As String
Dim NewRange As String
Dim isActive As Variant

isActive = ThisWorkbook.Sheets("Index").Range("D2:D251")

'Report file tab name
Dim tabName As Variant

tabName = ThisWorkbook.Sheets("Index").Range("E2:E251")

'Alerts report directory
Dim NewDir As String
NewDir = ThisWorkbook.Sheets("Data").Range("B2")
'Initialize data files
Dim data_sheets As Variant

data_sheets = ThisWorkbook.Sheets("Index").Range("C2:C251")

'Initialize report directory
Dim RepDir As String
RepDir = ThisWorkbook.Sheets("Data").Range("B1")
'Initialize report files
Dim reports As Variant

reports = ThisWorkbook.Sheets("Index").Range("B2:B251")

'Get report count
Dim rSize As Integer
rSize = 250 ' = UBound(data_sheets) - LBound(data_sheets)
Dim i As Integer
Dim Count As Integer
Count = 0

'Report WorkBook
Dim rwb As String

If MsgBox("You are running reports for: " & vbNewLine & vbNewLine & RepDir & vbNewLine & vbNewLine _
& "Continue?", vbYesNo) = vbNo Then Exit Sub

'Loop through each report
For i = 1 To rSize
 
  rwb = RepDir + reports(i, 1)
 
  If (rwb = RepDir Or isActive(i, 1) = 0) Then
    GoTo ContinueLoop
  End If
  
  Count = Count + 1
 
'Open report file
  Set Data_sht = Workbooks.Open(rwb, True).Sheets(1)
 
'Create data tab
  Sheets.Add After:=ActiveSheet
  ActiveSheet.Name = "Info"
  
  Dim wb As String
  wb = NewDir + data_sheets(i, 1)

'Open data sheet and copy data
  Set Pivot_sht = Workbooks.Open(wb, True).Sheets(1)
  Selection.CurrentRegion.Select
  Selection.Copy
  Windows(reports(i, 1)).Activate
  ActiveSheet.Paste

'Set Variables Equal to Data Sheet and Pivot Sheet
  Set Data_sht = ActiveWorkbook.Worksheets("Info")
  Set Pivot_sht = ActiveWorkbook.Worksheets(tabName(i, 1))

'Pivot Table Name
  PivotName = "PivotReport" + CStr(i - 1)

'Dynamically Retrieve Range Address of Data
  Set StartPoint = Data_sht.Range("A1")
  Set DataRange = Data_sht.Range(StartPoint, StartPoint.CurrentRegion)
  NewRange = Data_sht.Name & "!" & _
    DataRange.Address(ReferenceStyle:=xlR1C1)


'Make sure every column in data set has a heading and is not blank (error prevention)
  If WorksheetFunction.CountBlank(DataRange.Rows(1)) > 0 Then
    MsgBox "One of your data columns has a blank heading." & vbNewLine _
      & "Please fix and re-run! [" & WorksheetFunction.CountBlank(DataRange.Rows(1)) & "]", _
      vbCritical, "Column Heading Missing!" _
  
    Exit Sub
  End If

'Change Pivot Table Data Source Range Address
  Pivot_sht.PivotTables(PivotName).ChangePivotCache _
    ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=NewRange)
      
'Ensure Pivot Table is Refreshed
  Pivot_sht.PivotTables(PivotName).RefreshTable
  
'Delete data tab
  Application.DisplayAlerts = False 'Suppress save alert
  
  Sheets("Info").Delete
  
  Sheets(tabName(i, 1)).Activate
  
  Application.Workbooks(data_sheets(i, 1)).Close
  
'Save report file
  Application.Workbooks(reports(i, 1)).Save
  
  Application.Workbooks(reports(i, 1)).Close

  Application.DisplayAlerts = True 'Re-enable user alerts
 
 'Wait one second on thread sync
  Application.Wait (Now + TimeValue("0:00:01"))

ContinueLoop:

Next i

'Notify user of completion
MsgBox Count & " Reports Updated!"

ErrorHandler:
Select Case Err.Number
  Case 1004
    MsgBox "Report '" & data_sheets(i, 1) & "' does not contain data" & _
    vbNewLine & vbNewLine & "OR" & vbNewLine & _
    vbNewLine & "The application could not find the pivot table in question." & _
    "Please rename the pivot table to 'PivotReport' + [index_number] (i.e. PivotReport3)"
  End Select
Resume Next

End Sub
