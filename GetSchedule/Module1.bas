Attribute VB_Name = "Module1"
Dim bWeStartedOutlook As Boolean
Public UserDate As Date
Dim ItemCount As Long
    Dim success As Boolean
    Dim dteStart As Date
    Dim dteEnd As Date
Dim ThisAppt As Object ' Outlook.AppointmentItem
Dim MyItem As Object
Dim StringToCheck As String
Dim i As Long
    Dim MyBook As Excel.Workbook
    Dim xlSht As Excel.Worksheet
    Dim rngStart As Excel.Range
    Dim rngHeader As Excel.Range
Dim olApp As Object '  Outlook.Application
Dim olNS As Object ' Outlook.Namespace
Dim myCalItems As Object ' Outlook.Items
Dim ItemstoCheck As Object ' Outlook.Items
  Dim ColCount As Long
  
  Dim arrData As Variant
  Dim RaterData As Variant

Private Function Quote(MyText)
' from Sue Mosher's excellent book "Microsoft Outlook Programming"
  Quote = Chr(34) & MyText & Chr(34)
End Function

Function GetOutlookApp() As Object
On Error Resume Next
  Set GetOutlookApp = GetObject(, "Outlook.Application")
  If Err.Number <> 0 Then
    Set GetOutlookApp = CreateObject("Outlook.Application")
    bWeStartedOutlook = True
  End If
On Error GoTo 0
End Function

Sub GetSchedule()

    
      ' set up worksheet

    Set MyBook = Excel.Workbooks.Add 'ThisWorkbook 'Excel.Workbooks.Open("P:\0-Resources\GetSchedule.xlsm")
    Set xlSht = MyBook.Sheets(1)
    Set rngStart = xlSht.Range("A1")
    Set rngHeader = Range(rngStart, rngStart.Offset(0, 6))
 
  ' with assistance from Jon Peltier http://peltiertech.com/WordPress and
   ' http://support.microsoft.com/kb/306022
  rngHeader.Value = Array("Subject", "Start Date", "Start Time", "Location", "Categories", "Icon", "Region") ' "End Date", "End Time", "Body"
    dteStart = AskForDate 'InputBox("What is the start date?")
    'dteEnd = InputBox("What is the end date?")
    ItemCount = 0
    'success = GetCalData("Test1", "UserCal", dteStart)
    success = GetCalData("Testing Schedule", "PublicCal", dteStart)
    If success Then
        AddRaters
        Dim s As String
        s = Environ("USERPROFILE") & "\Desktop"
        ActiveWorkbook.SaveAs s & "\Schedule " & Month(UserDate) & "-" & Day(UserDate) & ".xlsx", FileFormat:=51
         ActiveWorkbook.Close
    End If
End Sub
Function AskForDate()
    AskDate.Show
    AskForDate = UserDate
End Function

Private Function GetCalData(CalName As String, CalType As String, StartDate As Date, Optional EndDate As Date) As Boolean
' Exports calendar information to Excel worksheet
' -------------------------------------------------
' Notes:
' If Outlook is not open, it still works, but much
' slower (~8 secs vs. 2 secs w/ Outlook open).
' End Date is optional, if you want to pull from
' only one day, use: Call GetCalData("7/14/2008")
' -------------------------------------------------

 
' if no end date was specified, then the requestor
' only wants one day, so set EndDate = StartDate
' this will let us return appts from multiple dates,
' if the requestor does in fact set an appropriate end date
If EndDate = "12:00:00 AM" Then
  EndDate = StartDate
End If
 
If EndDate < StartDate Then
  MsgBox "Those dates seem switched, please check" & _
      "them and try again.", vbInformation
  GoTo ExitProc
End If
 
' get Outlook

Set olApp = GetOutlookApp
If olApp Is Nothing Then
  MsgBox "Cannot start Outlook.", vbExclamation
  GoTo ExitProc
End If
 
' get default Calendar

Set olNS = olApp.GetNamespace("MAPI")
Select Case CalType
    Case "UserCal"
    MsgBox olNS.GetDefaultFolder(olFolderCalendar).Parent.Folders(CalName)
        Set myCalItems = olNS.GetDefaultFolder(olFolderCalendar).Parent.Folders(CalName).Items
    Case "PublicCal"
        Set myCalItems = olNS.GetDefaultFolder(olPublicFoldersAllPublicFolders).Folders(CalName).Items
    Case Else
        GoTo ExitProc
End Select

' ------------------------------------------------------------------
' the following code adapted from:
' http://www.outlookcode.com/article.aspx?id=30
'
With myCalItems
  .Sort "[Start]", False
  .IncludeRecurrences = False
End With
'
StringToCheck = "[Start] >= " & Quote(StartDate & " 12:00 AM") & _
    " AND [End] <= " & Quote(EndDate & " 11:59 PM")
Debug.Print StringToCheck
'

Set ItemstoCheck = myCalItems.Restrict(StringToCheck)
Debug.Print ItemstoCheck.Count
' ------------------------------------------------------------------
 
If ItemstoCheck.Count > 0 Then
  ' we found at least one appt
  ' check if there are actually any items in the collection,
  ' otherwise exit
  If ItemstoCheck.Item(1) Is Nothing Then GoTo ExitProc
 

 
  ' create/fill array with exported info

  
  ColCount = rngHeader.Columns.Count
    
  ReDim arrData(1 To ItemstoCheck.Count, 1 To ColCount)
  
  
  For i = 1 To ItemstoCheck.Count
    Set ThisAppt = ItemstoCheck.Item(i)
 
    arrData(i, 1) = ThisAppt.Subject
    arrData(i, 2) = Format(ThisAppt.Start, "MM/DD/YYYY")
    arrData(i, 3) = Format(ThisAppt.Start, "HH:MM AM/PM")
    'arrData(i, 4) = Format(ThisAppt.End, "MM/DD/YYYY")
    'arrData(i, 5) = Format(ThisAppt.End, "HH:MM AM/PM")
    arrData(i, 4) = ThisAppt.Location
    
    If ThisAppt.Categories <> "" Then
      arrData(i, 5) = ThisAppt.Categories
    Dim CatArray() As String
    Dim intCount As Integer
    
    CatArray = Split(ThisAppt.Categories, ",")
    
    For intCount = LBound(CatArray) To UBound(CatArray)
    Select Case CatArray(intCount)
        Case "Northern California"
            arrData(i, 6) = "small_red"
        Case "Central Valley"
            arrData(i, 6) = "small_purple"
        Case "Fresno Area"
            arrData(i, 6) = "small_green"
        Case "Southern California"
            arrData(i, 6) = "small_yellow"
        Case "Bakersfield Area"
            arrData(i, 6) = "measle_turquoise"
        Case "Bay Area & Coastal"
            arrData(i, 6) = "small_blue"
        Case "Las Vegas"
            arrData(i, 6) = "measle_brown"
        Case Else
            'arrData(i, 6) = "caution"
    End Select
    Next
    
    End If
    
    
    arrData(i, 7) = CalName
  Next i
 rngStart.Offset(ItemCount + 1, 0).Resize(ItemstoCheck.Count, ColCount).Value = arrData
 ItemCount = ItemCount + ItemstoCheck.Count
 

    
Else
    MsgBox "There are no original appointments or meetings during " & _
      "the time you specified. Exiting now.", vbCritical
    GoTo ExitProc
End If
 
' if we got this far, assume success
GetCalData = True

 
ExitProc:
If bWeStartedOutlook Then
  olApp.Quit
End If
'GetCalData = False
Set myCalItems = Nothing
Set ItemstoCheck = Nothing
'Set olNS = Nothing
'Set olApp = Nothing
'Set rngStart = Nothing
Set ThisAppt = Nothing
End Function


Sub AddRaters()
'Listing Rater Locations
    ReDim RaterData(1 To 20, 1 To 7)
    RaterData(1, 1) = "Will Barrett"
    RaterData(1, 4) = "North Highlands, CA 95660"
    RaterData(1, 6) = "man"
    
    RaterData(2, 1) = "Harry Williams"
    RaterData(2, 4) = "Sacramento, CA 95823"
    RaterData(2, 6) = "man"
    
    RaterData(3, 1) = "Dain Cilley"
    RaterData(3, 4) = "Manteca, CA 95336"
    RaterData(3, 6) = "man"
    
    RaterData(4, 1) = "Tony Souza"
    RaterData(4, 4) = "Manteca, CA 95336"
    RaterData(4, 6) = "man"
    
    RaterData(5, 1) = "Mark Souza"
    RaterData(5, 4) = "Manteca, CA 95336"
    RaterData(5, 6) = "man"
    
    RaterData(6, 1) = "Mike Lehr"
    RaterData(6, 4) = "Manteca, CA 95336"
    RaterData(6, 6) = "man"
    
    RaterData(7, 1) = "Benny Hatcher"
    RaterData(7, 4) = "Modesto, CA 95350"
    RaterData(7, 6) = "man"
    
    RaterData(8, 1) = "Sean O'Day"
    RaterData(8, 4) = "Ripon, CA 95366"
    RaterData(8, 6) = "man"
    
    RaterData(9, 1) = "Graham Ralfe"
    RaterData(9, 4) = "Reedley, CA 93654"
    RaterData(9, 6) = "man"
    
    RaterData(10, 1) = "Steve Blase"
    RaterData(10, 4) = "Clovis, CA 93612"
    RaterData(10, 6) = "man"
    
    RaterData(11, 1) = "Joe Silva"
    RaterData(11, 4) = "Livingston, CA 95334"
    RaterData(11, 6) = "man"
    
    RaterData(12, 1) = "Gabe Lopez"
    RaterData(12, 4) = "Patterson, CA 95363"
    RaterData(12, 6) = "man"
    
    RaterData(13, 1) = "Alex Gonzales"
    RaterData(13, 4) = "San Fernando, CA 91340"
    RaterData(13, 6) = "man"
    
    RaterData(14, 1) = "Ken Waggoner"
    RaterData(14, 4) = "Murrieta, CA 92563"
    RaterData(14, 6) = "man"
    
    RaterData(15, 1) = "Raymond Espudo"
    RaterData(15, 4) = "Beaumont, CA 92223"
    RaterData(15, 6) = "man"
    
    RaterData(16, 1) = "Sam Maimone"
    RaterData(16, 4) = "Lake Forest, CA 92630"
    RaterData(16, 6) = "man"
    
    RaterData(17, 1) = "Glen Spatt"
    RaterData(17, 4) = "Henderson, NV 89074"
    RaterData(17, 6) = "man"
    
    'Listing for helpers, not certified
    RaterData(18, 1) = "Gordon Mitchell"
    RaterData(18, 4) = "Keyes, CA 95328"
    RaterData(18, 6) = "woman"
    
    RaterData(19, 1) = "Chase Allen"
    RaterData(19, 4) = "Modesto, CA 95350"
    RaterData(19, 6) = "woman"
    
    RaterData(20, 1) = "Luis Calderon"
    RaterData(20, 4) = "Ripon, CA 95366"
    RaterData(20, 6) = "woman"
    
    
    rngStart.Offset(ItemCount + 1, 0).Resize(20, ColCount).Value = RaterData
End Sub











