VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AskDate 
   Caption         =   "Choose Date"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "AskDate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AskDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub OKbtn_Click()
    UserDate = MonthView.Value
    'MsgBox Selection
    Unload Me
    
End Sub

Private Sub MonthView_DateDblClick(ByVal DateDblClicked As Date)
    UserDate = MonthView.Value
    Unload Me
End Sub
