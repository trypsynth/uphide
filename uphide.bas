#include <wuapi.h>

BCX_Show_COM_Errors(True)

Dim As Object Session, Searcher, SearchResult, Update, Updates
Dim NumUpdates, I
Comset Session = Com("Microsoft.Update.Session")
Comset Searcher = Session.CreateupdateSearcher()
Searcher.ClientApplicationID = "uphide"
Print "Checking for updates..."
Comset SearchResult = Searcher.Search("IsInstalled=0 And IsHidden=0")
NumUpdates = SearchResult.Updates.Count
If NumUpdates = 0 Then
    Print "No updates were found."
    End
End If
Print "Enter the numbers of the updates you want to hide, seperated by spaces, and press enter. Leave blank to exit."
Dim Index = 1
Comset Updates = SearchResult.Updates
For Each Item In Updates
    Dim Title$
    Title$ = Item.Title
    Print Str$(Index, 1) & ": " & Title$
    Index++
Next
Dim Input$
Input Input$
If Input$ = "" Then End
Dim NumSelections
Dim ToHide[100] As String
NumSelections = Split(ToHide, Input$, " ", 0)
For I = 0 To NumSelections - 1
    Dim CurIndex
    CurIndex = Val(ToHide[I])
    If CurIndex <= 0 Or CurIndex > NumUpdates Then Continue
    Comset Update = Updates.Item(CurIndex - 1)
    Dim Title$
    Title$ = Update.Title
    Update.Ishidden = True
    Print "Hid " & Title$
Next
Pause
Uncom(Searcher)
Uncom(Session)
Uncom(SearchResult)
Uncom(Update)
Uncom(Updates)
