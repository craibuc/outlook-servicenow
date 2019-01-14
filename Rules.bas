Attribute VB_Name = "Rules"
'Usage:  add to a rule using "Run a script"
'Revisions:
'  2019/01/14 - Craig Buchanan - creating code
'Notes:
'  enable 'run a script' option: https://www.extendoffice.com/documents/outlook/4640-outlook-rule-run-a-script-missing.html

'
' controller for processing an Inbox message that may contain a RITM
'
Sub ProcessMailItem(Item As Outlook.MailItem)
    
    On Error GoTo ProcessMailItem_err
        
    Dim ticketNumber As String
    ticketNumber = ExtractTicketID(Item)
    
    If ticketNumber <> "" Then
            
        ' if there is NOT a folder that matches the RIMT, create one
        Dim Target As Outlook.MAPIFolder
        Set Target = Find_or_Create_Folder(ticketNumber)

        ' move message to folder
        Item.Move Target
    
    End If
    
ProcessMailItem_exit:
    Set Target = Nothing
    
    On Error GoTo 0
    Exit Sub
    
ProcessMailItem_err:
    Select Case Err.Number
    Case Else
        MsgBox Err.Description & " [" & Err.Number & "]", vbExclamation, "Error in ProcessMailItem"
    End Select
    
    Resume ProcessMailItem_exit
    
End Sub

'
' find a ticket number (matches pattern RITM9999999) embedded in message's subject or body
'
Function ExtractTicketID(Item As Outlook.MailItem) As String

    ' pattern
    Dim ticketPattern As String
    
    ' \d    Match any digit
    ' {}    Specifies how many times a token can be repeated
    ' $     Match must occur at the end of the string
    ticketPattern = "RITM\d{7}"

    Dim RegExp As RegExp: Set RegExp = New RegExp
    RegExp.Pattern = ticketPattern

    Dim Matches As MatchCollection
    
    If RegExp.Test(Item.Subject) Then
    
        ' extract RITM
        Set Matches = RegExp.Execute(Item.Subject)
        ExtractTicketID = Matches(0).Value
            
    ElseIf RegExp.Test(Item.Body) Then
            
        ' extract RITM
        Set Matches = RegExp.Execute(Item.Body)
        ExtractTicketID = Matches(0).Value
            
    Else
        ExtractTicketID = vbNullString
    
    End If

End Function

'
' locate a folder by its name; if it doesn't exist, create it as immediate child of Inbox node
'
Function Find_or_Create_Folder(FolderName As String)

    ' search Inbox and its children
    Dim Inbox As Outlook.MAPIFolder
    Set Inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    
    '
    Dim Folder As Outlook.MAPIFolder
    Set Folder = FindInFolders(Inbox.Folders, FolderName)
    
    If Folder Is Nothing Then
        Set Folder = Inbox.Folders.Add(FolderName, olFolderInbox)
    End If
    
    Set Find_or_Create_Folder = Folder
    
Find_or_Create_Folder_exit:
    Set Folder = Nothing
    Set Inbox = Nothing
    Exit Function
    
Find_or_Create_Folder_err:
    Select Case Err.Number
    Case Else
        MsgBox Err.Description, vbExclamation, "Error"
        
    End Select
    
    Resume Find_or_Create_Folder_exit
    
End Function

'
' Add a folder to the "Favorties" section of Outlook
'
Function AddToFavorites(Favorite As MAPIFolder)

    Dim MailModule As Outlook.MailModule
    Set MailModule = Application.ActiveExplorer().NavigationPane.Modules.GetNavigationModule(Outlook.OlNavigationModuleType.olModuleMail)
    
    Dim FavoriteGroup As Outlook.NavigationGroup
    Set FavoriteGroup = MailModule.NavigationGroups.GetDefaultNavigationGroup(Outlook.OlGroupType.olFavoriteFoldersGroup)
    
    Set AddToFavorites = FavoriteGroup.NavigationFolders.Add(Favorite)
    
    Set FavoriteGroup = Nothing
    Set MailModule = Nothing
    
End Function

'
' Purpose: recurse folder collection to find a folder
'
Private Function FindInFolders(TheFolders As Outlook.Folders, Name As String)
  Dim SubFolder As Outlook.MAPIFolder
   
  On Error Resume Next
   
  Set FindInFolders = Nothing
   
  For Each SubFolder In TheFolders
    If LCase(SubFolder.Name) Like LCase(Name) Then
      Set FindInFolders = SubFolder
      Exit For
    Else
      Set FindInFolders = FindInFolders(SubFolder.Folders, Name)
      If Not FindInFolders Is Nothing Then Exit For
    End If
  Next

End Function

