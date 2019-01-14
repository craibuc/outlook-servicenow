Attribute VB_Name = "RulesTests"
Sub Run_All_Tests()

    FindInFolders_Tests
    Find_or_Create_Folder_Tests
    AddToFavorites_Tests
    ExtractTicketID_Tests
    
End Sub

' --------------------------------------------------------------------------------
Sub FindInFolders_Tests()

    Const FUNCTION_NAME = "FindInFolders"
    Debug.Print FUNCTION_NAME

    FindInFolders_returns_reference_to_a_folder_that_matches_name
    FindInFolders_returns_nothing_if_folder_does_not_exist

End Sub

Sub FindInFolders_returns_reference_to_a_folder_that_matches_name()

    Const TEST_NAME = "returns_reference_to_a_folder_that_matches_name"
    Debug.Print vbTab + TEST_NAME
    
    ' arrange
    Dim Folder As Outlook.MAPIFolder
    Dim FolderName As String: FolderName = "RITM0071385"
    
    'act
    Set Folder = Find_or_Create_Folder(FolderName)
    
    ' assert
    Debug.Assert Folder.Name = FolderName
    
    '
    Set Folder = Nothing

End Sub

Sub FindInFolders_returns_nothing_if_folder_does_not_exist()

    Const TEST_NAME = "returns_nothing_if_folder_does_not_exist"
    Debug.Print vbTab + TEST_NAME
    
    ' arrange
    Dim Folder As Outlook.MAPIFolder
    Dim FolderName As String: FolderName = "RITM0000000"

    'act
    Set Folder = Find_or_Create_Folder(FolderName)
    
    ' assert
    Debug.Assert (Folder Is Nothing)
    
    '
    Set Folder = Nothing
    
End Sub

' --------------------------------------------------------------------------------
Sub Find_or_Create_Folder_Tests()

    Const FUNCTION_NAME = "Find_or_Create_Folder"
    Debug.Print FUNCTION_NAME

    Find_or_Create_Folder_returns_an_existing_folder
    Find_or_Create_Folder_returns_a_new_folder
    
End Sub

Sub Find_or_Create_Folder_returns_an_existing_folder()

    Const TEST_NAME = "Find_or_Create_Folder_returns_an_existing_folder"
    Debug.Print vbTab + TEST_NAME
    
    ' arrange
    Dim Folder As Outlook.MAPIFolder
    Dim FolderName As String: FolderName = "RITM0071385"

    'act
    Set Folder = Find_or_Create_Folder(FolderName)
    
    ' assert
    Debug.Assert Not Folder Is Nothing
    
    '
    Set Folder = Nothing
    
End Sub

Sub Find_or_Create_Folder_returns_a_new_folder()

    Const TEST_NAME = "Find_or_Create_Folder_returns_a_new_folder"
    Debug.Print vbTab + TEST_NAME
    
    ' arrange
    Dim Folder As Outlook.MAPIFolder
    Dim FolderName As String: FolderName = "RITM0000000"

    'act
    Set Folder = Find_or_Create_Folder(FolderName)
    
    ' assert
    Debug.Assert Not Folder Is Nothing
    
    '
    Set Folder = Nothing
    
End Sub

' --------------------------------------------------------------------------------
Sub AddToFavorites_Tests()

    Const FUNCTION_NAME = "AddToFavorites"
    Debug.Print FUNCTION_NAME

    AddToFavorites_creates_a_folder_favorite
    
End Sub

Sub AddToFavorites_creates_a_folder_favorite()

    Const TEST_NAME = "creates a folder favorite"
    Debug.Print vbTab + TEST_NAME
    
    ' arrange
    Dim Folder As Outlook.MAPIFolder
    Dim FolderName As String: FolderName = "RITM0000000"

    'act
    Set Folder = Find_or_Create_Folder(FolderName)
    Set Favorite = AddToFavorites(Folder)
    
    ' assert
    Debug.Assert Not Favorite Is Nothing
    
    '
    Set Folder = Nothing
    
End Sub

' --------------------------------------------------------------------------------
Sub ExtractTicketID_Tests()

    Const FUNCTION_NAME = "ExtractTicketID"
    Debug.Print FUNCTION_NAME

    ExtractTicketID_locates_a_RITM_in_the_mail_item
    
End Sub

Sub ExtractTicketID_locates_a_RITM_in_the_mail_item()

    Const TEST_NAME = "locates a RITM # in the mail item"
    Debug.Print vbTab + TEST_NAME

    'arrange
    Dim RITM As String: RITM = "RITM0123456"

    Dim scenario(2, 4)
    scenario(0, 0) = "RITM found in subject"
    scenario(0, 1) = RITM
    scenario(0, 2) = "lorem " & RITM & " ipsum"
    scenario(0, 3) = "lorem ipsum"
    
    scenario(1, 0) = "RITM found in body"
    scenario(1, 1) = RITM
    scenario(1, 2) = "lorem ipsum"
    scenario(1, 3) = "lorem " & RITM & " ipsum"
    
    scenario(2, 0) = "RITM NOT found in subject or body"
    scenario(2, 1) = vbNullString
    scenario(2, 2) = "lorem ipsum"
    scenario(2, 3) = "lorem ipsum"

    Dim i As Integer
    For i = 0 To 2
            
        'act
        Dim MailItem As Outlook.MailItem
        Dim Subject As String: Subject = scenario(i, 2)
        Dim Body As String: Body = scenario(i, 3)
        Set MailItem = CreateMailItem("first.last@csmc.org", Subject, Body)
        
        'act
        actual = ExtractTicketID(MailItem)
        
        'assert
        Debug.Print "-----" & i & "-----"
        Debug.Print "Test: " & scenario(i, 0)
        Debug.Print "Actual: " & actual
        Debug.Print "Expected: " & scenario(i, 1)
        Debug.Assert actual = scenario(i, 1)
        
        ' tidy
        MailItem.Delete
        Set MailItem = Nothing
            
    Next
        
End Sub

' --------------------------------------------------------------------------------
Sub ProcessMailItem_Tests()

    Const FUNCTION_NAME = "ProcessMailItem"
    Debug.Print FUNCTION_NAME

    ProcessMailItem_moves_a_message_with_a_RITM_in_the_subject_to_the_corresponding_folder
    
End Sub

Sub ProcessMailItem_moves_a_message_with_a_RITM_in_the_subject_to_the_corresponding_folder()

    Const TEST_NAME = "moves a message with a RITM # in the subject to the corresponding folder"
    Debug.Print vbTab + TEST_NAME

    'arrange
    Dim RITM As String: RITM = "RITM0000000"
    
    Dim Inbox As Outlook.MAPIFolder
    Set Inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    
    Dim Dummy As Outlook.MailItem
    Set Dummy = CreateMailItem("first.last@domain.tld", "ticket " & RITM & " was assigned to you.", "lorem ipsum")
    
    'Dim Target As Outlook.MAPIFolder
    'Set Target = Find_or_Create_Folder(RITM)
    
    'act
    'Dummy.Move Target
    ProcessMailItem Dummy
    
    'assert
    'Dim filter As String: filter = "[Subject] = '" + RITM + "'"
    'Debug.Assert Not Target.Items.Find(filter) Is Nothing
    
    'clean
    Set Target = Nothing
    Set Dummy = Nothing
    Set Inbox = Nothing

End Sub

Sub ProcessMailItem_moves_a_message_with_a_RITM_in_the_body_to_the_corresponding_folder()

    Const TEST_NAME = "moves a message with a RITM # in the body to the corresponding folder"
    Debug.Print vbTab + TEST_NAME

    'arrange
    Dim RITM As String: RITM = "RITM7777777"
    
    Dim Inbox As Outlook.MAPIFolder
    Set Inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    
    Dim Dummy As Outlook.MailItem
    Set Dummy = CreateMailItem("first.last@domain.tld", "lorem ipsum", RITM)
    
    Dim Target As Outlook.MAPIFolder
    Set Target = Find_or_Create_Folder(RITM)
    
    'act
    Dummy.Move Target
    
    'assert
    Dim filter As String: filter = "[Body] = '" + RITM + "'"
    Debug.Assert Not Target.Items.Find(filter) Is Nothing
    
    'clean
    Set Target = Nothing
    Set Dummy = Nothing
    Set Inbox = Nothing

End Sub

Function CreateMailItem(Recipient As String, Subject As String, Body As String)

    Dim oMailItem As Outlook.MailItem
    Set oMailItem = Application.CreateItem(olMailItem)
    With oMailItem
       .Subject = Subject
       .BodyFormat = olFormatPlain
       .Body = Body
       .Recipients.Add Recipient
       .Display
    End With

    Set CreateMailItem = oMailItem
    
End Function
