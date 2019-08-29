Option Explicit
'Global variable
Dim ImageName As String

Private Sub cmdBack_Click()
    'move up 1 cell
    ActiveCell.Offset(-1, 0).Select
    
    'check if it is header row
    If ActiveCell.Row = 1 Then
        MsgBox "First Row"
        'go down 1 cell
        ActiveCell.Offset(1, 0).Select
        Exit Sub
    Else
        Call GetTextBoxData
        Call GetOptionButtonValue
        Call GetImage
    End If
End Sub

Private Sub cmdLoad_Click()
    'Check if we have an image name in the Active Cell
    If ActiveCell.Column <> 1 Or ActiveCell.Row = 1 Or ActiveCell.Value = "" Then
        Cells(2, 1).Select
    End If
    
    'Get data from spreadsheet and put in TextBox of the UserForm
    Call GetTextBoxData
    
    'Select an Option button
    Call GetOptionButtonValue
    
    Call GetImage
    
    cmdBack.Enabled = True
    cmdNext.Enabled = True
End Sub



Private Sub cmdNext_Click()
    'move down 1 cell
    ActiveCell.Offset(1, 0).Select
    
    'check if it is empty cell
    If ActiveCell.Value = "" Then
        MsgBox "Last Row"
        'go back to previous cell
        ActiveCell.Offset(-1, 0).Select
        Exit Sub
    Else
        Call GetTextBoxData
        Call GetOptionButtonValue
        Call GetImage
    End If
End Sub

Private Sub UserForm_Initialize()
    'Switch off the button
    cmdBack.Enabled = False
    cmdNext.Enabled = False
End Sub

Private Sub GetTextBoxData()
    txtFileName.Text = ActiveCell.Value
    txtDate.Text = ActiveCell.Offset(, 1).Value
    txtInfo.Text = ActiveCell.Offset(, 2).Value
    txtDimensions.Text = ActiveCell.Offset(, 3).Value
    txtSize.Text = ActiveCell.Offset(, 4).Value
    txtCamera.Text = ActiveCell.Offset(, 5).Value
End Sub

Private Sub GetOptionButtonValue()
    Dim OB As Variant
    OB = ActiveCell.Offset(, 6).Value

    If OB = "Yes" Then
        OptionButton1.Value = True
    Else
        OptionButton2.Value = True
    End If
End Sub

Private Sub GetImage()
    Dim ImageFolder As String
    Dim FilePath As String
    Dim FullImagePath As String
    
    ImageName = ActiveCell.Value
    ImageFolder = "images\"
    
    FilePath = NavigateFromWorkBookPath()
    
    FullImagePath = FilePath & ImageFolder & ImageName
    
    If Dir(FullImagePath) <> "" Then
        Image1.Picture = LoadPicture(FullImagePath)
        Image1.PictureSizeMode = 3
    Else
        MsgBox "Could not load image - no such file"
    End If
    
End Sub

Private Function NavigateFromWorkBookPath() As String
    Dim WorkbookFolderPath As String
    Dim SlashPos As Integer
    Dim ImageFolderPath As String
    
    WorkbookFolderPath = ThisWorkbook.Path
    SlashPos = InStrRev(WorkbookFolderPath, "\")
    ImageFolderPath = Left(WorkbookFolderPath, SlashPos)
    
    NavigateFromWorkBookPath = ImageFolderPath
End Function

