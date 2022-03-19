Attribute VB_Name = "Module1"
Sub Parse()
    Dim rng As Range
    Dim items As New Collection
    Dim item As New Dictionary
    Dim i As Integer
    Dim cell As Variant
    Dim myFile As String
    Dim LastColumn As Long
    Dim sht As Worksheet
    Dim LastRow As Long
    Dim StartCell As Range
    Dim LastCell As Range
    
    Dim strPath As String
    Dim strFileExists As String

    Dim fso As Object, Cdrive As Object, objFile As Object, objFolder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    strPath = "C:\SwitchData\"
    strFileExists = Dir(strPath, vbDirectory)

    
    If strFileExists = "" Then
      fso.CreateFolder "C:\SwitchData\"
    End If
   
    Set objFile = fso.CreateTextFile("C:\SwitchData\StickerData.json", True, True)

    Set sht = Worksheets("StickerData")
    Set StartCell = Range("B3")
    
    LastRow = sht.Cells(sht.Rows.Count, StartCell.Column).End(xlUp).Row
    LastColumn = sht.Cells(StartCell.Row, sht.Columns.Count).End(xlToLeft).Column

    
    
        Set rng = Range("B3", Range("B3").End(xlDown))
        
        i = 0
        For Each cell In rng
            item("localeName") = cell.Value
            item("code") = cell.Offset(0, 1).Value
            item("theme") = cell.Offset(0, 2).Value
            item("grade") = cell.Offset(0, 3).Value
            item("imagePath") = cell.Offset(0, 4).Value
            item("localeContext") = cell.Offset(0, 5).Value
            item("hiddenImagePath") = cell.Offset(0, 6).Value
            items.Add item
            Set item = Nothing
            
            Next
            
 
              objFile.WriteLine (ConvertToJson(items, Whitespace:=2))
            objFile.Close
            MsgBox "StickerData 积己肯丰", vbOKOnly, "Json Create"
            
       
End Sub
Sub Parse2()
 Dim rng As Range
    Dim items As New Collection
    Dim item As New Dictionary
    Dim i As Integer
    Dim cell As Variant
    Dim myFile As String
    Dim LastColumn As Long
    Dim sht As Worksheet
    Dim LastRow As Long
    Dim StartCell As Range
    Dim LastCell As Range

    Dim strPath As String
    Dim strFileExists As String

    Dim fso As Object, Cdrive As Object, objFile As Object, objFolder As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    strPath = "C:\SwitchData\"
    strFileExists = Dir(strPath, vbDirectory)
    
    If strFileExists = "" Then
      fso.CreateFolder "C:\SwitchData\"
    End If
   
    Set objFile = fso.CreateTextFile("C:\SwitchData\MusicData.json", True, True)

    Set sht = Worksheets("MusicData")
    Set StartCell = Range("B3")
    
    LastRow = sht.Cells(sht.Rows.Count, StartCell.Column).End(xlUp).Row
    LastColumn = sht.Cells(StartCell.Row, sht.Columns.Count).End(xlToLeft).Column

        

        Set rng = Range("B3", Range("B3").End(xlDown))
        
        i = 0
        For Each cell In rng
            item("code") = cell.Value
            item("package") = cell.Offset(0, 1).Value
            item("category") = cell.Offset(0, 2).Value
            item("service") = cell.Offset(0, 3).Value
            item("noteGroupCode") = cell.Offset(0, 4).Value
            item("isLocked") = cell.Offset(0, 5).Value
            item("localeName") = cell.Offset(0, 6).Value
            item("localeDisplayGroupName") = cell.Offset(0, 7).Value
            item("albumBgColor") = cell.Offset(0, 8).Value
            item("albumFontColor") = cell.Offset(0, 9).Value
            item("analyticsData") = cell.Offset(0, 10).Value
            item("isHidden") = cell.Offset(0, 11).Value
            item("challengable") = cell.Offset(0, 12).Value
            item("secondOrderIndex") = cell.Offset(0, 13).Value
            item("indexAlphabet") = cell.Offset(0, 14).Value
            item("oneStarMaxMiss") = cell.Offset(0, 15).Value
            item("twoStarMaxMiss") = cell.Offset(0, 16).Value
            item("threeStarMaxMiss") = cell.Offset(0, 17).Value
            item("artistCode") = cell.Offset(0, 18).Value
            item("orderIndex") = cell.Offset(0, 19).Value
            item("isFavorte") = cell.Offset(0, 20).Value
            item("playCount") = cell.Offset(0, 21).Value
            item("player1Character") = cell.Offset(0, 22).Value
            item("player2Character") = cell.Offset(0, 23).Value
            item("musicState") = cell.Offset(0, 24).Value
            items.Add item
            Set item = Nothing
            
            Next
    
            objFile.WriteLine (ConvertToJson(items, Whitespace:=2))
            objFile.Close
            
            MsgBox "MusicData 积己 肯丰", vbOKOnly, "Json Create"

End Sub


Function FolderCreate(ByVal path As String) As Boolean

FolderCreate = True
Dim fso As New FileSystemObject

If Functions.FolderExists(path) Then
    Exit Function
Else
    On Error GoTo DeadInTheWater
    fso.CreateFolder path ' could there be any error with this, like if the path is really screwed up?
    Exit Function
End If

DeadInTheWater:
    MsgBox "A folder could not be created for the following path: " & path & ". Check the path name and try again."
    FolderCreate = False
    Exit Function

End Function

Function FolderExists(ByVal path As String) As Boolean

FolderExists = False
Dim fso As New FileSystemObject

If fso.FolderExists(path) Then FolderExists = True

End Function

Function CleanName(strName As String) As String
'will clean part # name so it can be made into valid folder name
'may need to add more lines to get rid of other characters

    CleanName = Replace(strName, "/", "")
    CleanName = Replace(CleanName, "*", "")
    etc...

End Function
