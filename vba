Option Explicit

' 需要引用 Microsoft Scripting Runtime

Public Function DetectJSONIndentation(filePath As String) As Integer
    Dim FSO As New FileSystemObject
    Dim JsonTS As TextStream
    Dim line As String
    Dim indentCount As Integer
    Dim i As Integer
    
    indentCount = -1 ' 初始化为-1，表示还未检测到缩进
    
    Set JsonTS = FSO.OpenTextFile(filePath, ForReading)
    
    Do While Not JsonTS.AtEndOfStream
        line = JsonTS.ReadLine
        
        ' 跳过空行
        If Trim(line) = "" Then
            GoTo ContinueLoop
        End If
        
        ' 检查是否以空格开始
        If Left(line, 1) = " " Then
            ' 计算开头的空格数
            For i = 1 To Len(line)
                If Mid(line, i, 1) <> " " Then
                    Exit For
                End If
            Next i
            
            ' 如果这是第一次检测到缩进，或者找到了更小的非零缩进
            If indentCount = -1 Or (i - 1 < indentCount And i > 1) Then
                indentCount = i - 1
            End If
            
            ' 如果已经找到了缩进，可以提前退出循环
            If indentCount > 0 Then
                Exit Do
            End If
        End If
        
ContinueLoop:
    Loop
    
    JsonTS.Close
    
    ' 如果没有检测到缩进，返回0
    If indentCount = -1 Then
        indentCount = 0
    End If
    
    DetectJSONIndentation = indentCount
End Function

Public Sub TestIndentationDetection()
    Dim filePath As String
    Dim indentation As Integer
    
    filePath = "C:\path\to\your\json\file.json"
    indentation = DetectJSONIndentation(filePath)
    
    Debug.Print "Detected indentation: " & indentation & " spaces"
End Sub
