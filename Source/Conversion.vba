' Andraz (@akavalar), March 2016
'
' This is free and unencumbered software released into the public domain.
'
' Anyone is free to copy, modify, publish, use, compile, sell, or
' distribute this software, either in source code form or as a compiled
' binary, for any purpose, commercial or non-commercial, and by any
' means.
'
' In jurisdictions that recognize copyright laws, the author or authors
' of this software dedicate any and all copyright interest in the
' software to the public domain. We make this dedication for the benefit
' of the public at large and to the detriment of our heirs and
' successors. We intend this dedication to be an overt act of
' relinquishment in perpetuity of all present and future rights to this
' software under copyright law.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
' EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
' MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
' OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
' ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
' OTHER DEALINGS IN THE SOFTWARE.
'
' For more information, please refer to <http://unlicense.org/>

'main conversion sub
Sub ConvertVersions(Target As String, VersionExcel As Integer, wbQueries As Workbook)

Dim DestinationFolder As String
Dim DestinationShellFolder As String
Dim DestinationFinalFolder As String
Dim Destination As Variant
Dim DestinationShell As Variant
Dim DestinationFinal As Variant
       
'comments assume Excel 2013 -> Excel 2010 conversion, i.e. VersionExcel = 2010 (interpret appropriately if convert Excel 2010 -> Excel 2013)

On Error GoTo ErrorConversion
    'set file system object and location of temporary folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    TempPath = Environ("Temp")
    
    'target file name
    TargetName = ActiveWorkbook.Name
    
    'create first randomly named folder (to hold extracted target file)
    Randomize
    DestinationFolder = TempPath & "\Jebiga" & Str(Int(((99999999 - 10000000 + 1) * Rnd + 10000000))) 'folder name
    If Len(Dir(DestinationFolder, vbDirectory)) > 0 Then Call fso.DeleteFolder(DestinationFolder, True) 'if folder exists, delete it
    MkDir DestinationFolder 'create folder
    
    'destination file name (remove xlsx extension and add zip instead)
    Destination = DestinationFolder & "\" & Left(TargetName, Len(TargetName) - 4) & "zip" 'remove "xlsx"
    
    'copy target to destination (use new destination file name above)
    Call fso.CopyFile(Target, Destination)
    
    'unzip copied file
    Call UnZip(DestinationFolder, Destination)
    
    'create second randomly named folder (to hold Excel 2010 shell file)
    Randomize
    DestinationShellFolder = TempPath & "\Jebiga" & Str(Int(((99999999 - 10000000 + 1) * Rnd + 10000000))) 'folder name
    If Len(Dir(DestinationShellFolder, vbDirectory)) > 0 Then Call fso.DeleteFolder(DestinationShellFolder, True) 'if folder exists, delete it
    MkDir DestinationShellFolder 'create folder
    
    'use correct conversion technique depending on the desired final file
    'create zipped Excel 2010 file (if VersionExcel = 2013, an empty Excel 2013 file is created)
    If VersionExcel = 2010 Then
        Call WriteZIP2010(DestinationShellFolder)
        DestinationShellName = "Workbook2010.zip"
    ElseIf VersionExcel = 2013 Then
        Call WriteZIP2013(DestinationShellFolder)
        DestinationShellName = "Workbook2013.zip"
    End If
    
    'extract contents of Excel 2010 file
    DestinationShell = DestinationShellFolder & "\" & DestinationShellName
    Call UnZip(DestinationShellFolder, DestinationShell)
    
    'move extracted PowerPivot model (part of target file) to the correct location within the second folder, rename it appropriately
    'version-dependent!
    If VersionExcel = 2010 Then
        Call fso.CopyFile(DestinationFolder & "xl\model\item.data", DestinationShellFolder + "\xl\customData\item1.data")
    ElseIf VersionExcel = 2013 Then
        Call fso.CopyFile(DestinationFolder & "xl\customData\item1.data", DestinationShellFolder + "\xl\model\item.data")
    End If
   
    'remove first folder
    Call fso.DeleteFolder(Left(DestinationFolder, Len(DestinationFolder) - 1), True)
    
    'remove Excel 2010 zip file created before
    Kill DestinationShell

    'create third randomly named folder (to hold the final file)
    Randomize
    DestinationFinalFolder = TempPath & "\Jebiga" & Str(Int(((99999999 - 10000000 + 1) * Rnd + 10000000))) 'folder name
    If Len(Dir(DestinationFinalFolder, vbDirectory)) > 0 Then Call fso.DeleteFolder(DestinationFinalFolder, True) 'if folder exists, delete it
    MkDir DestinationFinalFolder 'create folder

    'tell user to wait 30 seconds
    MsgBox "Wait 30 seconds to make sure final file created successfully."

    'create Excel 2010 file with embedded PowerPivot model
    DestinationFinalName = Left(TargetName, Len(TargetName) - 5) & "_Excel2010." 'create new name by removing ".xlsx"
    Call ArchiveFolder(DestinationShellFolder + "\", DestinationFinalFolder + "\" & DestinationFinalName & "zip") 'package everything into a zip file
    Application.Wait (Now + TimeValue("0:00:30")) 'wait 30 seconds
      
    'rename extension (zip -> xlsx)
    Call fso.MoveFile(DestinationFinalFolder + "\" & DestinationFinalName & "zip", DestinationFinalFolder + "\" & DestinationFinalName & "xlsx")
     
    'tell user where to find the new file
    MsgBox "Excel 2010 file with the injected 2013 PowerPivot model can be found at:" & Chr(10) & DestinationFinalFolder & Chr(10) & Chr(10) & _
        "File path used by the extraction tool will be automatically updated." & Chr(10) & Chr(10) & _
        "Don't forget to delete the new file(s) when done with data extraction!"

    'remove second folder
    Call fso.DeleteFolder(Left(DestinationShellFolder, Len(DestinationShellFolder) - 1), True)

    'update file path in the main sheet
    wbQueries.Worksheets("Extracting data from PowerPivot").Range("C6").Value = DestinationFinalFolder + "\" & DestinationFinalName & "xlsx"
    
    Exit Sub
    
ErrorConversion:
    MsgBox "Error: Conversion not successful."

End Sub
'extracted contents of zip archives
Sub UnZip(strTargetPath As String, Fname As Variant)

Dim oApp As Object
Dim FileNameFolder As Variant

'add backward slash to path if needed
If Right(strTargetPath, 1) <> Application.PathSeparator Then
    strTargetPath = strTargetPath & Application.PathSeparator
End If

'unzipping works on variant objects
FileNameFolder = strTargetPath

'unzip
Set oApp = CreateObject("Shell.Application")
oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(Fname).Items
  
End Sub
'archiving contents into zip archives
Sub ArchiveFolder(strTargetPath, Fname)

With CreateObject("Scripting.FileSystemObject")
    Fname = .GetAbsolutePathName(Fname)
    strTargetPath = .GetAbsolutePathName(strTargetPath)

    
    With .CreateTextFile(Fname, True)
        .Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, Chr(0))
    End With
End With

With CreateObject("Shell.Application")
    .Namespace(Fname).CopyHere .Namespace(strTargetPath).Items
End With

End Sub
'base64 representation of byte data, i.e. binary zip file corresponding to an empty Excel 2010 file with PowerPivot model removed
Sub WriteZIP2010(Path As String)

Dim Counter As Long
Dim ByteArray() As Byte

Base64_1 = "UEsDBBQAAgAIAAAAIQC1VTAj6wAAAEwCAAALAAAAX3JlbHMvLnJlbHOtks1qwzAMgO+DvYPRvVHawRijTi9j0NsY2QNotvJDEsvY" & _
            "bpe+/bzD2AJd6WFHy9KnT0Lb3TyN6sgh9uI0rIsSFDsjtnethrf6efUAKiZylkZxrOHEEXbV7c32lUdKuSh2vY8qU1zU0KXkHxGj" & _
            "6XiiWIhnl38aCROl/AwtejIDtYybsrzH8JsB1YKp9lZD2Ns7UPXJ8zVsaZre8JOYw8QunWmBPCd2lu3Kh1wfUp+nUTWFlpMGK+Yl" & _
            "hyOS90VGA5432lxv9Pe0OHEiS4nQSODLPl8Zl4TW/7miZcaPzTzih4ThXWT4dsHFDVSfUEsDBAoAAAAAAAiHe0gAAAAAAAAAAAAA" & _
            "AAAGAAAAX3JlbHMvUEsDBBQAAgAIAAAAIQBt9o6fyQAAAGgBAAATAAAAY3VzdG9tWG1sL2l0ZW0xLnhtbH2QTU/CUBREz08h7Est" & _
            "5UMMwoIEl250TUqhtYkUooUQf7x6+th0RV7SN52Ze2fy/n7nLLly4JMeF/Z88U3FkZpn+iQMePDuqdTk8jvVmjKo77yxJtI18W/J" & _
            "gjkvOg/BUzl121y7s/V/0HhOPBF7yo4zlq3MP6rnnPU34lb9IRPfGsXmZWzduOfVprvQtx9yV52pVXA3oXOjmqtc2NgyIZVNmTJz" & _
            "OhIXsploJHpkLNqqjnVGzhXikf5MZui3kEtNi+/mtXr3HRb8A1BLAwQUAAIACAAAACEADjToLtwAAACKAQAAFAAAAGN1c3RvbVht" & _
            "bC9pdGVtMTAueG1shZBNT8JAFEXPT0H3dcCCX0EIqdGN7jSaGENqsdhEitGKxh+vnikucGVeMr1zP9570++vIWM+WPBEhxUPvPBK" & _
            "xZKaY7bpsUPXb0elppCfqdbMW/WKS05JdO15GzNiyJnOReupTK071/aM/kca65kjgjXfcAbZyvlL9YI3/Y04qp/k4vVGgcxuVbtL" & _
            "w3Wbnam8c8MF506IG2Qb+azNNb+JqG5xK3vCxN0n4kLviqkv6JHqS9nn0K6JuJTNRX3RAQPRvepAZ2KuFPf15zK7nqVcyp0V54R/" & _
            "9gh//tSIH1BLAwQUAAIACAAAACEAt8ezgLMAAABEAQAAFAAAAGN1c3RvbVhtbC9pdGVtMTEueG1shY9PD8FAEMV/HwV3Vi8O4k+k" & _
            "QpxxEgelqSZsha00Pjzebi91kkl2Z95782bm8x4xpeLKhRZPUu48yCmwjOkQ0aOvvyXGchR+EmvJArtlw4KuVANVUyaMWEp5DZpc" & _
            "XbWzlafXn3GKG0OMImsojdBc8wvxR0rpnXLPvjgorzcymncgkWNKLKQU5ljpXQu32i0RWmmS3yRu+MSh34UrXGDb7ITOmclzpjxi" & _
            "r/CM+dNpfm6c8AVQSwMEFAACAAgAAAAhAK/cEBSwAAAAQgEAABQAAABjdXN0b21YbWwvaXRlbTEyLnhtbIWPPRPBQBCGn58Seo5G" & _
            "YYLJ+Oo0qIzC15AhJxNB+PF479JQmS12b9/n3dt9v0J6FCScCLixI+NCzBlLhypN6jSUAymWjfpbqZa9V+fMGFET1dKrR5eQscjE" & _
            "M7Fc5WSrmY4/kCtS2hjF/os03DU748ha+ay84SpXrtoxT1aqy70MU72sNnFswcR3h+JSMQ/94/bof/n7nsj9DblXKyzUHRDpgkh1" & _
            "k6XCKeaP0/xc2OUDUEsDBBQAAgAIAAAAIQDPW9xRsgAAAEoBAAAUAAAAY3VzdG9tWG1sL2l0ZW0xMy54bWyFj8sOAUEQRc+nMHua" & _
            "jYUMMvGKNVZiYRAEPUJ7fjxu92zGSm7SVV33karPO6bDgyMHStxYc+bCjgxLi4g6VWqqJTGWpeYrsZZNYKdMGFCRqqFfhzYxQymP" & _
            "QbOTK0+2yvT6LU440cQIm4LScFf2mT2paqa65CqXU+81Lxbq870MIzFjTay28foHfalS5a0E/0Zhl24hoxu8LtzhAltmpmmPRFck" & _
            "6p/iLswFz5o/bvNzaZsvUEsDBBQAAgAIAAAAIQATH2YUuAAAAFoBAAAUAAAAY3VzdG9tWG1sL2l0ZW0xNC54bWyFj0sPwVAUhL+f" & _
            "gj23jcRCPNIQtl1gIxZeoaGt1FXix2N6a1ErOck9c8/MmZx5v3oMeRBzpkbOnowrESkJfRr4tPDUa2IStprvxCYcHDtnxoSmVB39" & _
            "hgzoMZUydppIW6VzIs9Cf8SqLnQxqkNFabjLO+PERj1V33LTlhUuNE/WwuVdhlD97m4NNcv1syx+bm+4W0YVj5GbW5fDOrbOUtMx" & _
            "gVIEwv43bYu2Xs8hj5WqUJs/buYn+YAPUEsDBBQAAgAIAAAAIQB/4pjf+wAAAGACAAAUAAAAY3VzdG9tWG1sL2l0ZW0xNS54bWyt"

Base64_2 = "kUFPwlAQhL+fUrnDQ0k4kFpCIHLBE5iYEA8VSGmAQmgB448X5209VAVPZtO+7e7MdGff+SOkyxsb1gQcWbAnJ2VLxj01bmnQ1Bmo" & _
            "kzFTfa5uRmLdJyY8UBeqra8uESFDITeGScUqlTNpevySQrGjg1MkFaTjJO09K151bnXOOIhVKPeYd2Ll5VyOkbFW4s81QSzOWnmu" & _
            "P/gJ+hVm3xiFTV9Y94apqgN6YvaUh1fVqtN3lOfq/3ZxsmjoaentPSSq3mlrTe3F8cyj/jCWn6X5jW2bP5XTf1GumxPvPbbbWnxt" & _
            "5LJDX/XYQFoedX2vES+KEvP3dt23+4/4BFBLAwQUAAIACAAAACEAfvP828YAAACIAQAAFAAAAGN1c3RvbVhtbC9pdGVtMTYueG1s" & _
            "hZDNDgFBEIS/R1nuDBcHWStiw81BOIkDS9hgCeP34VEzElYikc6ke6qqq3vmcQ9pcmHDmoATc/YcSNmS0aBIlTIV5UBMRiJ8JjZj" & _
            "4dkhAzqUpKrp1iQipCvlxmtSdb2cM3k6/RKr2FHHKBY5peEs7z0rpspb5YSjuqxqp7kxUf3ay9BX5zqHHOScyjcWbnWSN9PTba6d" & _
            "3YmFfZiJpl21k9u5nZvV9qz177WeLTASGtPSa1uqw59OgfaKGCtcj/njab7+KeIJUEsDBBQAAgAIAAAAIQA4TnOIzAAAAH4BAAAU" & _
            "AAAAY3VzdG9tWG1sL2l0ZW0xNy54bWyFkM1OwkAURs+jVPc6umFhagmB6Ia4MLoyLrAWaIRCoAr68OCZITGYmJib5t4738980/0u" & _
            "p8uWOTMyPqhYsaZmQcM1p1xyzoU9E2koPX8VbZgk9JEHbjiT1XHrUpBzK3OeOLWqg3OjZ+RPaa0lVwRrcsQMbPRe8caLfWEveVfV" & _
            "OkfOFyPnQ67gnbXamVmGTpF17z5WX7lNubNXovEbyCh/tCM5nyaJSftHN/QT2qZXtgk94cnTAT3f2HPO/3TKTFPwbEVN+Mcz/Po7" & _
            "Bd9QSwMEFAACAAgAAAAhAIn+84rDAAAAagEAABQAAABjdXN0b21YbWwvaXRlbTE4LnhtbIWQy04CMRSGv0dB91LduDBcQkDdG1wR" & _
            "FjgQmDAMZqiX+PDC15KQcWVOTk97/ktPe/ztMeSbHRUdPlnRcKBkT02fa+7ocmvtiNQU9peiNeuMvjLliRtZ956GDOjxLHOXOaWq" & _
            "s3OtZ+JviMY7DwRj3WIGvvRu2PJm3VsLPlRF94nzw8L9ea7Ai8qq1TnoXOo7UpMUE/FoFhfGo7nQu7K7dJI06bh1wzizYn5lzOgV" & _
            "M7sTPafmzLWRv2JuJDz8ow9/fmLACVBLAwQUAAIACAAAACEAiqyKLFwNAAAIhgAAFAAAAGN1c3RvbVhtbC9pdGVtMTkueG1s7V3r" & _
            "c9u4EcfnzvR/UN2b+9CpTTvKJY6j6MbPnHvx4ywnd+0lo5EoxdZFkl2T8uP++LY/LAASJAE+FL2ug9FIIoHFYrFYLBa7IPjf/zTY" & _
            "9+yRjdiQ1dg967M7FrABu2Fj9oatsS22wTbxX0POmPlI7yF3zK4o9z27ZEdsHVAvcPc9a7IGewvIEcEMUEpgHgMnh79mIT63bId5" & _
            "+FxpkB57AO479oV18X+Df59NUCrENYf5nXVwLejy2AlKhkjpUWqHXeCe0ybof2LHgPpM+EZauTWib1/Du0/pIbUtpNy/sF+ResB2" & _
            "0bJdXH8NdybA+TnFnX1g6BNNfZTaBd4hSj2wM4n9Ad8B5XIMIe4muF4r4GSAuq+Jmx1cbxDHfJS9wd0NaAiR5lOLPUCM8R2CS7wl" & _
            "AT683nsq0cedx56hTZusjqs+tYW3qC+5d4b++Y24HYJLfeAW/ad4zGEOZK90iZp+gvYdXAfgUrYND/TZwLeOX953VxEtW7j6Bb3+" & _
            "jrW0lq4ZMA9mgnmdWhVQK8bEl3RdPXyGwLGcvuDpNoraS6XKTtkWlWnL/8VQuJ2hUNSu/m08XBSFW0RLmodFlLVXgsJ8SusLpHAr" & _
            "Q2G9kLL2SlCYT+nzBVL4LEPh80LK2itBYUwpn3uOMfs0ieObKPEdfsVsfgRdsAnIbXaIlF2k7QGmgdKiRIOdgooR8DYxG6SpbWN2" & _
            "+InmCEEb/20DS5LyVopyjj3G2iD4MTCGmk0SWHOSNLWITz2yjh7Zh4Qlkq7nA1E1keU43ZwfbZqvBcUDKqPD8XsbFV4O5bpe2iF7" & _
            "agQpEFBdfIdkGzyBhj7xbRj1ziZhnqZ8w6gNd9DWkGb3DlmWh5qsvCep6qEstwxPSDJvyFJM01ANWwN0dShnIuGa1LY6PhxzNje2" & _
            "jlrAPgF+LivH1GrRpzfUR4rLypIVsm2DOiHLk+MXNlkHmIfIFfJdXIrDTU/XAdJGJFtBQjay6fEI9ZF2D26/AL/qgKmzl+wVqFkn"

Base64_3 = "+l+QDfYcV9sYw+uQhJf438KVj7SXyKkD4jv0VQf3vuR2diSLWmY5Do8B+RONGS49h2TvCzlRUpE3Go/omsNVGX/FVJkpEvpoIEe8" & _
            "voJK0/g1dV+SpT/E9c/I75GchxXq1rnwLUnt6xmuSL9FOseoMOetTqu2ZC2D3cSt5BozXSK93qzCgbh1f2Z/ikqaW/GW2tDLac3/" & _
            "xypN50gNH8UV3gtcykaAbOH/ivgSUL/0wLMn0nM11FMGw5HmWyhbZhcU+qRZO/h/Kl1un8rcSWl/AvVPSO8SRFkc52TLiFU6l6dz" & _
            "ojwkWRiXxnIK+CspMdNjEVIZoheDDLxeRnhBRkYYHe4LoJ6sUDpkQP4UoSsUfCfK9Yz5tlq9gnoV3H2k3cpQOEjoiBeQ/ri+ZJ6d" & _
            "rvwadXx2/i62D7quD5beB/5K98HWEvoghiirvUza7oBWhre470i/NLe4H53my+3xTaf3ltwDW07rLbkHnq2Mziurw8yWMvej/U7W" & _
            "a3mL95oiTF/INi9vkw8olVujeWVa5EUSESyx8ilbUofjtu+dXMWGcoWWV/aY/IN81SDiZgGtV8bR2lWV+xx5COK+KFtWXwV6U68D" & _
            "FbZP+KRXq94UK1wvseZWubPyP+reqppcW+6Qx/BWxlBNfq0PgOpjpblHVCk+rln9cwpeeJiSvlib7yxZhkPpuTzlPbB8oXY9RF7J" & _
            "rpTCeCQr74TJS8kp3qU4800U6faxTptQXJnXaa+DU3BIK0suWaIPeTT3Sq5R9R74kbTdKfXAEbViLL2pV4T7jjykChev11ZC4TpA" & _
            "yi3Jni+j4U2K4t9KOd8leOGDvdUwmso1ohbrtcaxdeGrzcdeBgOHK8uxpPdYeVqr9KvJU23GOp0f2i4ZsXcuCZMvT0kPayhn0y5J" & _
            "Zz/yr6bTY4/wBeE8JW0uMJp9umY4vfbYm1Oe22asCpPgVxBxT0iJ4Eeg8VSk67NSYEnXNU1cy7HUo1cRBVkYXUMNaGZtsucpLaTS" & _
            "45FxTlIr4lKB1HfN1JjNhy2ja82RlCxv03qXz5xNbZSnOeVZOav6vZi3P5Nd0VkIX/9FrbxhZ9TWIY2+LyvDYxvPsqPzgmwlfc69" & _
            "xtWtdTTb4M3QYmx35Axvzm1EO6OE3dJHu/rUj/dUB5/HTqkuEVMphtZ1SccYCfm6Vk3Dg+4SedBdER74S+SB/9U88CqPHq4benKk" & _
            "PiXmlHSOiWc/kHV5R/qMR0SeyN4MCGJIbficiDVOg8Fczj6Pf01cVY+o5EUIG6nYi37HteNbmutFu4aGfZO6HpzF+kPXYVlbJatf" & _
            "lmUhrMvdLtXnsnOy7oNo50uZWYxTeS/Xlj32C/WAmn07tPIWK2ZhlY/kzl3Bw7Ily8yVOj9Na714jTy/XQnpOhoJqnT5yaabVo2r" & _
            "ZiNNL1nTW0lOvuYpXzZ5qTJjuZmj3MzRtcwcXTdzuJFdYmR33czhZo6VkS83cyxu5vCte17dzOFGdvHI9t3M4WaOlZGvec4cnjUe" & _
            "Jfq6I/cC3FP9HfKG6R75AbV0LJ+IGMv+Ej3KZfEc/NgCnGhZPrSI1IXRbt50+XRuuv60X7MaPcWl1VznS/8l5xfH0CusqXy5MlJe" & _
            "1DMc67XcwyD2cNvjHWV62bM8reJZn27h+7K7kSSpu3h+VtHZofUpMz1/Vs+nJJ//OaFnzQPidT/Xxmnjw/3nNcqNy/BdLb3oOW/1" & _
            "fMssrJhpnpz6Yz1nlKVrfjTZ6qoai7fFX8zxcXu05o8SsZguJjmbWroLqcWfshb7rFldVyb10FvSx5MoEmbLXY1RfJLSh9m0RdPZ" & _
            "JitpQvNqzfqMI3/C6S56YqeP+XlCVkCsn2MsoreLS5js9z3aAXUc4TFZ8Gl7bzosWWt/27oKKLYwhEZbHQvaZBWnU/I1pkkbZiXV" & _
            "M0r0overuZ2H89p5aFq/xpy/oPE9oZXvnWYd2+CrWF22/bfCc6Q8ZCFJglnjZ3GuJazrRVpURXUuy7JajDdvXn6Ist4DM1fm5a2K" & _
            "W/6WZHQcjRBx+kR6L+hirEznt3V+tUXvFVnMysZJtpPsRceyF7OadpLtJHvRsbbqHhzbiVeiH0w2b3Il19Rge7lnaJXHOI1fSTyF" & _
            "GWrnj9rS3clT7uQpd/KUO3nKnTzlTp5yJ0+5U4/cyVPuDBZ38pTTfO7kKXfylNN67uQpd/KUO3lq1idPxb6RsudMCT9Snd6c47Nt" & _
            "2hH5EqvNHq1G6+SZqsMq4Vev8O9T7jN614F4k0SXvFLbuH+Vc6q78FbF9Jnf5NNih1j38vclXKKFv0qf1l8B8YjvJ6yV/ybH9xG7" & _
            "YGdYJZugOBVFNZn8oIvekTLrfRNuh8t8d7i4Xftu175n8fR7OZEBJztOdkz7IJP7jr2CPcu2Ha3zPzXyBDgeyRt9R6cJhdGe5HS6"

Base64_4 = "/ixKOs/2TIoJTo+ZiDf0jGRbAkOaiLv2afZv0ruW3uH3PX55pOEQFliN7MV9SM2hTKuh5kN89/B7gTsOf0H5p8jndzxnI7NzdQP2" & _
            "RvknVz7RGxdb+N2KqODUXWq17hE15vpr+D+XERN+dwAOnhBMC1dn+K9lKPw7wenP5LRTJ32+IaupSis47SLyrLisosvJXvCMveVZ" & _
            "JMWzStZyLDGno52O9hLP1XmpZ+5M+tOeY3sKIHlGG+eFqbT+drlZrs+yOliN6toc9hAIu3qc8EGouP170KK3OC9armj8BivAb6R3" & _
            "QkXlrsHXfXD5o9yXwP0TH0Hbj3T6XkeuoD6i3eIM2SFJQY+gRIs35NpRYJ/Vet3U9pbmeWtKibyXo/AO7TgxvLdwF5j4XHRGa2P+" & _
            "HsENejshj4y/jmSnluDltBx5DZpUDD4gmWgR/ZNoZ3ItehPbG82D8xqUPUqN1ZO+LNE68Ta3kKwMtfdA9OVr9gMouEDaPymPe51O" & _
            "kMtpb6EHDmikDMifIsb2GmD+Qb6dmsaNHblWHki/5hPVrSSjJfn4IPcFfaQI8YP0bHE9wPvuYzR75fdYY4ZvzjPB9BMx8DDqvyrv" & _
            "1StH4aXc4XRD+7SapMM3pQ5P59lsR3tOUGijFkPE+i9ryWa1WBYm+WTssn1f+m4Ofv9IO0T095bWSA/wEXJK7Y+xcYk3n/KS3m0y" & _
            "j10sI9ofI2gRbxnneyB2Uu9cXTe8c3VdvnN1R+JbT+Ay1XMrdcYs6lG41jR+C3/4SI6uGo2OkZyZbVzXad4hD3USJv3e9CT8O/LZ" & _
            "daheDslrXse4biXo8uV7SIekG/Wdtyr/GhADaVcID/NZtMtG9PumrL2DMuk87rHrRp47oUvLcWXWuxeLOaQgVP8JnnOoK+kpHMp5" & _
            "JMnzGD7f+/uGSqjPbHzONQMVRzRviBlxSH7INFezZVTMQ62HYortJbpa/CNbh6IxWzIudanZp0r7VpFOPkf9W67L/EROnmR1KvKs" & _
            "Y2l9ct9WGl63vQVNIsK4VjCOvJKt6FZsRbdiK7oLaYVfWX6rtcKfQSu8HGnzCiXVM3IhWTbWsbPCGaRmfdMck3d6W9pq8AotJq+E" & _
            "1RXDdGlOVa05w/1v0vq1x832yTpV8RoeO43XyoF8M0wyTiry9dhok/0PUEsDBBQAAgAIAAAAIQBuqLc+OQIAAJYIAAATAAAAY3Vz" & _
            "dG9tWG1sL2l0ZW0yLnhtbM2We2/aUAzFz0fp9j8Nzz0qRoWoOlXapkm02qRpqkICNBqEKQld6Yfv9rsOsEAz1u4hoYjga59r+x4b" & _
            "X77ftXWsG0010YGuNVSiVJFmivVKT1XToap8H2CJFaAPscYam/VC5zpVBdQzVsfqqK3XIKeGidiVe47x6fBXyni+6kgez7iA9NBG" & _
            "xJ9hDzQHnyE766185Dwjj3i+Bngc6qPe6o0uQafsuySDmhroG3qul2RZQR6h9ZGaSC/UQhpgbYGssG+E3ATvo6nzHqFrkKU7Ra+Q" & _
            "Q89iZ8ZAZtYn+oT2RF3y6SK3C3l9INfQzpbBRWKrvvEagZlsnajI0BFyCvo+U9/sOeTT4D3D1xhtncpUOYu35KJP/lfGqW8V2/Yc" & _
            "/RPPFauY48a3jhiuGZsRbQ4mBj/Hj+MgNd5C+FkQc0hW3j30qcWdGitl9i5RXE8kIAL8lGF6Zk+WXboggwX6gSHK8O+RHSOZdd2Q" & _
            "tW+8uDrFpTveYRsvK/ewHXknZDCXmi0y9NTkL0gLkxxDyfI31cFnG0+bOqf5ib+2Lpqzzn3mPdmk0x0uWveot4X0NuI/JJfBHuUS" & _
            "/HEurUfl4u2sYNF2gi2lpyd4W+gMTYiXm/9e6ere1Lm2N1Wu/0WNd9Vxc0olrG5twpRPIDchA8s1/MWci5DzibGyu5thYvtWN0IZ" & _
            "qqhzc8dNwqndMPl9vcKdIfdtmmacy03fwPIN12yOjJcUrePh9/j2xp37uLuto888K8Z33afexr+Gjn4AUEsDBBQAAgAIAAAAIQAM" & _
            "dmRUnQEAAFAEAAAUAAAAY3VzdG9tWG1sL2l0ZW0yMC54bWy1U11PwkAQnJ+CvENbAYkESxDUF0hMwMTE+FBLRQItpq18+OPVuWmJ" & _
            "IITwYi7t7e3O7szt3X1/NdHCCiFmKGCBADESTDBHhCsU4aAMm3OBkQg+/SNGI4wVfcAQtygRdcFVCy6auCMyFGbCrKxyxJoG/4aU" & _
            "4x0NWBzjLaSFJWvHmOKF85yzjw9mpbQN5hMe7UyXhRsiY+E79PusG7C6Ye9sZXWETqU8VfQMT/R20abuNu0mbVPXQ187C6g1292A" & _
            "vogeo2bFHgzUmQm9M6kJGNvFHNJ0fP9JjgqJT8gxyrX4ue5YdsqIqTpmxjnPwkaVlo06/31q8cWa8HsVts1co3Ktc0xy5QshA/Fc"

Base64_5 = "q+6Uio364o7KBnH7SpcaZX6VPTUOrUcq6ZHpdz8l9THRfiIxZ+dzqEtdaducrsfY+o8m74TuhQd64esebJRWaDmcrb3T3L5bbSn0" & _
            "1L8i10Z1Txwp7lU/62OiOzBkXkjLFYd5ByUxlbiuM2o8DXpq+l9Sk8NR5VfjyhG6rphNHuskJoP733vr4pkjYzr+oqyd9+7iB1BL" & _
            "AwQUAAIACAAAACEAJ5x/tbcAAABOAQAAEwAAAGN1c3RvbVhtbC9pdGVtMy54bWyFj0EPwUAQhb+fUu4sFweh0lS4cKuTOFTbVENL" & _
            "aEX8eLzdXjjJHHZ2vjdvZt6vCTMelJzwuJNx5UbBmYopXYb0Gej1RCoS1VPRitzRDRELelKN9JvhM2EpZek0hbpa50qeVn+gVlwY" & _
            "YxT5l9KoWmj+WTyhkb5WbumTWHm7kWHl9Ed1ppods5d7pj0u+ltdxtrtmGma3Sb88gqdR+0uqR3tsFV1TiCvQHmk6xvxncJy86ff" & _
            "/Fzr8wFQSwMEFAACAAgAAAAhAC6d9hK5AAAATgEAABMAAABjdXN0b21YbWwvaXRlbTQueG1shY9LD8FQEIW/n0L3XDYW4pGmUrGw" & _
            "wkosGpq6SV+hRfx4nF6bWsksZu6cM9+deb8mzHmQkdLhRsyFK5aCnCkeQ/oMlDtSco7qn6TmJE7dsSWkJ9dIrzkzJizlzJzHaupL" & _
            "zsVs/GcqRckYo0haTqOu1f+F9CO1/JXqRn0Sqf5uZNiIUXBnJa0U2cptpa/FijRVa/9Y2XO7BC1S4AiVu6Nyape9ugt8XeGrDkVI" & _
            "5Y85KBqH+UMwP9fO+ABQSwMEFAACAAgAAAAhAB8w4tICAgAAegYAABMAAABjdXN0b21YbWwvaXRlbTUueG1sxZVLT9tQEIXPT0Fd" & _
            "sEscE2h4BBAKpZVKhVS3iB0yjhMsJQ6ynZT0x5d+d4Jrg1HKCmQ5njtzZubMw87Dn76Oda+pJtrQQrEy5Uo0U6pDfZCvtjo8N7Ck" & _
            "itAPsaYam/WnfuhMLVAfOR3rSH19Bjk1TILXKnJKTIe/VcF1p315XOMa0kObkH+GPdIcfIHsrL8VIq8YefqGT4h1Ds+YXBmWoQIQ" & _
            "DhWTw3EY1CIMzLMw/gXWTfgUOtAJvhleS11opK/Yl7rkPME3Np2LkD1WG2BJyXRDtHt9si64+Bn9+R+ngIqG+HS1o+9kWd+VHPSt" & _
            "dcbFbFsPImLOOM1gVaCLrDJPW8ymQ1yPCXT4DWx+iVXxvHNVvbnNs+KwD6bJ45ddbe4uv67S8b+MPtIVdZ9bbSXblvUqt7rdrrhp" & _
            "bOLlul12/S37/HLuUhsRaaFrNtfHI+buaQ//FvIIratnG2mXaC349Hj6SBG6HpYuiB36EXKOOJVZvGd5Sn1V89Peh6/YgaFVGNrc" & _
            "U+uWk4vGXLaROrBz70lza07wdRyW9obnj9uyMGRseQa2V1PL0pxdCLsvxugMxOotLTEji5wTp+rDOnw9ZmC7GsOv0Kl9B/bof4n0" & _
            "raKnUdd71GNfWqUJ85sYtxXCdXDe4PoytsJUE2za3nOvvXf8lpVc+rBY99V19vo/w5H+AlBLAwQUAAIACAAAACEAP1SUkzsBAACU" & _
            "AgAAEwAAAGN1c3RvbVhtbC9pdGVtNi54bWyFkkFPwkAUhOenIHcoqNGEIKQp6kFNTJTExHiotZYmdGvaguCPV799HAQ5kA3ldWfe" & _
            "m5nd/nwPNdZKheZqaalUlWrlKuV0obb66qrHfwvEKWH/DdQpM3SqR12pA+uMt7FGGuoaZmGcnK7NZMdMz5+pYX1ooICVbTEDdnP0" & _
            "S/BEC/gNtUe/FFNvHAW6x18Kz7971oQqZlLFs0DBO4i2+iPra8x9Y+iRntmdKMR7SD3UA72OXK9wV7q0hL67+jf9hilrMuX4957n" & _
            "lqkBXYC0D6StyTWzxDF1FxWfK7YzdTbF1w2IV87oOObkezql6umc5x2aCVhJf6l344b0xmiu7dZqsvgbXBozNZ3IzqEwlV2PA1j7" & _
            "Pj9tdfmd7HnpUz3h4xadvzQdu8Pa0jjTbcMb6YXlTzw4cCPBzjcz0i9QSwMEFAACAAgAAAAhAGoy1EepAAAAOAEAABMAAABjdXN0" & _
            "b21YbWwvaXRlbTcueG1shY+7DoJAEEXPpyC9rjYWBiEEo/ZiZSwMGCQRMApq/Hj1sjRYmVvs7Jw7r8/bI+BJwRmHO0eu3MipKJnj" & _
            "MmHEWK8jUpIon4qWZJZuiVkylGuqX4CPx0rOwnpyVXWdS/Vs/Sdq6cIMI2U9p1E21/xKPKGRv1bc0hcHxd1Gho16VDxYK5NK7Vau" & _
            "nRv1qiLrri2tLR2wU3ZBqI1DxbHubMT3UsvNn3rzc5fPF1BLAwQUAAIACAAAACEAJEbsil0GAADMYAAAEwAAAGN1c3RvbVhtbC9p"

Base64_6 = "dGVtOC54bWztXftv2zYQvj9F869DraQdNiDwXHh5dWgeQ51uBbZhcG3P8WZbhiUnaf/4bcejZIkURUuKSXk1YcSmRPLuu/v4Eh/K" & _
            "v/904DU8wRxm4MEDjGEFIUwhgAV8Dy04hjYc4a+HMQsY4v0Rxi5gQrHv4Q4u4AWm+havXkMXOnCJKeeUZoq5uOQFymTp7yHCzxJO" & _
            "wMfPJJPSx7tT1B9g/BDWmD7CMIv9DAMMc0Q+nGFogDlX+D3HVC3SeZrJcUopI8IbUexX8CvePYMeou1huIPfK5LwCW7hT0nmNX4v" & _
            "6Jr5og19+mUpZjGajxgaS7n0loZo0z1ZO8BwG33IbBqQPxnKFYUjjAkwPMEcL9HrR/ANho7gO/y+Rm1DjAswf4CYWdoe4ZyhFYyx" & _
            "MEb6QCnHpOeUPDInLS0B4wmmyuN8pE8b/17lsBxj6APiuEI9qTUviL+QrFmQXs7ILnzaJaaYr5bE5wotmCLSCC1e4jXDf42/zKdr" & _
            "jJXz90nrCGUHaLkoiaO822i+oRxjvDtEaQ8Y5ytj2X1RUt7auxhfV4uukyvNaT4uMcSrGZXStEw/4fc2L1ySb0daGUm9CTB2jboX" & _
            "ZO0E3mL8J4pLQn3iNqJS5cWpQvhNyMtsSXP6WrnnVPcCamd43CVdrxFpXrucgmnOSpA1l5N+QSnmeJf5pqrlYm4ZgV72G2JmRK1T" & _
            "Vb3ZvLJWndw31GpEQrkMKO2MSvAvVIdHeO8xliWn78CPaNGSausQ/yKpxPXJwyuq/1nur0jy35UtlbV5kj4vrtmeoNkTdHsb7bKv" & _
            "dmuLXlp1y0Oq37L9c0FqNYvyGFhfWxXZW/gJ9cia1bKe54Nrra162X3smVg5ztaHLvWw61iWOkUHfqbeaF0Db5pTxlos09e2zvke" & _
            "oKhPSdr6vMRs7eV68/2G2EOe09guovZLrYXLeU7fmS0jA6pRUfxbvX//Whurkt+KxxQz/HAkoaL/vkXkf2EsG5HlS4HcEuV7dZnl" & _
            "bZKraO/RfT4aZmXvLG7FI0XpM62bl1n+zMDaKFZyvIJ2wjSWpOw2pf8djYcD8kZTGBgfS6rn+hbUNI4e6mVlQUTxB7Y26wZqSBGa" & _
            "Xlxy+VPJvqA6pdEHa7X2BRFrR5/2CM1UMfptCg3vcSPCxHI0x17S/tlvc1bkkyb6n+RJa9Cgvz3qgUOam5nSnABPYxsPH82Oc21/" & _
            "aB3Je/y9bwyNar4gGQk3oz15bvQaxmFnTFJ1vsaG/vQpLMzMHIjzNM0gS54ozbeddWbxbGGwO1ptZs7JlhV2anmR9n68KrDYxKV1" & _
            "v7kSzmr/MK7nSU+1D8iK541sIdDPZNtCcY5PYtwTWX6m9L0kjibx6CFssHWoMy9rC51uZcCuh/jTwIg4ZCuTptFkW1+GYmBd40fr" & _
            "Goe1NPqKec9+PHs4pDHIaBPnb+ba83OofK58iuHHzPjlIl7BZu2JOFv93HXrdsFc9DYMfP69LE6VTaG0Yj4osbI/V9g7JGuSdfRX" & _
            "GDrGX1+xCp7uc0h3KCR7HAaokfOT9hv53Qt57llvnFr6GeMf4Qav2YrWB0ly/XluUQa7SlFWWa9ONeT5aG3WbvkOkWy96NIulGT9" & _
            "VRWfzfmOVoBUeZIYsW5wVs5p/KfWWi51sVQZ0bZ0KkncU6t4HqoMxqIceunbsebTsvUivraVti5iKZHLkblSbqf+VFmp2UXt+YHu" & _
            "jRX1hvXNopVZbC2Bjy/T91VWqhwXZrnYtlLn/G/W/2VWKh0H5tujMiu1jgezPFRbqXZsNMOGfqXesdIMK8Vr3Y6RZhhR79RwbDTF" & _
            "hmqnimOjGTbK79RxDNl5/nN+t/3MUbxTy/nerO91O9Wc723NN5Xbqef4MMtH+Z2KjgmzTFTbqenY2C0b1Xaq2va+jO5QGCizW9dx" & _
            "YYeL/VibOFTv63ZtOw7scFBv57pjxw47xbv3HQO7Z6DOCYZmeDgk7+/X+vXh+d/sqRXHn2n+9mmMe0h+r3JSybFhvg+vekbLcWKa" & _
            "k+LTac73pn2vP5fn/G/a/885kejYMT/mrXMW0/FimhfdKVTnfTu1Ytv5W8fDrnjYfvJ4F76+id/UWnQa8orOiLKzu7e0801812U+" & _
            "9ss/d7f9fLYNXsTzlMe5tyI77k5qnXS3z93L/yl3vvJEu+p9AfXfzu4beJN+F37HT/Iecd17/X3hfw104T9QSwMEFAACAAgAAAAh" & _
            "AOWKP5qvAAAAQgEAABMAAABjdXN0b21YbWwvaXRlbTkueG1shY9BD8FAFIS/n1LuLBcHoU1T4eTGSRyaaqqJLtEl4sdjdl3qJC/Z" & _
            "fTszb97s+zUj4UHDiYg7JVdaas5Y5vQZM2SkOxJjKYQfxFqqwG7ZsGQg1USvhJgZKymboKk19XW28vT6I051YYpRVR2lEVpr/1l8"

Base64_7 = "wU16p96zT3L130SGtV5WfC7XLJyFMJ+qlL/fn3WmszDlQnYX2B47oQtSJU/VL4NHK8Ve5RXmj4P5+WHMB1BLAwQUAAIACAAAACEA" & _
            "4bIRtMEAAADrAAAAGAAAAGN1c3RvbVhtbC9pdGVtUHJvcHMxLnhtbF1Oy2rDMBC8F/oPYu+KnDhVrGA5NNiGXEsLuQp5nRis3WAp" & _
            "oVD671XpradhZphHffgMs3jgEicmC+tVAQLJ8zDRxcLHey8rEDE5GtzMhBaI4dA8P9VD3A8uuZh4wVPCILIwZTy1Fr5K05l1udWy" & _
            "M5WW25e+lGZjdvK1P+pKFzvTlt03iDxNuSZauKZ02ysV/RWDiyu+IWVz5CW4lOlyUTyOk8eW/T0gJbUpCq38Pc+Hc5ih+f3zl37D" & _
            "MaqmVv8PNj9QSwMEFAACAAgAAAAhAF15dEzCAAAA6wAAABkAAABjdXN0b21YbWwvaXRlbVByb3BzMTAueG1sXU7LasMwELwX+g9i" & _
            "74ps1YnTYDnYuIJcSwu5CnmdGKzdYCmlUPrvVemtp2FmmEdz/AyL+MA1zkwGyk0BAsnzONPFwPublXsQMTka3cKEBojh2D4+NGM8" & _
            "jC65mHjFU8IgsjBnPA0GvqwdnrpnXcled7WsdL+V+66vpa13L1tty76uym8QeZpyTTRwTel2UCr6KwYXN3xDyubEa3Ap0/WieJpm" & _
            "jwP7e0BKShfFTvl7ng/nsED7++cv/YpTVG2j/h9sfwBQSwMEFAACAAgAAAAhAFIWE2/CAAAA6wAAABkAAABjdXN0b21YbWwvaXRl" & _
            "bVByb3BzMTEueG1sXU5NawIxFLwX+h/Cu8fsujZa2azYjYJXqdBryL7Vhc17sokilP73pvTmaZgZ5qPePMIo7jjFgclAOStAIHnu" & _
            "BjobOH3u5QpETI46NzKhAWLYNK8vdRfXnUsuJp7wkDCILAwZD9bA93tbFnP9ZuVyr1dysdSt3O7araz0R1UtrC2rXfkDIk9TrokG" & _
            "Lild10pFf8Hg4oyvSNnseQouZTqdFff94NGyvwWkpOZFoZW/5fnwFUZo/v78p4/YR9XU6vlg8wtQSwMEFAACAAgAAAAhAIDQsBzC" & _
            "AAAA6wAAABkAAABjdXN0b21YbWwvaXRlbVByb3BzMTIueG1sXU7LasMwELwX+g9i74qUxE2jYDlYMYZcSwu9CnmdGKzdYCmlUPrv" & _
            "Vemtp2FmmEd9/Iyz+MAlTUwW1isNAinwMNHFwttrL/cgUvY0+JkJLRDDsXl8qId0GHz2KfOC54xRFGEqeO4sfBlXVbtTe5Kmd5Ws" & _
            "3KaVe9O3Ujv93K2fjNua7TeIMk2lJlm45nw7KJXCFaNPK74hFXPkJfpc6HJRPI5TwI7DPSJltdF6p8K9zMf3OEPz++cv/YJjUk2t" & _
            "/h9sfgBQSwMEFAACAAgAAAAhAAyF34fBAAAA6wAAABkAAABjdXN0b21YbWwvaXRlbVByb3BzMTMueG1sXU7LasMwELwX8g9i74rk" & _
            "xhgrWA6pHUOupYVehbxODNZusJRQKP33qvTW0zAzzKM5fIZFPHCNM5OFYqtBIHkeZ7pYeH8bZA0iJkejW5jQAjEc2s1TM8b96JKL" & _
            "iVc8JwwiC3PGc2/h69QfK13WRnbGdLLc9YV8ORWl1GYwO1119TB03yDyNOWaaOGa0m2vVPRXDC5u+YaUzYnX4FKm60XxNM0ee/b3" & _
            "gJTUs9aV8vc8Hz7CAu3vn7/0K05RtY36f7D9AVBLAwQUAAIACAAAACEAnmMxc8EAAADrAAAAGQAAAGN1c3RvbVhtbC9pdGVtUHJv" & _
            "cHMxNC54bWxdTk2LwjAUvAv7H8K7x8S2Vleailu34FUUvIb0VQvNe9LEZWHxv5tlb3saZob5qLbffhRfOIWBycBirkEgOe4Guho4" & _
            "n1q5BhGipc6OTGiAGLb126zqwqaz0YbIEx4iepGEIeFhb+Bn+ZnnTdEUUu+apSz0opW7ssml/mgz/d7qbFXmTxBpmlJNMHCL8b5R" & _
            "KrgbehvmfEdKZs+TtzHR6aq47weHe3YPjxRVpnWp3CPN+4sfof7985c+Yh9UXan/B+sXUEsDBBQAAgAIAAAAIQAmkP8rwgAAAOsA" & _
            "AAAZAAAAY3VzdG9tWG1sL2l0ZW1Qcm9wczE1LnhtbF1Oy2rDMBC8F/oPYu+KlMZ23WA5uJUNuZYGehXyOjFYu8FSSqH036vSW0/D" & _
            "zDCP5vAZFvGBa5yZDGw3GgSS53Gms4HT2yBrEDE5Gt3ChAaI4dDe3zVj3I8uuZh4xWPCILIwZzxaA1/lY10P/XMni11RyaLsrHyq" & _
            "+l6Wgx22XfWird59g8jTlGuigUtK171S0V8wuLjhK1I2J16DS5muZ8XTNHu07G8BKakHrSvlb3k+vIcF2t8/f+lXnKJqG/X/YPsD"

Base64_8 = "UEsDBBQAAgAIAAAAIQDDPK5bwQAAAOsAAAAZAAAAY3VzdG9tWG1sL2l0ZW1Qcm9wczE2LnhtbF2OT4vCMBDF74LfIcy9Jpr6l6ai" & _
            "toJXUfAa0qkWmhlp4rKw7Hc3src9vXlvmHm/Yvvte/GFQ+iYDEwnCgSS46aju4Hr5ZitQIRoqbE9Exoghm05HhVN2DQ22hB5wFNE" & _
            "L1LQJT1VBn7yXV3neqazfX5QWV7XaVrrQzbXU60XlVbL1foXRKqm9CYYeMT43EgZ3AO9DRN+IqVly4O3MdnhLrltO4cVu5dHinKm" & _
            "1EK6V6r3N99D+eH5uz5jG2RZyP+A5RtQSwMEFAACAAgAAAAhAIBF6XbBAAAA6wAAABkAAABjdXN0b21YbWwvaXRlbVByb3BzMTcu" & _
            "eG1sXU7LasMwELwX+g9i74pk1zJOsBxqp4FcSwu5CnmdGKzdYCmlUPrvVemtp2FmmEe7/wyL+MA1zkwWio0GgeR5nOli4f3tKBsQ" & _
            "MTka3cKEFohh3z0+tGPcjS65mHjFU8IgsjBnPB0sfG1N3RflYGRvqq2s6icjm+cXLY2p+uFY6KY0wzeIPE25Jlq4pnTbKRX9FYOL" & _
            "G74hZXPiNbiU6XpRPE2zxwP7e0BKqtS6Vv6e58M5LND9/vlLv+IUVdeq/we7H1BLAwQUAAIACAAAACEAj5lfzMIAAADrAAAAGQAA" & _
            "AGN1c3RvbVhtbC9pdGVtUHJvcHMxOC54bWxdTstqwzAQvBf6D2LvimTHTtNgObhVArmWFnoV8joxWLvBUkoh5N+j0FtPw8wwj2b7" & _
            "Gybxg3McmQwUCw0CyXM/0tHA1+derkHE5Kh3ExMaIIZt+/zU9HHTu+Ri4hkPCYPIwpjxYA1cy/p9XxbdUta6s7J6W9ays4WWa613" & _
            "q8ruXqvu5QYiT1OuiQZOKZ03SkV/wuDigs9I2Rx4Di5lOh8VD8Po0bK/BKSkSq1Xyl/yfPgOE7SPP3/pDxyiahv1/2B7B1BLAwQU" & _
            "AAIACAAAACEAZnb4WcEAAADrAAAAGQAAAGN1c3RvbVhtbC9pdGVtUHJvcHMxOS54bWxdTstqwzAQvBf6D2LvimzjPBosByuqIdfS" & _
            "Qq5CXicGazdYSiiU/ntVeutpmBnm0Rw+wyweuMSJSUO5KkAgeR4mumj4eO/lDkRMjgY3M6EGYji0z0/NEPeDSy4mXvCUMIgsTBlP" & _
            "VsOX6bvarI+9LLtXI+uuNLKzppZbW/b2Zb3bVNX2G0SeplwTNVxTuu2Viv6KwcUV35CyOfISXMp0uSgex8mjZX8PSElVRbFR/p7n" & _
            "wznM0P7++Uu/4RhV26j/B9sfUEsDBBQAAgAIAAAAIQCvfPoRwQAAAOsAAAAYAAAAY3VzdG9tWG1sL2l0ZW1Qcm9wczIueG1sXU5N" & _
            "SwMxFLwL/ofw7mmy2Q9r2WxRQ6FXUeg1ZN+2C5v3yiYVQfzvRrx5GmaG+ej3n3ERH7immclCtdEgkAKPM50tvL8d5BZEyp5GvzCh" & _
            "BWLYD/d3/Zh2o88+ZV7xmDGKIswFj87CV2fqx1p3z7LRByeb7Usrn2rdyIe2cca0VdUa9w2iTFOpSRYuOV93SqVwwejThq9IxZx4" & _
            "jT4Xup4VT9Mc0HG4RaSsjNadCrcyH09xgeH3z1/6Faekhl79Pzj8AFBLAwQUAAIACAAAACEA9NxZFcEAAADrAAAAGQAAAGN1c3Rv" & _
            "bVhtbC9pdGVtUHJvcHMyMC54bWxdTstqwzAQvBf6D2Lviuw4dqRgORBiQ66lhV6FvE4M1m6wlFIo/feq9NbTMDPMoz1+hkV84Bpn" & _
            "JgvlpgCB5Hmc6Wrh7XWQGkRMjka3MKEFYjh2z0/tGA+jSy4mXvGSMIgszBkvZwtf5d5os6tPsqn7Ru6aspd6X9WyGnptKm22tRm+" & _
            "QeRpyjXRwi2l+0Gp6G8YXNzwHSmbE6/BpUzXq+Jpmj2e2T8CUlLbomiUf+T58B4W6H7//KVfcIqqa9X/g90PUEsDBBQAAgAIAAAA" & _
            "IQDQqE9pwgAAAOsAAAAYAAAAY3VzdG9tWG1sL2l0ZW1Qcm9wczMueG1sXU7LasMwELwX+g9i74qUOLGTYDmIuIZcQwq5CnmdGKzd" & _
            "YCmlUPrvVemtp2FmmEd9+AyT+MA5jkwGlgsNAslzP9LNwPulk1sQMTnq3cSEBojh0Ly+1H3c9y65mHjGU8IgsjBmPLUGvlpbFpUt" & _
            "t9LqjZXr3XEj7a6sZFcc7eqtKirdrb9B5GnKNdHAPaXHXqno7xhcXPADKZsDz8GlTOeb4mEYPbbsnwEpqZXWpfLPPB+uYYLm989f" & _
            "+oxDVE2t/h9sfgBQSwMEFAACAAgAAAAhADJlvo/BAAAA6wAAABgAAABjdXN0b21YbWwvaXRlbVByb3BzNC54bWxdTstqwzAQvBf6"

Base64_9 = "D2LvslTHde1gObQ2hlxLC7kKeZ0YrN1gySFQ+u9V6a2nYWaYR3O4+0XccA0zk4GnTINAcjzOdDbw+THICkSIlka7MKEBYji0jw/N" & _
            "GPajjTZEXvEY0YskzAmPvYGvoi/yXacr2XVDLwtdVPLtuazl0OW6etm96nKov0GkaUo1wcAlxuteqeAu6G3I+IqUzIlXb2Oi61nx" & _
            "NM0Oe3abR4oq17pUbkvz/uQXaH///KXfcQqqbdT/g+0PUEsDBBQAAgAIAAAAIQABdTobwQAAAOsAAAAYAAAAY3VzdG9tWG1sL2l0" & _
            "ZW1Qcm9wczUueG1sXU5Ni8IwFLwv+B/Cu8e0tXatNBVLLXhddsFrSF+10LwnTZSFxf9uxNuehplhPqrdr5vEHWc/MmlIlwkIJMv9" & _
            "SGcNP9+d3IDwwVBvJibUQAy7evFR9X7bm2B84BmPAZ2Iwhjx2Gr426SfaZOVuTysy0bmTZHLcr9u5SrvukOx2jdl1j1AxGmKNV7D" & _
            "JYTrVilvL+iMX/IVKZoDz86ESOez4mEYLbZsbw4pqCxJCmVvcd6d3AT16887/YWDV3Wl/h+sn1BLAwQUAAIACAAAACEAGjIP5sIA" & _
            "AADrAAAAGAAAAGN1c3RvbVhtbC9pdGVtUHJvcHM2LnhtbF1Oy2rDMBC8F/oPYu+KFNt1k2A5uI4DuZYGehXyOjFYu8FSSiDk36PS" & _
            "W0/DzDCPanvzk/jBOYxMBpYLDQLJcT/SycDxay9XIEK01NuJCQ0Qw7Z+fan6sOlttCHyjIeIXiRhTHjYGbh3usvbstEyK8tCFu3q" & _
            "Qzb7ZSHX6yxv23edd2/NA0SaplQTDJxjvGyUCu6M3oYFX5CSOfDsbUx0PikehtHhjt3VI0WVaV0qd03z/ttPUP/++Ut/4hBUXan/" & _
            "B+snUEsDBBQAAgAIAAAAIQCux2AowQAAAOsAAAAYAAAAY3VzdG9tWG1sL2l0ZW1Qcm9wczcueG1sXY7NqsIwEIX3gu8QZl/TqI0/" & _
            "NBWxCm4vCm5DOtVCMyNNlAsX392Iu7s5wznDzPnKza/vxROH0DEZUJMcBJLjpqOrgfPpkC1BhGipsT0TGiCGTTUelU1YNzbaEHnA" & _
            "Y0QvUtCleawN/K2U3qvDvM4WShXZXE9n2VbtkuhtkatC6Xq2f4FI1ZTeBAO3GO9rKYO7obdhwnektGx58DYmO1wlt23nsGb38EhR" & _
            "TvNcS/dI9f7ie6g+PN/rH2yDrEr5H7B6A1BLAwQUAAIACAAAACEAlb1IvsEAAADrAAAAGAAAAGN1c3RvbVhtbC9pdGVtUHJvcHM4" & _
            "LnhtbF1Oy2rDMBC8F/oPYu+KlNSJ5GA50DqGXEsLuQp5nRis3WAppVD671XpradhZphHc/iMs/jAJU1MDtYrDQIp8DDRxcH7Wy8t" & _
            "iJQ9DX5mQgfEcGgfH5oh7Qeffcq84CljFEWYCp46B1/HbWVsVT3L3ugnWW3NWtZdv5F1bY8v1hird903iDJNpSY5uOZ82yuVwhWj" & _
            "Tyu+IRVz5CX6XOhyUTyOU8COwz0iZbXReqfCvczHc5yh/f3zl37FMam2Uf8Ptj9QSwMEFAACAAgAAAAhAJC5XYnCAAAA6wAAABgA" & _
            "AABjdXN0b21YbWwvaXRlbVByb3BzOS54bWxdTk2LwjAUvC/4H8K7x6Sr/VCailoFr8sKew3pqxaa96SJy8Ky/92Itz0NM8N81Jsf" & _
            "P4pvnMLAZCCbaxBIjruBLgbOn0dZgQjRUmdHJjRADJtm9lZ3Yd3ZaEPkCU8RvUjCkPDUGvit9GrfLttc7hdFKZfHciV3RXGQWbbN" & _
            "q12Z6zLf/oFI05RqgoFrjLe1UsFd0dsw5xtSMnuevI2JThfFfT84bNndPVJU71oXyt3TvP/yIzTPP6/0B/ZBNbX6f7B5AFBLAwQK" & _
            "AAAAAAAIh3tIAAAAAAAAAAAAAAAAEAAAAGN1c3RvbVhtbC9fcmVscy9QSwMEFAACAAgAAAAhAHQ/OXq8AAAAKAEAAB4AAABjdXN0" & _
            "b21YbWwvX3JlbHMvaXRlbTEueG1sLnJlbHONz7GKwzAMBuD94N7BaG+c3FDKEadLKXQ7Sg66GkdJTGPLWGpp377mpit06CiJ//tR" & _
            "u72FRV0xs6dooKlqUBgdDT5OBn77/WoDisXGwS4U0cAdGbbd50d7xMVKCfHsE6uiRDYwi6RvrdnNGCxXlDCWy0g5WCljnnSy7mwn" & _
            "1F91vdb5vwHdk6kOg4F8GBpQ/T3hOzaNo3e4I3cJGOVFhXYXFgqnsPxkKo2qt3lCMeAFw9+qqYoJumv103/dA1BLAwQUAAIACAAA" & _
            "ACEAH9A/q70AAAApAQAAHwAAAGN1c3RvbVhtbC9fcmVscy9pdGVtMTAueG1sLnJlbHONz7GKwzAMBuC90Hcw2i9OOhxHidOlFLqV"

Base64_10 = "ksKtxlES09gylnpc376mUws33CiJ//tRu/sNi/rBzJ6igaaqQWF0NPg4Gbj0h48vUCw2DnahiAbuyLDr1qv2jIuVEuLZJ1ZFiWxg" & _
            "FklbrdnNGCxXlDCWy0g5WCljnnSy7mon1Ju6/tT51YDuzVTHwUA+Dg2o/p7wPzaNo3e4J3cLGOWPCu1uLBS+w3LKVBpVb/OEYsAL" & _
            "hueqqauCgu5a/fZg9wBQSwMEFAACAAgAAAAhADi1Giq9AAAAKQEAAB8AAABjdXN0b21YbWwvX3JlbHMvaXRlbTExLnhtbC5yZWxz" & _
            "jc+xisMwDAbgvdB3MNovTjocR4nTpRS6lZLCrcZREtPYMpZ6XN++plMLN9woif/7Ubv7DYv6wcyeooGmqkFhdDT4OBm49IePL1As" & _
            "Ng52oYgG7siw69ar9oyLlRLi2SdWRYlsYBZJW63ZzRgsV5QwlstIOVgpY550su5qJ9Sbuv7U+dWA7s1Ux8FAPg4NqP6e8D82jaN3" & _
            "uCd3CxjljwrtbiwUvsNyylQaVW/zhGLAC4bnqmmqgoLuWv32YPcAUEsDBBQAAgAIAAAAIQAQHARyvQAAACkBAAAfAAAAY3VzdG9t" & _
            "WG1sL19yZWxzL2l0ZW0xMi54bWwucmVsc43PsYrDMAwG4P2g72C0N046lOOI06UcdDtKCrcaR0lMY8tY6nF9+5pOLXToKIn/+1G7" & _
            "+w+L+sPMnqKBpqpBYXQ0+DgZOPXf609QLDYOdqGIBq7IsOtWH+0RFyslxLNPrIoS2cAskr60ZjdjsFxRwlguI+VgpYx50sm6s51Q" & _
            "b+p6q/OjAd2TqQ6DgXwYGlD9NeE7No2jd7gndwkY5UWFdhcWCr9h+clUGlVv84RiwAuG+6rZVAUF3bX66cHuBlBLAwQUAAIACAAA" & _
            "ACEAN3kh870AAAApAQAAHwAAAGN1c3RvbVhtbC9fcmVscy9pdGVtMTMueG1sLnJlbHONz8FqwzAMBuD7oO9gdG+crDDKiJNLGfRW" & _
            "Rga7GkdJTGPLWGpZ375mpxV26FES//ejtv8Jq7piZk/RQFPVoDA6Gn2cDXwNH9s9KBYbR7tSRAM3ZOi7zUv7iauVEuLFJ1ZFiWxg" & _
            "EUnvWrNbMFiuKGEsl4lysFLGPOtk3dnOqF/r+k3nvwZ0D6Y6jgbycWxADbeEz9g0Td7hgdwlYJR/KrS7sFD4DuspU2lUg80zigEv" & _
            "GH5Xza4qKOiu1Q8PdndQSwMEFAACAAgAAAAhAEBOOcK9AAAAKQEAAB8AAABjdXN0b21YbWwvX3JlbHMvaXRlbTE0LnhtbC5yZWxz" & _
            "jc/BasMwDAbg+6DvYHRvnIwyyoiTSxn0VkYGuxpHSUxjy1hqWd++ZqcVduhREv/3o7b/Cau6YmZP0UBT1aAwOhp9nA18DR/bPSgW" & _
            "G0e7UkQDN2Tou81L+4mrlRLixSdWRYlsYBFJ71qzWzBYrihhLJeJcrBSxjzrZN3Zzqhf6/pN578GdA+mOo4G8nFsQA23hM/YNE3e" & _
            "4YHcJWCUfyq0u7BQ+A7rKVNpVIPNM4oBLxh+V82uKijortUPD3Z3UEsDBBQAAgAIAAAAIQBnKxxDvQAAACkBAAAfAAAAY3VzdG9t" & _
            "WG1sL19yZWxzL2l0ZW0xNS54bWwucmVsc43PwWrDMAwG4Pug72B0b5wMOsqIk0sZ9FZGBrsaR0lMY8tYalnfvmanFXboURL/96O2" & _
            "/wmrumJmT9FAU9WgMDoafZwNfA0f2z0oFhtHu1JEAzdk6LvNS/uJq5US4sUnVkWJbGARSe9as1swWK4oYSyXiXKwUsY862Td2c6o" & _
            "X+v6Tee/BnQPpjqOBvJxbEANt4TP2DRN3uGB3CVglH8qtLuwUPgO6ylTaVSDzTOKAS8YflfNrioo6K7VDw92d1BLAwQUAAIACAAA" & _
            "ACEAT4ICG70AAAApAQAAHwAAAGN1c3RvbVhtbC9fcmVscy9pdGVtMTYueG1sLnJlbHONz7GKwzAMBuD9oO9gtDdObijHEadLOeh2" & _
            "lBRuNY6SmMaWsdSjffuaTi106CiJ//tRu72ERf1jZk/RQFPVoDA6GnycDBz7n/UXKBYbB7tQRANXZNh2q4/2gIuVEuLZJ1ZFiWxg" & _
            "FknfWrObMViuKGEsl5FysFLGPOlk3clOqD/reqPzowHdk6n2g4G8HxpQ/TXhOzaNo3e4I3cOGOVFhXZnFgp/YfnNVBpVb/OEYsAL" & _
            "hvuq2VQFBd21+unB7gZQSwMEFAACAAgAAAAhAGjnJ5q9AAAAKQEAAB8AAABjdXN0b21YbWwvX3JlbHMvaXRlbTE3LnhtbC5yZWxz" & _
            "jc/BasMwDAbg+6DvYHRvnOzQlREnlzLorYwMdjWOkpjGlrHUsr59zU4r7NCjJP7vR23/E1Z1xcyeooGmqkFhdDT6OBv4Gj62e1As"

Base64_11 = "No52pYgGbsjQd5uX9hNXKyXEi0+sihLZwCKS3rVmt2CwXFHCWC4T5WCljHnWybqznVG/1vVO578GdA+mOo4G8nFsQA23hM/YNE3e" & _
            "4YHcJWCUfyq0u7BQ+A7rKVNpVIPNM4oBLxh+V81bVVDQXasfHuzuUEsDBBQAAgAIAAAAIQCh7DJ5vQAAACkBAAAfAAAAY3VzdG9t" & _
            "WG1sL19yZWxzL2l0ZW0xOC54bWwucmVsc43PsYrDMAwG4P2g72C0N05uKOWI0+U46HaUFLoaR0nMxZax1KN9+5pOLXToKIn/+1G7" & _
            "u4RF/WNmT9FAU9WgMDoafJwMHPuf9RYUi42DXSiigSsy7LrVR3vAxUoJ8ewTq6JENjCLpC+t2c0YLFeUMJbLSDlYKWOedLLuz06o" & _
            "P+t6o/OjAd2TqfaDgbwfGlD9NeE7No2jd/hN7hwwyosK7c4sFE5h+c1UGlVv84RiwAuG+6rZVgUF3bX66cHuBlBLAwQUAAIACAAA" & _
            "ACEAhokX+L0AAAApAQAAHwAAAGN1c3RvbVhtbC9fcmVscy9pdGVtMTkueG1sLnJlbHONz8FqwzAMBuD7oO9gdG+c7FDWESeXMuit" & _
            "jAx2NY6SmMaWsdSyvn3NTivs0KMk/u9Hbf8TVnXFzJ6igaaqQWF0NPo4G/gaPrZvoFhsHO1KEQ3ckKHvNi/tJ65WSogXn1gVJbKB" & _
            "RSS9a81uwWC5ooSxXCbKwUoZ86yTdWc7o36t653Ofw3oHkx1HA3k49iAGm4Jn7FpmrzDA7lLwCj/VGh3YaHwHdZTptKoBptnFANe" & _
            "MPyumn1VUNBdqx8e7O5QSwMEFAACAAgAAAAhAFyWJyK8AAAAKAEAAB4AAABjdXN0b21YbWwvX3JlbHMvaXRlbTIueG1sLnJlbHON" & _
            "z8GKwjAQBuD7gu8Q5m5TPYgsTb0sgjeRLngN6bQN22RCZhR9e4OnFTx4nBn+72ea3S3M6oqZPUUDq6oGhdFR7+No4LfbL7egWGzs" & _
            "7UwRDdyRYdcuvpoTzlZKiCefWBUlsoFJJH1rzW7CYLmihLFcBsrBShnzqJN1f3ZEva7rjc7/DWhfTHXoDeRDvwLV3RN+YtMweIc/" & _
            "5C4Bo7yp0O7CQuEc5mOm0qg6m0cUA14wPFfrqpig20a//Nc+AFBLAwQUAAIACAAAACEATGbSnrwAAAApAQAAHwAAAGN1c3RvbVht" & _
            "bC9fcmVscy9pdGVtMjAueG1sLnJlbHONz8GKwjAQBuD7gu8Q5m5TPYgsTb0sgjeRLngN6bQN22RCZhR9e4OnFTx4nBn+72ea3S3M" & _
            "6oqZPUUDq6oGhdFR7+No4LfbL7egWGzs7UwRDdyRYdcuvpoTzlZKiCefWBUlsoFJJH1rzW7CYLmihLFcBsrBShnzqJN1f3ZEva7r" & _
            "jc7/DWhfTHXoDeRDvwLV3RN+YtMweIc/5C4Bo7yp0O7CQuEc5mOm0qg6m0cUA14wPFfruioo6LbRLw+2D1BLAwQUAAIACAAAACEA" & _
            "e/MCo7wAAAAoAQAAHgAAAGN1c3RvbVhtbC9fcmVscy9pdGVtMy54bWwucmVsc43PwYrCMBAG4PuC7xDmblMVFlmaelkEbyJd8BrS" & _
            "aRu2yYTMKPr2hj2t4MHjzPB/P9PsbmFWV8zsKRpYVTUojI56H0cDP91+uQXFYmNvZ4po4I4Mu3bx0ZxwtlJCPPnEqiiRDUwi6Utr" & _
            "dhMGyxUljOUyUA5WyphHnaz7tSPqdV1/6vzfgPbJVIfeQD70K1DdPeE7Ng2Dd/hN7hIwyosK7S4sFM5hPmYqjaqzeUQx4AXD32pT" & _
            "FRN02+in/9oHUEsDBBQAAgAIAAAAIQAMxBqSvAAAACgBAAAeAAAAY3VzdG9tWG1sL19yZWxzL2l0ZW00LnhtbC5yZWxzjc/BisIw" & _
            "EAbg+4LvEOZuU0UWWZp6WQRvIl3wGtJpG7bJhMwo+vaGPa3gwePM8H8/0+xuYVZXzOwpGlhVNSiMjnofRwM/3X65BcViY29nimjg" & _
            "jgy7dvHRnHC2UkI8+cSqKJENTCLpS2t2EwbLFSWM5TJQDlbKmEedrPu1I+p1XX/q/N+A9slUh95APvQrUN094Ts2DYN3+E3uEjDK" & _
            "iwrtLiwUzmE+ZiqNqrN5RDHgBcPfalMVE3Tb6Kf/2gdQSwMEFAACAAgAAAAhACuhPxO8AAAAKAEAAB4AAABjdXN0b21YbWwvX3Jl" & _
            "bHMvaXRlbTUueG1sLnJlbHONz8GKwjAQBuD7gu8Q5m5TBRdZmnpZBG8iXfAa0mkbtsmEzCj69oY9reDB48zwfz/T7G5hVlfM7Cka" & _
            "WFU1KIyOeh9HAz/dfrkFxWJjb2eKaOCODLt28dGccLZSQjz5xKookQ1MIulLa3YTBssVJYzlMlAOVsqYR52s+7Uj6nVdf+r834D2"

Base64_12 = "yVSH3kA+9CtQ3T3hOzYNg3f4Te4SMMqLCu0uLBTOYT5mKo2qs3lEMeAFw99qUxUTdNvop//aB1BLAwQUAAIACAAAACEAAwghS7wA" & _
            "AAAoAQAAHgAAAGN1c3RvbVhtbC9fcmVscy9pdGVtNi54bWwucmVsc43PwYrCMBAG4PuC7xDmblM9iCxNvSyCN5EueA3ptA3bZEJm" & _
            "FH17g6cVPHicGf7vZ5rdLczqipk9RQOrqgaF0VHv42jgt9svt6BYbOztTBEN3JFh1y6+mhPOVkqIJ59YFSWygUkkfWvNbsJguaKE" & _
            "sVwGysFKGfOok3V/dkS9ruuNzv8NaF9MdegN5EO/AtXdE35i0zB4hz/kLgGjvKnQ7sJC4RzmY6bSqDqbRxQDXjA8V5uqmKDbRr/8" & _
            "1z4AUEsDBBQAAgAIAAAAIQAkbQTKvAAAACgBAAAeAAAAY3VzdG9tWG1sL19yZWxzL2l0ZW03LnhtbC5yZWxzjc/BisIwEAbg+4Lv" & _
            "EOZuUz24sjT1sgjeRLrgNaTTNmyTCZlR9O0Ne1rBg8eZ4f9+ptndwqyumNlTNLCqalAYHfU+jgZ+uv1yC4rFxt7OFNHAHRl27eKj" & _
            "OeFspYR48olVUSIbmETSl9bsJgyWK0oYy2WgHKyUMY86WfdrR9Trut7o/N+A9slUh95APvQrUN094Ts2DYN3+E3uEjDKiwrtLiwU" & _
            "zmE+ZiqNqrN5RDHgBcPf6rMqJui20U//tQ9QSwMEFAACAAgAAAAhAO1mESm8AAAAKAEAAB4AAABjdXN0b21YbWwvX3JlbHMvaXRl" & _
            "bTgueG1sLnJlbHONz8GKwjAQBuC74DuEudtUDyLS1MsieBPpgteQTtuwTSZkRtG3N+xJYQ97nBn+72eawyPM6o6ZPUUD66oGhdFR" & _
            "7+No4Ls7rnagWGzs7UwRDTyR4dAuF80FZyslxJNPrIoS2cAkkvZas5swWK4oYSyXgXKwUsY86mTdjx1Rb+p6q/O7Ae2HqU69gXzq" & _
            "16C6Z8L/2DQM3uEXuVvAKH9UaHdjoXAN8zlTaVSdzSOKAS8Yfle7qpig20Z//Ne+AFBLAwQUAAIACAAAACEAygM0qLwAAAAoAQAA" & _
            "HgAAAGN1c3RvbVhtbC9fcmVscy9pdGVtOS54bWwucmVsc43PwYrCMBAG4PuC7xDmblM9yLo09bII3kS64DWk0zZskwmZUfTtDXta" & _
            "wYPHmeH/fqbZ3cKsrpjZUzSwqmpQGB31Po4Gfrr98hMUi429nSmigTsy7NrFR3PC2UoJ8eQTq6JENjCJpC+t2U0YLFeUMJbLQDlY" & _
            "KWMedbLu146o13W90fm/Ae2TqQ69gXzoV6C6e8J3bBoG7/Cb3CVglBcV2l1YKJzDfMxUGlVn84hiwAuGv9W2KibottFP/7UPUEsD" & _
            "BAoAAAAAAAiHe0gAAAAAAAAAAAAAAAAKAAAAY3VzdG9tWG1sL1BLAwQUAAIACAAAACEAHG5LdHABAAD9AgAAEAAAAGRvY1Byb3Bz" & _
            "L2FwcC54bWydkstOwzAQRfdI/EPkPXUKFUKVY4QKqAsQlfpgbZxJY9W1I88QtXw9TqqmKbBiN4+r6+Nri/vd1iY1BDTeZWw4SFkC" & _
            "TvvcuHXGlovnqzuWICmXK+sdZGwPyO7l5YWYBV9BIAOYRAuHGSuJqjHnqEvYKhzEtYubwoetotiGNfdFYTQ8ev25BUf8Ok1vOewI" & _
            "XA75VdUZsoPjuKb/muZeN3y4Wuyr6CfFQ1VZoxXFW8pXo4NHX1DytNNgBe8vRTSag/4MhvYyFbzfirlWFibRWBbKIgh+GogpqCa0" & _
            "mTIBpahpXIMmHxI0XzG2a5Z8KIQGJ2O1CkY5YgfZoWlrWyEF+e7DBksAQsG7YVv2tf3ajOSwFcTiXMg7kFifIy4MWcC3YqYC/UE8" & _
            "7BO3DKzHOG/4fuEdD/ph/WLcBpfVwj8qgmN050MxL1WAPKbdRdsNxDQyBNvoJ6Vya8iPmt+L5qFXh98sh6NBepOm7fseZ4Kf/q38" & _
            "BlBLAwQUAAIACAAAACEAujbtdxMBAAARAgAAEQAAAGRvY1Byb3BzL2NvcmUueG1slZGxTsMwEIZ3JN4h8p7YcaCAlaQDqBNISAQV" & _
            "sVn2NbWIHcs2pH173NCmFXRhPP3ffXf2lfON7pIvcF71pkJ5RlACRvRSmbZCr80ivUWJD9xI3vUGKrQFj+b15UUpLBO9g2fXW3BB" & _
            "gU+iyXgmbIXWIViGsRdr0NxnkTAxXPVO8xBL12LLxQdvAVNCZlhD4JIHjnfC1E5GtFdKMSntp+tGgRQYOtBggsd5luMjG8Bpf7Zh" & _
            "TE5IrcLWwln0EE70xqsJHIYhG4oRjfvn+O3p8WV8aqrM7q8EoLrcL8KEAx5AJlHAfsYdkmVx/9AsUE1JPktJkdKbhhaMUHZF30v8"

Base64_13 = "q/8o1PE4K/Uf4zUjdyfGg6Au8Z8j1t9QSwMECgAAAAAACId7SAAAAAAAAAAAAAAAAAkAAABkb2NQcm9wcy9QSwMEFAACAAgAAAAh" & _
            "AJ8XkB9aAgAA6QMAABIAAAB4bC9jb25uZWN0aW9ucy54bWyNUt9v2jAQfp+0/+Fk9XGQ0MJWGkIVAZMqgcoGq/ZWOfYFLBw7sh1+" & _
            "bNr/vktYO6a+7CXx+Xx33/fdN7o/lhr26LyyJmW9bswAjbBSmU3Kvq0/d24Z+MCN5NoaTNkJPbsfv383EtYYFIHKPFAP41O2DaG6" & _
            "iyIvtlhy37UVGsoU1pU8UOg2ka8ccum3iKHU0XUcf4xKrgwbX7QDJQkIgx1ilWm1xzYyvKTD0h7QLdXeBpjywBlI9MKpKrTo11vl" & _
            "4bKRh9qjhPwEs6NADQSF8mVZGyV4+yTHcEA0ELYIB+t2ubU7ILaAZY5SUvHFSEkjP7RZv7W1lmDoMkcouam51idAqQKV0BSJGunY" & _
            "ZRBOFQEfMHBYOCTq8ulFbdI652K3cbY2LefxSOZLd0GBGDu7VxJdulg9zrNld5Asm3IfYIWidiqc4MEUNl27GpMHo4LiGiYEVNtN" & _
            "ulDCWW+L8Lz6Ml+hoz0/Z4brEzVoQiXQJ42QsLK1E5hezf7QvkoW0+8wsWVFOuVK05y0l6x4gTTwsdXbp9fJ5Ix0bdMeWaetWSjv" & _
            "yTywaBR0sLAS05lz1iVNWal+IHxFX1E9pjfJBLWGNn1+ucZjeOK6RtYuije6NAn9Gq9bPXssGo+s5hXJ5dHIuRVcn53i7GHqlNYT" & _
            "UjXQTRzHzWNqPPeh/QPplrKf00/D4TDO4s4k6193+gO86dzefO7TaTKI6X6WZYNf7Gzuu2Ov/8bg5Yu8XcIW2aIgPd9afBgNX0xO" & _
            "Te4u/ClqHWpHqNmr3ZplPDSc/2NzDauI6Jy/Lbnob/d/Aj/+DVBLAwQKAAAAAAAIh3tIAAAAAAAAAAAAAAAADgAAAHhsL2N1c3Rv" & _
            "bURhdGEvUEsDBBQAAgAIAAAAIQCs1GIUnAAAALkAAAAcAAAAeGwvY3VzdG9tRGF0YS9pdGVtUHJvcHMxLnhtbDWOsQrCMBBAd8F/" & _
            "CLfbq05WTIuLIOgg1bmE5GoDTa7kgti/tw6Ob3iPd2w+YVRvSuI5atgWJSiKlp2PLw3Px3mzByXZRGdGjqRhJoGmXq+OzmQjmRNd" & _
            "MgW1VKJoGHKeDohiBwpGiuBtYuE+F5YDct97SyhTIuNkIMphxF1ZVlhhMD6C8k7D7e907f3aUlreulM04yxefrgkBLD+AlBLAwQK" & _
            "AAAAAAAIh3tIAAAAAAAAAAAAAAAAFAAAAHhsL2N1c3RvbURhdGEvX3JlbHMvUEsDBBQAAgAIAAAAIQC9HiMLuQAAABMBAAAnAAAA" & _
            "eGwvY3VzdG9tRGF0YS9fcmVscy9pdGVtUHJvcHMxLnhtbC5yZWxzZc/BSsQwEAbgu+A7hLnbtB5UpOleRNirrA8wJNM2bJMJmdnF" & _
            "fXujJ4vH+X/+D2Y8fKXNXKlK5Oxg6HowlD2HmBcHn6f3hxcwopgDbpzJwY0EDtP93fhBG2obyRqLmKZkcbCqlldrxa+UUDoulFsz" & _
            "c02o7ayLLejPuJB97PsnW/8aMO1McwwO6jEMYE63Qv/sFH1l4Vk7z8nyPEf/qz7vVesvopzeULFBWBdSB1EpDV34yew02t0r0zdQ" & _
            "SwMEFAACAAgAAAAhAAVv5QYoAgAA1QQAAA0AAAB4bC9zdHlsZXMueG1spZTbitswEIbvC30HoXtHjptsk2B7aTZrWNiWQlLorWLL" & _
            "jlgdjCSncUvfvSPbiRO20MLeWDP/jL4ZnRzfn6RAR2Ys1yrB00mIEVO5LriqEvxtlwULjKyjqqBCK5bglll8n75/F1vXCrY9MOYQ" & _
            "IJRN8MG5ekWIzQ9MUjvRNVMQKbWR1IFrKmJrw2hh/SQpSBSGd0RSrnBPWMn8fyCSmpemDnIta+r4ngvu2o6FkcxXT5XShu4FtHqa" & _
            "zmh+ZnfOK7zkudFWl24COKLLkufsdZdLsiRASuNSK2dRrhvlYK8A7aGrF6V/qMyHvNhnpbH9iY5UgBJiksa5FtogB1WZTwJFUcn6" & _
            "jE+GU+Glkkou2l6MvEB6VDdYiHMhLuUj3AtpDNvgmFEZOGiwd20NdRScWI/p8v6RXRnaTqP51YRugLp7bQq4IePCz1IaC1Y6mGB4" & _
            "dfCj0zXxQedgO9O44LTSigqPPM8YDMDmTIitv0Xfyxv2qUSqkZl0T0WC4T761Z9NaGgwe0zveP41rWe/GYtO5S3/gu4K3dAvKvIH" & _
            "m+Av/saKEYH2DReOq780DMziNPbaRZ2/wrdVgFGwkjbC7S7BBI/2Z1bwRkaXrK/8qN2QNdrP/qSmd74GO7ln67oRNYYn+Nfj+uNy"

Base64_14 = "85hFwSJcL4LZBzYPlvP1JpjPHtabTbYMo/Dh99WLesN76t49HMp0trICssyw2KH57agl+Mrp2+/2D9ruv90iyPg/Sv8AUEsDBAoA" & _
            "AAAAAAiHe0gAAAAAAAAAAAAAAAAJAAAAeGwvdGhlbWUvUEsDBBQAAgAIAAAAIQD7YqVttgUAAKcbAAATAAAAeGwvdGhlbWUvdGhl" & _
            "bWUxLnhtbO1ZT2/bNhS/D9h3IHRvZdmS6wRxitix2y1NGyRuhx5piZZYU6JA0kl9G9LjgAHDumGXAbvtMGwr0AK7dJ8mW4etA/oV" & _
            "9vTHNhXTbdKm2IbWB1skf+8/3+OjvHH1fszQIRGS8qRtOZdrFiKJzwOahG3r9qB/qWUhqXASYMYT0ramRFpXNz/8YAOvq4jEBAF9" & _
            "Itdx24qUStdtW/owjeVlnpIE1kZcxFjBUIR2IPAR8I2ZXa/VmnaMaWKhBMfA9tZoRH2CBhlLa3PGvMfgK1Eym/CZOPBziTpFjg3G" & _
            "TvYjp7LLBDrErG2BnIAfDch9ZSGGpYKFtlXLP5a9uWHPiZhaQavR9fNPSVcSBON6TifC4ZzQ6btrV7bn/OsF/2Vcr9fr9pw5vxyA" & _
            "fR8sdZawbr/ldGY8NVDxuMy7W/NqbhWv8W8s4dc6nY63VsE3Fnh3Cd+qNd2tegXvLvDesv6drW63WcF7C3xzCd+/stZ0q/gcFDGa" & _
            "jJfQWTznkZlDRpxdN8JbAG/NNsACZWu7q6BP1Kq9FuN7XPQBkAcXK5ogNU3JCPuA6+J4KCjOBOB1grWVYsqXS1OZLCR9QVPVtj5O" & _
            "MWTEAvLi6Y8vnj5GL54+Ojl+cnL8y8mDByfHPxsIr+Mk1Amff//F399+iv56/N3zh1+Z8VLH//7TZ7/9+qUZqHTgs68f/fHk0bNv" & _
            "Pv/zh4cG+JbAQx0+oDGR6CY5Qvs8BtsMAshQnI9iEGFaocARIA3AnooqwJtTzEy4Dqk6746AAmACXpvcq+h6EImJogbgThRXgLuc" & _
            "sw4XRnN2Mlm6OZMkNAsXEx23j/GhSXb3VGh7kxR2MjWx7EakouYeg2jjkCREoWyNjwkxkN2ltOLXXeoLLvlIobsUdTA1umRAh8pM" & _
            "dJ3GEJcpNoe64pvdO6jDmYn9NjmsIiEhMDOxJKzixmt4onBs1BjHTEfewCoyKXkwFX7F4VJBpEPCOOoFREoTzS0xrai7g6ESGcO+" & _
            "y6ZxFSkUHZuQNzDnOnKbj7sRjlOjzjSJdOxHcgxbFKM9roxK8GqGZGOIA05WhvsOJep8aX2bhpF5g2QrE2FKCcKr+ThlI0ySsr5X" & _
            "KnVMk5eVbUahbr8v2zP4Fhxi7AzFehXuf1iit/Ek2SOQFe8r9PsK/S5W6FW5fPF1eVGKbb3XztnEKxvvEWXsQE0ZuSHzIi7BvKAP" & _
            "k/kgJ5r3+WkEj6W4Ci4UOH9GgqtPqIoOIpyCGCeXEMqSdShRyiXcLqyVvPMrKgWb8zlvdq8ENFa7PCimG/p9c84mH4VSF9TIGJxV" & _
            "WOPKmwlzCuAZpTmeWZr3Umm25k3IG4SztwlOs16Iho2CGQkyvxcMZmG58BDJCAekjJFjNMRpnNFtrVd7TZO21ngzaWcJki7OXSHO" & _
            "u4Ao1ZaiZC+nI0uqI3QEWnl1z0I+TtvWCHoueIxT4CezUoVZmLQtX5WmvDKZTxts3pZObaXBFRGpkGoby6igypdmr2OShf51z838" & _
            "cDEG2K+rRaPl/Ita2KdDS0Yj4qsVM4thucYnioiDKDhCQzYR+xj0dovdFVAJR0V9NhCQoW658aqZX2bB6dc+ZXZglka4rEktLfYF" & _
            "PH+e65CPNPXsFbq/pimNCzTFe3dNyXYuNLiNIL96QRsgMMr2aNviQkUcqlAaUb8voHHIZYFeCNIiUwmx7CV2pis5XNStgkdR5MJI" & _
            "7dMQCQqVTkWCkD1V2vkKZk5dP19njMo6M1dXpsXvkBwSNsiyt5nZb6FoVk1KR+S400GzTdk1DPv/4c7Hrb1Oe7AQ5J6nF3G1oq8d" & _
            "BWtvpsI5j9q62eK6d+ajNoVrCsq+oHBT4bNFfzvg+xB9NO8oEWzES60y/eaTQ9C5pRmXsXq7bdQiBK3a228+NWc3Vji7Vns7zvYM" & _
            "vvZe7mp7OUVt7SKTj5b+zOLDeyB7G+5HE6Zk8d7pPlxKu7O/IYCPvSDd/AdQSwMEFAACAAgAAAAhAIHIQA05AQAAFAIAAA8AAAB4" & _
            "bC93b3JrYm9vay54bWyNUctOwzAQvCPxD5bv1GnaVKVqUgkBoheERGnPJt40Vv2IbIe0f886USjcOO3OenY0s15vzlqRL3BeWpPT" & _
            "6SShBExphTTHnH7snu+WlPjAjeDKGsjpBTzdFLc3686606e1J4ICxue0DqFZMebLGjT3E9uAwZfKOs0DQndkvnHAha8BglYsTZIF"

Base64_15 = "01waOiis3H80bFXJEh5t2WowYRBxoHhA+76WjafFupIK9kMiwpvmlWv0fVaUKO7Dk5ABRE4zhLaDPwPXNg+tVAjuZ8mMsuIn5Jsj" & _
            "AireqrBDa6M63iudp+kiMiNrL6Hz16UIyfkgjbBdpOJpL39Q14ODFKHOaTpbXmcvII91QB/LLIvq7Jd8f8CxEtOne489flQsW/Q/" & _
            "xTAriY3bimkvMG6VXJWYJpaeOM8W6cAYbRffUEsDBAoAAAAAAAiHe0gAAAAAAAAAAAAAAAAOAAAAeGwvd29ya3NoZWV0cy9QSwME" & _
            "FAACAAgAAAAhAGJT7zNfAQAAhgIAABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWyNkstuwjAQRfeV+g+W98RAHxREQJUQaheV" & _
            "qr72jjNJLGxPZA8F/r6TUKpKbNjN9XiO7h17vtx7J74hJoshl6NsKAUEg6UNdS4/P9aDBykS6VBqhwFyeYAkl4vrq/kO4yY1ACSY" & _
            "EFIuG6J2plQyDXidMmwhcKfC6DWxjLVKbQRd9kPeqfFweK+8tkEeCbN4CQOryhpYodl6CHSERHCa2H9qbJtONG8uwXkdN9t2YNC3" & _
            "jCiss3TooVJ4M3uuA0ZdOM69H91qc2L34gzvrYmYsKKMcb9GzzNP1VQxaTEvLSfo1i4iVLl8HEm1mPcXvyzs0r9akC7ewYEhKPmN" & _
            "pOh2XyBuuuYzHw27UXU2u+6DvkZRQqW3jt5w9wS2bogh42xyx3G6ILPysIJkeJ1MysZ/NlaaNNetruFFx9qGJBxU/aWJFPEI6mvC" & _
            "tq+YWCAR+pNqODvETt1IUSHSSXR+/37Q4gdQSwMECgAAAAAACId7SAAAAAAAAAAAAAAAAAkAAAB4bC9fcmVscy9QSwMEFAACAAgA" & _
            "AAAhAMBlKDiRAQAAcw4AABoAAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc72XTWvDMAyG74P9h+D74jj9Hk17KYPexuhgV+Mq" & _
            "bWhsB9vd1n8/Lx1NBp3YoegUJGP54ZXzRpkvP3WdvIPzlTUFE2nGEjDKbiuzK9jr5ulhyhIfpNnK2hoo2Ak8Wy7u7+YvUMsQN/l9" & _
            "1fgkVjG+YPsQmkfOvdqDlj61DZi4UlqnZYih2/FGqoPcAc+zbMxdvwZb/KqZrLcFc+ttPH9zauA/tW1ZVgpWVh01mHDlCK6OPlj9" & _
            "putYVLodhIKlaZflVQA9SGNZxq/TiAE1zhTFIVdHoPLcVh1rDKg27oB6SQwkF+TCjDGeCTVOjl6bnBpnguKQqyNQefIRebcylOeW" & _
            "OCHuhQ6lDc9JgfZoTN4jlCfPyHlGqD63xPmw7uD3AKHDuaQ8bx+oOPS9QqUh92LUivMhuTozjOcvs9GVctbbMqTK6h+Sb4LJVYKV" & _
            "DPLZ2ab/qbwstBTtKt4oct8TuO+RD1sCn7bIPQe3nBm5PEOM56bvlQ+nGnqX+Rxjx5OrgYohyF0GNZmcfOITl5GP//pVXHwBUEsD" & _
            "BAoAAAAAAAiHe0gAAAAAAAAAAAAAAAADAAAAeGwvUEsDBBQAAgAIAAAAIQBcyE35xwEAAJcPAAATAAAAW0NvbnRlbnRfVHlwZXNd" & _
            "LnhtbM1XS28TMRC+I/EfVr6iXccBSouy6YHCEXooElfXniRW/JLHKcm/x7tLI4TSDWkjNJe1dj3fY6zRp/Xseuts9QAJTfAtE82E" & _
            "VeBV0MYvW/b97kt9ySrM0mtpg4eW7QDZ9fz1q9ndLgJWBe2xZauc40fOUa3ASWxCBF92FiE5mctrWvIo1VougU8nkwuugs/gc507" & _
            "Djaf3cBCbmyuPm/L58FJAous+jQUdlotkzFao2Qu+/zB679U6t8KTUH2NbgyEd+UAsYPKnQ7Tws8jdMyyxHgvfEy7Trst3KsyWio" & _
            "bmXKX6UrhXxr+c+Q1vchrJtxAwc6DIuFUaCD2rgCaTAmkBpXANnZpl8bJ41/7HlEvy9G3i/izEb2/Ed85DIrMDxfbqGnOSJYps6D" & _
            "6kfjzC3/wXzEA+adhXPLD6THut9gDu6mDC83GdxtChH/5dwd1rBVUJrcE3RYSNmMaw71P5w9TW+88T3psz1MCXh4S8DDOwIe3hPw" & _
            "cEHAwwcCHi4JeLgi4EFMKJigkJSCQlQKClkpKISloJCWgkJcCgp5KSgEpqCQmNP/npgF1yuXX/oEp4s/3hQ7dB1PUyzUL+4Wusuk" & _
            "Bn1Am/fX6vkvUEsBAhQAFAACAAgAAAAhALVVMCPrAAAATAIAAAsAJAAAAAAAAQAAAAAAAAAAAF9yZWxzLy5yZWxzCgAgAAAAAAAB" & _
            "ABgAAFg0jtrnqAHwY7k/hIjRAfBjuT+EiNEBUEsBAhQACgAAAAAACId7SAAAAAAAAAAAAAAAAAYAJAAAAAAAAAAQAAAAFAEAAF9y"

Base64_16 = "ZWxzLwoAIAAAAAAAAQAYAPBjuT+EiNEB8GO5P4SI0QFwZ64/hIjRAVBLAQIUABQAAgAIAAAAIQBt9o6fyQAAAGgBAAATACQAAAAA" & _
            "AAAAAAAAADgBAABjdXN0b21YbWwvaXRlbTEueG1sCgAgAAAAAAABABgAAFg0jtrnqAGAjq4/hIjRAYCOrj+EiNEBUEsBAhQAFAAC" & _
            "AAgAAAAhAA406C7cAAAAigEAABQAJAAAAAAAAAAAAAAAMgIAAGN1c3RvbVhtbC9pdGVtMTAueG1sCgAgAAAAAAABABgAAFg0jtrn" & _
            "qAGQta4/hIjRAZC1rj+EiNEBUEsBAhQAFAACAAgAAAAhALfHs4CzAAAARAEAABQAJAAAAAAAAAAAAAAAQAMAAGN1c3RvbVhtbC9p" & _
            "dGVtMTEueG1sCgAgAAAAAAABABgAAFg0jtrnqAGg3K4/hIjRAaDcrj+EiNEBUEsBAhQAFAACAAgAAAAhAK/cEBSwAAAAQgEAABQA" & _
            "JAAAAAAAAAAAAAAAJQQAAGN1c3RvbVhtbC9pdGVtMTIueG1sCgAgAAAAAAABABgAAFg0jtrnqAGwA68/hIjRAbADrz+EiNEBUEsB" & _
            "AhQAFAACAAgAAAAhAM9b3FGyAAAASgEAABQAJAAAAAAAAAAAAAAABwUAAGN1c3RvbVhtbC9pdGVtMTMueG1sCgAgAAAAAAABABgA" & _
            "AFg0jtrnqAGwA68/hIjRAbADrz+EiNEBUEsBAhQAFAACAAgAAAAhABMfZhS4AAAAWgEAABQAJAAAAAAAAAAAAAAA6wUAAGN1c3Rv" & _
            "bVhtbC9pdGVtMTQueG1sCgAgAAAAAAABABgAAFg0jtrnqAHQUa8/hIjRAdBRrz+EiNEBUEsBAhQAFAACAAgAAAAhAH/imN/7AAAA" & _
            "YAIAABQAJAAAAAAAAAAAAAAA1QYAAGN1c3RvbVhtbC9pdGVtMTUueG1sCgAgAAAAAAABABgAAFg0jtrnqAHgeK8/hIjRAeB4rz+E" & _
            "iNEBUEsBAhQAFAACAAgAAAAhAH7z/NvGAAAAiAEAABQAJAAAAAAAAAAAAAAAAggAAGN1c3RvbVhtbC9pdGVtMTYueG1sCgAgAAAA" & _
            "AAABABgAAFg0jtrnqAHgeK8/hIjRAeB4rz+EiNEBUEsBAhQAFAACAAgAAAAhADhOc4jMAAAAfgEAABQAJAAAAAAAAAAAAAAA+ggA" & _
            "AGN1c3RvbVhtbC9pdGVtMTcueG1sCgAgAAAAAAABABgAAFg0jtrnqAHwn68/hIjRAfCfrz+EiNEBUEsBAhQAFAACAAgAAAAhAIn+" & _
            "84rDAAAAagEAABQAJAAAAAAAAAAAAAAA+AkAAGN1c3RvbVhtbC9pdGVtMTgueG1sCgAgAAAAAAABABgAAFg0jtrnqAEAx68/hIjR" & _
            "AQDHrz+EiNEBUEsBAhQAFAACAAgAAAAhAIqsiixcDQAACIYAABQAJAAAAAAAAQAAAAAA7QoAAGN1c3RvbVhtbC9pdGVtMTkueG1s" & _
            "CgAgAAAAAAABABgAAFg0jtrnqAEAx68/hIjRAQDHrz+EiNEBUEsBAhQAFAACAAgAAAAhAG6otz45AgAAlggAABMAJAAAAAAAAQAA" & _
            "AAAAexgAAGN1c3RvbVhtbC9pdGVtMi54bWwKACAAAAAAAAEAGAAAWDSO2ueoARDurz+EiNEBEO6vP4SI0QFQSwECFAAUAAIACAAA" & _
            "ACEADHZkVJ0BAABQBAAAFAAkAAAAAAAAAAAAAADlGgAAY3VzdG9tWG1sL2l0ZW0yMC54bWwKACAAAAAAAAEAGAAAWDSO2ueoASAV" & _
            "sD+EiNEBIBWwP4SI0QFQSwECFAAUAAIACAAAACEAJ5x/tbcAAABOAQAAEwAkAAAAAAAAAAAAAAC0HAAAY3VzdG9tWG1sL2l0ZW0z" & _
            "LnhtbAoAIAAAAAAAAQAYAABYNI7a56gBIBWwP4SI0QEgFbA/hIjRAVBLAQIUABQAAgAIAAAAIQAunfYSuQAAAE4BAAATACQAAAAA" & _
            "AAAAAAAAAJwdAABjdXN0b21YbWwvaXRlbTQueG1sCgAgAAAAAAABABgAAFg0jtrnqAEwPLA/hIjRATA8sD+EiNEBUEsBAhQAFAAC" & _
            "AAgAAAAhAB8w4tICAgAAegYAABMAJAAAAAAAAAAAAAAAhh4AAGN1c3RvbVhtbC9pdGVtNS54bWwKACAAAAAAAAEAGAAAWDSO2ueo" & _
            "AUBjsD+EiNEBQGOwP4SI0QFQSwECFAAUAAIACAAAACEAP1SUkzsBAACUAgAAEwAkAAAAAAAAAAAAAAC5IAAAY3VzdG9tWG1sL2l0" & _
            "ZW02LnhtbAoAIAAAAAAAAQAYAABYNI7a56gBUIqwP4SI0QFQirA/hIjRAVBLAQIUABQAAgAIAAAAIQBqMtRHqQAAADgBAAATACQA"

Base64_17 = "AAAAAAAAAAAAACUiAABjdXN0b21YbWwvaXRlbTcueG1sCgAgAAAAAAABABgAAFg0jtrnqAFQirA/hIjRAVCKsD+EiNEBUEsBAhQA" & _
            "FAACAAgAAAAhACRG7IpdBgAAzGAAABMAJAAAAAAAAQAAAAAA/yIAAGN1c3RvbVhtbC9pdGVtOC54bWwKACAAAAAAAAEAGAAAWDSO" & _
            "2ueoAXDYsD+EiNEBcNiwP4SI0QFQSwECFAAUAAIACAAAACEA5Yo/mq8AAABCAQAAEwAkAAAAAAAAAAAAAACNKQAAY3VzdG9tWG1s" & _
            "L2l0ZW05LnhtbAoAIAAAAAAAAQAYAABYNI7a56gBgP+wP4SI0QGA/7A/hIjRAVBLAQIUABQAAgAIAAAAIQDhshG0wQAAAOsAAAAY" & _
            "ACQAAAAAAAEAAAAAAG0qAABjdXN0b21YbWwvaXRlbVByb3BzMS54bWwKACAAAAAAAAEAGAAAWDSO2ueoAaBNsT+EiNEBoE2xP4SI" & _
            "0QFQSwECFAAUAAIACAAAACEAXXl0TMIAAADrAAAAGQAkAAAAAAABAAAAAABkKwAAY3VzdG9tWG1sL2l0ZW1Qcm9wczEwLnhtbAoA" & _
            "IAAAAAAAAQAYAABYNI7a56gBoE2xP4SI0QGgTbE/hIjRAVBLAQIUABQAAgAIAAAAIQBSFhNvwgAAAOsAAAAZACQAAAAAAAEAAAAA" & _
            "AF0sAABjdXN0b21YbWwvaXRlbVByb3BzMTEueG1sCgAgAAAAAAABABgAAFg0jtrnqAGwdLE/hIjRAbB0sT+EiNEBUEsBAhQAFAAC" & _
            "AAgAAAAhAIDQsBzCAAAA6wAAABkAJAAAAAAAAQAAAAAAVi0AAGN1c3RvbVhtbC9pdGVtUHJvcHMxMi54bWwKACAAAAAAAAEAGAAA" & _
            "WDSO2ueoAcCbsT+EiNEBwJuxP4SI0QFQSwECFAAUAAIACAAAACEADIXfh8EAAADrAAAAGQAkAAAAAAABAAAAAABPLgAAY3VzdG9t" & _
            "WG1sL2l0ZW1Qcm9wczEzLnhtbAoAIAAAAAAAAQAYAABYNI7a56gB0MKxP4SI0QHQwrE/hIjRAVBLAQIUABQAAgAIAAAAIQCeYzFz" & _
            "wQAAAOsAAAAZACQAAAAAAAEAAAAAAEcvAABjdXN0b21YbWwvaXRlbVByb3BzMTQueG1sCgAgAAAAAAABABgAAFg0jtrnqAHg6bE/" & _
            "hIjRAeDpsT+EiNEBUEsBAhQAFAACAAgAAAAhACaQ/yvCAAAA6wAAABkAJAAAAAAAAQAAAAAAPzAAAGN1c3RvbVhtbC9pdGVtUHJv" & _
            "cHMxNS54bWwKACAAAAAAAAEAGAAAWDSO2ueoAeDpsT+EiNEB4OmxP4SI0QFQSwECFAAUAAIACAAAACEAwzyuW8EAAADrAAAAGQAk" & _
            "AAAAAAABAAAAAAA4MQAAY3VzdG9tWG1sL2l0ZW1Qcm9wczE2LnhtbAoAIAAAAAAAAQAYAABYNI7a56gB8BCyP4SI0QHwELI/hIjR" & _
            "AVBLAQIUABQAAgAIAAAAIQCARel2wQAAAOsAAAAZACQAAAAAAAEAAAAAADAyAABjdXN0b21YbWwvaXRlbVByb3BzMTcueG1sCgAg" & _
            "AAAAAAABABgAAFg0jtrnqAEAOLI/hIjRAQA4sj+EiNEBUEsBAhQAFAACAAgAAAAhAI+ZX8zCAAAA6wAAABkAJAAAAAAAAQAAAAAA" & _
            "KDMAAGN1c3RvbVhtbC9pdGVtUHJvcHMxOC54bWwKACAAAAAAAAEAGAAAWDSO2ueoARBfsj+EiNEBEF+yP4SI0QFQSwECFAAUAAIA" & _
            "CAAAACEAZnb4WcEAAADrAAAAGQAkAAAAAAABAAAAAAAhNAAAY3VzdG9tWG1sL2l0ZW1Qcm9wczE5LnhtbAoAIAAAAAAAAQAYAABY" & _
            "NI7a56gBIIayP4SI0QEghrI/hIjRAVBLAQIUABQAAgAIAAAAIQCvfPoRwQAAAOsAAAAYACQAAAAAAAEAAAAAABk1AABjdXN0b21Y" & _
            "bWwvaXRlbVByb3BzMi54bWwKACAAAAAAAAEAGAAAWDSO2ueoATCtsj+EiNEBMK2yP4SI0QFQSwECFAAUAAIACAAAACEA9NxZFcEA" & _
            "AADrAAAAGQAkAAAAAAABAAAAAAAQNgAAY3VzdG9tWG1sL2l0ZW1Qcm9wczIwLnhtbAoAIAAAAAAAAQAYAABYNI7a56gBQNSyP4SI" & _
            "0QFA1LI/hIjRAVBLAQIUABQAAgAIAAAAIQDQqE9pwgAAAOsAAAAYACQAAAAAAAEAAAAAAAg3AABjdXN0b21YbWwvaXRlbVByb3Bz" & _
            "My54bWwKACAAAAAAAAEAGAAAWDSO2ueoAVD7sj+EiNEBUPuyP4SI0QFQSwECFAAUAAIACAAAACEAMmW+j8EAAADrAAAAGAAkAAAA"

Base64_18 = "AAABAAAAAAAAOAAAY3VzdG9tWG1sL2l0ZW1Qcm9wczQueG1sCgAgAAAAAAABABgAAFg0jtrnqAFgIrM/hIjRAWAisz+EiNEBUEsB" & _
            "AhQAFAACAAgAAAAhAAF1OhvBAAAA6wAAABgAJAAAAAAAAQAAAAAA9zgAAGN1c3RvbVhtbC9pdGVtUHJvcHM1LnhtbAoAIAAAAAAA" & _
            "AQAYAABYNI7a56gBYCKzP4SI0QFgIrM/hIjRAVBLAQIUABQAAgAIAAAAIQAaMg/mwgAAAOsAAAAYACQAAAAAAAEAAAAAAO45AABj" & _
            "dXN0b21YbWwvaXRlbVByb3BzNi54bWwKACAAAAAAAAEAGAAAWDSO2ueoAXBJsz+EiNEBcEmzP4SI0QFQSwECFAAUAAIACAAAACEA" & _
            "rsdgKMEAAADrAAAAGAAkAAAAAAABAAAAAADmOgAAY3VzdG9tWG1sL2l0ZW1Qcm9wczcueG1sCgAgAAAAAAABABgAAFg0jtrnqAGA" & _
            "cLM/hIjRAYBwsz+EiNEBUEsBAhQAFAACAAgAAAAhAJW9SL7BAAAA6wAAABgAJAAAAAAAAQAAAAAA3TsAAGN1c3RvbVhtbC9pdGVt" & _
            "UHJvcHM4LnhtbAoAIAAAAAAAAQAYAABYNI7a56gBkJezP4SI0QGQl7M/hIjRAVBLAQIUABQAAgAIAAAAIQCQuV2JwgAAAOsAAAAY" & _
            "ACQAAAAAAAEAAAAAANQ8AABjdXN0b21YbWwvaXRlbVByb3BzOS54bWwKACAAAAAAAAEAGAAAWDSO2ueoAZCXsz+EiNEBkJezP4SI" & _
            "0QFQSwECFAAKAAAAAAAIh3tIAAAAAAAAAAAAAAAAEAAkAAAAAAAAABAAAADMPQAAY3VzdG9tWG1sL19yZWxzLwoAIAAAAAAAAQAY" & _
            "AKAvtj+EiNEBoC+2P4SI0QEgpK0/hIjRAVBLAQIUABQAAgAIAAAAIQB0Pzl6vAAAACgBAAAeACQAAAAAAAEAAAAAAPo9AABjdXN0" & _
            "b21YbWwvX3JlbHMvaXRlbTEueG1sLnJlbHMKACAAAAAAAAEAGAAAWDSO2ueoAaC+sz+EiNEBoL6zP4SI0QFQSwECFAAUAAIACAAA" & _
            "ACEAH9A/q70AAAApAQAAHwAkAAAAAAABAAAAAADyPgAAY3VzdG9tWG1sL19yZWxzL2l0ZW0xMC54bWwucmVscwoAIAAAAAAAAQAY" & _
            "AABYNI7a56gBsOWzP4SI0QGw5bM/hIjRAVBLAQIUABQAAgAIAAAAIQA4tRoqvQAAACkBAAAfACQAAAAAAAEAAAAAAOw/AABjdXN0" & _
            "b21YbWwvX3JlbHMvaXRlbTExLnhtbC5yZWxzCgAgAAAAAAABABgAAFg0jtrnqAHQM7Q/hIjRAdAztD+EiNEBUEsBAhQAFAACAAgA" & _
            "AAAhABAcBHK9AAAAKQEAAB8AJAAAAAAAAQAAAAAA5kAAAGN1c3RvbVhtbC9fcmVscy9pdGVtMTIueG1sLnJlbHMKACAAAAAAAAEA" & _
            "GAAAWDSO2ueoAdAztD+EiNEB0DO0P4SI0QFQSwECFAAUAAIACAAAACEAN3kh870AAAApAQAAHwAkAAAAAAABAAAAAADgQQAAY3Vz" & _
            "dG9tWG1sL19yZWxzL2l0ZW0xMy54bWwucmVscwoAIAAAAAAAAQAYAABYNI7a56gB8IG0P4SI0QHwgbQ/hIjRAVBLAQIUABQAAgAI" & _
            "AAAAIQBATjnCvQAAACkBAAAfACQAAAAAAAEAAAAAANpCAABjdXN0b21YbWwvX3JlbHMvaXRlbTE0LnhtbC5yZWxzCgAgAAAAAAAB" & _
            "ABgAAFg0jtrnqAHwgbQ/hIjRAfCBtD+EiNEBUEsBAhQAFAACAAgAAAAhAGcrHEO9AAAAKQEAAB8AJAAAAAAAAQAAAAAA1EMAAGN1" & _
            "c3RvbVhtbC9fcmVscy9pdGVtMTUueG1sLnJlbHMKACAAAAAAAAEAGAAAWDSO2ueoAQCptD+EiNEBAKm0P4SI0QFQSwECFAAUAAIA" & _
            "CAAAACEAT4ICG70AAAApAQAAHwAkAAAAAAABAAAAAADORAAAY3VzdG9tWG1sL19yZWxzL2l0ZW0xNi54bWwucmVscwoAIAAAAAAA" & _
            "AQAYAABYNI7a56gBENC0P4SI0QEQ0LQ/hIjRAVBLAQIUABQAAgAIAAAAIQBo5yeavQAAACkBAAAfACQAAAAAAAEAAAAAAMhFAABj" & _
            "dXN0b21YbWwvX3JlbHMvaXRlbTE3LnhtbC5yZWxzCgAgAAAAAAABABgAAFg0jtrnqAEg97Q/hIjRASD3tD+EiNEBUEsBAhQAFAAC" & _
            "AAgAAAAhAKHsMnm9AAAAKQEAAB8AJAAAAAAAAQAAAAAAwkYAAGN1c3RvbVhtbC9fcmVscy9pdGVtMTgueG1sLnJlbHMKACAAAAAA"

Base64_19 = "AAEAGAAAWDSO2ueoATAetT+EiNEBMB61P4SI0QFQSwECFAAUAAIACAAAACEAhokX+L0AAAApAQAAHwAkAAAAAAABAAAAAAC8RwAA" & _
            "Y3VzdG9tWG1sL19yZWxzL2l0ZW0xOS54bWwucmVscwoAIAAAAAAAAQAYAABYNI7a56gBQEW1P4SI0QFARbU/hIjRAVBLAQIUABQA" & _
            "AgAIAAAAIQBclicivAAAACgBAAAeACQAAAAAAAEAAAAAALZIAABjdXN0b21YbWwvX3JlbHMvaXRlbTIueG1sLnJlbHMKACAAAAAA" & _
            "AAEAGAAAWDSO2ueoAUBFtT+EiNEBQEW1P4SI0QFQSwECFAAUAAIACAAAACEATGbSnrwAAAApAQAAHwAkAAAAAAABAAAAAACuSQAA" & _
            "Y3VzdG9tWG1sL19yZWxzL2l0ZW0yMC54bWwucmVscwoAIAAAAAAAAQAYAABYNI7a56gBUGy1P4SI0QFQbLU/hIjRAVBLAQIUABQA" & _
            "AgAIAAAAIQB78wKjvAAAACgBAAAeACQAAAAAAAEAAAAAAKdKAABjdXN0b21YbWwvX3JlbHMvaXRlbTMueG1sLnJlbHMKACAAAAAA" & _
            "AAEAGAAAWDSO2ueoAWCTtT+EiNEBYJO1P4SI0QFQSwECFAAUAAIACAAAACEADMQakrwAAAAoAQAAHgAkAAAAAAABAAAAAACfSwAA" & _
            "Y3VzdG9tWG1sL19yZWxzL2l0ZW00LnhtbC5yZWxzCgAgAAAAAAABABgAAFg0jtrnqAFwurU/hIjRAXC6tT+EiNEBUEsBAhQAFAAC" & _
            "AAgAAAAhACuhPxO8AAAAKAEAAB4AJAAAAAAAAQAAAAAAl0wAAGN1c3RvbVhtbC9fcmVscy9pdGVtNS54bWwucmVscwoAIAAAAAAA" & _
            "AQAYAABYNI7a56gBcLq1P4SI0QFwurU/hIjRAVBLAQIUABQAAgAIAAAAIQADCCFLvAAAACgBAAAeACQAAAAAAAEAAAAAAI9NAABj" & _
            "dXN0b21YbWwvX3JlbHMvaXRlbTYueG1sLnJlbHMKACAAAAAAAAEAGAAAWDSO2ueoAYDhtT+EiNEBgOG1P4SI0QFQSwECFAAUAAIA" & _
            "CAAAACEAJG0EyrwAAAAoAQAAHgAkAAAAAAABAAAAAACHTgAAY3VzdG9tWG1sL19yZWxzL2l0ZW03LnhtbC5yZWxzCgAgAAAAAAAB" & _
            "ABgAAFg0jtrnqAGQCLY/hIjRAZAItj+EiNEBUEsBAhQAFAACAAgAAAAhAO1mESm8AAAAKAEAAB4AJAAAAAAAAQAAAAAAf08AAGN1" & _
            "c3RvbVhtbC9fcmVscy9pdGVtOC54bWwucmVscwoAIAAAAAAAAQAYAABYNI7a56gBoC+2P4SI0QGgL7Y/hIjRAVBLAQIUABQAAgAI" & _
            "AAAAIQDKAzSovAAAACgBAAAeACQAAAAAAAEAAAAAAHdQAABjdXN0b21YbWwvX3JlbHMvaXRlbTkueG1sLnJlbHMKACAAAAAAAAEA" & _
            "GAAAWDSO2ueoAaAvtj+EiNEBoC+2P4SI0QFQSwECFAAKAAAAAAAIh3tIAAAAAAAAAAAAAAAACgAkAAAAAAAAABAAAABvUQAAY3Vz" & _
            "dG9tWG1sLwoAIAAAAAAAAQAYAJCXsz+EiNEBkJezP4SI0QEgpK0/hIjRAVBLAQIUABQAAgAIAAAAIQAcbkt0cAEAAP0CAAAQACQA" & _
            "AAAAAAEAAAAAAJdRAABkb2NQcm9wcy9hcHAueG1sCgAgAAAAAAABABgAAFg0jtrnqAGwVrY/hIjRAbBWtj+EiNEBUEsBAhQAFAAC" & _
            "AAgAAAAhALo27XcTAQAAEQIAABEAJAAAAAAAAQAAAAAANVMAAGRvY1Byb3BzL2NvcmUueG1sCgAgAAAAAAABABgAAFg0jtrnqAGw" & _
            "VrY/hIjRAbBWtj+EiNEBUEsBAhQACgAAAAAACId7SAAAAAAAAAAAAAAAAAkAJAAAAAAAAAAQAAAAd1QAAGRvY1Byb3BzLwoAIAAA" & _
            "AAAAAQAYALBWtj+EiNEBsFa2P4SI0QEwy60/hIjRAVBLAQIUABQAAgAIAAAAIQCfF5AfWgIAAOkDAAASACQAAAAAAAEAAAAAAJ5U" & _
            "AAB4bC9jb25uZWN0aW9ucy54bWwKACAAAAAAAAEAGAAAWDSO2ueoAcB9tj+EiNEBwH22P4SI0QFQSwECFAAKAAAAAAAIh3tIAAAA" & _
            "AAAAAAAAAAAADgAkAAAAAAAAABAAAAAoVwAAeGwvY3VzdG9tRGF0YS8KACAAAAAAAAEAGADAfbY/hIjRAcB9tj+EiNEBQPKtP4SI" & _
            "0QFQSwECFAAUAAIACAAAACEArNRiFJwAAAC5AAAAHAAkAAAAAAABAAAAAABUVwAAeGwvY3VzdG9tRGF0YS9pdGVtUHJvcHMxLnht"

Base64_20 = "bAoAIAAAAAAAAQAYAABYNI7a56gBwH22P4SI0QHAfbY/hIjRAVBLAQIUAAoAAAAAAAiHe0gAAAAAAAAAAAAAAAAUACQAAAAAAAAA" & _
            "EAAAACpYAAB4bC9jdXN0b21EYXRhL19yZWxzLwoAIAAAAAAAAQAYANCktj+EiNEB0KS2P4SI0QFA8q0/hIjRAVBLAQIUABQAAgAI" & _
            "AAAAIQC9HiMLuQAAABMBAAAnACQAAAAAAAEAAAAAAFxYAAB4bC9jdXN0b21EYXRhL19yZWxzL2l0ZW1Qcm9wczEueG1sLnJlbHMK" & _
            "ACAAAAAAAAEAGAAAWDSO2ueoAdCktj+EiNEB0KS2P4SI0QFQSwECFAAUAAIACAAAACEABW/lBigCAADVBAAADQAkAAAAAAABAAAA" & _
            "AABaWQAAeGwvc3R5bGVzLnhtbAoAIAAAAAAAAQAYAABYNI7a56gB4Mu2P4SI0QHgy7Y/hIjRAVBLAQIUAAoAAAAAAAiHe0gAAAAA" & _
            "AAAAAAAAAAAJACQAAAAAAAAAEAAAAK1bAAB4bC90aGVtZS8KACAAAAAAAAEAGADgy7Y/hIjRAeDLtj+EiNEBUBmuP4SI0QFQSwEC" & _
            "FAAUAAIACAAAACEA+2KlbbYFAACnGwAAEwAkAAAAAAABAAAAAADUWwAAeGwvdGhlbWUvdGhlbWUxLnhtbAoAIAAAAAAAAQAYAABY" & _
            "NI7a56gB4Mu2P4SI0QHgy7Y/hIjRAVBLAQIUABQAAgAIAAAAIQCByEANOQEAABQCAAAPACQAAAAAAAEAAAAAALthAAB4bC93b3Jr" & _
            "Ym9vay54bWwKACAAAAAAAAEAGAAAWDSO2ueoAfDytj+EiNEB8PK2P4SI0QFQSwECFAAKAAAAAAAIh3tIAAAAAAAAAAAAAAAADgAk" & _
            "AAAAAAAAABAAAAAhYwAAeGwvd29ya3NoZWV0cy8KACAAAAAAAAEAGADw8rY/hIjRAfDytj+EiNEBUBmuP4SI0QFQSwECFAAUAAIA" & _
            "CAAAACEAYlPvM18BAACGAgAAGAAkAAAAAAABAAAAAABNYwAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1sCgAgAAAAAAABABgAAFg0" & _
            "jtrnqAHw8rY/hIjRAfDytj+EiNEBUEsBAhQACgAAAAAACId7SAAAAAAAAAAAAAAAAAkAJAAAAAAAAAAQAAAA4mQAAHhsL19yZWxz" & _
            "LwoAIAAAAAAAAQAYAAAatz+EiNEBABq3P4SI0QFgQK4/hIjRAVBLAQIUABQAAgAIAAAAIQDAZSg4kQEAAHMOAAAaACQAAAAAAAEA" & _
            "AAAAAAllAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVscwoAIAAAAAAAAQAYAABYNI7a56gBABq3P4SI0QEAGrc/hIjRAVBLAQIU" & _
            "AAoAAAAAAAiHe0gAAAAAAAAAAAAAAAADACQAAAAAAAAAEAAAANJmAAB4bC8KACAAAAAAAAEAGADw8rY/hIjRAfDytj+EiNEBMMut" & _
            "P4SI0QFQSwECFAAUAAIACAAAACEAXMhN+ccBAACXDwAAEwAkAAAAAAABAAAAAADzZgAAW0NvbnRlbnRfVHlwZXNdLnhtbAoAIAAA" & _
            "AAAAAQAYAABYNI7a56gBABq3P4SI0QEAGrc/hIjRAVBLBQYAAAAAUgBSAHwhAADraAAAAAA="

Base64 = Base64_1 & Base64_2 & Base64_3 & Base64_4 & Base64_5 & Base64_6 & Base64_7 & Base64_8 & Base64_9 & Base64_10 & _
        Base64_11 & Base64_12 & Base64_13 & Base64_14 & Base64_15 & Base64_16 & Base64_17 & Base64_18 & Base64_19 & Base64_20

ByteArray() = Base64Decode(Base64)

Excel2010zip = Path + "\Workbook2010.zip"

Open Excel2010zip For Binary Lock Read Write As #1

For Counter = 0 To UBound(ByteArray)
    Put #1, LOF(1) + 1, ByteArray(Counter)
Next

Close #1

End Sub
'base64 representation of byte data, i.e. binary zip file corresponding to an empty Excel 2013 file with PowerPivot model removed
Sub WriteZIP2013(Path As String)

Dim Counter As Long
Dim ByteArray() As Byte

Base64_1 = "UEsDBBQAAgAIAAAIIQC1VTAj6wAAAEwCAAALAAAAX3JlbHMvLnJlbHOtks1qwzAMgO+DvYPRvVHawRijTi9j0NsY2QNotvJDEsvY" & _
            "bpe+/bzD2AJd6WFHy9KnT0Lb3TyN6sgh9uI0rIsSFDsjtnethrf6efUAKiZylkZxrOHEEXbV7c32lUdKuSh2vY8qU1zU0KXkHxGj" & _
            "6XiiWIhnl38aCROl/AwtejIDtYybsrzH8JsB1YKp9lZD2Ns7UPXJ8zVsaZre8JOYw8QunWmBPCd2lu3Kh1wfUp+nUTWFlpMGK+Yl" & _
            "hyOS90VGA5432lxv9Pe0OHEiS4nQSODLPl8Zl4TW/7miZcaPzTzih4ThXWT4dsHFDVSfUEsDBAoAAAAAAFiJe0gAAAAAAAAAAAAA" & _
            "AAAGAAAAX3JlbHMvUEsDBBQAAgAIAAAIIQA8Wa26xwAAAGgBAAATAAAAY3VzdG9tWG1sL2l0ZW0xLnhtbH2QwW6CYBCEv0cx3hFK" & _
            "1VaDejBpj73o2VBQS6LYtNQYH7714/fCqdkAszuzOxP+fjMWXDhyoMeZLV98U3GiZkafBwYkfnsyNYXzUrZmH9g1K16IVI3tFszJ" & _
            "eFV5DJrKrfvl2put/oPG+mRKbO07ythppf9JvuBHfSNu2Su5+J4o1i/n3Ytb3kxahrz94LvsbC2DugmZG9lC5szGfujOk+9HRj4R" & _
            "Kc9MTB85y+UTUW7txO2sUJ/aj0WJO6WzkW7xv34t3/0Pc25QSwMEFAACAAgAAAghACecf7W3AAAATgEAABQAAABjdXN0b21YbWwv" & _
            "aXRlbTEwLnhtbIWPQQ/BQBCFv59S7iwXB6HSVLhwq5M4VNtUQ0toRfx4vN1eOMkcdna+N29m3q8JMx6UnPC4k3HlRsGZiildhvQZ" & _
            "6PVEKhLVU9GK3NENEQt6Uo30m+EzYSll6TSFulrnSp5Wf6BWXBhjFPmX0qhaaP5ZPKGRvlZu6ZNYebuRYeX0R3Wmmh2zl3umPS76" & _
            "W13G2u2YaZrdJvzyCp1H7S6pHe2wVXVOIK9AeaTrG/GdwnLzp9/8XOvzAVBLAwQUAAIACAAACCEACMzEL60AAAA8AQAAFAAAAGN1" & _
            "c3RvbVhtbC9pdGVtMTEueG1shY9ND8FAEIafn4I7qxcHoSL1cXPBSRwooYluG5aUH4+320udZA47M+8zM+9+3gNGFKRcaPDgyJUb" & _
            "CRmWIS0COnT1NqRYYvUPUi0nr65ZMaMtqqdqRMiAucjUM4mmqs1WO0v+jFPk9DGKU4006ia6n0mPuYt3ykv1xU555ciwVGXlYa+6" & _
            "YOG7U3G5mKculA6i2nzkCefdO6822ag7YSzvY+UBW0WpmD+T5udvIV9QSwMEFAACAAgAAAghADr+1MmyAAAARAEAABQAAABjdXN0" & _
            "b21YbWwvaXRlbTEyLnhtbIWPTQ/BQBCGn59S7iwXB6lKUx9xxkkctJVqoku0pPx4vF0XTjLJ7uz7vDM783r6jKkpOOJxY8+FkpwT" & _
            "lhFt+nTp6fZELIn0VNSSObpmxYyOXAO9xgT4zOUsnCdX1aezVc/Gf6BSnBliFNmX00jN9f9JPOEqf6W8oQ92yj8TGRYiSylWc8TS" & _
            "aqZyxeqUKpqz7aaIvnpErrZyG1SOtthInRBq/lD5Xaxkq2io+VNtfnYMeANQSwMEFAACAAgAAAghAIZDXiK3AAAAXgEAABQAAABj" & _
            "dXN0b21YbWwvaXRlbTEzLnhtbIWPTw/BUBDEfx8Fd15VgkiRpsLVARdxkBKa0Ar1Jz48pisRTrLJ23k78+bNPh8BfW7s2VHiwpoj" & _
            "JxIyUrpUqFPDUy+JSYk1X4lN2Rg7ZcKQqlRN3fr0CBhJuTdNoldv51SehX5LrjrQwak2X0qnaaL/M/ExZ+lz4YK9sxR+J3KM1a+W" & _
            "cvzRz35SVyxF9OUR2Ty3DXJjy8w1HRAqfyjsa8e67VqcnrpPm4ahlmqhKt65P77uZ/seL1BLAwQUAAIACAAACCEAppMiY70AAABk" & _
            "AQAAFAAAAGN1c3RvbVhtbC9pdGVtMTQueG1shZBND8FAFEXPTyl7ho2FUJH62Asrsahqqkk7RIeIH4/bkUit5OV1Zt49982bvp4j" & _
            "JtwpKQi4kXKhIueEZUybPl16WgMplkT1g1RL5tUNaxZ0RA10mhAyYimy9Ewu16ezVc+aP+IUZ4YYRdYgjaq57j9JT7iKd9rX6oNY" & _
            "+89EhpU8RaNSqWcu71Se2jGT7pTJl5grY/ZypZq97WeMGjdEnnL+fc6rLbaqztRzrdzqexGfslPUuvnjNz//IOQNUEsDBBQAAgAI" & _
            "AAAIIQBRJjBrmwEAAEoEAAAUAAAAY3VzdG9tWG1sL2l0ZW0xNS54bWy1U11PwkAQnJ+CvJe2gh8hCKmgvkBiAiYmxodaKjbSYtqK"

Base64_2 = "4I9X56YlghDCi7m0t7c7tzO3e/f91UIHC8SYooI5QqTIEGGGBBeowkUNDucKIwkC+seMJpgoeocRrmERdcpVB220cENkLEzEXUXm" & _
            "hDkN/gU5xxuasDkma0ib3oj8M8YDvBOf0zbRT/i0C0U2rqgvpZ2iS3/AjCHzGt7u2q6u0Lk054oe4YHeHjwq9mi3aJu8PgY6U0iV" & _
            "xbmG9CX0PNG/4OmHqklE71RqQsY2Mbs07T95VqJi4jNyjEstQak7lZ0zYrJOuOOYXXDQoOXgjP8BtQRizfg9C+txr1G5VAezUvlc" & _
            "yFA8l8r7SsVGfXVDZZO4baUfGjV+9S01Lq17KumT6fc8luqY6TyJmIv+7KpST9pW3fUZW/7R5B9QvXhHLQLdg5XSOi2Xs73VzfW7" & _
            "5Umhr/pVuTaq++LIcav8RR0z3YER98W02uIwL8ASk8X1GaMu/011q8m+negdNbR2cS6cUyIc8tgHMRnc/97bNh45Cqb9L8reeOlt" & _
            "/ABQSwMEFAACAAgAAAghAMBjTJfbAAAAigEAABMAAABjdXN0b21YbWwvaXRlbTIueG1shZBNS8NAFEXPT6nu48Rq6we1pUR0oztF" & _
            "QURCojVgU9FYxR+vnpm6qCt5JHPnfrz3ku+vERM+mPNEjyX3vPBKw4KWIzbZZovcs6fSUsnXqi2zpF5ywQmZrqG3CWNGnOqcJ09j" & _
            "atW5tWf0P9JZzxwSrNmaM8g2zl+oV7zp78RR/aQUrzYKFHZr0i4dVylbq7xzzTlnTogbFGv5IuW630RUN7iRPWbq7lNxpXfJnY5d" & _
            "e+353mHgk9FnnwO/LJMr1XNRaT2II1fp73sfinIztdyAWyvOCf/sEf78qTE/UEsDBBQAAgAIAAAIIQDlij+arwAAAEIBAAATAAAA" & _
            "Y3VzdG9tWG1sL2l0ZW0zLnhtbIWPQQ/BQBSEv59S7iwXB6FNU+HkxkkcmmqqiS7RJeLHY3Zd6iQv2X07M2/e7Ps1I+FBw4mIOyVX" & _
            "WmrOWOb0GTNkpDsSYymEH8RaqsBu2bBkINVEr4SYGSspm6CpNfV1tvL0+iNOdWGKUVUdpRFaa/9ZfMFNeqfes09y9d9EhrVeVnwu" & _
            "1yychTCfqpS/3591prMw5UJ2F9geO6ELUiVP1S+DRyvFXuUV5o+D+flhzAdQSwMEFAACAAgAAAghAGoy1EepAAAAOAEAABMAAABj" & _
            "dXN0b21YbWwvaXRlbTQueG1shY+7DoJAEEXPpyC9rjYWBiEEo/ZiZSwMGCQRMApq/Hj1sjRYmVvs7Jw7r8/bI+BJwRmHO0eu3Mip" & _
            "KJnjMmHEWK8jUpIon4qWZJZuiVkylGuqX4CPx0rOwnpyVXWdS/Vs/Sdq6cIMI2U9p1E21/xKPKGRv1bc0hcHxd1Gho16VDxYK5NK" & _
            "7VaunRv1qiLrri2tLR2wU3ZBqI1DxbHubMT3UsvNn3rzc5fPF1BLAwQUAAIACAAACCEALp32ErkAAABOAQAAEwAAAGN1c3RvbVht" & _
            "bC9pdGVtNS54bWyFj0sPwVAQhb+fQvdcNhbikaZSsbDCSiwamrpJX6FF/HicXptaySxm7pwz3515vybMeZCR0uFGzIUrloKcKR5D" & _
            "+gyUO1JyjuqfpOYkTt2xJaQn10ivOTMmLOXMnMdq6kvOxWz8ZypFyRijSFpOo67V/4X0I7X8lepGfRKp/m5k2IhRcGclrRTZym2l" & _
            "r8WKNFVr/1jZc7sELVLgCJW7o3Jql726C3xd4asORUjljzkoGof5QzA/1874AFBLAwQUAAIACAAACCEAxcQvSTkCAACWCAAAEwAA" & _
            "AGN1c3RvbVhtbC9pdGVtNi54bWzNln1v00AMxp+PMvi/S9q9AFPpVHUamgQIqSCQEJrSpOsimhQl6Vj34YHfOc1Iu1A2XqTq1M5n" & _
            "P2f7Hru+ff/W1bGulWiqHV1prEy5Ys2U6rkeq61d+fzdwZIqRB9hTTUx6zu91alaoA7ZHaunrl6ATAwTc6r0nOLT4S9VsL7oSB5r" & _
            "UkN6aGPiz7CHmoMvkJ31RgFymZFHvEAjPI71Qa/0Uuegc86do9kntyd87+mAT0sdPdUzMmuhC7D7SAHrAtnpQvAd9odIPmcidAdk" & _
            "6W4xqOUwsNiFMVCY9ZE+oj1Rn3z6yN1aXu/JNbK7FXCR2W5ovMZgpms3qjN0hOy4v8vUV1u7fPb4nuFrgrZD1j7ce0suhtzj0jgN" & _
            "uF9sPgtkV7ex1XA9VvRPYlWMzfA9R5NinXPKcZAbbxH8LIgwJgfvDvrUoiTGSpO9TyzXExmIED9NmIHZs2WXLshggX5kiCb8G2TH" & _
            "dWFdN2YfGAuuTmnjidfYJsvK3e9E2QkFPOVmiw2dmPwZaWGSYyhb/qZ6+OziaVXnND/xV9ZFc/alz7In96mQw8W3PeqtIb2V+PfJ"

Base64_3 = "ZbRFuYR/nEv7Qbl4GytYt51gy+npKd4WOkMT4eX6v1fa35o6t7emyp2/qPGmOq5OqYzdjU2Y5gnk5mFouUa/mHMxcjkxKrt7GaZ2" & _
            "rnoRmlB1nZs7bhIm9sKU73WFO0Me2jQtuJebvqHlG92yeWG85GgdD7/Hd1fe3Ie9bT19YlWMb3pPvZX/Gnr6AVBLAwQUAAIACAAA" & _
            "CCEAt8ezgLMAAABEAQAAEwAAAGN1c3RvbVhtbC9pdGVtNy54bWyFj08PwUAQxX8fBXdWLw7iT6RCnHESB6WpJmyFrTQ+PN5uL3WS" & _
            "SXZn3nvzZubzHjGl4sqFFk9S7jzIKbCM6RDRo6+/JcZyFH4Sa8kCu2XDgq5UA1VTJoxYSnkNmlxdtbOVp9efcYobQ4wiayiN0Fzz" & _
            "C/FHSumdcs++OCivNzKadyCRY0ospBTmWOldC7faLRFaaZLfJG74xKHfhStcYNvshM6ZyXOmPGKv8Iz502l+bpzwBVBLAwQUAAIA" & _
            "CAAACCEA7RWtsVIGAAAMaAAAEwAAAGN1c3RvbVhtbC9pdGVtOC54bWztXWuP2kYUvT+F8j14N6laaUWI0L4SZR9V2bSR2qoisGVp" & _
            "eVjY7G7y49ucuWNje/wAG2ZMwwiBjWfm3nPvmZfnYf/3b5ve0DNNaUINeqR7WpBHY5rTjF5Tk46pRUc4NhAyowGuDxE6oxGHfqA7" & _
            "uqAXiPUD/r2hDrXpEjGnHGeMVFLyDDJF/Afy8XHphBx8RrGYDq6OoX+O8AEtEd/HuQj9Qn2cS0QOneGsj5QL/E4Rq8k6T2MpTjmm" & _
            "z3h9Dv2OfsPVM+oCbRfnbfwuWMJnuqW/FJnX+J3xf+GLFvX4KGJMAjSfcHavpCq21INND2xtH+ct+FDY1Gd/CpQLPvcRMsf5CCle" & _
            "wutH9D3OjuhH/F5D2wBhc6SfA7OI22WcE1ghGPMCpI8c8571nLJHpqylmcB4glhpnE/8aeH7KoXlGGcfgeMKeiJrXjB/HlszY72S" & _
            "kV34tMNMCV+5zOcCFoyB1IfFLv4L/Nc4Cp8uEaqm77HWIWTPYXlSkkR5t9J8wynucXUAaY8IczJDxfWkpLS1dwG+TiG6dio3R+mk" & _
            "RA//JpxLozz9jN91Xrhk3w4LZYTlZo7QJXTP2NoRvUf4Zw4Lz3rMrc+5qhHE8uj3RFphS5TSKZR7zmVvzvWMDLvk/0sgTWtXYwjN" & _
            "cQmq5s2kX3CMKa4K35S1PJlaRVAs+y0zM+TaqazeeFpVa5Hct1xr+Il8Oee4E87Bv3IZHuLaUyBLjd+md7DI5dI6wNdXclyPPbzg" & _
            "8h/n/ool/1PaUlVbQ9HXCEp2I6G5kdDdWGlXfbVbW4qllbfc4/Kt2j9NSC1nURqDaGvLIntPP0GPqjlb1nY+uC60tVj2L9ymLCto" & _
            "jVKqGvNlOoV1bLoez2sZwho7LTFeBqXedO2fbOfOuYfmcy2UrUXK2aYFjDPd53LhB8fyrXSrMDRLfjPoGUzwkUi8jFb4Fsj/Rqjo" & _
            "V6VzgVqfpNtmleV1ksto7/J12acVee8sqIv9jNynW7fMs7LnL2oakXMaOaVdN5Yw79al/2fu1c7ZG3VhEHy4XM6L60HdOLrQK/JC" & _
            "EsWfqG2WNZSQPDTdIOfKe4t9QXXKfQhRa+0LIlGPPu8RmnFGH7YuNLLF9RmTSLF/7Ik+y5Bbice9xOTuDSrRV1vsGRq3tnbUfNu1" & _
            "YA/U0Y8J77v7Ndgd9VvOeORtwD1tdxXHNB45GnCf6kN4xpF8wPGhNjRZo0fhHVU92sNRhEbNOMz0bcuO3pnQH93Ne7FxpOSoXT3I" & _
            "wpEJ/XVnlTFdUxjM3vXUMwJpygozpTxPey+YI5qtwqKyX18OF6V/EJTzsKXaB2T544+mEBTPa5hCcY77CemJOD9j/nWZo1HQe/Bq" & _
            "rB2qjNKbQlc0T2QKww3P5Ee8yXUFbiBnzKM29bIn71SGnL/EHLpuNPGWQaDoG9f4ybjGQSWNTsbYfi8YIR9w/2i4CnNWs0LpeQI5" & _
            "HzTG+VOsb3URrLUQdV1yRmbbFRatnPmWdRjkHNOmOLNs8pS1Hf0N1qBMM+wdsDXhio9XODvG0clYrxGtyInW0oSrcfrQKPmJ2rT0" & _
            "Ops096KnEFn6BeFPqEfueO71oyK5+lxOUob4F6Ess7Ii0pDmI5xbk/OAUS5NalPx6POWGR7KzGrtgoW0398ptsURNRMsfJseLzOX" & _
            "ZxnQwcC6GUzrdR1e32Te1npeV42zyWy19b4O75ebo7ccmOSgeGWC5cIkF/kz+pYHkzxkr0KxHJjlIGvtjeXAJAebrziyvJjtN+Wv" & _
            "s7JM1MWEa7molYvsNXWWA/Mc2JJgdhTPetvMGFL+qlHrcR0eL1oraz2ud15gsxXClgUdLGy+Ltr6X4f/y60GtxzsgoNya+DN+FzF"

Base64_4 = "dCh+32T1v2VAJwN1zg0fqs+LdnxYz+v0fLW9LpYTnZzk7/Kxft+V36vsbzLp/UPy+T6sCzo8r+vdv2ZZ08Na/b3TQ/J2mf2JlgNd" & _
            "LXLZ/ZiWCT1M5O8/tR7X4/Hi/bbW63q8vs3+YsuJrt5qlf3Ulg09bBTtH7c+1+PzKvvlLRc6a6N1zwew3t/O++ufh7ALD98ETzrP" & _
            "26N9xTvXxRMFbnnlXYffxhD2gtOh3/4u7vVPjTDBSxxFh45TbxWw3J1Uev6Gee5e/k+5czKfs5H1FJPqbzdxNLyJpkN/4BO+h6Po" & _
            "vThO4l09HfoKUEsDBBQAAgAIAAAIIQAozjA39AEAAOYFAAATAAAAY3VzdG9tWG1sL2l0ZW05LnhtbMWUbU/bUAyFz09h+94mtMAG" & _
            "6oKq8jJpTJOWDU1CCGVNKJHaFCXhpfvx2x470Ba6MT4xXfXW1z62j+178+tnT7u61URjrelamUpVyjVVoXd6rXW1FfK/hqXQEH2K" & _
            "tdDIrV/1RQdqgdritKtIPR2CnDgmx6uJXBDT8BeqWZfaUcAaLSEDtDn5p9iHugJfI5v1hxLkhlGgj/gkWK/gmZGrxJIqBmGojBzG" & _
            "YbAUYeCetfOv3fpKJ2j31Id9H7nHXrISzfRJ5/oAdqZjzmPiZK6zaOVd5TGWgqzfiXyrfe+I5Srp1b/4xVSX4tPVpj6T5ekOVaAv" & _
            "vEsWs+39GBJzymkKqxrd0KsM1GFOIXEDphGyxz7L3Kt43MVFvZXPdsFhB8wqjxtfbX5ddqt0NM+4jvSNuo+8tnu2Le9V5XXbvbmf" & _
            "zEv2dpEvwlJxu844bYB5w26oLjw7eqtt7m8LXYI9REpY58ims5gdzltIIT4puk0iB/PYvaV6HvYyecZMU2ef+BwL74TJ9UqfN5BC" & _
            "mNsbWL0FfXyNw8xfb3U3/WtHZp5n4Pdk4lmaWSQweu8sDtA2ry5yBlaJVfhnROMZ+63KyFzzluz1Whcj9u25798wTYRjZ5oz6zGY" & _
            "1cyP7aZfdDpamsHL3afgP34rIp2yGhZPf+GCB1/hSL8BUEsDBBQAAgAIAAAIIQAcE7QGwQAAAOsAAAAYAAAAY3VzdG9tWG1sL2l0" & _
            "ZW1Qcm9wczEueG1sXU7LasMwELwX+g9i74rk1E6VYDkkEYZcSwu9CnmdGKzdYCmlUPrvVemtp2FmmEe7/4yz+MAlTUwWqpUGgRR4" & _
            "mOhi4e21lwZEyp4GPzOhBWLYd48P7ZB2g88+ZV7wnDGKIkwFz87ClzNHc6ibZ7ltXCPrzelJmtO6l1pXlaldvz0e3DeIMk2lJlm4" & _
            "5nzbKZXCFaNPK74hFXPkJfpc6HJRPI5TQMfhHpGyWmu9UeFe5uN7nKH7/fOXfsExqa5V/w92P1BLAwQUAAIACAAACCEA5YJ8jMEA" & _
            "AADrAAAAGQAAAGN1c3RvbVhtbC9pdGVtUHJvcHMxMC54bWxdTstqwzAQvBfyD2LvimQrDxEsBztOINfSQq5CXicGaxUspRRC/r0q" & _
            "vfU0zAzzqPbffmJfOMcxkIFiKYEhudCPdDXw+XHiGlhMlno7BUIDFGBfL96qPu56m2xMYcZzQs+yMGY8dwae7VHKRivNm5NSfKXV" & _
            "luvt+sBbVXZN2bXFelO8gOVpyjXRwC2l+06I6G7obVyGO1I2hzB7mzKdryIMw+iwC+7hkZIopdwI98jz/uInqH///KXfcYiirsT/" & _
            "g/UPUEsDBBQAAgAIAAAIIQDG7oQjwQAAAOsAAAAZAAAAY3VzdG9tWG1sL2l0ZW1Qcm9wczExLnhtbF1Oy2rDMBC8F/oPYu+Kto4T" & _
            "28FySIkDuZYWehXyOjFYu8FSSqH036vSW0/DzDCPdv8ZZvVBS5yELTytEBSxl2Hii4W315OuQcXkeHCzMFlggX33+NAOcTe45GKS" & _
            "hc6JgsrClPF8tPC1PlRr7A+NLrBAXWK91U3ZlLqq+2rzXOAGT/03qDzNuSZauKZ02xkT/ZWCiyu5EWdzlCW4lOlyMTKOk6ej+Hsg" & _
            "TqZA3Bp/z/PhPczQ/f75S7/QGE3Xmv8Hux9QSwMEFAACAAgAAAghAHMogKjBAAAA6wAAABkAAABjdXN0b21YbWwvaXRlbVByb3Bz" & _
            "MTIueG1sXU5Ni8IwFLwv+B/Cu8dUXRORplJbBa/LLngN6asWmvekiSIs/vfNsrc9DTPDfJS7ZxjFA6c4MFlYzAsQSJ67gS4Wvj6P" & _
            "cgMiJkedG5nQAjHsqtlb2cVt55KLiSc8JQwiC0PGU2vhW7dmU+umlofjqpDv9VrLvW4W0jTr1ui9aWqzeoHI05RrooVrSretUtFf" & _
            "Mbg45xtSNnuegkuZThfFfT94bNnfA1JSy6LQyt/zfDiHEarfP3/pD+yjqkr1/2D1A1BLAwQUAAIACAAACCEAbmnhNsEAAADrAAAA"

Base64_5 = "GQAAAGN1c3RvbVhtbC9pdGVtUHJvcHMxMy54bWxdTstqwzAQvBf6D2LviiTHpHawHOokhlxLC7kKeZ0YrN1gKaVQ+u9R6a2nYWaY" & _
            "R7P7CrP4xCVOTBbMSoNA8jxMdLHw8d7LCkRMjgY3M6EFYti1z0/NELeDSy4mXvCUMIgsTBlPBwvfpqteTFH1sjyWa1nqrpB1bfby" & _
            "VW/Wdd+ZfW+qHxB5mnJNtHBN6bZVKvorBhdXfEPK5shLcCnT5aJ4HCePB/b3gJRUofVG+XueD+cwQ/v75y/9hmNUbaP+H2wfUEsD" & _
            "BBQAAgAIAAAIIQDaA1kowQAAAOsAAAAZAAAAY3VzdG9tWG1sL2l0ZW1Qcm9wczE0LnhtbF1OTWsCMRS8F/wP4d1jsqvYRDYr6hrw" & _
            "WlrwGrJvdWHzIptYCqX/3ZTeehpmhvlodl9hYp84pzGSgWopgSH52I90NfDxbrkClrKj3k2R0ABF2LWLl6ZP295ll3Kc8ZwxsCKM" & _
            "Bc+dge+VtVbJTvODrmq+lqc9V+r1yNdWWa2q/anW+gdYmaZSkwzccr5vhUj+hsGlZbwjFXOIc3C50Pkq4jCMHrvoHwEpi1rKjfCP" & _
            "Mh8uYYL2989f+g2HJNpG/D/YPgFQSwMEFAACAAgAAAghAMg6WWLCAAAA6wAAABkAAABjdXN0b21YbWwvaXRlbVByb3BzMTUueG1s" & _
            "XU7LasMwELwX8g9i74rkPGw1WA5xkkKupYFehbxODNZusJRQKP33KvTW0zAzzKPefoVRPHCKA5OFYq5BIHnuBrpYOH+8SQMiJked" & _
            "G5nQAjFsm9lL3cVN55KLiSc8JQwiC0PG08HCd7szZrk/alma5Vqu2raUbVEtpNHHfbUrqteyWv2AyNOUa6KFa0q3jVLRXzG4OOcb" & _
            "UjZ7noJLmU4XxX0/eDywvwekpBZal8rf83z4DCM0zz9/6Xfso2pq9f9g8wtQSwMEFAACAAgAAAghAHsWQgHBAAAA6wAAABgAAABj" & _
            "dXN0b21YbWwvaXRlbVByb3BzMi54bWxdTstqwzAQvBfyD2LvilTjuHawHBo/INfSQq5CXicGazdYSiiU/ntVeutpmBnmUR8+/SIe" & _
            "uIaZycDzVoNAcjzOdDHw8T7IEkSIlka7MKEBYjg0m6d6DPvRRhsir3iK6EUS5oSnzsBXv6vyvn05ymJoc5nvdC+r16GSOj9mhS67" & _
            "NiurbxBpmlJNMHCN8bZXKrgrehu2fENK5sSrtzHR9aJ4mmaHHbu7R4oq07pQ7p7m/dkv0Pz++Uu/4RRUU6v/B5sfUEsDBBQAAgAI" & _
            "AAAIIQDqYiG9wgAAAOsAAAAYAAAAY3VzdG9tWG1sL2l0ZW1Qcm9wczMueG1sXU7LasMwELwX+g9i74qstLHjYDm4dgy5lhZ6FfI6" & _
            "MVi7wVJKIeTfo9JbT8PMMI9q/+Nn8Y1LmJgM6FUGAsnxMNHJwOdHL7cgQrQ02JkJDRDDvn5+qoawG2y0IfKCx4heJGFKeOwM3PJW" & _
            "d33ZbmXTr7V8LVoty/6lkU2u34rNoTsURXkHkaYp1QQD5xgvO6WCO6O3YcUXpGSOvHgbE11Oisdxctixu3qkqNZZlit3TfP+y89Q" & _
            "//75S7/jGFRdqf8H6wdQSwMEFAACAAgAAAghAHpfu5LCAAAA6wAAABgAAABjdXN0b21YbWwvaXRlbVByb3BzNC54bWxdTstqwzAQ" & _
            "vBfyD2LvihTbieNgOSRVA7mWFnoV8joxWLvBUkqh9N+rkltPw8wwj3b/FSbxiXMcmQyslhoEkud+pIuB97eT3IKIyVHvJiY0QAz7" & _
            "bvHU9nHXu+Ri4hnPCYPIwpjxbA18v6yronkuS1lu11pW+rCRTV030h4PtS6qamWt/gGRpynXRAPXlG47paK/YnBxyTekbA48B5cy" & _
            "nS+Kh2H0aNnfA1JShdYb5e95PnyECbq/P4/0Kw5Rda36f7D7BVBLAwQUAAIACAAACCEAydPslcAAAADrAAAAGAAAAGN1c3RvbVht" & _
            "bC9pdGVtUHJvcHM1LnhtbF1OTYvCMBS8C/sfwrvXpNqtRZqK9gO8isJeQ/qqheY9aeKyIP73zbI3T8PMMB/l7sdN4htnPzJpSJcK" & _
            "BJLlfqSrhsu5SwoQPhjqzcSEGohhV30syt5vexOMDzzjMaATURgjHhsNz3XdqVXatInqsjbJskOd7PeRpptCHdZtnRf55wtEnKZY" & _
            "4zXcQrhvpfT2hs74Jd+Rojnw7EyIdL5KHobRYsP24ZCCXCmVS/uI8+7LTVD9/flPn3Dwsirl+8HqF1BLAwQUAAIACAAACCEAAZmH" & _
            "ScAAAADrAAAAGAAAAGN1c3RvbVhtbC9pdGVtUHJvcHM2LnhtbF1OTYvCMBS8C/sfwrvHdKu7RGkqrm3A66Kw15C+aqF5T5q4LIj/"

Base64_6 = "fSPePA0zw3xUm78wil+c4sBk4H1egEDy3A10MnA8WKlBxOSocyMTGiCGTf02q7q47lxyMfGE+4RBZGHIuG8M3Har1bLd7haybD60" & _
            "XH61Vm51Y2W7aIvMSq2tvYPI05RrooFzSpe1UtGfMbg45wtSNnuegkuZTifFfT94bNhfA1JSZVF8Kn/N8+EnjFA//jzT39hHVVfq" & _
            "9WD9D1BLAwQUAAIACAAACCEAxXve78IAAADrAAAAGAAAAGN1c3RvbVhtbC9pdGVtUHJvcHM3LnhtbF1Oy2rDMBC8F/oPYu+KbDlO" & _
            "nWA5pFUCuZYGchXyOjFYu8FSSqD036vSW0/DzDCPdvsIk/jEOY5MBspFAQLJcz/SxcDp4yAbEDE56t3EhAaIYds9P7V93PQuuZh4" & _
            "xmPCILIwZjxaA191+VbZVV3J/b6p5VLrRq6bFyvtujy86p3WRVV/g8jTlGuigWtKt41S0V8xuLjgG1I2B56DS5nOF8XDMHq07O8B" & _
            "KSldFCvl73k+nMME3e+fv/Q7DlF1rfp/sPsBUEsDBBQAAgAIAAAIIQCrK+5EwgAAAOsAAAAYAAAAY3VzdG9tWG1sL2l0ZW1Qcm9w" & _
            "czgueG1sXU7LasMwELwX+g9i74pU23FNsBzqyIFcSwu5CnmdGKzdYCmlUPrvVemtp2FmmEe7/wyL+MA1zkwGnjYaBJLncaaLgfe3" & _
            "o2xAxORodAsTGiCGfff40I5xN7rkYuIVTwmDyMKc8WQNfG3Lquprq2Xx3BeysuUg+2Z4kdvmcLSNHg5FXX6DyNOUa6KBa0q3nVLR" & _
            "XzG4uOEbUjYnXoNLma4XxdM0e7Ts7wEpqULrWvl7ng/nsED3++cv/YpTVF2r/h/sfgBQSwMEFAACAAgAAAghALhZq7jCAAAA6wAA" & _
            "ABgAAABjdXN0b21YbWwvaXRlbVByb3BzOS54bWxdTstqwzAQvAf6D2LviuzUta1gObR2ArmWFHoV8joxWLvBUkoh5N+r0ltPw8ww" & _
            "j2b37WfxhUuYmAzk6wwEkuNhorOBj9NB1iBCtDTYmQkNEMOufVo1Q9gONtoQecFjRC+SMCU89gbubwddFVXdybzab2TxmpeyLnot" & _
            "e10W5fNLt9e6e4BI05RqgoFLjNetUsFd0Nuw5itSMkdevI2JLmfF4zg57NndPFJUmywrlbulef/pZ2h///yl33EMqm3U/4PtD1BL" & _
            "AwQKAAAAAABYiXtIAAAAAAAAAAAAAAAAEAAAAGN1c3RvbVhtbC9fcmVscy9QSwMEFAACAAgAAAghAHQ/OXq8AAAAKAEAAB4AAABj" & _
            "dXN0b21YbWwvX3JlbHMvaXRlbTEueG1sLnJlbHONz7GKwzAMBuD94N7BaG+c3FDKEadLKXQ7Sg66GkdJTGPLWGpp377mpit06CiJ" & _
            "//tRu72FRV0xs6dooKlqUBgdDT5OBn77/WoDisXGwS4U0cAdGbbd50d7xMVKCfHsE6uiRDYwi6RvrdnNGCxXlDCWy0g5WCljnnSy" & _
            "7mwn1F91vdb5vwHdk6kOg4F8GBpQ/T3hOzaNo3e4I3cJGOVFhXYXFgqnsPxkKo2qt3lCMeAFw9+qqYoJumv103/dA1BLAwQUAAIA" & _
            "CAAACCEAH9A/q70AAAApAQAAHwAAAGN1c3RvbVhtbC9fcmVscy9pdGVtMTAueG1sLnJlbHONz7GKwzAMBuC90Hcw2i9OOhxHidOl" & _
            "FLqVksKtxlES09gylnpc376mUws33CiJ//tRu/sNi/rBzJ6igaaqQWF0NPg4Gbj0h48vUCw2DnahiAbuyLDr1qv2jIuVEuLZJ1ZF" & _
            "iWxgFklbrdnNGCxXlDCWy0g5WCljnnSy7mon1Ju6/tT51YDuzVTHwUA+Dg2o/p7wPzaNo3e4J3cLGOWPCu1uLBS+w3LKVBpVb/OE" & _
            "YsALhueqqauCgu5a/fZg9wBQSwMEFAACAAgAAAghADi1Giq9AAAAKQEAAB8AAABjdXN0b21YbWwvX3JlbHMvaXRlbTExLnhtbC5y" & _
            "ZWxzjc+xisMwDAbgvdB3MNovTjocR4nTpRS6lZLCrcZREtPYMpZ6XN++plMLN9woif/7Ubv7DYv6wcyeooGmqkFhdDT4OBm49IeP" & _
            "L1AsNg52oYgG7siw69ar9oyLlRLi2SdWRYlsYBZJW63ZzRgsV5QwlstIOVgpY550su5qJ9Sbuv7U+dWA7s1Ux8FAPg4NqP6e8D82" & _
            "jaN3uCd3CxjljwrtbiwUvsNyylQaVW/zhGLAC4bnqmmqgoLuWv32YPcAUEsDBBQAAgAIAAAIIQAQHARyvQAAACkBAAAfAAAAY3Vz" & _
            "dG9tWG1sL19yZWxzL2l0ZW0xMi54bWwucmVsc43PsYrDMAwG4P2g72C0N046lOOI06UcdDtKCrcaR0lMY8tY6nF9+5pOLXToKIn/"

Base64_7 = "+1G7+w+L+sPMnqKBpqpBYXQ0+DgZOPXf609QLDYOdqGIBq7IsOtWH+0RFyslxLNPrIoS2cAskr60ZjdjsFxRwlguI+VgpYx50sm6" & _
            "s51Qb+p6q/OjAd2TqQ6DgXwYGlD9NeE7No2jd7gndwkY5UWFdhcWCr9h+clUGlVv84RiwAuG+6rZVAUF3bX66cHuBlBLAwQUAAIA" & _
            "CAAACCEAN3kh870AAAApAQAAHwAAAGN1c3RvbVhtbC9fcmVscy9pdGVtMTMueG1sLnJlbHONz8FqwzAMBuD7oO9gdG+crDDKiJNL" & _
            "GfRWRga7GkdJTGPLWGpZ375mpxV26FES//ejtv8Jq7piZk/RQFPVoDA6Gn2cDXwNH9s9KBYbR7tSRAM3ZOi7zUv7iauVEuLFJ1ZF" & _
            "iWxgEUnvWrNbMFiuKGEsl4lysFLGPOtk3dnOqF/r+k3nvwZ0D6Y6jgbycWxADbeEz9g0Td7hgdwlYJR/KrS7sFD4DuspU2lUg80z" & _
            "igEvGH5Xza4qKOiu1Q8PdndQSwMEFAACAAgAAAghAEBOOcK9AAAAKQEAAB8AAABjdXN0b21YbWwvX3JlbHMvaXRlbTE0LnhtbC5y" & _
            "ZWxzjc/BasMwDAbg+6DvYHRvnIwyyoiTSxn0VkYGuxpHSUxjy1hqWd++ZqcVduhREv/3o7b/Cau6YmZP0UBT1aAwOhp9nA18DR/b" & _
            "PSgWG0e7UkQDN2Tou81L+4mrlRLixSdWRYlsYBFJ71qzWzBYrihhLJeJcrBSxjzrZN3Zzqhf6/pN578GdA+mOo4G8nFsQA23hM/Y" & _
            "NE3e4YHcJWCUfyq0u7BQ+A7rKVNpVIPNM4oBLxh+V82uKijortUPD3Z3UEsDBBQAAgAIAAAIIQBnKxxDvQAAACkBAAAfAAAAY3Vz" & _
            "dG9tWG1sL19yZWxzL2l0ZW0xNS54bWwucmVsc43PwWrDMAwG4Pug72B0b5wMOsqIk0sZ9FZGBrsaR0lMY8tYalnfvmanFXboURL/" & _
            "96O2/wmrumJmT9FAU9WgMDoafZwNfA0f2z0oFhtHu1JEAzdk6LvNS/uJq5US4sUnVkWJbGARSe9as1swWK4oYSyXiXKwUsY862Td" & _
            "2c6oX+v6Tee/BnQPpjqOBvJxbEANt4TP2DRN3uGB3CVglH8qtLuwUPgO6ylTaVSDzTOKAS8YflfNrioo6K7VDw92d1BLAwQUAAIA" & _
            "CAAACCEAXJYnIrwAAAAoAQAAHgAAAGN1c3RvbVhtbC9fcmVscy9pdGVtMi54bWwucmVsc43PwYrCMBAG4PuC7xDmblM9iCxNvSyC" & _
            "N5EueA3ptA3bZEJmFH17g6cVPHicGf7vZ5rdLczqipk9RQOrqgaF0VHv42jgt9svt6BYbOztTBEN3JFh1y6+mhPOVkqIJ59YFSWy" & _
            "gUkkfWvNbsJguaKEsVwGysFKGfOok3V/dkS9ruuNzv8NaF9MdegN5EO/AtXdE35i0zB4hz/kLgGjvKnQ7sJC4RzmY6bSqDqbRxQD" & _
            "XjA8V+uqmKDbRr/81z4AUEsDBBQAAgAIAAAIIQB78wKjvAAAACgBAAAeAAAAY3VzdG9tWG1sL19yZWxzL2l0ZW0zLnhtbC5yZWxz" & _
            "jc/BisIwEAbg+4LvEOZuUxUWWZp6WQRvIl3wGtJpG7bJhMwo+vaGPa3gwePM8H8/0+xuYVZXzOwpGlhVNSiMjnofRwM/3X65BcVi" & _
            "Y29nimjgjgy7dvHRnHC2UkI8+cSqKJENTCLpS2t2EwbLFSWM5TJQDlbKmEedrPu1I+p1XX/q/N+A9slUh95APvQrUN094Ts2DYN3" & _
            "+E3uEjDKiwrtLiwUzmE+ZiqNqrN5RDHgBcPfalMVE3Tb6Kf/2gdQSwMEFAACAAgAAAghAAzEGpK8AAAAKAEAAB4AAABjdXN0b21Y" & _
            "bWwvX3JlbHMvaXRlbTQueG1sLnJlbHONz8GKwjAQBuD7gu8Q5m5TRRZZmnpZBG8iXfAa0mkbtsmEzCj69oY9reDB48zwfz/T7G5h" & _
            "VlfM7CkaWFU1KIyOeh9HAz/dfrkFxWJjb2eKaOCODLt28dGccLZSQjz5xKookQ1MIulLa3YTBssVJYzlMlAOVsqYR52s+7Uj6nVd" & _
            "f+r834D2yVSH3kA+9CtQ3T3hOzYNg3f4Te4SMMqLCu0uLBTOYT5mKo2qs3lEMeAFw99qUxUTdNvop//aB1BLAwQUAAIACAAACCEA" & _
            "K6E/E7wAAAAoAQAAHgAAAGN1c3RvbVhtbC9fcmVscy9pdGVtNS54bWwucmVsc43PwYrCMBAG4PuC7xDmblMFF1maelkEbyJd8BrS" & _
            "aRu2yYTMKPr2hj2t4MHjzPB/P9PsbmFWV8zsKRpYVTUojI56H0cDP91+uQXFYmNvZ4po4I4Mu3bx0ZxwtlJCPPnEqiiRDUwi6Utr"

Base64_8 = "dhMGyxUljOUyUA5WyphHnaz7tSPqdV1/6vzfgPbJVIfeQD70K1DdPeE7Ng2Dd/hN7hIwyosK7S4sFM5hPmYqjaqzeUQx4AXD32pT" & _
            "FRN02+in/9oHUEsDBBQAAgAIAAAIIQADCCFLvAAAACgBAAAeAAAAY3VzdG9tWG1sL19yZWxzL2l0ZW02LnhtbC5yZWxzjc/BisIw" & _
            "EAbg+4LvEOZuUz2ILE29LII3kS54Dem0DdtkQmYUfXuDpxU8eJwZ/u9nmt0tzOqKmT1FA6uqBoXRUe/jaOC32y+3oFhs7O1MEQ3c" & _
            "kWHXLr6aE85WSognn1gVJbKBSSR9a81uwmC5ooSxXAbKwUoZ86iTdX92RL2u643O/w1oX0x16A3kQ78C1d0TfmLTMHiHP+QuAaO8" & _
            "qdDuwkLhHOZjptKoOptHFANeMDxXm6qYoNtGv/zXPgBQSwMEFAACAAgAAAghACRtBMq8AAAAKAEAAB4AAABjdXN0b21YbWwvX3Jl" & _
            "bHMvaXRlbTcueG1sLnJlbHONz8GKwjAQBuD7gu8Q5m5TPbiyNPWyCN5EuuA1pNM2bJMJmVH07Q17WsGDx5nh/36m2d3CrK6Y2VM0" & _
            "sKpqUBgd9T6OBn66/XILisXG3s4U0cAdGXbt4qM54WylhHjyiVVRIhuYRNKX1uwmDJYrShjLZaAcrJQxjzpZ92tH1Ou63uj834D2" & _
            "yVSH3kA+9CtQ3T3hOzYNg3f4Te4SMMqLCu0uLBTOYT5mKo2qs3lEMeAFw9/qsyom6LbRT/+1D1BLAwQUAAIACAAACCEA7WYRKbwA" & _
            "AAAoAQAAHgAAAGN1c3RvbVhtbC9fcmVscy9pdGVtOC54bWwucmVsc43PwYrCMBAG4LvgO4S521QPItLUyyJ4E+mC15BO27BNJmRG" & _
            "0bc37ElhD3ucGf7vZ5rDI8zqjpk9RQPrqgaF0VHv42jguzuudqBYbOztTBENPJHh0C4XzQVnKyXEk0+sihLZwCSS9lqzmzBYrihh" & _
            "LJeBcrBSxjzqZN2PHVFv6nqr87sB7YepTr2BfOrXoLpnwv/YNAze4Re5W8Aof1Rod2OhcA3zOVNpVJ3NI4oBLxh+V7uqmKDbRn/8" & _
            "174AUEsDBBQAAgAIAAAIIQDKAzSovAAAACgBAAAeAAAAY3VzdG9tWG1sL19yZWxzL2l0ZW05LnhtbC5yZWxzjc/BisIwEAbg+4Lv" & _
            "EOZuUz3IujT1sgjeRLrgNaTTNmyTCZlR9O0Ne1rBg8eZ4f9+ptndwqyumNlTNLCqalAYHfU+jgZ+uv3yExSLjb2dKaKBOzLs2sVH" & _
            "c8LZSgnx5BOrokQ2MImkL63ZTRgsV5QwlstAOVgpYx51su7XjqjXdb3R+b8B7ZOpDr2BfOhXoLp7wndsGgbv8JvcJWCUFxXaXVgo" & _
            "nMN8zFQaVWfziGLAC4a/1bYqJui20U//tQ9QSwMECgAAAAAAWIl7SAAAAAAAAAAAAAAAAAoAAABjdXN0b21YbWwvUEsDBBQAAgAI" & _
            "AAAIIQAhC9B9fAEAABADAAAQAAAAZG9jUHJvcHMvYXBwLnhtbJ2ST2/bMAzF7wP2HQzdGzndHwyBrGJIN/SwYQGStmdOpmOhsiSI" & _
            "rJHs0092ENdZd+rtkXx4/pmUujl0rugxkQ2+EstFKQr0JtTW7ytxv/t+9UUUxOBrcMFjJY5I4ka/f6c2KURMbJGKHOGpEi1zXElJ" & _
            "psUOaJHHPk+akDrgXKa9DE1jDd4G89yhZ3ldlp8lHhh9jfVVnALFKXHV81tD62AGPnrYHWPO0+prjM4a4PyX+qc1KVBouPh2MOiU" & _
            "nA9VDtqieU6Wj7pUcl6qrQGH6xysG3CESr401B3CsLQN2ERa9bzq0XBIBdk/eW3XovgNhANOJXpIFjyLk+1UjNpF4qQfQ3qiFpFJ" & _
            "yak5yrl3ru1HvRwNWVwa5QSS9SXizrJD+tVsIPF/iJdz4pFBzBi3A98rvPOH/olehy6Cz/uTk/ph/RPdx124BcbzNi+battCwjof" & _
            "YNr21FB3GSu5wb9uwe+xPnteD4bbP5weuF5+WpQfynI8+bmn5MtT1n8BUEsDBBQAAgAIAAAIIQCCHkg3EgEAABECAAARAAAAZG9j" & _
            "UHJvcHMvY29yZS54bWyVkT1PwzAQhnck/kPkPbGToFKsJB1AnUBCIoiKzbKvqUX8IduQ9t/jhjatoAvj6X3uubOvWmxVn3yB89Lo" & _
            "GuUZQQloboTUXY1e22U6R4kPTAvWGw012oFHi+b6quKWcuPg2RkLLkjwSTRpT7mt0SYESzH2fAOK+SwSOoZr4xQLsXQdtox/sA5w" & _
            "QcgMKwhMsMDwXpjayYgOSsEnpf10/SgQHEMPCnTwOM9yfGIDOOUvNozJGalk2Fm4iB7Did56OYHDMGRDOaJx/xyvnh5fxqemUu//"

Base64_9 = "igNqqsMilDtgAUQSBfRn3DF5K+8f2iVqCpLPUlKmxbwlhJJbSsh7hX/1n4QqHmct/2e8uTszHgVNhf8csfkGUEsDBAoAAAAAAFiJ" & _
            "e0gAAAAAAAAAAAAAAAAJAAAAZG9jUHJvcHMvUEsDBBQAAgAIAAAIIQC5hj0iqgEAAIwDAAASAAAAeGwvY29ubmVjdGlvbnMueG1s" & _
            "zVLLbtswELwX6D8QvMsiJVt+wFLg2s0pBYrATc4UubIJ8yGQjOug6L+XkpxGqIueeyGxD87OzHJ9d9EKncF5aU2J6YRgBIZbIc2h" & _
            "xN/298kCIx+YEUxZAyV+BY/vqo8f1twaAzzEZx5FDONLfAyhXaWp50fQzE9sCyZWGus0CzF0h9S3DpjwR4CgVZoRUqSaSYOrERyS" & _
            "IhLByDAd5+3hEhD3Z4zCaxtjSiJDB42DiCKe3oiTCBE7H3zob/TiZIl/7D5nM0LzIlksP+2SaZ5vkwXNimRHltvNbJ6Tzf3yJx7o" & _
            "ry50diNBS+6st02YcKtT2zSSw40ISlJK33RElNUfWgQrYAHTLGGETJNpM4ekJjSGNS2gLvKcF3OcVus0Eh/OXkb6DnNjT4bRCaDd" & _
            "KHmGsVlH6Z+tO9XWnnYssC9WgMJIgOdOtqF3qsuja2GwdPY3Q2NSS/M45FmtYFypGT8dnH0x/aaqtai/OvROcTwEbX+ncWzRmnWP" & _
            "rvOv8X7YbOeBVayNWB6MeLCcqUGds993Tiq1jSND/wdI1/y/Ljxa1+m7SvrXWkeBr34BUEsDBAoAAAAAAFuJe0gAAAAAAAAAAAAA" & _
            "AAAJAAAAeGwvbW9kZWwvUEsDBBQAAgAIAAAIIQBwCbZPYwIAAJoFAAANAAAAeGwvc3R5bGVzLnhtbKWUXW+bMBSG7yftP1i+pwYa" & _
            "siQCqqYpUqVumtRM2q0DJrHqD2SbLtm0/75jSAJRp21ar3y+/JzXPob0Zi8FemHGcq0yHF2FGDFV6oqrbYa/rItghpF1VFVUaMUy" & _
            "fGAW3+Tv36XWHQR72jHmECCUzfDOuWZBiC13TFJ7pRumIFNrI6kD12yJbQyjlfWbpCBxGE6JpFzhnrCQ5b9AJDXPbROUWjbU8Q0X" & _
            "3B06FkayXDxslTZ0I0DqPprQ8sTunFd4yUujra7dFeCIrmtestcq52ROgJSntVbOolK3ysFdAdpDF89Kf1OFT/lgX5Wn9jt6oQIi" & _
            "ISZ5WmqhDXLQlfkiiCgqWV9xazgVPlRTycWhD8Y+QHpUt1jIcyHO7WPcB/IUrsExowpw0NFeHxroo2BiPaar+0v11tBDFCejDd0C" & _
            "fTfaVPBChoOfQnkqWO1gg+HbnV+dbohPOgfXmacVp1utqPDI046jAdiSCfHkX9HX+oK9r5FqZSHdQ5VheI/+9CcTBB3NHtM7nj+m" & _
            "9ew3Y9G+vuSf0V2jC/o5ivxgM/zJv1gxINCm5cJx9RvBwKz2g9Yu6/wTvuwCjIrVtBVufU5meLA/soq3Mj5XfeYv2h2rBvvRTyqa" & _
            "+h5s7x6t61bUGp7hH/fLD/PVfREHs3A5CybXLAnmyXIVJJO75WpVzMM4vPs5+qLe8D113z0MJZosrIAqczzsUfzTEMvwyOnld/cH" & _
            "ssfa5/E0vE2iMCiuwyiYTOksmE2vk6BIong1nSzvkyIZaU/+T3sUkigaxCcLxyUTXLFL+etxFIYE7h8OQU6TIMNPNf8FUEsDBAoA" & _
            "AAAAAFiJe0gAAAAAAAAAAAAAAAAJAAAAeGwvdGhlbWUvUEsDBBQAAgAIAAAIIQCLgm5Y6AUAAI4aAAATAAAAeGwvdGhlbWUvdGhl" & _
            "bWUxLnhtbO1ZT48bNRS/I/EdrLmn838mWTVbJZOkhe62VXdb1KMzcTLuesbR2NndqKqE2iMSEqIgLkjcOCCgUitxKZ9moQiK1K+A" & _
            "x5M/nsRhW0ilIppIyfj5955/fs9+zzNz8dJpSsAxyhmmWdOwL1gGQFlMBzgbNY1bh71a3QCMw2wACc1Q05giZlzaff+9i3CHJyhF" & _
            "QOhnbAc2jYTz8Y5psliIIbtAxygTfUOap5CLZj4yBzk8EXZTYjqWFZgpxJkBMpgKs9eHQxwjcFiYNHbnxrtE/GScFYKY5AexHFHV" & _
            "kNjBkV38sSmLSA6OIWkaYpwBPTlEp9wABDIuOpqGJT+GuXvRXCgRvkFX0evJz0xvpjA4cqRePuovFD3P94LWwr5T2l/HdcNu0A0W" & _
            "9iQAxrGYqb2G9duNdsefYRVQeamx3Qk7rl3BK/bdNXzLL74VvLvEe2v4Xi9a+lABlZe+xiehE3kVvL/EB2v40Gp1vLCCl6CE4Oxo" & _
            "DW35gRvNZ7uADCm5ooU3fK8XOjP4EmUqq6vUz/imtZbCuzTvCYAMLuQ4A3w6RkMYC1wECe7nGOzhUSIW3hhmlAmx5Vg9yxW/xdeT"

Base64_10 = "V9IjcAdBRbsUxWxNVPABLM7xmDeND4VVQ4G8fPb9y2dPwMtnj88ePD178NPZw4dnD37UKF6B2UhVfPHtZ39+/TH448k3Lx59occz" & _
            "Ff/rD5/88vPneiBXgc+/fPzb08fPv/r09+8eaeCtHPZV+CFOEQPX0Am4SVMxN80AqJ+/nsZhAnFFAyYCqQF2eVIBXptCosO1UdV5" & _
            "t3ORJHTAy5O7Fa4HST7hWAO8mqQV4D6lpE1z7XSuFmOp05lkI/3g+UTF3YTwWDd2tBLa7mQsVjvWmYwSVKF5g4howxHKEAdFHz1C" & _
            "SKN2B+OKX/dxnFNGhxzcwaANsdYlh7jP9UpXcCriMoX6UFd8s38btCnRme+g4ypSbAhIdCYRqbjxMpxwmGoZw5SoyD3IEx3Jg2ke" & _
            "VxzOuIj0CBEKugPEmE7nej6t0L0qkos+7PtkmlaROcdHOuQepFRFduhRlMB0rOWMs0TFfsCOxBKF4AblWhK0ukOKtogDzDaG+zZG" & _
            "/PW29S2RV/ULpOiZ5LotgWh1P07JEKJsVgMq2TzF2bmpfSWp+++Suj6pt3Ks3VqrqXwT7j+YwDtwkt1AYs+8y9/v8vf/MX9v2svb" & _
            "z9rLRG2qp3VpJt14dB9iQg74lKA9JlM8E9Mb9IRQNqTS4k5hnIjL2XAV3CiH8hrklH+EeXKQwLEYxpYjjNjM9IiBMWWiSBgbbcsi" & _
            "M0n36aCU2vb85lQoQL6UiyIzl4uSxEtpEC7vwhbmZWvEVAK+NPrqJJTBqiRcDYnQfTUStrUtFg0Ni7r9dyxMJSpi/wFYPNfwvZKR" & _
            "WG+QoEERp1J/Ht2tR3qTM6vTdjTTa3hbi3SFhLLcqiSUZZjAAVoVbznWjYY+1I6WRlh/E7E213MDyaotcCL2nOsLMzEcN42hOB6K" & _
            "y3Qs7LEib0IyyppGzGeO/ieZZZwz3oEsKWGyq5x/ijnKAcGpWOtqGEi25GY7ofX2kmtYb5/nzNUgo+EQxXyDZNkUfaURbe+/BBcN" & _
            "OhGkD5LBCeiTSX4TCkf5oV04cIAZX3hzgHNlcS+9uJKuZlux8tBsuUUhGSdwVlHUZF7C5fWCjjIPyXR1VqbOhf1RbxtV93yllaS5" & _
            "oYCEG7PYmyvyCitXz8rX5rpG/Zwq8e8LgkKtrqfm6qltqh1bPBAowwXuOTVi29VgddWayrlSttbeTtD+XbHyO+K4OiGclY8BTsU9" & _
            "QjR/rlxmAimdZ5dTDiY5bhr3LL/lRY4f1ay63615rmfV6n7LrbV837W7vm112s594RSepLZfjt0T9zNkOnv5IuVrL2DS+TH7QkxT" & _
            "k8pzsCmV5QsY29n8AgZg4Zl7gdNruI12UGu4rV7N67TrtUYUtGudIAo7vU7k1xu9+wY4lmCv5UZe0K3XAjuKal5gFfTrjVroOU7L" & _
            "C1v1rte6P/O1mPn8f+5eyWv3L1BLAwQUAAIACAAACCEArdgGF3oCAACWBQAADwAAAHhsL3dvcmtib29rLnhtbLVUW2/TMBR+R+I/" & _
            "RH7PcmmStlHTqetFTAI0jbK9IKFT56Sx6tjBdmknxH/HTugo28skxkt87Jx8/i52JpfHhnvfUWkmRUGii5B4KKgsmdgW5PN65Y+I" & _
            "pw2IErgUWJAH1ORy+vbN5CDVbiPlzrMAQhekNqbNg0DTGhvQF7JFYd9UUjVg7FRtA90qhFLXiKbhQRyGWdAAE6RHyNVLMGRVMYoL" & _
            "SfcNCtODKORgLH1ds1af0Br6ErgG1G7f+lQ2rYXYMM7MQwdKvIbm11shFWy4lX2M0hOyLZ9BN4wqqWVlLizUb5LP9EZhEEW95Omk" & _
            "Yhzvets9aNuP0LhdOPE4aLMsmcGyIJYGlwf8a0Ht26s943YSJUkckmD6GMWN8kqsYM/N2tI6wdvGdBDHseu0ombcoBJgcC6FsR6+" & _
            "kl8d9ryWVrh3i9/2TKHubZtO7BNoDht9A6b29ooXZJF/OfMT6L84CtRJCx737+unOqcTZ9Edw4P+45ibesd7Jkp5KIg9+w9n9aEr" & _
            "71lp6oLESRg+rr1Dtq1NQcbDURdAcAbdsTuNnuhy/eRqe4/ccO2SszHmzBbquow6gNNXFDi1Mbqha0zjNOo68Gjea9ON1kFWkB9R" & _
            "Es6G4Tjxw+Ug9ZPROPZHySD258kiXqbD5WJ5lf583UNrUfKzw0ZrUGatgO7s3+IWqyvQ6MQ5QZbnOdnVfBnPFunCX2Xp3E9Ws8yf" & _
            "hWnmp/NBNouG2XCejf4D2RIMfJAl8n7auHLtbrR+uuC5OKj+/hWTcpgM0oEfj8aZnwCGPkCFfkLLGDIapmVifwV9rrafeFQKgdR0" & _
            "N23tFLtVZ8GzHYMnlHqTglO0wcna6S9QSwMECgAAAAAAWIl7SAAAAAAAAAAAAAAAAA4AAAB4bC93b3Jrc2hlZXRzL1BLAwQUAAIA"

Base64_11 = "CAAACCEAYlPvM18BAACGAgAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbI2Sy27CMBBF95X6D5b3xEAfFERAlRBqF5WqvvaO" & _
            "M0ksbE9kDwX+vpNQqkps2M31eI7uHXu+3HsnviEmiyGXo2woBQSDpQ11Lj8/1oMHKRLpUGqHAXJ5gCSXi+ur+Q7jJjUAJJgQUi4b" & _
            "onamVDINeJ0ybCFwp8LoNbGMtUptBF32Q96p8XB4r7y2QR4Js3gJA6vKGlih2XoIdIREcJrYf2psm040by7BeR0323Zg0LeMKKyz" & _
            "dOihUngze64DRl04zr0f3WpzYvfiDO+tiZiwooxxv0bPM0/VVDFpMS8tJ+jWLiJUuXwcSbWY9xe/LOzSv1qQLt7BgSEo+Y2k6HZf" & _
            "IG665jMfDbtRdTa77oO+RlFCpbeO3nD3BLZuiCHjbHLHcbogs/KwgmR4nUzKxn82Vpo0162u4UXH2oYkHFT9pYkU8Qjqa8K2r5hY" & _
            "IBH6k2o4O8RO3UhRIdJJdH7/ftDiB1BLAwQKAAAAAABYiXtIAAAAAAAAAAAAAAAACQAAAHhsL19yZWxzL1BLAwQUAAIACAAACCEA" & _
            "7X4f02cBAAClCwAAGgAAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzvZY9a8MwEIb3Qv+D0V7Ldr5LnCyhkK2UFLoK+WKbWJKR" & _
            "Lknz76smtE4gPToYTeZO6NXDcwh5vvxUTXQA62qjc5bGCYtAS1PUuszZ++blacoih0IXojEacnYCx5aLx4f5GzQC/SZX1a2LfIp2" & _
            "OasQ22fOnaxACRebFrRf2RqrBPrSlrwVcidK4FmSjLm9zmCLm8xoXeTMrgt//ubUwn+yzXZbS1gZuVeg8c4RXO4dGvWhGh8qbAmY" & _
            "szjuurxGUIPYxzJ+nyYdhMaZkjjB7aSknn7tGK1BnusO6KpJgUxCe8nIMWWhcSYkTnA7KamnVzvo90KHci4vzZR0Mg7uhOTJkuA8" & _
            "I9JPnzhHY3euAsAO57fl+PlDygk/K1JNGhpnTOGM+qRpzRHsa30wuBIoOiRlCriwxMX3wl9qRsEnlZCjCn6r6Es1C65nSPEM+8Rx" & _
            "eGrg6uW+1NTxwW2QMtJhaJzZDw6/+blefAFQSwMECgAAAAAAWIl7SAAAAAAAAAAAAAAAAAMAAAB4bC9QSwMEFAACAAgAAAghAJky" & _
            "tI2gAQAAsAwAABMAAABbQ29udGVudF9UeXBlc10ueG1szZfLTgIxFIb3Jr7DpFvClMEbGgYWXpbqAhO3tT1AQ29pDwpvb6cIMUYh" & _
            "BGK6mclMe/7vm6Y56fSHC62Kd/BBWlOTquyQAgy3QppJTV5GD+0eKQIyI5iyBmqyhECGg9OT/mjpIBSx2oSaTBHdDaWBT0GzUFoH" & _
            "Jo6MrdcM46OfUMf4jE2AdjudS8qtQTDYxiaDDPp3MGZzhcX9Ir5emXhQgRS3q4kNqybMOSU5wzhO3434QWl/EcpYmeaEqXShFScQ" & _
            "+iuhGfkb8HedYMj2M7PjseQgLJ/rWFJqK0C1UkwkPMXF91JA8cw8PjId4+hC0Q/rZ2/WzsrtmrtpwXlgIkwBUKsy3UvNpFmvzBZ+" & _
            "mhxoulVHFtnk7/DAuKNgdT1cIcXsAMa9aYCnDXTkT/6WvMMh4FLBsfGr0G1kPg9o9atWVCLoZ29dOHzNN6FNHniUezt0M3A4y8Dh" & _
            "PAOHiwwcLjNwuMrAoZeBw3UGDlUnB4kcOmWVQ6uscuiVVQ7Nsvr3bhnrEjmeYjzsD18foZvqttuPGKMP/lpoTtkCxC9smv43Bp9Q" & _
            "SwECFAAUAAIACAAACCEAtVUwI+sAAABMAgAACwAkAAAAAAABAAAAAAAAAAAAX3JlbHMvLnJlbHMKACAAAAAAAAEAGAAAwPjv4ueo" & _
            "AbCV9UaGiNEBsJX1RoaI0QFQSwECFAAKAAAAAABYiXtIAAAAAAAAAAAAAAAABgAkAAAAAAAAABAAAAAUAQAAX3JlbHMvCgAgAAAA" & _
            "AAABABgAsJX1RoaI0QGwlfVGhojRAbCV9UaGiNEBUEsBAhQAFAACAAgAAAghADxZrbrHAAAAaAEAABMAJAAAAAAAAAAAAAAAOAEA" & _
            "AGN1c3RvbVhtbC9pdGVtMS54bWwKACAAAAAAAAEAGAAAwPjv4ueoAUDX+0aGiNEBQNf7RoaI0QFQSwECFAAUAAIACAAACCEAJ5x/" & _
            "tbcAAABOAQAAFAAkAAAAAAAAAAAAAAAwAgAAY3VzdG9tWG1sL2l0ZW0xMC54bWwKACAAAAAAAAEAGAAAwPjv4ueoAWCW/kaGiNEB" & _
            "YJb+RoaI0QFQSwECFAAUAAIACAAACCEACMzEL60AAAA8AQAAFAAkAAAAAAAAAAAAAAAZAwAAY3VzdG9tWG1sL2l0ZW0xMS54bWwK"

Base64_12 = "ACAAAAAAAAEAGAAAwPjv4ueoAUBI/kaGiNEBQEj+RoaI0QFQSwECFAAUAAIACAAACCEAOv7UybIAAABEAQAAFAAkAAAAAAAAAAAA" & _
            "AAD4AwAAY3VzdG9tWG1sL2l0ZW0xMi54bWwKACAAAAAAAAEAGAAAwPjv4ueoASD6/UaGiNEBIPr9RoaI0QFQSwECFAAUAAIACAAA" & _
            "CCEAhkNeIrcAAABeAQAAFAAkAAAAAAAAAAAAAADcBAAAY3VzdG9tWG1sL2l0ZW0xMy54bWwKACAAAAAAAAEAGAAAwPjv4ueoAZAL" & _
            "/0aGiNEBkAv/RoaI0QFQSwECFAAUAAIACAAACCEAppMiY70AAABkAQAAFAAkAAAAAAAAAAAAAADFBQAAY3VzdG9tWG1sL2l0ZW0x" & _
            "NC54bWwKACAAAAAAAAEAGAAAwPjv4ueoAQDK+EaGiNEBAMr4RoaI0QFQSwECFAAUAAIACAAACCEAUSYwa5sBAABKBAAAFAAkAAAA" & _
            "AAAAAAAAAAC0BgAAY3VzdG9tWG1sL2l0ZW0xNS54bWwKACAAAAAAAAEAGAAAwPjv4ueoAeB7+EaGiNEB4Hv4RoaI0QFQSwECFAAU" & _
            "AAIACAAACCEAwGNMl9sAAACKAQAAEwAkAAAAAAAAAAAAAACBCAAAY3VzdG9tWG1sL2l0ZW0yLnhtbAoAIAAAAAAAAQAYAADA+O/i" & _
            "56gB8IT9RoaI0QHwhP1GhojRAVBLAQIUABQAAgAIAAAIIQDlij+arwAAAEIBAAATACQAAAAAAAAAAAAAAI0JAABjdXN0b21YbWwv" & _
            "aXRlbTMueG1sCgAgAAAAAAABABgAAMD47+LnqAHQNv1GhojRAdA2/UaGiNEBUEsBAhQAFAACAAgAAAghAGoy1EepAAAAOAEAABMA" & _
            "JAAAAAAAAAAAAAAAbQoAAGN1c3RvbVhtbC9pdGVtNC54bWwKACAAAAAAAAEAGAAAwPjv4ueoAbDo/EaGiNEBsOj8RoaI0QFQSwEC" & _
            "FAAUAAIACAAACCEALp32ErkAAABOAQAAEwAkAAAAAAAAAAAAAABHCwAAY3VzdG9tWG1sL2l0ZW01LnhtbAoAIAAAAAAAAQAYAADA" & _
            "+O/i56gBcEz8RoaI0QFwTPxGhojRAVBLAQIUABQAAgAIAAAIIQDFxC9JOQIAAJYIAAATACQAAAAAAAEAAAAAADEMAABjdXN0b21Y" & _
            "bWwvaXRlbTYueG1sCgAgAAAAAAABABgAAMD47+LnqAGAc/xGhojRAYBz/EaGiNEBUEsBAhQAFAACAAgAAAghALfHs4CzAAAARAEA" & _
            "ABMAJAAAAAAAAAAAAAAAmw4AAGN1c3RvbVhtbC9pdGVtNy54bWwKACAAAAAAAAEAGAAAwPjv4ueoAfCE/UaGiNEB8IT9RoaI0QFQ" & _
            "SwECFAAUAAIACAAACCEA7RWtsVIGAAAMaAAAEwAkAAAAAAABAAAAAAB/DwAAY3VzdG9tWG1sL2l0ZW04LnhtbAoAIAAAAAAAAQAY" & _
            "AADA+O/i56gBENP9RoaI0QEQ0/1GhojRAVBLAQIUABQAAgAIAAAIIQAozjA39AEAAOYFAAATACQAAAAAAAAAAAAAAAIWAABjdXN0" & _
            "b21YbWwvaXRlbTkueG1sCgAgAAAAAAABABgAAMD47+LnqAFwvf5GhojRAXC9/kaGiNEBUEsBAhQAFAACAAgAAAghABwTtAbBAAAA" & _
            "6wAAABgAJAAAAAAAAQAAAAAAJxgAAGN1c3RvbVhtbC9pdGVtUHJvcHMxLnhtbAoAIAAAAAAAAQAYAADA+O/i56gBUBz3RoaI0QFQ" & _
            "HPdGhojRAVBLAQIUABQAAgAIAAAIIQDlgnyMwQAAAOsAAAAZACQAAAAAAAEAAAAAAB4ZAABjdXN0b21YbWwvaXRlbVByb3BzMTAu" & _
            "eG1sCgAgAAAAAAABABgAAMD47+LnqAFQb/5GhojRAVBv/kaGiNEBUEsBAhQAFAACAAgAAAghAMbuhCPBAAAA6wAAABkAJAAAAAAA" & _
            "AQAAAAAAFhoAAGN1c3RvbVhtbC9pdGVtUHJvcHMxMS54bWwKACAAAAAAAAEAGAAAwPjv4ueoATAh/kaGiNEBMCH+RoaI0QFQSwEC" & _
            "FAAUAAIACAAACCEAcyiAqMEAAADrAAAAGQAkAAAAAAABAAAAAAAOGwAAY3VzdG9tWG1sL2l0ZW1Qcm9wczEyLnhtbAoAIAAAAAAA" & _
            "AQAYAADA+O/i56gBYCX8RoaI0QFgJfxGhojRAVBLAQIUABQAAgAIAAAIIQBuaeE2wQAAAOsAAAAZACQAAAAAAAEAAAAAAAYcAABj" & _
            "dXN0b21YbWwvaXRlbVByb3BzMTMueG1sCgAgAAAAAAABABgAAMD47+LnqAFQ/vtGhojRAVD++0aGiNEBUEsBAhQAFAACAAgAAAgh"

Base64_13 = "ANoDWSjBAAAA6wAAABkAJAAAAAAAAQAAAAAA/hwAAGN1c3RvbVhtbC9pdGVtUHJvcHMxNC54bWwKACAAAAAAAAEAGAAAwPjv4ueo" & _
            "AQDK+EaGiNEBAMr4RoaI0QFQSwECFAAUAAIACAAACCEAyDpZYsIAAADrAAAAGQAkAAAAAAABAAAAAAD2HQAAY3VzdG9tWG1sL2l0" & _
            "ZW1Qcm9wczE1LnhtbAoAIAAAAAAAAQAYAADA+O/i56gB4Hv4RoaI0QHge/hGhojRAVBLAQIUABQAAgAIAAAIIQB7FkIBwQAAAOsA" & _
            "AAAYACQAAAAAAAEAAAAAAO8eAABjdXN0b21YbWwvaXRlbVByb3BzMi54bWwKACAAAAAAAAEAGAAAwPjv4ueoAeBd/UaGiNEB4F39" & _
            "RoaI0QFQSwECFAAUAAIACAAACCEA6mIhvcIAAADrAAAAGAAkAAAAAAABAAAAAADmHwAAY3VzdG9tWG1sL2l0ZW1Qcm9wczMueG1s" & _
            "CgAgAAAAAAABABgAAMD47+LnqAHAD/1GhojRAcAP/UaGiNEBUEsBAhQAFAACAAgAAAghAHpfu5LCAAAA6wAAABgAJAAAAAAAAQAA" & _
            "AAAA3iAAAGN1c3RvbVhtbC9pdGVtUHJvcHM0LnhtbAoAIAAAAAAAAQAYAADA+O/i56gBoMH8RoaI0QGgwfxGhojRAVBLAQIUABQA" & _
            "AgAIAAAIIQDJ0+yVwAAAAOsAAAAYACQAAAAAAAEAAAAAANYhAABjdXN0b21YbWwvaXRlbVByb3BzNS54bWwKACAAAAAAAAEAGAAA" & _
            "wPjv4ueoAXBM/EaGiNEBcEz8RoaI0QFQSwECFAAUAAIACAAACCEAAZmHScAAAADrAAAAGAAkAAAAAAABAAAAAADMIgAAY3VzdG9t" & _
            "WG1sL2l0ZW1Qcm9wczYueG1sCgAgAAAAAAABABgAAMD47+LnqAGQmvxGhojRAZCa/EaGiNEBUEsBAhQAFAACAAgAAAghAMV73u/C" & _
            "AAAA6wAAABgAJAAAAAAAAQAAAAAAwiMAAGN1c3RvbVhtbC9pdGVtUHJvcHM3LnhtbAoAIAAAAAAAAQAYAADA+O/i56gBAKz9RoaI" & _
            "0QEArP1GhojRAVBLAQIUABQAAgAIAAAIIQCrK+5EwgAAAOsAAAAYACQAAAAAAAEAAAAAALokAABjdXN0b21YbWwvaXRlbVByb3Bz" & _
            "OC54bWwKACAAAAAAAAEAGAAAwPjv4ueoAYDk/kaGiNEBgOT+RoaI0QFQSwECFAAUAAIACAAACCEAuFmruMIAAADrAAAAGAAkAAAA" & _
            "AAABAAAAAACyJQAAY3VzdG9tWG1sL2l0ZW1Qcm9wczkueG1sCgAgAAAAAAABABgAAMD47+LnqAFwvf5GhojRAXC9/kaGiNEBUEsB" & _
            "AhQACgAAAAAAWIl7SAAAAAAAAAAAAAAAABAAJAAAAAAAAAAQAAAAqiYAAGN1c3RvbVhtbC9fcmVscy8KACAAAAAAAAEAGAAwsPtG" & _
            "hojRATCw+0aGiNEBcGr3RoaI0QFQSwECFAAUAAIACAAACCEAdD85erwAAAAoAQAAHgAkAAAAAAABAAAAAADYJgAAY3VzdG9tWG1s" & _
            "L19yZWxzL2l0ZW0xLnhtbC5yZWxzCgAgAAAAAAABABgAAMD47+LnqAFwavdGhojRAXBq90aGiNEBUEsBAhQAFAACAAgAAAghAB/Q" & _
            "P6u9AAAAKQEAAB8AJAAAAAAAAQAAAAAA0CcAAGN1c3RvbVhtbC9fcmVscy9pdGVtMTAueG1sLnJlbHMKACAAAAAAAAEAGAAAwPjv" & _
            "4ueoAQA7+0aGiNEBADv7RoaI0QFQSwECFAAUAAIACAAACCEAOLUaKr0AAAApAQAAHwAkAAAAAAABAAAAAADKKAAAY3VzdG9tWG1s" & _
            "L19yZWxzL2l0ZW0xMS54bWwucmVscwoAIAAAAAAAAQAYAADA+O/i56gB8BP7RoaI0QHwE/tGhojRAVBLAQIUABQAAgAIAAAIIQAQ" & _
            "HARyvQAAACkBAAAfACQAAAAAAAEAAAAAAMQpAABjdXN0b21YbWwvX3JlbHMvaXRlbTEyLnhtbC5yZWxzCgAgAAAAAAABABgAAMD4" & _
            "7+LnqAHg7PpGhojRAeDs+kaGiNEBUEsBAhQAFAACAAgAAAghADd5IfO9AAAAKQEAAB8AJAAAAAAAAQAAAAAAvioAAGN1c3RvbVht" & _
            "bC9fcmVscy9pdGVtMTMueG1sLnJlbHMKACAAAAAAAAEAGAAAwPjv4ueoAdDF+kaGiNEB0MX6RoaI0QFQSwECFAAUAAIACAAACCEA" & _
            "QE45wr0AAAApAQAAHwAkAAAAAAABAAAAAAC4KwAAY3VzdG9tWG1sL19yZWxzL2l0ZW0xNC54bWwucmVscwoAIAAAAAAAAQAYAADA"

Base64_14 = "+O/i56gBwJ76RoaI0QHAnvpGhojRAVBLAQIUABQAAgAIAAAIIQBnKxxDvQAAACkBAAAfACQAAAAAAAEAAAAAALIsAABjdXN0b21Y" & _
            "bWwvX3JlbHMvaXRlbTE1LnhtbC5yZWxzCgAgAAAAAAABABgAAMD47+LnqAGwd/pGhojRAbB3+kaGiNEBUEsBAhQAFAACAAgAAAgh" & _
            "AFyWJyK8AAAAKAEAAB4AJAAAAAAAAQAAAAAArC0AAGN1c3RvbVhtbC9fcmVscy9pdGVtMi54bWwucmVscwoAIAAAAAAAAQAYAADA" & _
            "+O/i56gBgJH3RoaI0QGAkfdGhojRAVBLAQIUABQAAgAIAAAIIQB78wKjvAAAACgBAAAeACQAAAAAAAEAAAAAAKQuAABjdXN0b21Y" & _
            "bWwvX3JlbHMvaXRlbTMueG1sLnJlbHMKACAAAAAAAAEAGAAAwPjv4ueoAaDf90aGiNEBoN/3RoaI0QFQSwECFAAUAAIACAAACCEA" & _
            "DMQakrwAAAAoAQAAHgAkAAAAAAABAAAAAACcLwAAY3VzdG9tWG1sL19yZWxzL2l0ZW00LnhtbC5yZWxzCgAgAAAAAAABABgAAMD4" & _
            "7+LnqAEQ8fhGhojRARDx+EaGiNEBUEsBAhQAFAACAAgAAAghACuhPxO8AAAAKAEAAB4AJAAAAAAAAQAAAAAAlDAAAGN1c3RvbVht" & _
            "bC9fcmVscy9pdGVtNS54bWwucmVscwoAIAAAAAAAAQAYAADA+O/i56gBkCn6RoaI0QGQKfpGhojRAVBLAQIUABQAAgAIAAAIIQAD" & _
            "CCFLvAAAACgBAAAeACQAAAAAAAEAAAAAAIwxAABjdXN0b21YbWwvX3JlbHMvaXRlbTYueG1sLnJlbHMKACAAAAAAAAEAGAAAwPjv" & _
            "4ueoAaBQ+kaGiNEBoFD6RoaI0QFQSwECFAAUAAIACAAACCEAJG0EyrwAAAAoAQAAHgAkAAAAAAABAAAAAACEMgAAY3VzdG9tWG1s" & _
            "L19yZWxzL2l0ZW03LnhtbC5yZWxzCgAgAAAAAAABABgAAMD47+LnqAEwsPtGhojRATCw+0aGiNEBUEsBAhQAFAACAAgAAAghAO1m" & _
            "ESm8AAAAKAEAAB4AJAAAAAAAAQAAAAAAfDMAAGN1c3RvbVhtbC9fcmVscy9pdGVtOC54bWwucmVscwoAIAAAAAAAAQAYAADA+O/i" & _
            "56gBMLD7RoaI0QEwsPtGhojRAVBLAQIUABQAAgAIAAAIIQDKAzSovAAAACgBAAAeACQAAAAAAAEAAAAAAHQ0AABjdXN0b21YbWwv" & _
            "X3JlbHMvaXRlbTkueG1sLnJlbHMKACAAAAAAAAEAGAAAwPjv4ueoASCJ+0aGiNEBIIn7RoaI0QFQSwECFAAKAAAAAABYiXtIAAAA" & _
            "AAAAAAAAAAAACgAkAAAAAAAAABAAAABsNQAAY3VzdG9tWG1sLwoAIAAAAAAAAQAYAJAL/0aGiNEBkAv/RoaI0QFA9fZGhojRAVBL" & _
            "AQIUABQAAgAIAAAIIQAhC9B9fAEAABADAAAQACQAAAAAAAEAAAAAAJQ1AABkb2NQcm9wcy9hcHAueG1sCgAgAAAAAAABABgAAMD4" & _
            "7+LnqAHALfhGhojRAcAt+EaGiNEBUEsBAhQAFAACAAgAAAghAIIeSDcSAQAAEQIAABEAJAAAAAAAAQAAAAAAPjcAAGRvY1Byb3Bz" & _
            "L2NvcmUueG1sCgAgAAAAAAABABgAAMD47+LnqAHQVPhGhojRAdBU+EaGiNEBUEsBAhQACgAAAAAAWIl7SAAAAAAAAAAAAAAAAAkA" & _
            "JAAAAAAAAAAQAAAAfzgAAGRvY1Byb3BzLwoAIAAAAAAAAQAYANBU+EaGiNEB0FT4RoaI0QGwBvhGhojRAVBLAQIUABQAAgAIAAAI" & _
            "IQC5hj0iqgEAAIwDAAASACQAAAAAAAEAAAAAAKY4AAB4bC9jb25uZWN0aW9ucy54bWwKACAAAAAAAAEAGAAAwPjv4ueoAVAc90aG" & _
            "iNEBUBz3RoaI0QFQSwECFAAKAAAAAABbiXtIAAAAAAAAAAAAAAAACQAkAAAAAAAAABAAAACAOgAAeGwvbW9kZWwvCgAgAAAAAAAB" & _
            "ABgA8A3kSoaI0QHwDeRKhojRAaAy/0aGiNEBUEsBAhQAFAACAAgAAAghAHAJtk9jAgAAmgUAAA0AJAAAAAAAAQAAAAAApzoAAHhs" & _
            "L3N0eWxlcy54bWwKACAAAAAAAAEAGAAAwPjv4ueoAQBZ9kaGiNEBAFn2RoaI0QFQSwECFAAKAAAAAABYiXtIAAAAAAAAAAAAAAAA" & _
            "CQAkAAAAAAAAABAAAAA1PQAAeGwvdGhlbWUvCgAgAAAAAAABABgAIKf2RoaI0QEgp/ZGhojRARCA9kaGiNEBUEsBAhQAFAACAAgA"

Base64_15 = "AAghAIuCbljoBQAAjhoAABMAJAAAAAAAAQAAAAAAXD0AAHhsL3RoZW1lL3RoZW1lMS54bWwKACAAAAAAAAEAGAAAwPjv4ueoASCn" & _
            "9kaGiNEBIKf2RoaI0QFQSwECFAAUAAIACAAACCEArdgGF3oCAACWBQAADwAkAAAAAAABAAAAAAB1QwAAeGwvd29ya2Jvb2sueG1s" & _
            "CgAgAAAAAAABABgAAMD47+LnqAHwMfZGhojRAfAx9kaGiNEBUEsBAhQACgAAAAAAWIl7SAAAAAAAAAAAAAAAAA4AJAAAAAAAAAAQ" & _
            "AAAAHEYAAHhsL3dvcmtzaGVldHMvCgAgAAAAAAABABgAMM72RoaI0QEwzvZGhojRATDO9kaGiNEBUEsBAhQAFAACAAgAAAghAGJT" & _
            "7zNfAQAAhgIAABgAJAAAAAAAAQAAAAAASEYAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbAoAIAAAAAAAAQAYAADA+O/i56gBMM72" & _
            "RoaI0QEwzvZGhojRAVBLAQIUAAoAAAAAAFiJe0gAAAAAAAAAAAAAAAAJACQAAAAAAAAAEAAAAN1HAAB4bC9fcmVscy8KACAAAAAA" & _
            "AAEAGADgCvZGhojRAeAK9kaGiNEB0OP1RoaI0QFQSwECFAAUAAIACAAACCEA7X4f02cBAAClCwAAGgAkAAAAAAABAAAAAAAESAAA" & _
            "eGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHMKACAAAAAAAAEAGAAAwPjv4ueoAeAK9kaGiNEB4Ar2RoaI0QFQSwECFAAKAAAAAABY" & _
            "iXtIAAAAAAAAAAAAAAAAAwAkAAAAAAAAABAAAACjSQAAeGwvCgAgAAAAAAABABgAoDL/RoaI0QGgMv9GhojRAdDj9UaGiNEBUEsB" & _
            "AhQAFAACAAgAAAghAJkytI2gAQAAsAwAABMAJAAAAAAAAQAAAAAAxEkAAFtDb250ZW50X1R5cGVzXS54bWwKACAAAAAAAAEAGAAA" & _
            "wPjv4ueoAXD59EaGiNEBcPn0RoaI0QFQSwUGAAAAAEAAQADgGQAAlUsAAAAA"

Base64 = Base64_1 & Base64_2 & Base64_3 & Base64_4 & Base64_5 & Base64_6 & Base64_7 & Base64_8 & Base64_9 & Base64_10 & _
        Base64_11 & Base64_12 & Base64_13 & Base64_14 & Base64_15

ByteArray() = Base64Decode(Base64)

Excel2013zip = Path + "\Workbook2013.zip"

Open Excel2013zip For Binary Lock Read Write As #1

For Counter = 0 To UBound(ByteArray)
    Put #1, LOF(1) + 1, ByteArray(Counter)
Next

Close #1

End Sub
'remove all new converted files
Sub DeleteConvertedFiles()

Set fso = CreateObject("Scripting.FileSystemObject")
On Error GoTo ErrorFileOpen
    Call fso.DeleteFolder("C:\Users\AKavalar\AppData\Local\Temp\Jebiga *", True) 'if "Jebiga *" folders exist, delete them
    MsgBox "All files successfully deleted."
    End

ErrorFileOpen:
    If Err.Number = 70 Then
        MsgBox "To delete all files, make sure to close them first."
    ElseIf Err.Number = 76 Then
        MsgBox "No files found."
    End If

End Sub
