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

'declare public objects
Public wbQueries As Workbook
Public wbTarget As Workbook
Public ConnCommand As String
Public TableName As String
Public Flag As Integer
'make filepath accessible from every subroutine
Public Property Get TargetPath() As String

TargetPath = ThisWorkbook.Worksheets("Extracting data from PowerPivot").Range("C6")
    
End Property
'make file name accessible from every subroutine
Public Property Get TargetFileName() As String

Dim wbTargetFileNameArray() As String

'find file name in file path
wbTargetFileNameArray = Split(TargetPath, "\")
On Error GoTo ErrorTargetFileName
    TargetFileName = wbTargetFileNameArray(UBound(wbTargetFileNameArray) - LBound(wbTargetFileNameArray))
    Exit Property

'error: file not found, wrong spelling, etc.
ErrorTargetFileName:
    MsgBox "Can't open target file."
    End 'stop everything
    
End Property
'make PowerPivot version accessible from every subroutine
Public Property Get Version() As Integer

Dim XMLpart As CustomXMLPart
Dim XMLnode As CustomXMLNode
Dim wbVersion As String
Dim Injected As Integer

'for details on version numbers, see
'https://blogs.msdn.microsoft.com/analysisservices/2012/10/01/more-on-multiple-powerpivot-versions/
'https://blogs.msdn.microsoft.com/analysisservices/2012/07/18/supporting-multiple-powerpivot-versions/
On Error Resume Next
    'first check if model injected into an empty workbook (see OpenInitialize() sub below)
    Set XMLpart = wbTarget.CustomXMLParts("JebigaCustomXMLPart")
    
    If Len(XMLpart.XML) = 0 Then 'default value, model has not been injected into another workbook
        Injected = 0
    ElseIf XMLpart.SelectSingleNode("//ns0:JebigaCustomElement").Text = "InjectedModel" Then 'model injected
        Injected = 1
    End If

    'now check for vintage of PowerPivot model
    Set XMLpart = Nothing 'reset
    
    '2013 PowerPivot model
    Set XMLpart = wbTarget.CustomXMLParts("http://gemini/pivotcustomization/PowerPivotVersion")
    
    '2012/2008 PowerPivot model
    If Len(XMLpart.XML) = 0 Then
        Set XMLpart = wbTarget.CustomXMLParts("http://gemini/workbookcustomization/PowerPivotVersion")
    End If
    
    '2008 RTM PowerPivot model (no explicit version number found)
    If Len(XMLpart.XML) = 0 Then
        Set XMLpart = wbTarget.CustomXMLParts("http://gemini/workbookcustomization/SandboxNonEmpty")
        Version = 2007 'PowerPivot 2008 RTM (see below for other versions)
    End If
    
    'PowerPivot model not found
    If Len(XMLpart.XML) = 0 Then
        MsgBox "No PowerPivot model detected in target file."
        End 'stop everything
    End If
    
    'child node of XML part contains the PowerPivot version
    Set XMLnode = XMLpart.SelectSingleNode("//ns0:CustomContent")
    wbVersion = XMLnode.Text
    
    If Split(wbVersion, ".")(0) = "2011" Then
        Version = 2013 'PowerPivot 2013
    ElseIf Split(wbVersion, ".")(0) = "11" And Injected = 1 Then
        Version = 2010 'PowerPivot 2013 model injected into Excel 2010 file
    ElseIf Split(wbVersion, ".")(0) = "11" And Injected = 0 Then
        Version = 2012 'standard PowerPivot 2012 model (in an Excel 2010 file)
    ElseIf Split(wbVersion, ".")(0) = "10" Then
        Version = 2008 'PowerPivot 2008
    End If
    
End Property
'public subroutine that opens the target file (and repairs it)
Public Sub OpenInitialize()

Set wbQueries = ThisWorkbook

On Error GoTo ErrorWorkbook:
    'open file and maximize the window
    Set wbTarget = Workbooks.Open(TargetPath)
    Application.WindowState = xlMaximized
    
    'check if 2013 model used with Excel 2010 -> offer to inject 2013 model into 2010 file
    If (Application.Version() = 14) And (Version = 2013) Then
        If MsgBox("Excel 2010 cannot work with PowerPivot 2013 models directly." & Chr(10) & Chr(10) & _
            "To access and query the data, the PowerPivot model can be converted, i.e. injected into an empty Excel 2010 file. No other data (worksheets, etc.) will be transferred. Proceed?", vbYesNo, "Migrate PowerPivot model to empty 2010 Excel file?") = vbYes Then
            
            'convert PowerPivot 2013 model to PowerPivot 2010 model
            Call ConvertVersions(TargetPath, 2010, wbQueries)
            
            'shift focus back to query workbook
            wbQueries.Activate
            
            'close original target file
            wbTarget.Close SaveChanges:=False
            
            'open converted file
            Set wbTarget = Workbooks.Open(TargetPath)
            Application.WindowState = xlMaximized
            
            'indicate in the converted file that the model has been injected (effect on queries), save changes
            wbTarget.CustomXMLParts.Add ("<JebigaCustomElement xmlns=""JebigaCustomXMLPart""><![CDATA[InjectedModel]]></JebigaCustomElement>")
            wbTarget.Save
        
        'no conversion
        Else
            End 'stop everything
        End If
    End If
    
    'shift focus back to query workbook
    wbQueries.Activate
    
    Exit Sub

'error: file not found, wrong spelling, etc. / file needs repair
ErrorWorkbook:
    'if file exists & error 1004: might need to repair file
    If (Err.Number = 1004) And (Dir(TargetPath) <> "") Then
        'open and repair file, no alerts
        Application.DisplayAlerts = False 'turn off alerts
        Set wbTarget = Workbooks.Open(TargetPath, CorruptLoad:=xlRepairFile)
        Application.WindowState = xlMaximized
        Application.DisplayAlerts = True 'turn alerts back on

        'determine if file contains a 2008 RTM PowerPivot model
        Call RTM2008model
        
        'go back to where the error occured
        Resume Next
    
    'if file not found, wrong spelling, etc.
    Else
        MsgBox "Can't open target file."
        End 'stop everything
    End If
    
End Sub
'sub that determines whether file contains a 2008 RTM PowerPivot model (file has to be resaved in Excel 2010, see message)
Sub RTM2008model()

On Error GoTo ErrorNoRTMModel
    'RTM 2008 model present, can't do anything in Excel 2013
    If Len(ActiveWorkbook.CustomXMLParts("http://gemini/workbookcustomization/SandboxNonEmpty").XML) > 0 Then
        MsgBox "This file contains a 2008 RTM PowerPivot model. These models CANNOT be queried using Excel 2013 (in fact, Excel just ""repaired"" the file and by doing so completely removed the PowerPivot model because of the malformed XML structure of these files)." _
            & Chr(10) & Chr(10) & "To query these models, open the file in Excel 2010 (PowerPivot add-in does not have to be installed) and immediately save it - the XML structure will be repaired in the process but the PowerPivot model will not be removed." _
            & Chr(10) & Chr(10) & "Resaved files containing 2008 RTM PowerPivot models can then be queried using either Excel 2013 or Excel 2010."
        End 'stop everything
    End If

'RTM 2008 model not present
ErrorNoRTMModel:
    Exit Sub 'go back to ErrorWorkbook error handler

End Sub
'subroutine that establishes the connection to the data model and queries it given the command (argument)
Sub Query(Command As String)

Dim ConnString As String
Dim ConnStringAll As String
Dim InitialCatalog As String
Dim ConnTable As QueryTable

Set wbQueries = ThisWorkbook
Set wbTarget = Application.Workbooks(TargetFileName)

'connection string for 2008 RTM models is different from other versions (basically, initial catalog property is unique for each model)
If Version = 2007 Then
    For i = 1 To wbTarget.Connections.Count
        ConnStringAll = wbTarget.Connections.Item(i).OLEDBConnection.Connection
        If InStr(1, UCase(ConnStringAll), "MSOLAP") And InStr(1, UCase(ConnStringAll), "EMBEDDED") Then 'assume there is only one MSOLAP connection
            InitialCatalog = Split(ConnStringAll, "Initial Catalog=")(1)
            InitialCatalog = Split(InitialCatalog, ";")(0)
            ConnString = "OLEDB;Provider=MSOLAP.5;Persist Security Info=True;Initial Catalog=" & InitialCatalog & ";Data Source=$Embedded$;MDX Compatibility=1;Safety Options=2;ConnectTo=11.0;MDX Missing Member Mode=Error;Optimize Response=3;Cell Error Mode=TextValue;SQLQueryMode=DataKeys"
        End If
    Next
Else
    ConnString = "OLEDB;Provider=MSOLAP.5;Persist Security Info=True;Initial Catalog=Microsoft_SQLServer_AnalysisServices;Data Source=$Embedded$;MDX Compatibility=1;Safety Options=2;ConnectTo=11.0;MDX Missing Member Mode=Error;Optimize Response=3;Cell Error Mode=TextValue;SQLQueryMode=DataKeys"
End If

'add a new worksheet to the target workbook, modify its name
wbTarget.Activate
wbTarget.Sheets.Add
wbTarget.ActiveSheet.Name = Replace(ActiveSheet.Name, "Sheet", "Query Results ")

'insert new query table, set its properties incl. unique name (same as worksheet name), don't run the query just yet
Set ConnTable = ActiveSheet.ListObjects.Add(SourceType:=0, Source:=ConnString, Destination:=Range("$A$1")).QueryTable
With ConnTable
        .CommandType = xlCmdDefault
        .CommandText = Command
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshStyle = xlInsertDeleteCells
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = Replace(wbTarget.ActiveSheet.Name, " ", "_")
End With

'run the query (i.e. refresh the query table)
On Error GoTo ErrorQuery
    ConnTable.Refresh BackgroundQuery:=False
    Flag = 1 'indicating query successful
    Exit Sub
    
'query problem (memory allocation problem, typo in query, etc.)
ErrorQuery:
    MsgBox "Query Error: manually refresh the query table to identify the problem (Alt+F5)."
    Flag = 0 'indicating query unsuccessful

End Sub
'subroutine that shows basic table information
Sub GetTables()

Dim i As Long

Set wbQueries = ThisWorkbook
On Error GoTo ErrorTarget
    Set wbTarget = Application.Workbooks(TargetFileName)

'file not open and model not initialized, run OpenInitialize sub
ErrorTarget:
    If Err.Number = 9 Then
        Call OpenInitialize
    End If

'check if 2013 model used with Excel 2010
If (Application.Version() = 14) And (Version = 2013) Then
    MsgBox "Excel 2010 cannot work with PowerPivot 2013 models (unless first converted)."
    End 'stop everything
End If

'query command (DMV)
ConnCommand = "select dimension_name, table_id, rows_count from $system.discover_storage_tables"

'run query
Call Query(ConnCommand)

'modify output if query successful (Flag=1), otherwise leave as is
If Flag = 0 Then
    End 'stop everything
Else
    If IsEmpty(Cells(2, 1)) = True Then
        'if model empty, don't do anything
        MsgBox "PowerPivot model contains no data."
    Else
        Application.ScreenUpdating = False
        
        'remove all references to columns (H$)/calculated columns (H$)/measures (R$) from the output (rows where table_id values contain a "$")
        i = 2
        Do While (Cells(i, 1) <> "")
            If InStr(1, Cells(i, 2), "$") > 0 Then
                Rows(i).Delete
            Else
                i = i + 1
            End If
        Loop
        
        'remove table_id column, not needed anymore
        Columns(2).Delete
        
        'sort the query table alphabetically based on table names
        TableName = Replace(wbTarget.ActiveSheet.Name, " ", "_")
        With wbTarget.ActiveSheet.ListObjects(TableName).Sort
            .SortFields.Clear
            .SortFields.Add Key:=Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        'rename columns
        Cells(1, 1) = "Tables"
        Cells(1, 2) = "Number of rows"
        
        'clean up
        Columns("A:B").Columns.AutoFit
        Columns("B:B").NumberFormat = "#,##0"
        Cells(1, 1).Select
        
        Application.ScreenUpdating = True
    End If
End If

End Sub
'subroutine that shows table and column information incl. correct column names (by running a 2nd query)
Sub GetTablesColumns()

Dim i As Long
Dim TableNameAux As String
Dim SheetTable As Worksheet
Dim SheetTableAux As Worksheet

Set wbQueries = ThisWorkbook
On Error GoTo ErrorTarget
    Set wbTarget = Application.Workbooks(TargetFileName)
    
'file not open and model not initialized, run OpenInitialize sub
ErrorTarget:
    If Err.Number = 9 Then
        Call OpenInitialize
    End If

'check if 2013 model used with Excel 2010
If (Application.Version() = 14) And (Version = 2013) Then
    MsgBox "Excel 2010 cannot work with PowerPivot 2013 models (unless first converted)."
    End 'stop everything
End If

'query command (DMV)
ConnCommand = "select table_id, dimension_name, rows_count, rows_count from $system.discover_storage_tables"

'run query
Call Query(ConnCommand)

'disable screen updating (2nd query table not observed)
Application.ScreenUpdating = False

'run 2nd query and modify output if main query successful (Flag=1), otherwise leave as is
If Flag = 0 Then
    End 'stop everything
Else
    If IsEmpty(Cells(2, 1)) = True Then
        'if model empty, don't do anything
        MsgBox "PowerPivot model contains no data."
        Application.ScreenUpdating = True
    Else
        'define main worksheet
        TableName = Replace(wbTarget.ActiveSheet.Name, " ", "_")
        Set SheetTable = wbTarget.ActiveSheet
        
        'query command 2 (DMV)
        ConnCommand = "select table_id, dimension_name, attribute_name from $SYSTEM.DISCOVER_STORAGE_TABLE_COLUMNS where column_id='POS_TO_ID'"
        
        'run query 2
        Call Query(ConnCommand)
    
        'define 2nd worksheet
        TableNameAux = Replace(wbTarget.ActiveSheet.Name, " ", "_")
        Set SheetTableAux = wbTarget.ActiveSheet
        SheetTable.Activate
        
        'join query tables and modify output if 2nd query successful (Flag=1), otherwise delete 2nd query table
        If Flag = 0 Then
            'delete 2nd worksheet
            Application.DisplayAlerts = False
            SheetTableAux.Delete
            Application.DisplayAlerts = True
            
            'message
            MsgBox "Main query successful, but the auxiliary query (to get correct column names) failed."
            
            'reset screen updating
            Application.ScreenUpdating = True
            End 'stop everything
        Else
            'define an aux column to use when sorting tables and columns (tables on top, associated columns below them)
            i = 2
            Do While (Cells(i, 1) <> "")
                If InStr(1, Cells(i, 1), "$") > 0 Then 'columns/calc columns/measures
                    Cells(i, 5) = 1
                ElseIf InStr(1, Cells(i, 1), "$") = 0 Then 'tables
                    Cells(i, 5) = 0
                End If
                i = i + 1
            Loop
    
            'sort: sort table names alphabetically, making sure the "table name" row is on top and all associated "column name" rows are below it (latter not sorted yet)
            With SheetTable.ListObjects(TableName).Sort
                .SortFields.Clear
                .SortFields.Add Key:=Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'here, table names in column B
                .SortFields.Add Key:=Range("E:E"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal '0's on top (tables), 1's below (columns)
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
    
            'join both query tables (joins not allowed when querying DMV tables, need to do two data pulls and them combine the results)
            i = 2
            Do While (Cells(i, 1) <> "")
                'pull correct column names from the 2nd query table using VLOOKUP on hashed names, put them in another aux column
                Cells(i, 6).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5]," & TableNameAux & ",3,FALSE),"""")"
                
                'for rows which correspond to table names BUT NOT measures (not matched with 2nd query table)
                If (Cells(i, 6) = "") And (InStr(1, Cells(i, 1), "$") = 0) Then 'not matched but table name
                    Cells(i, 1) = Cells(i, 2) 'move user-friendly table name to column A
                    Cells(i, 2).ClearContents 'clear column B (reserved for column names)
                    Cells(i, 4).ClearContents 'clear column D (reserved for approx. column unique values)
                
                'for rows which correspond to columns/calculated columns (matched with 2nd query table)
                ElseIf (Cells(i, 6) <> "") Then 'matched
                    Cells(i, 2) = Cells(i, 6) 'move matched, i.e. correct column name to column B
                    Cells(i, 1) = Cells(i - 1, 1) 'copy table name (in column A) from the row above (guaranteed to be table name)
                    Cells(i, 3).ClearContents 'clear column C (reserved for table row count)
                End If
                
                'for rows which correspond to measures (not matched with 2nd query table, but not table)
                If InStr(1, Cells(i, 1), "R$") = 1 Then
                    Rows(i).Delete
                Else
                    'increment row counter
                    i = i + 1
                End If
            Loop
                
            'sort again: first sort tables, then sort columns alphabetically
            With SheetTable.ListObjects(TableName).Sort
                .SortFields.Clear
                .SortFields.Add Key:=Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'table names
                .SortFields.Add Key:=Range("E:E"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal '0's on top (hierarchy)
                .SortFields.Add Key:=Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'column names
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            
            'remove table names from rows showing column names (retained them so far for easier sorting), highlight table names
            i = 2
            Do While (Cells(i, 1) <> "")
                If Cells(i, 2) <> "" Then Cells(i, 1).ClearContents
                If Cells(i, 2) = "" Then Range(Cells(i, 1), Cells(i, 4)).Interior.Color = 65535
                i = i + 1
            Loop
            
            'remove both aux columns
            Columns("E:F").Delete Shift:=xlToLeft
            
            'rename columns
            Cells(1, 1) = "Tables"
            Cells(1, 2) = "Columns"
            Cells(1, 3) = "Number of rows"
            Cells(1, 4) = "Number of unique values"
            
            'clean up
            Columns("A:D").Columns.AutoFit
            Columns("C:D").NumberFormat = "#,##0"
            Cells(1, 1).Select
            
            'delete 2nd worksheet
            Application.DisplayAlerts = False
            SheetTableAux.Delete
            Application.DisplayAlerts = True
            
            'reset screen updating
            Application.ScreenUpdating = True
            
            'unique values for some reason sometimes not correct (up to +3 greater than they should've been)
            MsgBox "Caution: unique values in each column are sometimes overstated by up to +3."
        End If
    End If
End If

End Sub
'sub that gets actual number of unique values in column (+ unique values themselves)
Sub GetUniqueColumnValues()

Dim TablePP As String
Dim ColumnPP As String
Dim i As Long

Set wbQueries = ThisWorkbook
On Error GoTo ErrorTarget
    Set wbTarget = Application.Workbooks(TargetFileName)
    
'file not open and model not initialized, run OpenInitialize sub
ErrorTarget:
    If Err.Number = 9 Then
        Call OpenInitialize
    End If

'define inputs (table and column)
TablePP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C17")
ColumnPP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C18")

'check if 2013 model used with Excel 2010
If (Application.Version() = 14) And (Version = 2013) Then
    MsgBox "Excel 2010 cannot work with PowerPivot 2013 models (unless first converted)."
    End 'stop everything
End If

'query command (DAX)
ConnCommand = "EVALUATE(ADDCOLUMNS(SUMMARIZE(" & TablePP & "," & TablePP & "[" & ColumnPP & "]),""Observations"", CALCULATE(COUNTROWS(" & TablePP & ")),""Unique Values"", DISTINCTCOUNT(" & TablePP & "[" & ColumnPP & "]),""Duplicates"",COUNTROWS(" & TablePP & ")-DISTINCTCOUNT(" & TablePP & "[" & ColumnPP & "])))"

'run query
Call Query(ConnCommand)

'modify output if query successful (Flag=1), otherwise leave as is
If Flag = 0 Then
    End 'stop everything
Else
    Application.ScreenUpdating = False
    
    'sort unique values alphabetically
    TableName = Replace(wbTarget.ActiveSheet.Name, " ", "_")
    With wbTarget.ActiveSheet.ListObjects(TableName).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'unique values
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'check if duplicates present (column 4), use binary flag
    If Cells(2, 4) > 0 Then
        Cells(2, 4) = "Yes"
    Else
        Cells(2, 4) = "No"
    End If
    Cells(2, 4).HorizontalAlignment = xlRight
    
    'retain only one row for duplicates (column 4) and count of unique values (column 3), delete the rest (starting from i=3)
    i = 3
    Do While (Cells(i, 2) <> "")
        Range(Cells(i, 3), Cells(i, 4)).ClearContents
        i = i + 1
    Loop
    
    'rename columns
    Cells(1, 2) = "Number of observations"
    Cells(1, 3) = "Number of unique values"
    Cells(1, 4) = "Duplicates?"
    
    'clean up
    Columns("A:D").Columns.AutoFit
    Columns("B:C").NumberFormat = "#,##0"
    Cells(1, 1).Select
    
    Application.ScreenUpdating = True
End If

End Sub
'sub that gets the entire table, no chunks, no numeric ranges, etc.
Sub GetFullTable()

Dim TablePP As String

Set wbQueries = ThisWorkbook
On Error GoTo ErrorTarget
    Set wbTarget = Application.Workbooks(TargetFileName)
    
'file not open and model not initialized, run OpenInitialize sub
ErrorTarget:
    If Err.Number = 9 Then
        Call OpenInitialize
    End If

'define input (table)
TablePP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C22")

'check if 2013 model used with Excel 2010
If (Application.Version() = 14) And (Version = 2013) Then
    MsgBox "Excel 2010 cannot work with PowerPivot 2013 models (unless first converted)."
    End 'stop everything
End If

'query command (DAX)
ConnCommand = "EVALUATE(" & TablePP & ")"

'run query
Call Query(ConnCommand)

'clean up
Cells(1, 1).Select

End Sub
'sub that gets a chunk of the table associated with a specific factor value
Sub GetSubsetTableFactor()

Dim TablePP As String
Dim ColumnPP As String
Dim ColumnValuePP As String

Set wbQueries = ThisWorkbook
On Error GoTo ErrorTarget
    Set wbTarget = Application.Workbooks(TargetFileName)
    
'file not open and model not initialized, run OpenInitialize sub
ErrorTarget:
    If Err.Number = 9 Then
        Call OpenInitialize
    End If

'define inputs (table, column, factor)
TablePP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C25")
ColumnPP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C26")
ColumnValuePP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C27")

'check if 2013 model used with Excel 2010
If (Application.Version() = 14) And (Version = 2013) Then
    MsgBox "Excel 2010 cannot work with PowerPivot 2013 models (unless first converted)."
    End 'stop everything
End If

'query command (DAX)
ConnCommand = "EVALUATE(CALCULATETABLE(" & TablePP & "," & TablePP & "[" & ColumnPP & "]=""" & ColumnValuePP & """))"

'run query
Call Query(ConnCommand)

'clean up
Cells(1, 1).Select

End Sub
'sub that cycels through chunks of the table associated with specific factor values (given the starting value)
Sub GetSubsetTableFactorSequence()

Dim TablePP As String
Dim ColumnPP As String
Dim ArraySheet As Worksheet
Dim ArrayValue As String
Dim StartValue As String
Dim i As Long

Set wbQueries = ThisWorkbook
On Error GoTo ErrorTarget
    Set wbTarget = Application.Workbooks(TargetFileName)
    
'file not open and model not initialized, run OpenInitialize sub
ErrorTarget:
    If Err.Number = 9 Then
        Call OpenInitialize
    End If

'define inputs (table, column, optional starting value for the factor)
TablePP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C29")
ColumnPP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C30")
StartValue = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C31")

'check if 2013 model used with Excel 2010
If (Application.Version() = 14) And (Version = 2013) Then
    MsgBox "Excel 2010 cannot work with PowerPivot 2013 models (unless first converted)."
    End 'stop everything
End If

'query command for the helper query (DAX)
ConnCommand = "EVALUATE(SUMMARIZE(" & TablePP & "," & TablePP & "[" & ColumnPP & "]))"

'run query (helper query)
Call Query(ConnCommand)

'if helper query successful (Flag=1), find the position of starting value and start the sequence of actual queries
If Flag = 0 Then
    End 'stop everything
Else
    Application.ScreenUpdating = False
    
    'sort unique values alphabetically
    TableName = Replace(wbTarget.ActiveSheet.Name, " ", "_")
    With wbTarget.ActiveSheet.ListObjects(TableName).Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'rename and clean up
    Cells(1, 1) = "Unique values of " & Cells(1, 1)
    Columns("A").Columns.AutoFit
    Cells(1, 1).Select
    
    Application.ScreenUpdating = True
    
    'find starting value (optional)
    'starting value provided
    If StartValue <> "" Then
        'find the position of the starting value
        On Error GoTo ErrorIndex:
            'if found, set i to starting value's position in the helper table
            i = Application.Match(StartValue, Columns("A:A"), 0)
    
    'starting value not provided, start at the top (i=2)
    Else
        i = 2
    End If
    
    'start cycling through factors
    Set ArraySheet = ActiveSheet
    Do While (ArraySheet.Cells(i, 1) <> "")
        'set current factor value
        ArrayValue = ArraySheet.Cells(i, 1)
        
        'query command (DAX)
        ConnCommand = "EVALUATE(CALCULATETABLE(" & TablePP & "," & TablePP & "[" & ColumnPP & "]=""" & ArrayValue & """))"
        
        'run query
        Call Query(ConnCommand)
        
        'error message, indicate last processed factor
        If Flag = 0 Then
            MsgBox "Data extraction failed. Last successfully processed factor was """ + ArraySheet.Cells(i - 1, 1) + """. Check " + Replace(TableName, "_", " ") + " sheet."
            End 'stop everything
        End If
        
        'clean up
        Cells(1, 1).Select
        
        'increment counter
        i = i + 1
    Loop
End If

'error handler: starting value provided, but its position not found (wrong spelling, etc.)
ErrorIndex:
    If Err.Number = 13 Then
        MsgBox "Error: Starting value not found."
        End 'stop everything
    End If

End Sub
'sub that fetches a numeric range of values
Sub GetSubsetTableNumeric()

Dim TablePP As String
Dim ColumnPP As String
Dim LValuePP As String
Dim UValuePP As String

Set wbQueries = ThisWorkbook
On Error GoTo ErrorTarget
    Set wbTarget = Application.Workbooks(TargetFileName)
    
'file not open and model not initialized, run OpenInitialize sub
ErrorTarget:
    If Err.Number = 9 Then
        Call OpenInitialize
    End If

'define inputs (table, column, lower bound, upper bound)
TablePP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C33")
ColumnPP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C34")
LValuePP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C35")
UValuePP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C36")

'check if 2013 model used with Excel 2010
If (Application.Version() = 14) And (Version = 2013) Then
    MsgBox "Excel 2010 cannot work with PowerPivot 2013 models (unless first converted)."
    End 'stop everything
End If

'decide if using DAX/DMV query or SQL-like query
If ColumnPP = "RowNumber" Then
    'query command (SQL-like)
    'if (Excel 2013 AND 2013 PowerPivot model) OR (Excel 2010 AND injected 2013 PowerPivot model)
    If ((Application.Version() > 14) And (Version = 2013)) Or ((Application.Version() = 14) And (Version = 2010)) Then
        ConnCommand = "SELECT * FROM [Model].[$" & TablePP & "] WHERE [Model].[$" & TablePP & "].[RowNumber]>" & LValuePP & " AND [Model].[$" & TablePP & "].[RowNumber]<=" & UValuePP
    
    'if Excel 2010 (not injected) OR Excel 2013 w/ 2012/2008 PowerPivot model
    Else
        ConnCommand = "SELECT * FROM [Sandbox].[$" & TablePP & "] WHERE [Sandbox].[$" & TablePP & "].[RowNumber]>" & LValuePP & " AND [Sandbox].[$" & TablePP & "].[RowNumber]<=" & UValuePP
    End If
    
    'similar internal SQL-like queries (replace Model with Sandbox for Excel 2010)
        'SELECT TOP 1000 * FROM [Model].[$dbo_DimDate]
        'SELECT COUNT(*) FROM [Model].[$dbo_DimDate]
        'SELECT DISTINCT(CalendarYearLabel) FROM [Model].[$dbo_DimDate]
    
    'run query
    Call Query(ConnCommand)
Else
    'query command (DAX)
    ConnCommand = "EVALUATE(CALCULATETABLE(" & TablePP & "," & TablePP & "[" & ColumnPP & "]>" & LValuePP & "," & TablePP & "[" & ColumnPP & "]<=" & UValuePP & "))"
    
    'run query
    Call Query(ConnCommand)
End If

'clean up
Cells(1, 1).Select

End Sub
'sub that fetches a sequence of chunks based on a numeric range of values
Sub GetSubsetTableNumericSequence()

Dim TablePP As String
Dim ColumnPP As String
Dim LValuePP As Long
Dim UValuePP As Long
Dim StepPP As Long
Dim NumberPP As Long
Dim i As Long
Dim LValueStep As Long
Dim UValueStep As Long

Set wbQueries = ThisWorkbook
On Error GoTo ErrorTarget
    Set wbTarget = Application.Workbooks(TargetFileName)
    
'file not open and model not initialized, run OpenInitialize sub
ErrorTarget:
    If Err.Number = 9 Then
        Call OpenInitialize
    End If

'define inputs (table, column, lower bound, upper bound, step size + number of steps)
TablePP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C38")
ColumnPP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C39")
LValuePP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C40")
UValuePP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C41")
StepPP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C42")
NumberPP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("C43")

'check if 2013 model used with Excel 2010
If (Application.Version() = 14) And (Version = 2013) Then
    MsgBox "Excel 2010 cannot work with PowerPivot 2013 models (unless first converted)."
    End 'stop everything
End If

'cycle through chunks
For i = 1 To NumberPP
    'set lower bound for step i
    LValueStep = LValuePP + (i - 1) * StepPP
    
    'set upper bound for step i
    UValueStep = LValuePP + i * StepPP
    
    'if upper bound of last step greater than user-provided upper bound, set upper bound for that step to the user-provided value
    If UValueStep > UValuePP Then UValueStep = UValuePP

    'decide if using DAX/DMV query or SQL-like query
    If ColumnPP = "RowNumber" Then
        'query command (SQL-like)
        'if (Excel 2013 AND 2013 PowerPivot model) OR (Excel 2010 AND injected 2013 PowerPivot model)
        If ((Application.Version() > 14) And (Version = 2013)) Or ((Application.Version() = 14) And (Version = 2010)) Then
            ConnCommand = "SELECT * FROM [Model].[$" & TablePP & "] WHERE [Model].[$" & TablePP & "].[RowNumber]>" & Str(LValueStep) & " AND [Model].[$" & TablePP & "].[RowNumber]<=" & Str(UValueStep)
        
        'if Excel 2010 (not injected) OR Excel 2013 w/ 2012/2008 PowerPivot model
        Else
            ConnCommand = "SELECT * FROM [Sandbox].[$" & TablePP & "] WHERE [Sandbox].[$" & TablePP & "].[RowNumber]>" & Str(LValueStep) & " AND [Sandbox].[$" & TablePP & "].[RowNumber]<=" & Str(UValueStep)
        End If
        
        'run query
        Call Query(ConnCommand)
    Else
        'query command (DAX)
        ConnCommand = "EVALUATE(CALCULATETABLE(" & TablePP & "," & TablePP & "[" & ColumnPP & "]>" & Str(LValueStep) & "," & TablePP & "[" & ColumnPP & "]<=" & Str(UValueStep) & "))"
        
        'run query
        Call Query(ConnCommand)
    End If
    
    'error message, indicate last processed chunk
    If Flag = 0 Then
        MsgBox "Data extraction failed. Data successfully extracted up to the chunk ending at" + Str(LValueStep) + " (inclusive)."
        End 'stop everything
    End If
    
    'clean up
    Cells(1, 1).Select
Next i

End Sub
'get random rows from a table (not used)
Sub GetRandomRows()

Dim TablePP As String
Dim RowsPP As Long

Set wbQueries = ThisWorkbook
On Error GoTo ErrorTarget
    Set wbTarget = Application.Workbooks(TargetFileName)
    
'file not open and model not initialized, run OpenInitialize sub
ErrorTarget:
    If Err.Number = 9 Then
        Call OpenInitialize
    End If

'define inputs (table, desired number of rows)
'TablePP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("fill_in")
'RowsPP = wbQueries.Worksheets("Extracting data from PowerPivot").Range("fill_in")
TablePP = "dbo_DimDate" 'remove if sub actually used
RowsPP = 1000 'remove if sub actually used

'check if 2013 model used with Excel 2010
If (Application.Version() = 14) And (Version = 2013) Then
    MsgBox "Excel 2010 cannot work with PowerPivot 2013 models (unless first converted)."
    End 'stop everything
End If

'query command (DAX)
ConnCommand = "EVALUATE(SAMPLE(" & Str(RowsPP) & "," & TablePP & ",RAND()))"

'run query
Call Query(ConnCommand)

'clean up
Cells(1, 1).Select

End Sub
'read Cartridges path from Windows registry
Sub ReadFromRegistry()

Dim LocationCartridge As String
Dim TargetKey As String
Dim WinScript As Object

Set WinScript = CreateObject("WScript.Shell") 'Windows Scripting object

'define input (specific registry key)
TargetKey = "HKCU\Software\Microsoft\Office\Excel\Addins\Microsoft.AnalysisServices.Modeler.FieldList\CartridgePath"

'if key exists
On Error GoTo ErrorCartridge
    'get key value
    LocationCartridge = WinScript.RegRead(TargetKey)
    MsgBox "Cartridges folder exists in Windows registry:" & Chr(10) & LocationCartridge
    End 'stop everything

'if key doesn't exist
ErrorCartridge:
    MsgBox "Required settings don't exist in Windows registry."

End Sub
'write to Windows registry
Sub WriteToRegistry()

Dim LocationCartridge As String
Dim LocationCartridgeParent As String
Dim TargetKey As String
Dim ArrayLocations() As String
Dim LenArray As Integer
Dim WinScript As Object
Dim wbQueries As Workbook
Dim i As Integer

Set wbQueries = ThisWorkbook
Set WinScript = CreateObject("WScript.Shell") 'Windows Scripting object

'alternative possible locations with IDENTICAL (confirmed) files
ArrayLocations = Split("root\office15\ADDINS\PowerPivot Excel Add-in\Cartridges\,root\vfs\ProgramFilesX86\Microsoft Analysis Services\AS OLEDB\110\Cartridges\,root\vfs\ProgramFilesCommonX86\Microsoft Shared\OFFICE15\DataModel\Cartridges\,root\vfs\ProgramFilesX64\Microsoft Analysis Services\AS OLEDB\110\Cartridges\,root\vfs\ProgramFilesCommonX64\Microsoft Shared\OFFICE15\DataModel\Cartridges\", ",")

'length of string array
LenArray = UBound(ArrayLocations) - LBound(ArrayLocations) + 1

'yes/no dialog box
If MsgBox("Are you sure you want to write to Windows registry?", vbYesNo, "Confirm") = vbYes Then
    'define inputs (optional pre-supplied Cartridges path and parent registry key)
    LocationCartridge = wbQueries.Worksheets("Registry settings (Excel 2013)").Range("C7")
    TargetKey = "HKCU\Software\Microsoft\Office\Excel\Addins\Microsoft.AnalysisServices.Modeler.FieldList\"
    
    'get path to the Cartridges folder if not supplied
    If LocationCartridge = "" Then
        'find location of Excel executable
        LocationCartridgeParent = WinScript.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe\")
        
        'split on root to get parent folder
        LocationCartridgeParent = Split(LCase(LocationCartridgeParent), "root")(0)
        
        'add the (first) 2nd part
        LocationCartridge = LocationCartridgeParent & ArrayLocations(0)
        
        'cycle through other 2nd parts if first location doesn't exist
        i = 1
        Do While (Len(Dir(LocationCartridge)) = 0) And (i <= (LenArray - 1))
            LocationCartridge = LocationCartridgeParent & ArrayLocations(i)
            i = i + 1
        Loop
        
        'if no location exists, let user know and exit
        If Len(Dir(LocationCartridge)) = 0 Then
            MsgBox "Required files not found. Use custom path for Cartridges folder."
            End 'stop everything
        Else
            MsgBox "Cartridges folder found:" & Chr(10) & LocationCartridge
        End If
        
    End If
    
    'write required settings to registry
    WinScript.RegWrite TargetKey & "CartridgePath", LocationCartridge, "REG_SZ"
    WinScript.RegWrite TargetKey & "Description", "Microsoft SQL Server PowerPivot for Microsoft Excel", "REG_SZ"
    WinScript.RegWrite TargetKey & "FriendlyName", "PowerPivot for Excel", "REG_SZ"
    WinScript.RegWrite TargetKey & "LoadBehavior", 3, "REG_DWORD"
    
    'confirmation message
    MsgBox "Required settings successfully written to Windows registry."
End If

End Sub
