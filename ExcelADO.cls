VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public UseHeader As Boolean   'By default: don't use header
Public UseIMEX As Boolean     'By default: IMEX=0
Public WorkbookPath As String 'By default: empty (current workbook)
Private m_aliases As Dictionary
Private m_connection_string As String
Private m_sql As String
Private m_error_message As String
Private rs As ADODB.Recordset

Private Sub Class_Initialize()
    'Set m_aliases = New Dictionary
    UseHeader = True
End Sub

Sub RunQuery(Query As String, Optional Aliases As Variant)

    Dim i As Integer
    Dim anAlias As Variant, alias_value As Variant
    
    'Replace aliases with their values
    m_sql = Query
    If IsMissing(Aliases) Then
        If Not m_aliases Is Nothing Then
            For Each anAlias In m_aliases
                m_sql = Replace(m_sql, anAlias, m_aliases(anAlias))
            Next
        End If
    Else
        For i = 0 To UBound(Aliases)
            anAlias = Split(Aliases(i), "=")
            m_sql = Replace(m_sql, anAlias(0), anAlias(1))
        Next
    End If
    
    Call SetConnectionString
    Call InitializeRecordset
    
    On Error Resume Next
    rs.Open m_sql, m_connection_string, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Err <> 0 Then
        Set rs = Nothing
        m_error_message = Err.description
        Err.Clear
    End If
    
End Sub

Private Sub InitializeDictionary()
    If m_aliases Is Nothing Then
        Set m_aliases = New Dictionary
    End If
End Sub

Private Sub InitializeRecordset()
    If Not rs Is Nothing Then
        If rs.State <> 0 Then
            rs.Close
        End If
    Else
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient 'To make RecordCount property work
    End If
End Sub

Private Sub SetConnectionString()

    Dim sProvider As String
    Dim sDataSource As String
    Dim sIMEX As String
    Dim sHeader As String
    Dim sExcelVersion As String
    
    '1. Choose provider
    Select Case Val(Application.Version)
        Case Is < 12
            sProvider = "Provider=Microsoft.Jet.OLEDB.4.0;"
            sExcelVersion = "Excel 8.0;"
        Case 12, 14
            sProvider = "Provider=Microsoft.ACE.OLEDB.12.0;"
            sExcelVersion = "Excel 12.0;"
        Case 15
            sProvider = "Provider=Microsoft.ACE.OLEDB.15.0;"
            sExcelVersion = "Excel 12.0;"
    End Select
    
    '2. Data Source
    sDataSource = "Data Source=" & IIf(Len(WorkbookPath) = 0, ActiveWorkbook.FullName, WorkbookPath) & ";"
    
    '3. IMEX
    sIMEX = "IMEX=" & IIf(UseIMEX, "1", "0") & ";"
    
    '4. Header
    sHeader = "HDR=" & IIf(UseHeader, "Yes", "No") & ";"
    
    '5. Create connection string
    m_connection_string = sProvider & _
                          sDataSource & _
                          "Extended Properties=""" & _
                          sExcelVersion & _
                          sHeader & _
                          sIMEX & _
                          """;"
End Sub

'========================================
'          HELPER METHODS
'========================================
Function PrintHeaders(TopLeftCell As Range) As Range
    Dim i As Integer
    For i = 0 To rs.Fields.Count - 1
        TopLeftCell.Offset(, i).Value = rs(i).Name
    Next
    Set PrintHeaders = TopLeftCell.Resize(, rs.Fields.Count)
End Function

Sub ExportRecordset(TopLeftCell As Range)
    TopLeftCell.CopyFromRecordset rs
End Sub

Function AlterRange(DataRange As Range, Optional SkipRows As Integer, Optional UseNumOfColumns As Integer) As Range
    If UseNumOfColumns = 0 Then UseNumOfColumns = DataRange.Columns.Count
    With DataRange
        Set AlterRange = .Offset(SkipRows).Resize(.Rows.Count - SkipRows, UseNumOfColumns)
    End With
End Function

Function FormatRangeAddress(DataRange As Range, Optional SkipRows As Integer, Optional UseNumOfColumns As Integer) As String
    If UseNumOfColumns = 0 Then UseNumOfColumns = DataRange.Columns.Count
    With DataRange
        FormatRangeAddress = "[" & .Parent.Name & "$" & .Offset(SkipRows).Resize(.Rows.Count - SkipRows, UseNumOfColumns).Address(0, 0) & "]"
    End With
End Function


'========================================
'             PROPERTIES
'========================================
Property Get Aliases(Alias As Variant) As String
    If m_aliases Is Nothing Then
        Aliases = vbNullString
        Exit Property
    End If
    Aliases = IIf(m_aliases.Exists(Alias), m_aliases(Alias), vbNullString)
End Property
Property Let Aliases(Alias As Variant, Value As String)
    Call InitializeDictionary
    m_aliases(Alias) = Value
End Property

Property Get ConnectionString() As String
    'Must ALWAYS call this method because user can change
    'workbook path while working with class
    Call SetConnectionString
    ConnectionString = m_connection_string
End Property

Property Get Recordset() As ADODB.Recordset
Attribute Recordset.VB_UserMemId = 0
    Set Recordset = rs
End Property

Property Get SqlString() As String
    SqlString = m_sql
End Property

Property Get IsQuerySuccessful() As Boolean
    IsQuerySuccessful = Not rs Is Nothing
End Property

Property Get IsEmptyRecordset() As Boolean
    IsEmptyRecordset = rs.BOF = rs.EOF
End Property

Property Get ErrorMessage() As String
    ErrorMessage = m_error_message
End Property


'===========================================
Private Sub Class_Terminate()
    If rs Is Nothing Then Exit Sub
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
End Sub
