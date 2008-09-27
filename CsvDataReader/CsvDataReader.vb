Imports Microsoft.VisualBasic.FileIO
Imports System.Data
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Collections.ObjectModel
Imports log4net

''' <summary>
''' Implements an IDataReader fo Csv data files.
''' </summary>
''' <remarks><seealso cref="IDataReader">IDataReader</seealso>, <seealso cref="IDataRecord">IDataRecord</seealso></remarks>
Public Class CsvDataReader
    Implements IDataReader

#Region "Privates"

    Private Const DEFAULT_FIELD_SEPARATOR As String = ","
    Private Const DEFAULT_FIELD_DELIMITER As String = """"
    Private Const DEFAULT_FIELD_TYPE As FieldType = FieldType.Delimited
    Private Const DEFAULT_SCHEMA_FILE As String = "scheme.ini"

    Private Const SCHEMA_CHARACTER_SET_OEM As String = "OEM"
    Private Const SCHEMA_CHARACTER_SET_ANSI As String = "ANSI"

    Private Const SCHEMA_FORMAT_TAB_DELIMITED As String = "TabDelimited"
    Private Const SCHEMA_FORMAT_CSV_DELIMITED As String = "CsvDelimited"
    Private Const SCHEMA_FORMAT_FIXED_LENGTH As String = "FixedLength"
    Private Const SCHEMA_FORMAT_DELIMITED As String = "Delimited\((.*)\)"

    REM .NET Type => schema.ini type
    Private Const SCHEMA_COLUMN_TYPE_STRING As String = "(Char|Text|LongChar|Memo|String)"
    Private Const SCHEMA_COLUMN_TYPE_DATETIME As String = "(Date|DateTime)"
    Private Const SCHEMA_COLUMN_TYPE_BOOLEAN As String = "(Bit|Boolean)"
    Private Const SCHEMA_COLUMN_TYPE_GUID As String = "(Guid|Uuid)"
    Private Const SCHEMA_COLUMN_TYPE_DOUBLE As String = "(Double|Float)"
    Private Const SCHEMA_COLUMN_TYPE_SINGLE As String = "Single"
    Private Const SCHEMA_COLUMN_TYPE_LONG As String = "Int64"
    Private Const SCHEMA_COLUMN_TYPE_INTEGER As String = "(Long|Int32)"
    Private Const SCHEMA_COLUMN_TYPE_SHORT As String = "(Short|Integer|Int16)"
    Private Const SCHEMA_COLUMN_TYPE_DECIMAL As String = "(Decimal|Currency)"
    Private Const SCHEMA_COLUMN_TYPE_BYTE As String = "Byte"

    Private Shared ReadOnly Log As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod.DeclaringType)

    Private _disposed As Boolean = False
    Private _dataTable As New DataTable
    Private _dataRow As DataRow = Nothing
    Private _encoding As Encoding = Encoding.UTF8
    Private _isClosed As Boolean = True
    Private _parser As TextFieldParser = Nothing
    Private _path As String = String.Empty
    Private _fieldSeparator As String = DEFAULT_FIELD_SEPARATOR
    Private _fieldType As FieldType = DEFAULT_FIELD_TYPE
    Private _schemaFile As String = DEFAULT_SCHEMA_FILE
    Private _schemaSection As String = String.Empty
    Private _stream As Stream = Nothing

    Private Declare Unicode Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionW" ( _
        ByVal lpApplicationName As String, _
        ByVal lpReturnedString() As Char, _
        ByVal nSize As Int32, _
        ByVal lpFileName As String) As Int32

#End Region

#Region "Constructors"

#Region "Stream Constructors"

    ''' <summary>
    ''' Creates a new CsvDataReader for the stream specified.
    ''' </summary>
    ''' <param name="stream">Stream. The stream of data to read.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal stream As Stream)
        Me.Stream = stream
    End Sub

    ''' <summary>
    ''' Creates a new CsvDataReader for the stream specified.
    ''' </summary>
    ''' <param name="stream">Stream. The stream of data to read.</param>
    ''' <param name="encoding">Encoding. The encoding of the specified file.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal stream As Stream, ByVal encoding As Encoding)
        Me.Stream = stream
        Me.Encoding = encoding
    End Sub

    ''' <summary>
    ''' Creates a new CsvDataReader for the file specified.
    ''' </summary>
    ''' <param name="stream">Stream. The stream of data to read.</param>
    ''' <param name="columns">Collection(Of DataColumn). The collection of column definitions for the specified files columns.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal stream As Stream, ByVal columns As Collection(Of CsvDataColumn))
        Me.Stream = stream

        For Each column As DataColumn In columns
            Me.DataTable.Columns.Add(column)
        Next
    End Sub

    ''' <summary>
    ''' Creates a new CsvDataReader for the file specified.
    ''' </summary>
    ''' <param name="stream">Stream. The stream of data to read.</param>
    ''' <param name="columns">Collection(Of DataColumn). The collection of column definitions for the specified files columns.</param>
    ''' <param name="encoding">Encoding. The encoding of the specified file.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal stream As Stream, ByVal columns As Collection(Of CsvDataColumn), ByVal encoding As Encoding)
        Me.Stream = stream
        Me.Encoding = encoding

        For Each column As DataColumn In columns
            Me.DataTable.Columns.Add(column)
        Next
    End Sub

#End Region

#Region "Path Constructors"

    ''' <summary>
    ''' Creates a new CsvDataReader for the file specified.
    ''' </summary>
    ''' <param name="path">String. The full path to the file to read.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal path As String)
        Me.Path = path
    End Sub

    ''' <summary>
    ''' Creates a new CsvDataReader for the file specified.
    ''' </summary>
    ''' <param name="path">String. The full path to the file to read.</param>
    ''' <param name="encoding">Encoding. The encoding of the specified file.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal path As String, ByVal encoding As Encoding)
        Me.Path = path
        Me.Encoding = encoding
    End Sub

    ''' <summary>
    ''' Creates a new CsvDataReader for the file specified.
    ''' </summary>
    ''' <param name="path">String. The full path to the file to read.</param>
    ''' <param name="columns">Collection(Of DataColumn). The collection of column definitions for the specified files columns.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal path As String, ByVal columns As Collection(Of CsvDataColumn))
        Me.Path = path

        For Each column As DataColumn In columns
            Me.DataTable.Columns.Add(column)
        Next
    End Sub

    ''' <summary>
    ''' Creates a new CsvDataReader for the file specified.
    ''' </summary>
    ''' <param name="path">String. The full path to the file to read.</param>
    ''' <param name="columns">Collection(Of DataColumn). The collection of column definitions for the specified files columns.</param>
    ''' <param name="encoding">Encoding. The encoding of the specified file.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal path As String, ByVal columns As Collection(Of CsvDataColumn), ByVal encoding As Encoding)
        Me.Path = path
        Me.Encoding = encoding

        For Each column As DataColumn In columns
            Me.DataTable.Columns.Add(column)
        Next
    End Sub

#End Region

#End Region

#Region "Properties"

    ''' <summary>
    ''' Creates a new TextFieldParser instance using the current settings.
    ''' </summary>
    ''' <value></value>
    ''' <returns>TextFieldParser</returns>
    ''' <remarks></remarks>
    Protected Overridable ReadOnly Property Parser() As TextFieldParser
        Get
            If _parser Is Nothing Then
                If Me.DataTable.Columns.Count = 0 And Me.FieldType = FileIO.FieldType.FixedWidth Then
                    Throw New MalformedLineException("Unable to parse fixed width columns with no column definitions")
                End If

                If Me.Stream IsNot Nothing Then
                    _parser = New TextFieldParser(Me.Stream, Me.Encoding, True)
                Else
                    _parser = New TextFieldParser(Me.Path, Me.Encoding, True)
                End If
                _parser.TextFieldType = Me.FieldType
                _parser.HasFieldsEnclosedInQuotes = True
                _parser.TrimWhiteSpace = True

                If Me.FieldType = FileIO.FieldType.FixedWidth Then
                    Dim columns As DataColumnCollection = Me.DataTable.Columns
                    Dim lengths(columns.Count - 1) As Integer

                    For i As Integer = 0 To columns.Count - 1
                        lengths(i) = DirectCast(columns.Item(i), CsvDataColumn).FieldWidth
                    Next

                    _parser.FieldWidths = lengths
                Else
                    _parser.SetDelimiters(Me.FieldSeparator)
                End If
            End If

            Return _parser
        End Get
    End Property

    ''' <summary>
    ''' Gets/sets the full path to the file to read.
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Overridable Property Path() As String
        Get
            Return _path
        End Get
        Set(ByVal value As String)
            If Not File.Exists(value) Then
                Throw New FileNotFoundException
            Else
                _path = value.Trim
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets/sets the encoding to use when reading the file.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Encoding</returns>
    ''' <remarks>The default is UTF8.</remarks>
    Public Overridable Property Encoding() As Encoding
        Get
            Return _encoding
        End Get
        Set(ByVal value As Encoding)
            _encoding = value
        End Set
    End Property

    ''' <summary>
    ''' Gets/sets the character(s) used to separate fields from one another.
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property FieldSeparator() As String
        Get
            Return _fieldSeparator
        End Get
        Set(ByVal value As String)
            If Not String.IsNullOrEmpty(value) Then
                Me.FieldType = FileIO.FieldType.Delimited
            End If

            _fieldSeparator = value
        End Set
    End Property

    ''' <summary>
    ''' Gets/sets the type if fields in the file (Delimited or FixedWidth)
    ''' </summary>
    ''' <value></value>
    ''' <returns>FieldType</returns>
    ''' <remarks></remarks>
    Public Property FieldType() As FieldType
        Get
            Return _fieldType
        End Get
        Set(ByVal value As FieldType)
            If value = FileIO.FieldType.FixedWidth Then
                Me.FieldSeparator = String.Empty
            End If

            _fieldType = value
        End Set
    End Property

    ''' <summary>
    ''' Gets the DataTable for the current file.
    ''' </summary>
    ''' <value></value>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Protected Overridable ReadOnly Property DataTable() As DataTable
        Get
            Return _dataTable
        End Get
    End Property

    ''' <summary>
    ''' Gets the current row from the DataTable for the current file.
    ''' </summary>
    ''' <value></value>
    ''' <returns>DataRow</returns>
    ''' <remarks></remarks>
    Protected Overridable ReadOnly Property CurrentDataRow() As DataRow
        Get
            Return _dataRow
        End Get
    End Property

    ''' <summary>
    ''' Gets a collection of data column definitions from the first row of the file.
    ''' </summary>
    ''' <returns>Collection(Of DataColumn)</returns>
    ''' <remarks></remarks>
    Protected Overridable Function GetColumnsFromFirstRow() As Collection(Of DataColumn)
        Dim fields() As String = Me.Parser.ReadFields

        Dim columns As New Collection(Of DataColumn)
        For Each field As String In fields
            Log.DebugFormat("Adding column {0}", field.Trim)

            Dim column As New DataColumn(field.Trim, GetType(String))
            column.AllowDBNull = True
            columns.Add(column)
        Next

        Return columns
    End Function

    Protected Overridable Function GetColumnsFromSchemaFile() As Collection(Of DataColumn)
        Dim columns As New Collection(Of DataColumn)

        Dim section As String = Me.SchemaSection
        If String.IsNullOrEmpty(section) Then
            Dim file As FileInfo = New FileInfo(Me.Path)

            section = file.Name
        End If

        Dim buffer(10240) As Char
        Dim count As Integer = GetPrivateProfileSection(section, buffer, buffer.Length, Me.SchemaFile)

        If count > 0 Then
            Dim result As New String(buffer)
            Dim separators() As Char = {vbNullChar}
            Dim fields() As String = result.Trim("").Split(separators)
            Dim settings As New Dictionary(Of String, String)
            Dim cols As New Collection(Of String)

            For Each f In fields
                Dim pair() As String = f.Split("=")

                If pair(0).Substring(0, 3).ToLower = "col" And pair(0).Trim.ToLower <> "colnameheader" Then
                    Log.DebugFormat("Column {0}", pair(1))
                    cols.Add(pair(1).Trim)
                Else
                    Log.DebugFormat("Setting {0}", f)
                    settings(pair(0).Trim.ToLower) = pair(1).Trim
                End If
            Next

            Dim format As String = settings("format")
            If Regex.IsMatch(format, SCHEMA_FORMAT_TAB_DELIMITED, RegexOptions.IgnoreCase) Then
                Me.FieldType = FileIO.FieldType.Delimited
                Me.FieldSeparator = vbTab
            ElseIf Regex.IsMatch(format, SCHEMA_FORMAT_CSV_DELIMITED, RegexOptions.IgnoreCase) Then
                Me.FieldType = FileIO.FieldType.Delimited
                Me.FieldSeparator = ","
            ElseIf Regex.IsMatch(format, SCHEMA_FORMAT_FIXED_LENGTH, RegexOptions.IgnoreCase) Then
                Me.FieldType = FileIO.FieldType.FixedWidth
            ElseIf Regex.IsMatch(format, SCHEMA_FORMAT_DELIMITED, RegexOptions.IgnoreCase) Then
                Me.FieldType = FileIO.FieldType.Delimited
                Me.FieldSeparator = Regex.Match(format, SCHEMA_FORMAT_DELIMITED, RegexOptions.IgnoreCase).Groups(1).Value
            End If

            Dim charset As String = settings("characterset").Trim
            If Regex.IsMatch(charset, SCHEMA_CHARACTER_SET_ANSI, RegexOptions.IgnoreCase) Then
                Me.Encoding = Encoding.GetEncoding(System.Globalization.CultureInfo.CurrentCulture.TextInfo.ANSICodePage)
            ElseIf Regex.IsMatch(charset, SCHEMA_CHARACTER_SET_OEM, RegexOptions.IgnoreCase) Then
                Me.Encoding = Encoding.GetEncoding(System.Globalization.CultureInfo.CurrentCulture.TextInfo.OEMCodePage)
            ElseIf IsNumeric(charset) Then
                Me.Encoding = Encoding.GetEncoding(Convert.ToInt32(charset))
            Else
                Me.Encoding = Encoding.GetEncoding(charset)
            End If

            Dim split As New Regex("\s+", RegexOptions.Compiled)
            For Each col As String In cols
                Dim parts() As String = split.Split(col)

                For Each part As String In parts
                    Log.DebugFormat("Column Part: {0}", part)
                Next

                Dim column As New CsvDataColumn(parts(0).Trim)
                Dim type As String = parts(1)
                If Regex.IsMatch(type, SCHEMA_COLUMN_TYPE_STRING, RegexOptions.IgnoreCase) Then
                    column.DataType = GetType(String)
                ElseIf Regex.IsMatch(type, SCHEMA_COLUMN_TYPE_DATETIME, RegexOptions.IgnoreCase) Then
                    column.DataType = GetType(DateTime)
                ElseIf Regex.IsMatch(type, SCHEMA_COLUMN_TYPE_DOUBLE, RegexOptions.IgnoreCase) Then
                    column.DataType = GetType(Double)
                ElseIf Regex.IsMatch(type, SCHEMA_COLUMN_TYPE_SINGLE, RegexOptions.IgnoreCase) Then
                    column.DataType = GetType(Single)
                ElseIf Regex.IsMatch(type, SCHEMA_COLUMN_TYPE_BOOLEAN, RegexOptions.IgnoreCase) Then
                    column.DataType = GetType(Boolean)
                ElseIf Regex.IsMatch(type, SCHEMA_COLUMN_TYPE_GUID, RegexOptions.IgnoreCase) Then
                    column.DataType = GetType(Guid)
                ElseIf Regex.IsMatch(type, SCHEMA_COLUMN_TYPE_LONG, RegexOptions.IgnoreCase) Then
                    column.DataType = GetType(Long)
                ElseIf Regex.IsMatch(type, SCHEMA_COLUMN_TYPE_INTEGER, RegexOptions.IgnoreCase) Then
                    column.DataType = GetType(Integer)
                ElseIf Regex.IsMatch(type, SCHEMA_COLUMN_TYPE_SHORT, RegexOptions.IgnoreCase) Then
                    column.DataType = GetType(Short)
                ElseIf Regex.IsMatch(type, SCHEMA_COLUMN_TYPE_DECIMAL, RegexOptions.IgnoreCase) Then
                    column.DataType = GetType(Decimal)
                ElseIf Regex.IsMatch(type, SCHEMA_COLUMN_TYPE_BYTE, RegexOptions.IgnoreCase) Then
                    column.DataType = GetType(Byte)
                End If

                If parts.Length = 4 Then
                    column.FieldWidth = parts(3)
                End If

                columns.Add(column)
            Next
        End If

        Return columns
    End Function

    ''' <summary>
    ''' Gets/sets the name and path of the schema file containing the column definitions.
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks>The default file name is schema.ini in the same directory as the csv file.</remarks>
    Public Overridable Property SchemaFile() As String
        Get
            Return _schemaFile
        End Get
        Set(ByVal value As String)
            _schemaFile = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Gets/sets the section in the schema file containing the column definitions.
    ''' </summary>
    ''' <value></value>
    ''' <returns>String</returns>
    ''' <remarks>If no section is given, the name of the file will be used.</remarks>
    Public Overridable Property SchemaSection() As String
        Get
            Return _schemaSection
        End Get
        Set(ByVal value As String)
            _schemaSection = value.Trim
        End Set
    End Property

    ''' <summary>
    ''' Gets/sets the stream of text data to parse.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Stream</returns>
    ''' <remarks></remarks>
    Public Overridable Property Stream() As Stream
        Get
            Return _stream
        End Get
        Set(ByVal value As Stream)
            _stream = value
        End Set
    End Property

#End Region

#Region "IDataReader"

    ''' <summary>
    ''' Closes the data reader object.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Close() Implements System.Data.IDataReader.Close
        Me.Parser.Close()
    End Sub

    ''' <summary>
    ''' Gets a value indicating the depth of nesting for the current row.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Integer</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Depth() As Integer Implements System.Data.IDataReader.Depth
        Get
            Return 0
        End Get
    End Property

    ''' <summary>
    ''' Returns a DataTable that describes the column metadata of the reader.
    ''' </summary>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetSchemaTable() As System.Data.DataTable Implements System.Data.IDataReader.GetSchemaTable
        Return Me.DataTable
    End Function

    ''' <summary>
    ''' Gets a value indicating whether the data reader is closed.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Boolean</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsClosed() As Boolean Implements System.Data.IDataReader.IsClosed
        Get
            Return _isClosed
        End Get
    End Property

    ''' <summary>
    ''' Advances the data reader to the next result, when reading the results of batched files.
    ''' </summary>
    ''' <returns>Boolean</returns>
    ''' <remarks></remarks>
    Public Function NextResult() As Boolean Implements System.Data.IDataReader.NextResult
        Return False
    End Function

    ''' <summary>
    ''' Advances the reader to the next record.
    ''' </summary>
    ''' <returns>Boolean</returns>
    ''' <remarks></remarks>
    Public Function Read() As Boolean Implements System.Data.IDataReader.Read
        Me.EnsureColumns()

        If Not Me.Parser.EndOfData Then
            _isClosed = False

            Dim fields() As String = Me.Parser.ReadFields
            Debug.WriteLine(fields.Length)
            REM Turn empty strings into Nothing, which Row turns to DbNull if the Column allows it
            REM Row.Add will throw an exception if the column does not allow nulls
            For f = 0 To fields.Length - 1
                Dim value As String = fields(f)

                Log.DebugFormat("Field({0})={1}", f, value)
                Log.DebugFormat("Field({0}).Length={1}", f, value.Length)

                If String.IsNullOrEmpty(value) Then
                    fields(f) = Nothing
                Else
                    fields(f) = value
                End If

                Log.DebugFormat("Field({0})={1}", f, value)
                Log.DebugFormat("Field({0}).Length={1}", f, value.Length)
            Next

            Me.DataTable.Rows.Add(fields)

            _dataRow = Me.DataTable.Rows(Me.DataTable.Rows.Count - 1)

            Return True
        Else
            _isClosed = True

            Return False
        End If
    End Function

    ''' <summary>
    ''' Gets the number of rows changed, inserted, or deleted.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Integer</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property RecordsAffected() As Integer Implements System.Data.IDataReader.RecordsAffected
        Get
            Return -1
        End Get
    End Property

    ''' <summary>
    ''' Gets the number of columns in the current row. 
    ''' </summary>
    ''' <value></value>
    ''' <returns>Integer</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FieldCount() As Integer Implements System.Data.IDataRecord.FieldCount
        Get
            Me.EnsureColumns()

            Return Me.DataTable.Columns.Count
        End Get
    End Property

    ''' <summary>
    ''' Gets the value of the specified column as a Boolean. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>Boolean</returns>
    ''' <remarks></remarks>
    Public Function GetBoolean(ByVal i As Integer) As Boolean Implements System.Data.IDataRecord.GetBoolean
        Return Me.CurrentDataRow.Item(i)
    End Function

    ''' <summary>
    ''' Gets the 8-bit unsigned integer value of the specified column. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>Byte</returns>
    ''' <remarks></remarks>
    Public Function GetByte(ByVal i As Integer) As Byte Implements System.Data.IDataRecord.GetByte
        Return Convert.ToByte(Me.GetChar(i))
    End Function

    ''' <summary>
    ''' Reads a stream of bytes from the specified column offset into the buffer as an array, starting at the given buffer offset.
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <param name="fieldOffset">Integer. The offset in the field to start reading bytes from.</param>
    ''' <param name="buffer">Byte(). The buffer to copy the fields bytes into.</param>
    ''' <param name="bufferoffset">Integer. The offset in the buffer to start copying to.</param>
    ''' <param name="length">Integer. The number of bytes to copy into the buffer.</param>
    ''' <returns>Integer</returns>
    ''' <remarks></remarks>
    Public Function GetBytes(ByVal i As Integer, ByVal fieldOffset As Long, ByVal buffer() As Byte, ByVal bufferoffset As Integer, ByVal length As Integer) As Long Implements System.Data.IDataRecord.GetBytes
        Dim encoding As New Text.UTF8Encoding(False)
        Dim bytes() As Byte = encoding.GetBytes(Me.GetString(i).ToCharArray)

        If buffer Is Nothing Then
            Return bytes.Length
        End If

        Array.Copy(bytes, buffer, buffer.Length)

        Return buffer.Length
    End Function

    ''' <summary>
    ''' Gets the character value of the specified column. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>Char</returns>
    ''' <remarks></remarks>
    Public Function GetChar(ByVal i As Integer) As Char Implements System.Data.IDataRecord.GetChar
        Dim value As String = Me.GetString(i)

        If value.Length >= 1 Then
            Return Convert.ToChar(value.Substring(0, 1))
        Else
            Return Char.MinValue
        End If
    End Function

    ''' <summary>
    ''' Reads a stream of characters from the specified column offset into the buffer as an array, starting at the given buffer offset. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <param name="fieldOffset">Integer. The offset in the field to start reading bytes from.</param>
    ''' <param name="buffer">Byte(). The buffer to copy the fields bytes into.</param>
    ''' <param name="bufferoffset">Integer. The offset in the buffer to start copying to.</param>
    ''' <param name="length">Integer. The number of bytes to copy into the buffer.</param>
    ''' <returns>Integer</returns>
    ''' <remarks></remarks>
    Public Function GetChars(ByVal i As Integer, ByVal fieldoffset As Long, ByVal buffer() As Char, ByVal bufferoffset As Integer, ByVal length As Integer) As Long Implements System.Data.IDataRecord.GetChars
        Dim value As String = Me.GetString(i)

        If buffer Is Nothing Then
            Return value.Length
        End If

        Array.Copy(value.ToCharArray(bufferoffset, length), buffer, buffer.Length)

        Return buffer.Length
    End Function

    ''' <summary>
    ''' Returns an IDataReader for the specified column ordinal. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>IDataReader</returns>
    ''' <remarks></remarks>
    Public Function GetData(ByVal i As Integer) As System.Data.IDataReader Implements System.Data.IDataRecord.GetData
        Throw New NotImplementedException
    End Function

    ''' <summary>
    ''' Gets the data type information for the specified field. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Function GetDataTypeName(ByVal i As Integer) As String Implements System.Data.IDataRecord.GetDataTypeName
        Me.EnsureColumns()

        Return Me.DataTable.Columns(i).DataType.Name
    End Function

    ''' <summary>
    ''' Gets the date and time data value of the specified field. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>DateTime</returns>
    ''' <remarks></remarks>
    Public Function GetDateTime(ByVal i As Integer) As DateTime Implements System.Data.IDataRecord.GetDateTime
        Return Me.GetValue(i)
    End Function

    ''' <summary>
    ''' Gets the fixed-position numeric value of the specified field. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>Decimal</returns>
    ''' <remarks></remarks>
    Public Function GetDecimal(ByVal i As Integer) As Decimal Implements System.Data.IDataRecord.GetDecimal
        Return Me.GetValue(i)
    End Function

    ''' <summary>
    ''' Gets the double-precision floating point number of the specified field. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>Double</returns>
    ''' <remarks></remarks>
    Public Function GetDouble(ByVal i As Integer) As Double Implements System.Data.IDataRecord.GetDouble
        Return Me.GetValue(i)
    End Function

    ''' <summary>
    ''' Gets the Type information corresponding to the type of Object that would be returned from GetValue. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>Type</returns>
    ''' <remarks></remarks>
    Public Function GetFieldType(ByVal i As Integer) As System.Type Implements System.Data.IDataRecord.GetFieldType
        Return Me.GetValue(i).GetType
    End Function

    ''' <summary>
    ''' Gets the single-precision floating point number of the specified field. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>Float</returns>
    ''' <remarks></remarks>
    Public Function GetFloat(ByVal i As Integer) As Single Implements System.Data.IDataRecord.GetFloat
        Return Me.GetValue(i)
    End Function

    ''' <summary>
    ''' Returns the GUID value of the specified field. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>Guid</returns>
    ''' <remarks></remarks>
    Public Function GetGuid(ByVal i As Integer) As System.Guid Implements System.Data.IDataRecord.GetGuid
        If Me.GetFieldType(i) Is GetType(String) Then
            Return New Guid(Me.GetString(i))
        Else
            Return Me.GetValue(i)
        End If
    End Function

    ''' <summary>
    ''' Gets the 16-bit signed integer value of the specified field. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>Int16</returns>
    ''' <remarks></remarks>
    Public Function GetInt16(ByVal i As Integer) As Short Implements System.Data.IDataRecord.GetInt16
        Return Me.GetValue(i)
    End Function

    ''' <summary>
    ''' Gets the 32-bit signed integer value of the specified field. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>Int32</returns>
    ''' <remarks></remarks>
    Public Function GetInt32(ByVal i As Integer) As Integer Implements System.Data.IDataRecord.GetInt32
        Return Me.GetValue(i)
    End Function

    ''' <summary>
    ''' Gets the 64-bit signed integer value of the specified field.
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>Int64</returns>
    ''' <remarks></remarks>
    Public Function GetInt64(ByVal i As Integer) As Long Implements System.Data.IDataRecord.GetInt64
        Return Me.GetValue(i)
    End Function

    ''' <summary>
    ''' Gets the name for the field to find. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Function GetName(ByVal i As Integer) As String Implements System.Data.IDataRecord.GetName
        Me.EnsureColumns()

        Return Me.DataTable.Columns(i).ColumnName
    End Function

    ''' <summary>
    ''' Return the index of the named field. 
    ''' </summary>
    ''' <param name="name">String. The name of the column to return the index of.</param>
    ''' <returns>Integer</returns>
    ''' <remarks></remarks>
    Public Function GetOrdinal(ByVal name As String) As Integer Implements System.Data.IDataRecord.GetOrdinal
        Me.EnsureColumns()

        Return Me.DataTable.Columns(name).Ordinal
    End Function

    ''' <summary>
    ''' Gets the string value of the specified field. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Function GetString(ByVal i As Integer) As String Implements System.Data.IDataRecord.GetString
        Return Me.GetValue(i).ToString
    End Function

    ''' <summary>
    ''' Return the value of the specified field. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>Object</returns>
    ''' <remarks></remarks>
    Public Function GetValue(ByVal i As Integer) As Object Implements System.Data.IDataRecord.GetValue
        Return Me.CurrentDataRow.Item(i)
    End Function

    ''' <summary>
    ''' Gets all the attribute fields in the collection for the current record and returns the number of values read.
    ''' </summary>
    ''' <param name="values">Object(). The array of values to fill.</param>
    ''' <returns>Integer</returns>
    ''' <remarks></remarks>
    Public Function GetValues(ByVal values() As Object) As Integer Implements System.Data.IDataRecord.GetValues
        Dim count As Integer = IIf(values.Length < Me.FieldCount, values.Length, Me.FieldCount)

        Array.Copy(Me.CurrentDataRow.ItemArray, values, count)

        Return count
    End Function

    ''' <summary>
    ''' Return whether the specified field is set to null. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>Boolean</returns>
    ''' <remarks></remarks>
    Public Function IsDBNull(ByVal i As Integer) As Boolean Implements System.Data.IDataRecord.IsDBNull
        Return Me.CurrentDataRow.IsNull(i)
    End Function

    ''' <summary>
    ''' Gets the column located at the specified index. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <value></value>
    ''' <returns>Object</returns>
    ''' <remarks></remarks>
    Default Public Overloads ReadOnly Property Item(ByVal i As Integer) As Object Implements System.Data.IDataRecord.Item
        Get
            Return Me.GetValue(i)
        End Get
    End Property

    ''' <summary>
    ''' Gets the column with the specified name. 
    ''' </summary>
    ''' <param name="name">String. The name of the column to return.</param>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Default Public Overloads ReadOnly Property Item(ByVal name As String) As Object Implements System.Data.IDataRecord.Item
        Get
            Return Me.GetValue(Me.GetOrdinal(name))
        End Get
    End Property

    ''' <summary>
    ''' Ensures that file is read to determine columns from the header line if columns are not already defined.
    ''' </summary>
    ''' <remarks></remarks>
    Protected Overridable Sub EnsureColumns()
        If Me.DataTable.Columns.Count = 0 Then
            If Not String.IsNullOrEmpty(Me.SchemaFile) And File.Exists(Me.SchemaFile) Then
                For Each column As DataColumn In Me.GetColumnsFromSchemaFile
                    Me.DataTable.Columns.Add(column)
                Next
            Else
                For Each column As DataColumn In Me.GetColumnsFromFirstRow
                    Me.DataTable.Columns.Add(column)
                Next
            End If
        End If
    End Sub

#End Region

#Region "IDisposable"

    ''' <summary>
    ''' Disposes the current reader.
    ''' </summary>
    ''' <param name="disposing">Boolean. Flag whether we're dispoosing or getting collected..</param>
    ''' <remarks></remarks>
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not _disposed Then
            If disposing Then
                If _parser IsNot Nothing Then
                    _parser.Dispose()
                End If
            End If
            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        _disposed = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class
