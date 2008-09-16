Imports System.Data
Imports System.Text
Imports System.IO
Imports System.Collections.ObjectModel
Imports log4net

''' <summary>
''' Implements an IDataReader fo Csv data files.
''' </summary>
''' <remarks><seealso cref="IDataReader">IDataReader</seealso>, <seealso cref="IDataRecord">IDataRecord</seealso></remarks>
Public Class CsvDataReader
    Inherits StreamReader
    Implements IDataReader

#Region "Privates"

    Private Const DEFAULT_FIELD_SEPARATOR As String = ","
    Private Shared ReadOnly Log As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod.DeclaringType)

    Private _disposed As Boolean = False
    Private _dataTable As New DataTable
    Private _dataRow As DataRow
    Private _isClosed As Boolean = True

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Creates a new CsvDataReader for the file specified.
    ''' </summary>
    ''' <param name="path">String. The full path to the file to reader.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal path As String)
        Me.New(path, Encoding.UTF8)
    End Sub

    ''' <summary>
    ''' Creates a new CsvDataReader for the file specified.
    ''' </summary>
    ''' <param name="path">String. The full path to the file to reader.</param>
    ''' <param name="columns">Collection(Of DataColumn). The collection of column definitions for the specified files columns.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal path As String, ByVal columns As Collection(Of DataColumn))
        Me.New(path, columns, Encoding.UTF8)
    End Sub

    ''' <summary>
    ''' Creates a new CsvDataReader for the file specified.
    ''' </summary>
    ''' <param name="path">String. The full path to the file to reader.</param>
    ''' <param name="encoding">Encoding. The encoding of the specified file.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal path As String, ByVal encoding As Encoding)
        MyBase.New(path, encoding)

        For Each column As DataColumn In Me.GetColumnsFromFirstRow
            Me.DataTable.Columns.Add(column)
        Next
    End Sub

    ''' <summary>
    ''' Creates a new CsvDataReader for the file specified.
    ''' </summary>
    ''' <param name="path">String. The full path to the file to reader.</param>
    ''' <param name="columns">Collection(Of DataColumn). The collection of column definitions for the specified files columns.</param>
    ''' <param name="encoding">Encoding. The encoding of the specified file.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal path As String, ByVal columns As Collection(Of DataColumn), ByVal encoding As Encoding)
        MyBase.New(path, encoding)

        For Each column As DataColumn In columns
            Me.DataTable.Columns.Add(column)
        Next
    End Sub

#End Region

#Region "Properties"

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
        Dim header As String = Me.ReadLine
        Dim fields() As String = header.Split(DEFAULT_FIELD_SEPARATOR)

        Dim columns As New Collection(Of DataColumn)
        For Each field As String In fields
            Log.DebugFormat("Adding column {0}", field.Trim)

            Dim column As New DataColumn(field.Trim, GetType(String))
            column.AllowDBNull = True
            columns.Add(column)
        Next

        Return columns
    End Function

#End Region

#Region "IDataReader"

    ''' <summary>
    ''' Closes the data reader object.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overrides Sub Close() Implements System.Data.IDataReader.Close
        MyBase.Close()

        Dim s As StreamReader = Nothing
        Dim r As IDataReader = s

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
    Public Overloads Function Read() As Boolean Implements System.Data.IDataReader.Read
        If Me.Peek >= 0 Then
            _isClosed = False

            Dim line As String = Me.ReadLine.Trim

            If String.IsNullOrEmpty(line) Then
                If Me.Peek < 0 Then
                    _isClosed = True
                End If

                Return False
            End If

            Dim fields() As Object = line.Split(DEFAULT_FIELD_SEPARATOR)

            REM Turn empty strings into Nothing, which Row turns to DbNull if the Column allows it
            REM Row.Add will throw an exception if the column does not allow nulls
            For f = 0 To fields.Length - 1
                Dim value As String = fields(f).Trim

                If String.IsNullOrEmpty(value) Then
                    fields(f) = Nothing
                Else
                    fields(f) = value
                End If
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
        Return Me.DataTable.Columns(i).ColumnName
    End Function

    ''' <summary>
    ''' Return the index of the named field. 
    ''' </summary>
    ''' <param name="name">String. The name of the column to return the index of.</param>
    ''' <returns>Integer</returns>
    ''' <remarks></remarks>
    Public Function GetOrdinal(ByVal name As String) As Integer Implements System.Data.IDataRecord.GetOrdinal
        Return Me.DataTable.Columns(name).Ordinal
    End Function

    ''' <summary>
    ''' Gets the string value of the specified field. 
    ''' </summary>
    ''' <param name="i">Integer. The ordinal of the column to return.</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Function GetString(ByVal i As Integer) As String Implements System.Data.IDataRecord.GetString
        Return Me.CurrentDataRow.Item(i).ToString
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

#End Region

#Region "IDisposable"

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not _disposed Then
            If disposing Then
                Me.DataTable.Dispose()
            End If

            _disposed = True
        End If

        MyBase.Dispose(disposing)
    End Sub

#End Region

End Class
