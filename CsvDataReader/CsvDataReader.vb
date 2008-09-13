Imports System.Data
Imports System.Text
Imports System.IO
Imports System.Collections.ObjectModel
Imports log4net

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

    Public Sub New(ByVal path As String)
        MyBase.New(path)

        For Each column As DataColumn In Me.GetColumnsFromFirstRow
            Me.DataTable.Columns.Add(column)
        Next
    End Sub

    Public Sub New(ByVal path As String, ByVal columns As Collection(Of DataColumn))
        MyBase.New(path)

        For Each column As DataColumn In columns
            column.AllowDBNull = True
            Me.DataTable.Columns.Add(column)
        Next
    End Sub

    Public Sub New(ByVal path As String, ByVal encoding As Encoding)
        MyBase.New(path, encoding)

        For Each column As DataColumn In Me.GetColumnsFromFirstRow
            Me.DataTable.Columns.Add(column)
        Next
    End Sub

    Public Sub New(ByVal path As String, ByVal columns As Collection(Of DataColumn), ByVal encoding As Encoding)
        MyBase.New(path, encoding)

        For Each column As DataColumn In columns
            Me.DataTable.Columns.Add(column)
        Next
    End Sub

#End Region

#Region "Properties"

    Protected Overridable ReadOnly Property DataTable() As DataTable
        Get
            Return _dataTable
        End Get
    End Property

    Protected Overridable ReadOnly Property CurrentDataRow() As DataRow
        Get
            Return _dataRow
        End Get
    End Property

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

    Public Overrides Sub Close() Implements System.Data.IDataReader.Close
        MyBase.Close()
    End Sub

    Public ReadOnly Property Depth() As Integer Implements System.Data.IDataReader.Depth
        Get
            Return 0
        End Get
    End Property

    Public Function GetSchemaTable() As System.Data.DataTable Implements System.Data.IDataReader.GetSchemaTable
        Return Me.DataTable
    End Function

    Public ReadOnly Property IsClosed() As Boolean Implements System.Data.IDataReader.IsClosed
        Get
            Return _isClosed
        End Get
    End Property

    Public Function NextResult() As Boolean Implements System.Data.IDataReader.NextResult
        Return False
    End Function

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

    Public ReadOnly Property RecordsAffected() As Integer Implements System.Data.IDataReader.RecordsAffected
        Get
            Return 0
        End Get
    End Property

    Public ReadOnly Property FieldCount() As Integer Implements System.Data.IDataRecord.FieldCount
        Get
            Return Me.DataTable.Columns.Count
        End Get
    End Property

    Public Function GetBoolean(ByVal i As Integer) As Boolean Implements System.Data.IDataRecord.GetBoolean
        Return Me.CurrentDataRow.Item(i)
    End Function

    Public Function GetByte(ByVal i As Integer) As Byte Implements System.Data.IDataRecord.GetByte
        'Return Convert.ToByte(_row(i).Trim)
    End Function

    Public Function GetBytes(ByVal i As Integer, ByVal fieldOffset As Long, ByVal buffer() As Byte, ByVal bufferoffset As Integer, ByVal length As Integer) As Long Implements System.Data.IDataRecord.GetBytes

    End Function

    Public Function GetChar(ByVal i As Integer) As Char Implements System.Data.IDataRecord.GetChar
        Return Me.GetString(i)
    End Function

    Public Function GetChars(ByVal i As Integer, ByVal fieldoffset As Long, ByVal buffer() As Char, ByVal bufferoffset As Integer, ByVal length As Integer) As Long Implements System.Data.IDataRecord.GetChars
        Throw New NotImplementedException
    End Function

    Public Function GetData(ByVal i As Integer) As System.Data.IDataReader Implements System.Data.IDataRecord.GetData
        Throw New NotImplementedException
    End Function

    Public Function GetDataTypeName(ByVal i As Integer) As String Implements System.Data.IDataRecord.GetDataTypeName
        Return Me.DataTable.Columns(i).DataType.Name
    End Function

    Public Function GetDateTime(ByVal i As Integer) As Date Implements System.Data.IDataRecord.GetDateTime
        Return Me.GetValue(i)
    End Function

    Public Function GetDecimal(ByVal i As Integer) As Decimal Implements System.Data.IDataRecord.GetDecimal
        Return Me.GetValue(i)
    End Function

    Public Function GetDouble(ByVal i As Integer) As Double Implements System.Data.IDataRecord.GetDouble
        Return Me.GetValue(i)
    End Function

    Public Function GetFieldType(ByVal i As Integer) As System.Type Implements System.Data.IDataRecord.GetFieldType
        Return Me.GetValue(i).GetType
    End Function

    Public Function GetFloat(ByVal i As Integer) As Single Implements System.Data.IDataRecord.GetFloat
        Return Me.GetValue(i)
    End Function

    Public Function GetGuid(ByVal i As Integer) As System.Guid Implements System.Data.IDataRecord.GetGuid
        If Me.GetFieldType(i) Is GetType(String) Then
            Return New Guid(Me.GetString(i))
        Else
            Return Me.GetValue(i)
        End If
    End Function

    Public Function GetInt16(ByVal i As Integer) As Short Implements System.Data.IDataRecord.GetInt16
        Return Me.GetValue(i)
    End Function

    Public Function GetInt32(ByVal i As Integer) As Integer Implements System.Data.IDataRecord.GetInt32
        Return Me.GetValue(i)
    End Function

    Public Function GetInt64(ByVal i As Integer) As Long Implements System.Data.IDataRecord.GetInt64
        Return Me.GetValue(i)
    End Function

    Public Function GetName(ByVal i As Integer) As String Implements System.Data.IDataRecord.GetName
        Return Me.DataTable.Columns(i).ColumnName
    End Function

    Public Function GetOrdinal(ByVal name As String) As Integer Implements System.Data.IDataRecord.GetOrdinal
        Return Me.DataTable.Columns(name).Ordinal
    End Function

    Public Function GetString(ByVal i As Integer) As String Implements System.Data.IDataRecord.GetString
        Return Me.CurrentDataRow.Item(i).ToString
    End Function

    Public Function GetValue(ByVal i As Integer) As Object Implements System.Data.IDataRecord.GetValue
        Return Me.CurrentDataRow.Item(i)
    End Function

    Public Function GetValues(ByVal values() As Object) As Integer Implements System.Data.IDataRecord.GetValues
        Dim count As Integer = IIf(values.Length < Me.FieldCount, values.Length, Me.FieldCount)

        Array.Copy(Me.CurrentDataRow.ItemArray, values, count)

        Return count
    End Function

    Public Function IsDBNull(ByVal i As Integer) As Boolean Implements System.Data.IDataRecord.IsDBNull
        Return Me.CurrentDataRow.IsNull(i)
    End Function

    Default Public Overloads ReadOnly Property Item(ByVal i As Integer) As Object Implements System.Data.IDataRecord.Item
        Get
            Return Me.GetValue(i)
        End Get
    End Property

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
