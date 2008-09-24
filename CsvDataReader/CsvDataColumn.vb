Imports System.Data
Imports log4net

''' <summary>
''' Defines a column for use in the CsvDataReader.
''' </summary>
''' <remarks></remarks>
Public Class CsvDataColumn
    Inherits DataColumn

#Region "Privates"

    Private Shared ReadOnly Log As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod.DeclaringType)

    Private _fieldWidth As Integer = 0

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Creates a new CsvDataColumn
    ''' </summary>
    ''' <param name="columnName">String. The name of the column.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal columnName As String)
        MyBase.New(columnName)
    End Sub

    ''' <summary>
    ''' Creates a new CsvDataColumn
    ''' </summary>
    ''' <param name="columnName">String. The name of the column.</param>
    ''' <param name="dataType">Type. The data type of the column value.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal columnName As String, ByVal dataType As System.Type)
        MyBase.New(columnName, dataType)
    End Sub

    ''' <summary>
    ''' Creates a new CsvDataColumn
    ''' </summary>
    ''' <param name="columnName">String. The name of the column.</param>
    ''' <param name="dataType">Type. The data type of the column value.</param>
    ''' <param name="fieldWidth">Integer. The width of the field if coming from a fied width file.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal columnName As String, ByVal dataType As System.Type, ByVal fieldWidth As Integer)
        MyBase.New(columnName, dataType)
        Me.FieldWidth = FieldWidth
    End Sub

#End Region

#Region "Properties"

    ''' <summary>
    ''' Gets/sets the width a fixed length field.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property FieldWidth() As Integer
        Get
            Return _fieldWidth
        End Get
        Set(ByVal value As Integer)
            _fieldWidth = value
        End Set
    End Property

#End Region

End Class
