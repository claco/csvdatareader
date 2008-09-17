Imports System.IO
Imports System.Collections.ObjectModel
Imports NUnit.Framework
Imports ChrisLaco.Data

<TestFixture(Description:="CsvDataReader Tests")> _
Public Class CsvDataReaderTests

#Region "Constructor Tests"

    <TestFixtureSetUp()> _
    Public Sub TestFixtureSetupUp()
        log4net.Config.XmlConfigurator.Configure()
    End Sub

    <Test(Description:="New throws exception for bogus file")> _
    <ExpectedException(GetType(FileNotFoundException))> _
    Public Sub NewBadFileName()
        Using reader As New CsvDataReader("crapfile")

        End Using
    End Sub

#End Region

    <Test(Description:="Test Cav file with header")> _
    Public Sub CsvWithHeader()
        Using reader As IDataReader = New CsvDataReader("data\header.csv")
            Assert.IsInstanceOfType(GetType(CsvDataReader), reader, "Is instance of CsvDataReader")
            Assert.AreEqual(6, reader.FieldCount, "Has correct number of fields")
            Assert.AreEqual(0, reader.Depth, "Depth is always zero")

            Assert.AreEqual("STRING", reader.GetName(0), "Has first column name")
            Assert.AreEqual("String", reader.GetDataTypeName(0), "First column type name is String")
            Assert.AreEqual(0, reader.GetOrdinal("String"), "First column ordinal is correct")

            Assert.AreEqual("INTEGER", reader.GetName(1), "Has second column name")
            Assert.AreEqual("String", reader.GetDataTypeName(1), "Second column type name is String")
            Assert.AreEqual(1, reader.GetOrdinal("INTEGER"), "Second column ordinal is correct")

            Assert.AreEqual("DATETIME", reader.GetName(2), "Has third column name")
            Assert.AreEqual("String", reader.GetDataTypeName(2), "Third column type name is String")
            Assert.AreEqual(2, reader.GetOrdinal("DATETIME"), "Third column ordinal is correct")

            Assert.AreEqual("DECIMAL", reader.GetName(3), "Has fourth column name")
            Assert.AreEqual("String", reader.GetDataTypeName(3), "Fourth column type name is String")
            Assert.AreEqual(3, reader.GetOrdinal("DECIMAL"), "Fourth column ordinal is correct")

            Assert.AreEqual("GUID", reader.GetName(4), "Has fifth column name")
            Assert.AreEqual("String", reader.GetDataTypeName(4), "Fifth column type name is String")
            Assert.AreEqual(4, reader.GetOrdinal("GUID"), "Fifth column ordinal is correct")

            Assert.AreEqual("BOOLEAN", reader.GetName(5), "Has sixth column name")
            Assert.AreEqual("String", reader.GetDataTypeName(5), "Sixth column type name is String")
            Assert.AreEqual(5, reader.GetOrdinal("BOOLEAN"), "Sixth column ordinal is correct")


            Assert.IsTrue(reader.Read, "Read first line")
            Dim values(reader.FieldCount - 1) As Object
            Assert.AreEqual(6, reader.GetValues(values), "GetValues returns the number of columns")
            Assert.AreEqual(6, values.Length, "Values array has corrent length")

            Assert.AreEqual("Christopher", reader.GetString(0), "First column value is correct")
            Assert.AreEqual(GetType(String), reader.GetFieldType(0), "First column is System.String")
            Assert.IsFalse(reader.IsDBNull(0), "First column is not null")
            Assert.AreEqual(values(0), reader.GetValue(0), "GetValues item sames as GetValue")
            Assert.AreEqual(values(0), reader.Item(0), "GetValues item sames as Item")
            Assert.AreEqual(values(0), reader.Item(reader.GetName(0)), "GetValues item sames as Item")
            Assert.AreEqual("C", reader.GetChar(0).ToString, "GetChar returns first character")
            Assert.AreEqual(11, reader.GetChars(0, Nothing, Nothing, Nothing, Nothing), "GetChars returns length when no buffer is passed")
            Dim chars(reader.GetString(0).Length - 1) As Char
            Assert.AreEqual(11, reader.GetChars(0, 0, chars, 0, 11), "Read characters from column")
            Assert.AreEqual(11, chars.Length, "Read correct number of character")
            Assert.AreEqual("C", chars(0).ToString)
            Assert.AreEqual("h", chars(1).ToString)
            Assert.AreEqual("r", chars(2).ToString)
            Assert.AreEqual("i", chars(3).ToString)
            Assert.AreEqual("s", chars(4).ToString)
            Assert.AreEqual("t", chars(5).ToString)
            Assert.AreEqual("o", chars(6).ToString)
            Assert.AreEqual("p", chars(7).ToString)
            Assert.AreEqual("h", chars(8).ToString)
            Assert.AreEqual("e", chars(9).ToString)
            Assert.AreEqual("r", chars(10).ToString)

            Assert.AreEqual(11, reader.GetBytes(0, Nothing, Nothing, Nothing, Nothing), "GetBytes returns length when no buffer is passed")
            Assert.AreEqual(Convert.ToByte(Convert.ToChar("C")), reader.GetByte(0), "GetByte returns first")
            Dim bytes(reader.GetString(0).Length - 1) As Byte
            Assert.AreEqual(11, reader.GetBytes(0, 0, bytes, 0, 11), "Read bytes from column")
            Assert.AreEqual(11, bytes.Length, "Read correct number of bytes")
            Assert.AreEqual(67, bytes(0))
            Assert.AreEqual(104, bytes(1))
            Assert.AreEqual(114, bytes(2))
            Assert.AreEqual(105, bytes(3))
            Assert.AreEqual(115, bytes(4))
            Assert.AreEqual(116, bytes(5))
            Assert.AreEqual(111, bytes(6))
            Assert.AreEqual(112, bytes(7))
            Assert.AreEqual(104, bytes(8))
            Assert.AreEqual(101, bytes(9))
            Assert.AreEqual(114, bytes(10))





            Assert.AreEqual("35", reader.GetString(1), "Second column value is correct")
            Assert.AreEqual(Convert.ToInt16(35), reader.GetInt16(1), "Second column value is correct as Int16")
            Assert.AreEqual(Convert.ToInt32(35), reader.GetInt32(1), "Second column value is correct as Int32")
            Assert.AreEqual(Convert.ToInt64(35), reader.GetInt64(1), "Second column value is correct as Int64")
            Assert.AreEqual(Convert.ToSingle(35), reader.GetFloat(1), "Second column value is correct as Float/Single")
            Assert.AreEqual(Convert.ToDouble(35), reader.GetDouble(1), "Second column value is correct as Double")
            Assert.AreEqual(Convert.ToDecimal(35), reader.GetDecimal(1), "Second column value is correct as Decimal")
            Assert.AreEqual(GetType(String), reader.GetFieldType(1), "Second column is System.String")
            Assert.IsFalse(reader.IsDBNull(1), "Second column is not null")
            Assert.AreEqual(values(1), reader.GetValue(1), "GetValues item sames as GetValue")
            Assert.AreEqual(values(1), reader.Item(1), "GetValues item sames as Item")
            Assert.AreEqual(values(1), reader.Item(reader.GetName(1)), "GetValues item sames as Item")
            Assert.AreEqual("3", reader.GetChar(1).ToString, "GetChar returns first character")

            Assert.AreEqual("1/2/2003 2:34:56", reader.GetString(2), "Third column value is correct")
            Assert.AreEqual(DateTime.Parse("1/2/2003 2:34:56"), reader.GetDateTime(2), "Third column value is correct as DateTime")
            Assert.AreEqual(GetType(String), reader.GetFieldType(2), "Third column is System.String")
            Assert.IsFalse(reader.IsDBNull(2), "Third column is not null")
            Assert.AreEqual(values(2), reader.GetValue(2), "GetValues item sames as GetValue")
            Assert.AreEqual(values(2), reader.Item(2), "GetValues item sames as Item")
            Assert.AreEqual(values(2), reader.Item(reader.GetName(2)), "GetValues item sames as Item")
            Assert.AreEqual("1", reader.GetChar(2).ToString, "GetChar returns first character")

            Assert.AreEqual("1.23", reader.GetString(3), "Fourth column value is correct")
            Assert.AreEqual(Convert.ToSingle(1.23), reader.GetFloat(3), "Fourth column value is correct as Float/Single")
            Assert.AreEqual(Convert.ToDouble(1.23), reader.GetDouble(3), "Fourth column value is correct as Double")
            Assert.AreEqual(Convert.ToDecimal(1.23), reader.GetDecimal(3), "Fourth column value is correct as Decimal")
            Assert.AreEqual(GetType(String), reader.GetFieldType(3), "Fourth column is System.String")
            Assert.IsFalse(reader.IsDBNull(3), "Fourth column is not null")
            Assert.AreEqual(values(3), reader.GetValue(3), "GetValues item sames as GetValue")
            Assert.AreEqual(values(3), reader.Item(3), "GetValues item sames as Item")
            Assert.AreEqual(values(3), reader.Item(reader.GetName(3)), "GetValues item sames as Item")
            Assert.AreEqual("1", reader.GetChar(3).ToString, "GetChar returns first character")

            Assert.AreEqual("11111111-2222-3333-4444-555555555555", reader.GetString(4), "Fifth column value is correct")
            Assert.AreEqual(New Guid("11111111-2222-3333-4444-555555555555"), reader.GetGuid(4), "Fifth column value is correct as Decimal")
            Assert.AreEqual(GetType(String), reader.GetFieldType(4), "Fifth column is System.String")
            Assert.IsFalse(reader.IsDBNull(4), "Fifth column is not null")
            Assert.AreEqual(values(4), reader.GetValue(4), "GetValues item sames as GetValue")
            Assert.AreEqual(values(4), reader.Item(4), "GetValues item sames as Item")
            Assert.AreEqual(values(4), reader.Item(reader.GetName(4)), "GetValues item sames as Item")
            Assert.AreEqual("1", reader.GetChar(4).ToString, "GetChar returns first character")

            Assert.AreEqual("True", reader.GetString(5), "Sixth column value is correct")
            Assert.IsTrue(reader.GetBoolean(5), "Sixth column value is correct as Boolean")
            Assert.AreEqual(GetType(String), reader.GetFieldType(5), "Sixth column is System.String")
            Assert.IsFalse(reader.IsDBNull(5), "Sixth column is not null")
            Assert.AreEqual(values(5), reader.GetValue(5), "GetValues item sames as GetValue")
            Assert.AreEqual(values(5), reader.Item(5), "GetValues item sames as Item")
            Assert.AreEqual(values(5), reader.Item(reader.GetName(5)), "GetValues item sames as Item")
            Assert.AreEqual("T", reader.GetChar(5).ToString, "GetChar returns first character")


            Assert.IsTrue(reader.Read, "Read second line")
            Array.Clear(values, 0, values.Length)
            Array.Resize(values, 2)
            Assert.AreEqual(2, reader.GetValues(values), "GetValues returns the number value slize")
            Assert.AreEqual(2, values.Length, "Values array has corrent length")

            Assert.IsEmpty(reader.GetString(0), "First column value is empty")
            Assert.AreEqual(GetType(System.DBNull), reader.GetFieldType(0), "First column is System.DBNull")
            Assert.IsTrue(reader.IsDBNull(0), "First column is null")
            Assert.AreEqual(values(0), reader.GetValue(0), "GetValues item sames as GetValue")
            Assert.AreEqual(Char.MinValue, reader.GetChar(0), "GetChar returns Char.MinValue for DbNull")

            Assert.AreEqual("23", reader.GetString(1), "Second column value is correct")
            Assert.AreEqual(Convert.ToInt16(23), reader.GetInt16(1), "Second column value is correct as Int16")
            Assert.AreEqual(Convert.ToInt32(23), reader.GetInt32(1), "Second column value is correct as Int32")
            Assert.AreEqual(Convert.ToInt64(23), reader.GetInt64(1), "Second column value is correct as Int64")
            Assert.AreEqual(Convert.ToSingle(23), reader.GetFloat(1), "Second column value is correct as Float/Single")
            Assert.AreEqual(Convert.ToDouble(23), reader.GetDouble(1), "Second column value is correct as Double")
            Assert.AreEqual(Convert.ToDecimal(23), reader.GetDecimal(1), "Second column value is correct as Decimal")
            Assert.AreEqual(GetType(String), reader.GetFieldType(1), "Second column is System.String")
            Assert.IsFalse(reader.IsDBNull(1), "Second column is not null")
            Assert.AreEqual(values(1), reader.GetValue(1), "GetValues item sames as GetValue")

            Assert.AreEqual("2/3/2004", reader.GetString(2), "Third column value is correct")
            Assert.AreEqual(DateTime.Parse("2/3/2004"), reader.GetDateTime(2), "Third column value is correct as DateTime")
            Assert.AreEqual(GetType(String), reader.GetFieldType(2), "Third column is System.String")
            Assert.IsFalse(reader.IsDBNull(2), "Third column is not null")

            Assert.AreEqual("1.342", reader.GetString(3), "Fourth column value is correct")
            Assert.AreEqual(Convert.ToSingle(1.342), reader.GetFloat(3), "Fourth column value is correct as Float/Single")
            Assert.AreEqual(Convert.ToDouble(1.342), reader.GetDouble(3), "Fourth column value is correct as Double")
            Assert.AreEqual(Convert.ToDecimal(1.342), reader.GetDecimal(3), "Fourth column value is correct as Decimal")
            Assert.AreEqual(GetType(String), reader.GetFieldType(3), "Fourth column is System.String")
            Assert.IsFalse(reader.IsDBNull(3), "Fourth column is not null")

            Assert.IsEmpty(reader.GetString(4), "Fifth column value is empty")
            Assert.AreEqual(GetType(System.DBNull), reader.GetFieldType(4), "Fifth column is System.DBNull")
            Assert.IsTrue(reader.IsDBNull(4), "Fifth column is null")

            Assert.AreEqual("False", reader.GetString(5), "Sixth column value is correct")
            Assert.IsFalse(reader.GetBoolean(5), "Sixth column value is correct as Boolean")
            Assert.AreEqual(GetType(String), reader.GetFieldType(5), "Sixth column is System.String")
            Assert.IsFalse(reader.IsDBNull(5), "Sixth column is not null")

            Assert.IsFalse(reader.IsClosed, "Reader is still open until last read")
            Assert.IsFalse(reader.Read, "End of file")
            Assert.IsFalse(reader.NextResult, "NextResult always returns false")
            Assert.AreEqual(-1, reader.RecordsAffected, "RecordsAffected always returns -1")
            Assert.IsTrue(reader.IsClosed, "Reader is closed after final read")

            Try
                reader.GetData(0)
                Assert.Fail("No exception thrown for GetData")
            Catch ex As Exception
                Assert.IsInstanceOfType(GetType(NotImplementedException), ex, "GetData threw NotImplementedException")
            End Try
        End Using
    End Sub

    <Test(Description:="Test Cav file with defined columns")> _
    Public Sub CsvWithColumns()
        Dim columns As New Collection(Of DataColumn)
        columns.Add(New DataColumn("STRING", GetType(String)))
        columns.Add(New DataColumn("INTEGER", GetType(Integer)))
        columns.Add(New DataColumn("DATETIME", GetType(DateTime)))
        columns.Add(New DataColumn("DECIMAL", GetType(Decimal)))
        columns.Add(New DataColumn("GUID", GetType(Guid)))
        columns.Add(New DataColumn("BOOLEAN", GetType(Boolean)))

        Using reader As IDataReader = New CsvDataReader("data\noheader.csv", columns)
            Assert.IsInstanceOfType(GetType(CsvDataReader), reader, "Is instance of CsvDataReader")
            Assert.AreEqual(6, reader.FieldCount, "Has correct number of fields")


            Assert.AreEqual("STRING", reader.GetName(0), "Has first column name")
            Assert.AreEqual("String", reader.GetDataTypeName(0), "First column type name is String")
            Assert.AreEqual(0, reader.GetOrdinal("String"), "First column ordinal is correct")


            Dim schemaTable As DataTable = reader.GetSchemaTable
            For c As Integer = 0 To schemaTable.Columns.Count - 1
                Me.AreEqual(columns.Item(c), schemaTable.Columns.Item(c))
            Next


            Assert.AreEqual("INTEGER", reader.GetName(1), "Has second column name")
            Assert.AreEqual("Int32", reader.GetDataTypeName(1), "Second column type name is String")
            Assert.AreEqual(1, reader.GetOrdinal("INTEGER"), "Second column ordinal is correct")

            Assert.AreEqual("DATETIME", reader.GetName(2), "Has third column name")
            Assert.AreEqual("DateTime", reader.GetDataTypeName(2), "Third column type name is String")
            Assert.AreEqual(2, reader.GetOrdinal("DATETIME"), "Third column ordinal is correct")

            Assert.AreEqual("DECIMAL", reader.GetName(3), "Has fourth column name")
            Assert.AreEqual("Decimal", reader.GetDataTypeName(3), "Fourth column type name is String")
            Assert.AreEqual(3, reader.GetOrdinal("DECIMAL"), "Fourth column ordinal is correct")

            Assert.AreEqual("GUID", reader.GetName(4), "Has fifth column name")
            Assert.AreEqual("Guid", reader.GetDataTypeName(4), "Fifth column type name is String")
            Assert.AreEqual(4, reader.GetOrdinal("GUID"), "Fifth column ordinal is correct")

            Assert.AreEqual("BOOLEAN", reader.GetName(5), "Has sixth column name")
            Assert.AreEqual("Boolean", reader.GetDataTypeName(5), "Sixth column type name is Boolean")
            Assert.AreEqual(5, reader.GetOrdinal("BOOLEAN"), "Sixth column ordinal is correct")


            Assert.IsTrue(reader.Read, "Read first line")
            Assert.AreEqual("Christopher", reader.GetString(0), "First column value is correct")
            Assert.AreEqual(GetType(String), reader.GetFieldType(0), "First column is System.String")
            Assert.IsFalse(reader.IsDBNull(0), "First column is not null")
            Assert.AreEqual("C", reader.GetChar(0).ToString, "GetChar returns first character")
            Assert.AreEqual(11, reader.GetChars(0, Nothing, Nothing, Nothing, Nothing), "GetChars returns length when no buffer is passed")
            Dim chars(reader.GetString(0).Length - 1) As Char
            Assert.AreEqual(11, reader.GetChars(0, 0, chars, 0, 11), "Read characters from column")
            Assert.AreEqual(11, chars.Length, "Read correct number of character")
            Assert.AreEqual("C", chars(0).ToString)
            Assert.AreEqual("h", chars(1).ToString)
            Assert.AreEqual("r", chars(2).ToString)
            Assert.AreEqual("i", chars(3).ToString)
            Assert.AreEqual("s", chars(4).ToString)
            Assert.AreEqual("t", chars(5).ToString)
            Assert.AreEqual("o", chars(6).ToString)
            Assert.AreEqual("p", chars(7).ToString)
            Assert.AreEqual("h", chars(8).ToString)
            Assert.AreEqual("e", chars(9).ToString)
            Assert.AreEqual("r", chars(10).ToString)

            Assert.AreEqual(11, reader.GetBytes(0, Nothing, Nothing, Nothing, Nothing), "GetBytes returns length when no buffer is passed")
            Assert.AreEqual(Convert.ToByte(Convert.ToChar("C")), reader.GetByte(0), "GetByte returns first")
            Dim bytes(reader.GetString(0).Length - 1) As Byte
            Assert.AreEqual(11, reader.GetBytes(0, 0, bytes, 0, 11), "Read bytes from column")
            Assert.AreEqual(11, bytes.Length, "Read correct number of bytes")
            Assert.AreEqual(67, bytes(0))
            Assert.AreEqual(104, bytes(1))
            Assert.AreEqual(114, bytes(2))
            Assert.AreEqual(105, bytes(3))
            Assert.AreEqual(115, bytes(4))
            Assert.AreEqual(116, bytes(5))
            Assert.AreEqual(111, bytes(6))
            Assert.AreEqual(112, bytes(7))
            Assert.AreEqual(104, bytes(8))
            Assert.AreEqual(101, bytes(9))
            Assert.AreEqual(114, bytes(10))

            Assert.AreEqual("35", reader.GetString(1), "Second column value is correct")
            Assert.AreEqual(Convert.ToInt16(35), reader.GetInt16(1), "Second column value is correct as Int16")
            Assert.AreEqual(Convert.ToInt32(35), reader.GetInt32(1), "Second column value is correct as Int32")
            Assert.AreEqual(Convert.ToInt64(35), reader.GetInt64(1), "Second column value is correct as Int64")
            Assert.AreEqual(Convert.ToSingle(35), reader.GetFloat(1), "Second column value is correct as Float/Single")
            Assert.AreEqual(Convert.ToDouble(35), reader.GetDouble(1), "Second column value is correct as Double")
            Assert.AreEqual(Convert.ToDecimal(35), reader.GetDecimal(1), "Second column value is correct as Decimal")
            Assert.AreEqual(GetType(Int32), reader.GetFieldType(1), "Second column is System.String")
            Assert.IsFalse(reader.IsDBNull(1), "Second column is not null")

            Assert.AreEqual("1/2/2003 2:34:56 AM", reader.GetString(2), "Third column value is correct")
            Assert.AreEqual(DateTime.Parse("1/2/2003 2:34:56"), reader.GetDateTime(2), "Third column value is correct as DateTime")
            Assert.AreEqual(GetType(DateTime), reader.GetFieldType(2), "Third column is System.String")
            Assert.IsFalse(reader.IsDBNull(2), "Third column is not null")

            Assert.AreEqual("1.23", reader.GetString(3), "Fourth column value is correct")
            Assert.AreEqual(Convert.ToSingle(1.23), reader.GetFloat(3), "Fourth column value is correct as Float/Single")
            Assert.AreEqual(Convert.ToDouble(1.23), reader.GetDouble(3), "Fourth column value is correct as Double")
            Assert.AreEqual(Convert.ToDecimal(1.23), reader.GetDecimal(3), "Fourth column value is correct as Decimal")
            Assert.AreEqual(GetType(Decimal), reader.GetFieldType(3), "Fourth column is System.String")
            Assert.IsFalse(reader.IsDBNull(3), "Fourth column is not null")

            Assert.AreEqual("11111111-2222-3333-4444-555555555555", reader.GetString(4), "Fifth column value is correct")
            Assert.AreEqual(New Guid("11111111-2222-3333-4444-555555555555"), reader.GetGuid(4), "Fifth column value is correct as Decimal")
            Assert.AreEqual(GetType(Guid), reader.GetFieldType(4), "Fifth column is System.String")
            Assert.IsFalse(reader.IsDBNull(4), "Fifth column is not null")

            Assert.AreEqual("True", reader.GetString(5), "Sixth column value is correct")
            Assert.IsTrue(reader.GetBoolean(5), "Sixth column value is correct as Boolean")
            Assert.AreEqual(GetType(Boolean), reader.GetFieldType(5), "Sixth column is System.Boolean")
            Assert.IsFalse(reader.IsDBNull(5), "Sixth column is not null")


            Assert.IsTrue(reader.Read, "Read second line")
            Assert.IsEmpty(reader.GetString(0), "First column value is empty")
            Assert.AreEqual(GetType(System.DBNull), reader.GetFieldType(0), "First column is System.DBNull")
            Assert.IsTrue(reader.IsDBNull(0), "First column is null")

            Assert.AreEqual("23", reader.GetString(1), "Second column value is correct")
            Assert.AreEqual(Convert.ToInt16(23), reader.GetInt16(1), "Second column value is correct as Int16")
            Assert.AreEqual(Convert.ToInt32(23), reader.GetInt32(1), "Second column value is correct as Int32")
            Assert.AreEqual(Convert.ToInt64(23), reader.GetInt64(1), "Second column value is correct as Int64")
            Assert.AreEqual(Convert.ToSingle(23), reader.GetFloat(1), "Second column value is correct as Float/Single")
            Assert.AreEqual(Convert.ToDouble(23), reader.GetDouble(1), "Second column value is correct as Double")
            Assert.AreEqual(Convert.ToDecimal(23), reader.GetDecimal(1), "Second column value is correct as Decimal")
            Assert.AreEqual(GetType(Int32), reader.GetFieldType(1), "Second column is System.String")
            Assert.IsFalse(reader.IsDBNull(1), "Second column is not null")

            Assert.AreEqual("2/3/2004 12:00:00 AM", reader.GetString(2), "Third column value is correct")
            Assert.AreEqual(DateTime.Parse("2/3/2004"), reader.GetDateTime(2), "Third column value is correct as DateTime")
            Assert.AreEqual(GetType(DateTime), reader.GetFieldType(2), "Third column is System.String")
            Assert.IsFalse(reader.IsDBNull(2), "Third column is not null")

            Assert.AreEqual("1.342", reader.GetString(3), "Fourth column value is correct")
            Assert.AreEqual(Convert.ToSingle(1.342), reader.GetFloat(3), "Fourth column value is correct as Float/Single")
            Assert.AreEqual(Convert.ToDouble(1.342), reader.GetDouble(3), "Fourth column value is correct as Double")
            Assert.AreEqual(Convert.ToDecimal(1.342), reader.GetDecimal(3), "Fourth column value is correct as Decimal")
            Assert.AreEqual(GetType(Decimal), reader.GetFieldType(3), "Fourth column is System.String")
            Assert.IsFalse(reader.IsDBNull(3), "Fourth column is not null")

            Assert.IsEmpty(reader.GetString(4), "Fifth column value is empty")
            Assert.AreEqual(GetType(System.DBNull), reader.GetFieldType(4), "Fifth column is System.DBNull")
            Assert.IsTrue(reader.IsDBNull(4), "Fifth column is null")

            Assert.AreEqual("False", reader.GetString(5), "Sixth column value is correct")
            Assert.IsFalse(reader.GetBoolean(5), "Sixth column value is correct as Boolean")
            Assert.AreEqual(GetType(Boolean), reader.GetFieldType(5), "Sixth column is System.Boolean")
            Assert.IsFalse(reader.IsDBNull(5), "Sixth column is not null")

            Assert.IsFalse(reader.IsClosed, "Reader is still open until last read")
            Assert.IsFalse(reader.Read, "End of file")
            Assert.IsTrue(reader.IsClosed, "Reader is closed after final read")

            reader.Close()
        End Using
    End Sub

    Private Sub AreEqual(ByVal expected As DataColumn, ByVal actual As DataColumn)
        Assert.AreEqual(expected, actual, "Columns objects are equal")
        Assert.AreEqual(expected.ColumnName, actual.ColumnName, "Column names are equal")
        Assert.AreEqual(expected.DataType, actual.DataType, "Column data types are equal")
        Assert.AreEqual(expected.AllowDBNull, actual.AllowDBNull, "Column nulls are equal")
    End Sub

End Class

