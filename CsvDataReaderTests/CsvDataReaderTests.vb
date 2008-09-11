Imports System.IO
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

End Class

