' NOTE: You can use the "Rename" command on the context menu to change the interface name "IService1" in both code and config file together.
<ServiceContract()>
Public Interface IService1

    <OperationContract()>
    Function SelectSql(ByVal sqlstr As String, ByVal FieldNo As String) As CompositeType

    <OperationContract()>
    Function SelectScalar(ByVal sqlstr As String) As Integer

    <OperationContract()>
    Function UpdateSql(ByVal sqlstr As String) As Integer

    <OperationContract()>
    Function DeleteSql(ByVal sqlstr As String) As Integer

    <OperationContract()>
    Function InsertSql(ByVal sqlstr As String) As Integer

    <OperationContract()>
    Function TableExists(ByVal TableName As String) As Boolean

    <OperationContract()>
    Function FindNextID(ByVal TableName As String, ByVal par As String) As Integer

    <OperationContract()>
    Function CloneTableRow(ByVal TableName As String, ByVal ID As Integer, ByRef NewID As Integer, ByVal ArrayParameter() As String, ByVal InitialValue() As String) As Boolean

    <OperationContract()>
    Function CloneTableRowWithoutInit(ByVal TableName As String, ByVal ID As Integer, ByRef NewID As Integer) As Boolean

    <OperationContract()>
    Function GetDataInZIPFile(ByVal sqlstr As String, ByVal FieldNo As String, ByVal TextFile As String) As Boolean

    ' TODO: Add your service operations here

End Interface

' Use a data contract as illustrated in the sample below to add composite types to service operations.
' You can add XSD files into the project. After building the project, you can directly use the data types defined there, with the namespace "WcfServiceLibrary1.ContractType".

<DataContract()>
Public Class CompositeType

    <DataMember()>
    Public Property StrResults() As String()

End Class
© 2022 GitHub, Inc.
Terms
Privacy
Security
Status
Docs
Contact GitHub
Pricing
API
Training
Blog
