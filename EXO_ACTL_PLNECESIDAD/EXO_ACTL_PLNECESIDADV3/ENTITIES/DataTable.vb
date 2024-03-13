Imports System.Xml.Serialization

<XmlRoot(ElementName:="Cell")>
Public Class Cell
    <XmlElement(ElementName:="ColumnUid")>
    Public Property ColumnUid As String
    <XmlElement(ElementName:="Value")>
    Public Property Value As String
End Class

<XmlRoot(ElementName:="Cells")>
Public Class Cells
    <XmlElement(ElementName:="Cell")>
    Public Property Cell As List(Of Cell)
End Class

<XmlRoot(ElementName:="Row")>
Public Class Row
    <XmlElement(ElementName:="Cells")>
    Public Property Cells As Cells
End Class

<XmlRoot(ElementName:="Rows")>
Public Class Rows
    <XmlElement(ElementName:="Row")>
    Public Property Row As List(Of Row)
End Class

<XmlRoot(ElementName:="DataTable")>
Public Class DataTable
    <XmlElement(ElementName:="Rows")>
    Public Property Rows As Rows
    <XmlAttribute(AttributeName:="Uid")>
    Public Property Uid As String
    <XmlText>
    Public Property Text As String
End Class

Public Class MembersPurchaseRequest
    Public Property ItemCode As String
    Public Property WarehouseCode As String
    Public Property Quantity As Double
    Public Property RequiredDate As Date
    Public Property ShipDate As Date
    Public Property RequiredQuantity As Double
    Public Property Price As Double
    Public Property UnitPrice As Double
    Public Property CostingCode As String
End Class

Public Class MembersTransferRequest
    Public Property ItemCode As String
    Public Property WarehouseCode As String
    Public Property FromWarehouseCode As String
    Public Property Quantity As Double
End Class