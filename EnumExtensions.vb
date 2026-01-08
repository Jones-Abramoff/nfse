Imports System.Reflection
Imports System.Xml.Serialization

Public Module EnumExtensions

    <System.Runtime.CompilerServices.Extension>
    Public Function GetXmlEnumValue(Of T)(value As T) As Integer
        Dim fi = GetType(T).GetField(value.ToString())
        Dim attr = CType(fi.GetCustomAttribute(GetType(XmlEnumAttribute)), XmlEnumAttribute)
        Return Integer.Parse(attr.Name)
    End Function

End Module

