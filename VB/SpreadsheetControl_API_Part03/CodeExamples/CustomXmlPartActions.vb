Imports DevExpress.Spreadsheet
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Xml

Namespace SpreadsheetControl_API_Part03.CodeExamples
    Friend Class CustomXmlPartActions
        Private Shared Sub StoreCustomXmlPart(ByVal workbook As IWorkbook)
'            #Region "#StoreCustomXmlPart"
            workbook.Worksheets(0).Cells("A1").Value = "Custom Xml Test"

            ' Add an empty custom XML part.
            Dim part As ICustomXmlPart = workbook.CustomXmlParts.Add()
            Dim elem As XmlElement = part.CustomXmlPartDocument.CreateElement("Person")
            elem.InnerText = "Stephen Edwards"
            part.CustomXmlPartDocument.AppendChild(elem)

            ' Add an XML part created from string.
            Dim xmlString As String = "<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf & _
"                                    <whitepaper>" & ControlChars.CrLf & _
"                                       <contact>" & ControlChars.CrLf & _
"                                          <firstname>Roger</firstname>" & ControlChars.CrLf & _
"                                          <lastname>Edwards</lastname>" & ControlChars.CrLf & _
"                                          <phone>832-433-0025</phone>" & ControlChars.CrLf & _
"                                          <address>1657 Wines Lane Houston, TX 77099</address>" & ControlChars.CrLf & _
"                                       </contact>" & ControlChars.CrLf & _
"                                       <date>2016-05-18</date>" & ControlChars.CrLf & _
"                                    </whitepaper>"
            workbook.CustomXmlParts.Add(xmlString)

            ' Add an XML part loaded from a file.
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load("Documents\fishes.xml")
            workbook.CustomXmlParts.Add(xmlDoc)
            workbook.SaveDocument("Documents\CustomXmlTest.xlsx")
            System.IO.File.Copy("Documents\CustomXmlTest.xlsx", "Documents\CustomXmlTest.xlsx.zip", True)
            System.Diagnostics.Process.Start("Documents\CustomXmlTest.xlsx.zip")
'            #End Region ' #StoreCustomXmlPart
        End Sub

        Private Shared Sub ObtainCustomXmlPart(ByVal workbook As IWorkbook)
'            #Region "#ObtainCustomXmlPart"
            workbook.LoadDocument("Documents\CustomXml.xlsx")
            Dim xmlDoc As XmlDocument = workbook.CustomXmlParts(0).CustomXmlPartDocument
            Dim nsmgr As New XmlNamespaceManager(xmlDoc.NameTable)
            Dim xPathString As String = "//Fish[Category='Cod']/ScientificClassification/Reference"
            Dim xmlNode As XmlNode = xmlDoc.DocumentElement.SelectSingleNode(xPathString, nsmgr)
            Dim hLink As String = xmlNode.InnerText
            workbook.Worksheets(0).Hyperlinks.Add(workbook.Worksheets(0).Cells("A2"), hLink, True)
'            #End Region ' #ObtainCustomXmlPart
        End Sub

        Private Shared Sub ModifyCustomXmlPart(ByVal workbook As IWorkbook)
'            #Region "#ModifyCustomXmlPart"
            workbook.LoadDocument("Documents\CustomXml.xlsx")
            Dim xmlDoc As XmlDocument = workbook.CustomXmlParts(1).CustomXmlPartDocument
            Dim nsmgr As New XmlNamespaceManager(xmlDoc.NameTable)
            Dim xPathString As String = "//whitepaper/contact[firstname='Roger']/firstname"
            Dim xmlNodes As XmlNodeList = xmlDoc.DocumentElement.SelectNodes(xPathString, nsmgr)
            For Each node As XmlNode In xmlNodes
                node.InnerText = "Stephen"
            Next node
            workbook.SaveDocument("Documents\CustomXmlRogerStephen.xlsx")
            workbook.Worksheets(0).Cells("A2").Value = xmlDoc.FirstChild.FirstChild.FirstChild.InnerText
'            #End Region ' #ModifyCustomXmlPart
        End Sub
    End Class
End Namespace
