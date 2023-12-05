Imports DevExpress.Spreadsheet
Imports System.Xml

Namespace SpreadsheetControl_API_Part03.CodeExamples

    Friend Class CustomXmlPartActions

        Private Shared Sub StoreCustomXmlPart(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
#Region "#StoreCustomXmlPart"
            workbook.Worksheets(CInt((0))).Cells(CStr(("A1"))).Value = "Custom Xml Test"
            ' Add an empty custom XML part.
            Dim part As DevExpress.Spreadsheet.ICustomXmlPart = workbook.CustomXmlParts.Add()
            Dim elem As System.Xml.XmlElement = part.CustomXmlPartDocument.CreateElement("Person")
            elem.InnerText = "Stephen Edwards"
            part.CustomXmlPartDocument.AppendChild(elem)
            ' Add an XML part created from string.
            Dim xmlString As String = "<?xml version=""1.0"" encoding=""UTF-8""?>
                                    <whitepaper>
                                       <contact>
                                          <firstname>Roger</firstname>
                                          <lastname>Edwards</lastname>
                                          <phone>832-433-0025</phone>
                                          <address>1657 Wines Lane Houston, TX 77099</address>
                                       </contact>
                                       <date>2016-05-18</date>
                                    </whitepaper>"
            workbook.CustomXmlParts.Add(xmlString)
            ' Add an XML part loaded from a file.
            Dim xmlDoc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
            xmlDoc.Load("Documents\fishes.xml")
            workbook.CustomXmlParts.Add(xmlDoc)
            workbook.SaveDocument("Documents\CustomXmlTest.xlsx")
            System.IO.File.Copy("Documents\CustomXmlTest.xlsx", "Documents\CustomXmlTest.xlsx.zip", True)
            System.Diagnostics.Process.Start("Documents\CustomXmlTest.xlsx.zip")
#End Region  ' #StoreCustomXmlPart
        End Sub

        Private Shared Sub ObtainCustomXmlPart(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
#Region "#ObtainCustomXmlPart"
            workbook.LoadDocument("Documents\CustomXml.xlsx")
            Dim xmlDoc As System.Xml.XmlDocument = workbook.CustomXmlParts(CInt((0))).CustomXmlPartDocument
            Dim nsmgr As System.Xml.XmlNamespaceManager = New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
            Dim xPathString As String = "//Fish[Category='Cod']/ScientificClassification/Reference"
            Dim xmlNode As System.Xml.XmlNode = xmlDoc.DocumentElement.SelectSingleNode(xPathString, nsmgr)
            Dim hLink As String = xmlNode.InnerText
            workbook.Worksheets(CInt((0))).Hyperlinks.Add(workbook.Worksheets(CInt((0))).Cells("A2"), hLink, True)
#End Region  ' #ObtainCustomXmlPart
        End Sub

        Private Shared Sub ModifyCustomXmlPart(ByVal workbook As DevExpress.Spreadsheet.IWorkbook)
#Region "#ModifyCustomXmlPart"
            workbook.LoadDocument("Documents\CustomXml.xlsx")
            Dim xmlDoc As System.Xml.XmlDocument = workbook.CustomXmlParts(CInt((1))).CustomXmlPartDocument
            Dim nsmgr As System.Xml.XmlNamespaceManager = New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
            Dim xPathString As String = "//whitepaper/contact[firstname='Roger']/firstname"
            Dim xmlNodes As System.Xml.XmlNodeList = xmlDoc.DocumentElement.SelectNodes(xPathString, nsmgr)
            For Each node As System.Xml.XmlNode In xmlNodes
                node.InnerText = "Stephen"
            Next

            workbook.SaveDocument("Documents\CustomXmlRogerStephen.xlsx")
            workbook.Worksheets(CInt((0))).Cells(CStr(("A2"))).Value = xmlDoc.FirstChild.FirstChild.FirstChild.InnerText
#End Region  ' #ModifyCustomXmlPart
        End Sub
    End Class
End Namespace
