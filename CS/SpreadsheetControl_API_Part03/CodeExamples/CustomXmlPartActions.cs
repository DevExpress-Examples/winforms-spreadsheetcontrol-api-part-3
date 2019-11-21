using DevExpress.Spreadsheet;
using System.Xml;

namespace SpreadsheetControl_API_Part03.CodeExamples
{
    class CustomXmlPartActions
    {
        static void StoreCustomXmlPart(IWorkbook workbook)
        {
            #region #StoreCustomXmlPart
            workbook.Worksheets[0].Cells["A1"].Value = "Custom Xml Test";

            // Add an empty custom XML part.
            ICustomXmlPart part = workbook.CustomXmlParts.Add();
            XmlElement elem = part.CustomXmlPartDocument.CreateElement("Person");
            elem.InnerText = "Stephen Edwards";
            part.CustomXmlPartDocument.AppendChild(elem);

            // Add an XML part created from string.
            string xmlString = @"<?xml version=""1.0"" encoding=""UTF-8""?>
                                    <whitepaper>
                                       <contact>
                                          <firstname>Roger</firstname>
                                          <lastname>Edwards</lastname>
                                          <phone>832-433-0025</phone>
                                          <address>1657 Wines Lane Houston, TX 77099</address>
                                       </contact>
                                       <date>2016-05-18</date>
                                    </whitepaper>";
            workbook.CustomXmlParts.Add(xmlString);

            // Add an XML part loaded from a file.
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load("Documents\\fishes.xml");
            workbook.CustomXmlParts.Add(xmlDoc);
            workbook.SaveDocument("Documents\\CustomXmlTest.xlsx");
            System.IO.File.Copy("Documents\\CustomXmlTest.xlsx", "Documents\\CustomXmlTest.xlsx.zip", true);
            System.Diagnostics.Process.Start("Documents\\CustomXmlTest.xlsx.zip");
            #endregion #StoreCustomXmlPart
        }

        static void ObtainCustomXmlPart(IWorkbook workbook)
        {
            #region #ObtainCustomXmlPart
            workbook.LoadDocument("Documents\\CustomXml.xlsx");
            XmlDocument xmlDoc = workbook.CustomXmlParts[0].CustomXmlPartDocument;
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            string xPathString = "//Fish[Category='Cod']/ScientificClassification/Reference";
            XmlNode xmlNode = xmlDoc.DocumentElement.SelectSingleNode(xPathString, nsmgr);
            string hLink = xmlNode.InnerText;
            workbook.Worksheets[0].Hyperlinks.Add(workbook.Worksheets[0].Cells["A2"], hLink, true);
            #endregion #ObtainCustomXmlPart
        }

        static void ModifyCustomXmlPart(IWorkbook workbook)
        {
            #region #ModifyCustomXmlPart
            workbook.LoadDocument("Documents\\CustomXml.xlsx");
            XmlDocument xmlDoc = workbook.CustomXmlParts[1].CustomXmlPartDocument;
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            string xPathString = "//whitepaper/contact[firstname='Roger']/firstname";
            XmlNodeList xmlNodes = xmlDoc.DocumentElement.SelectNodes(xPathString, nsmgr);
            foreach (XmlNode node in xmlNodes) node.InnerText = "Stephen";
            workbook.SaveDocument("Documents\\CustomXmlRogerStephen.xlsx");
            workbook.Worksheets[0].Cells["A2"].Value = xmlDoc.FirstChild.FirstChild.FirstChild.InnerText;
            #endregion #ModifyCustomXmlPart
        }
    }
}
