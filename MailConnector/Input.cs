using System;
using System.Xml;
using System.Reflection;
using System.IO;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Globalization;

namespace MailConnector
{
    public class Input
    {
        private static string pathXML = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName +
            @System.Configuration.ConfigurationManager.AppSettings["Local_XMLPath"];
        private static string pathXLSX = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName +
            @System.Configuration.ConfigurationManager.AppSettings["Local_XLSXPath"];
        // Get the column position in excel
        private static int[] colNum = {
        Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Program_Col1"]),
        Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Program_Col2"]),
        Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Program_Col3"]),
        Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Program_Col4"]),
        Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Program_Col5"]),
        Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Program_Col6"]),
        Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Program_Col7"])
    };

        //
        private static List<string> _action;
        private static List<string> _objet;
        private static List<string> _expediteur;
        private static List<string> _choix_types_destinataire;
        private static List<string> _destinataire;
        //
        private static List<string> _types_actions;
        private static List<string> _types_destinataire;
        // Excel
        public static List<string> ReadExcel(int col)
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(pathXLSX);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            List<string> list = new List<string>();
            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= max(xlRange); i++)
            {
                string str;
                //write the value to the console
                if (Convert.ToString(xlRange.Cells[i, col].Value2) != null)
                {
                    str = Convert.ToString(xlRange.Cells[i, col].Value2);
                    // only normalize the column title and the 2 drop down menus arrays 
                    // as well as their corresponding selection in the main array
                    if (col == colNum[0] || i == 1 || col == colNum[3] || col == colNum[5] || col == colNum[6])

                    {
                        str = RemoveDiacritics(str);
                        str = str.ToLower();
                        str = str.Replace(" ", "_");
                        // Replacement or the '<' and '>' characters to avoid XML misinterpretation
                        str = str.Replace("<", "&lt");
                        str = str.Replace(">", "&gt");
                    }
                    list.Add(str);
                }
                
                else if(xlRange.Cells[i, col].Locked == false)
                {
                    str = "";
                    list.Add(str);
                }
                //Console.WriteLine(str);
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return list;
        }
        public static void WriteXML()
        {
            Encoding utf8noBOM = new UTF8Encoding(false);
            var settings = new XmlWriterSettings
            {
                Indent = true, 
                Encoding = utf8noBOM, 
            };

            var sb = new StringBuilder();
            using (var writer = XmlWriter.Create(pathXML, settings))
            {
                // Read the excel file and setup the XML parameters
                _action = ReadExcel(colNum[0]);
                _objet = ReadExcel(colNum[1]);
                _expediteur = ReadExcel(colNum[2]);
                _choix_types_destinataire = ReadExcel(colNum[3]);
                _destinataire = ReadExcel(colNum[4]);
                //
                _types_actions = ReadExcel(colNum[5]);
                _types_destinataire = ReadExcel(colNum[6]);

                writer.WriteStartDocument();
                const string xsiNamespace = "http://www.w3.org/2001/XMLSchema-instance";
                const string schemaLocation = "../data-set.xsd";

                writer.WriteStartElement("root");
                writer.WriteAttributeString("xmlns", "xsi", null, xsiNamespace);
                writer.WriteAttributeString("xsi", "noNamespaceSchemaLocation", null, schemaLocation);

                // types de destinataires 
                writer.WriteComment(" types de destinataires ");
                writer.WriteStartElement("types_destinataires");

                for (int i = 1; i < _types_destinataire.Count; ++i)
                {
                    writer.WriteStartElement("destinataire");
                    writer.WriteAttributeString("num", _types_destinataire[i]);
                    writer.WriteString(_types_destinataire[i]);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();

                // types d'actions
                writer.WriteComment(" types d'actions ");
                writer.WriteStartElement("types_actions");

                for (int i = 1; i < _types_actions.Count; ++i)
                {
                    writer.WriteStartElement("action");
                    writer.WriteAttributeString("num", _types_actions[i]);
                    writer.WriteString(_types_actions[i]);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();

                // records
                writer.WriteComment(" records ");
                writer.WriteStartElement("records");

                for (int i = 1; i < _action.Count; ++i)
                {
                    if(_action[i] != "")
                        writer.WriteStartElement("record");

                    // - action
                    if (_action.Count > 1 && _action[i] != "")
                    {
                        writer.WriteStartElement(_action[0]);
                        writer.WriteAttributeString(_types_actions[0], _action[i]);
                        writer.WriteEndElement();
                    }
                    // - expediteur
                    if (_expediteur.Count > 1 && _expediteur[i] != "")
                    {
                        writer.WriteStartElement(_expediteur[0]);
                        writer.WriteString(_expediteur[i]);
                        writer.WriteEndElement();
                    }
                    // - objet
                    if (_objet.Count > 1 && _objet[i] != "")
                    {
                        writer.WriteStartElement(_objet[0]);
                        writer.WriteString(_objet[i]);
                        writer.WriteEndElement();
                    }
                    // - destinataire
                    if (_destinataire.Count > 1 && _destinataire[i] != "")
                    {
                        writer.WriteStartElement(_destinataire[0]);
                        writer.WriteAttributeString(_choix_types_destinataire[0], _choix_types_destinataire[i]);
                        writer.WriteString(_destinataire[i]);
                        writer.WriteEndElement();
                    }
                    if (_action[i] != "")
                        writer.WriteEndElement();
                }
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }
        public static void ReadFile()
        {
            // If there isn't any XLSX path specified or if the path isn't linking to any excel file
            // the XML is written, we generate it from the XLSX file
            // otherwise, we read it
            if (pathXLSX != "" || !File.Exists(pathXLSX))
                WriteXML();

            Output.WriteInFile("[OK]", "Lecture du document XML"+ pathXML);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(pathXML);
            XmlNodeList records = xmlDoc.SelectNodes("//root/records/record");
            foreach (XmlNode record in records)
            {
                Output.WriteInFile("====", " TRAITEMENT DE LA LIGNE ====");
                Output.WriteInFile("[INFO]","Lecture du noeud");
                // Read each record
                XmlNode action = record.SelectSingleNode("action");
                if ((action != null))
                    Output.WriteInFile("[INFO]", action.Attributes["type_action"]?.InnerText);
                else
                    Output.WriteInFile("[INFO]", "Pas d'action");
                XmlNode objet = record.SelectSingleNode("objet");
                if ((objet != null))
                    Output.WriteInFile("[INFO]", objet.InnerText);
                else
                    Output.WriteInFile("[INFO]", "Pas d'objet");

                XmlNode expediteur = record.SelectSingleNode("expediteur");
                if ((expediteur != null))
                    Output.WriteInFile("[INFO]", expediteur.InnerText);
                else
                    Output.WriteInFile("[INFO]", "Pas d'expéditeur");

                XmlNode destinataire = record.SelectSingleNode("destinataire");
                if ((destinataire != null))
                {
                    Output.WriteInFile("[INFO]", destinataire.Attributes["type_destinataire"]?.InnerText);
                    Output.WriteInFile("[INFO]", destinataire.InnerText);
                }
                else
                    Output.WriteInFile("[INFO]", "Pas de destinataire");

                // Interpret the data
                //if multiple actions
                string[] actionList;
                actionList = action.Attributes["type_action"]?.InnerText.Split(',');
                foreach(string act in actionList)
                {
                    InterpretData(act ?? "",
                    objet?.InnerText ?? "",
                    expediteur?.InnerText ?? "",
                    destinataire?.Attributes["type_destinataire"]?.InnerText ?? "",
                    destinataire?.InnerText ?? "");
                } 
            }
        }
        public static void InterpretData(
            string action = "", 
            string subject = "", 
            string sender = "", 
            string recipientType = "", 
            string recipient = "")
        {
            Output.WriteInFile("[OK]", "Interpretation du noeud");
            // Use reflection to call a function from its string name
            ExchangerClient p = new ExchangerClient(sender, subject, recipient, recipientType);

            // Verify if the corresponding emails exist
            if (!p.isEmpty())
            {
                // Send to the correct method
                Type t = p.GetType();
                MethodInfo mi = t.GetMethod(action, BindingFlags.Public | BindingFlags.Instance);
                mi.Invoke(p, null);
            }
            else
            {
                Output.WriteInFile("[WAR]", "Pas d'email correspondant à cette description");
            }
        }
        public static string RemoveDiacritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalizedString) 
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }
            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
        }
        public static int max(Excel.Range xlRange)
        {
            int max = 0;
            do
            {
                ++max;
            } while(xlRange.Cells[max, 1].Locked == false
            || xlRange.Cells[max, 7].Locked == false
            || xlRange.Cells[max, 9].Locked == false);
            return max;
        }
    }
}