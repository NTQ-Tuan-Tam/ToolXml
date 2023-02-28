using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Reflection;
using System.IO;
using System.Windows.Forms;
using System.Xml.Serialization;
using System.Xml;
using ToolXML.Model;
using ToolXML.Controll;
using ToolXML.Constants;

namespace ToolXML
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        internal void SaveToXML()
        {
            var mehtodList = (Excel.Worksheet)Application.ActiveWorkbook.Sheets[1];
			var MethodListsd = new MethodLists();
			MethodListsd.LoadMethodList(mehtodList);
			var xmlSerializer = new XmlSerializer(typeof(MethodLists));
			using (StreamWriter streamWriter = new StreamWriter("C:\\Users\\tuan.vu2\\Desktop\\tuantam1.xml", false, System.Text.Encoding.UTF8))
			{
				xmlSerializer.Serialize(streamWriter, MethodListsd , new XmlSerializerNamespaces(new[] { XmlQualifiedName.Empty }));
			}
			MessageBox.Show("Done!");
		}
        internal void LoadToXml()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "XML files (*.xml)|*.xml",
                FilterIndex = 0,
                DefaultExt = "xml",
                RestoreDirectory = true

			};
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                using (TextReader reader = new StringReader(File.ReadAllText(openFileDialog.FileName)))
                {
					var generateMethod = reader.SerializerObject<MethodLists>();
					// Close current workbook
					if (Application.ActiveWorkbook != null)
					{
						Application.ActiveWorkbook.Close();
					}

					// Add new workbook
					Excel._Workbook workbook = Application.Workbooks.Add(Type.Missing);
					Application.Visible = true;

					// Rename current default sheet to 設定
					workbook.ActiveSheet.Name = ExcelSheet.MethodList;

					Excel._Worksheet worksheetMethod = workbook.ActiveSheet;
					worksheetMethod.FillMethodList(generateMethod);
				}
            }
		}


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
