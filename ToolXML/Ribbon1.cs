using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ToolXML
{
	public partial class Ribbon1
	{
		private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void button1_Click(object sender, RibbonControlEventArgs e)
		{

		}

		private void SaveToXMl_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.SaveToXML();
		}

		private void ReadFromXml_Click(object sender, RibbonControlEventArgs e)
		{

		}
	}
}
