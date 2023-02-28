using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ToolXML.Constants;
using ToolXML.Model;

using Excel = Microsoft.Office.Interop.Excel;


namespace ToolXML.Controll
{
	internal static class GenerateMethodBuilder
	{
		internal static void LoadMethodList(this MethodLists methodList, Excel.Worksheet WorksheetMethodList)
		{
			var MethodNameRowStart = 2;
			Excel.Range range = WorksheetMethodList.UsedRange;
			var startRow = range.Row;
			var endRow = startRow + range.Rows.Count - 1;
			methodList.Method = new List<Method>();
			for (var i = MethodNameRowStart; i <= endRow; i++)
			{
				var Method = new Method();

				Method.Name = new Name();

				Method.Name.Value = WorksheetMethodList.Cells[i, ExcelMethodListColumns.MethodName].Value ?? string.Empty;
				Method.OutputType = WorksheetMethodList.Cells[i, ExcelMethodListColumns.OutputType].Value;
				Method.OutputDes = WorksheetMethodList.Cells[i, ExcelMethodListColumns.OutputDes].Value;

				var listInput = new List<Input>();
				var MethodInputRowStart = i;
				while (!string.IsNullOrWhiteSpace(WorksheetMethodList.Cells[MethodInputRowStart, ExcelMethodListColumns.InputType].Value)
						&& !string.IsNullOrWhiteSpace(WorksheetMethodList.Cells[MethodInputRowStart, ExcelMethodListColumns.InputName].Value)
						&& (WorksheetMethodList.Cells[MethodInputRowStart, ExcelMethodListColumns.MethodName].Value == Method.Name.Value
							|| string.IsNullOrWhiteSpace(WorksheetMethodList.Cells[MethodInputRowStart, ExcelMethodListColumns.MethodName].Value)
							)
						)
				{
					var input = new Input();
					input.Type = WorksheetMethodList.Cells[MethodInputRowStart, ExcelMethodListColumns.InputType].Value;
					input.Name = WorksheetMethodList.Cells[MethodInputRowStart, ExcelMethodListColumns.InputName].Value;
					listInput.Add(input);
					MethodInputRowStart++;
				}
				if (listInput.Count > 0)
				{
					Method.Inputs = new Inputs() { ListInput = listInput };
				}
				methodList.Method.Add(Method);
			}
		}
	}
}
