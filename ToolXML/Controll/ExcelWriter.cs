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
	internal static class ExcelWriter
	{
		internal static void FillMethodList(this Excel._Worksheet worksheetMethodList, List<Method> methods)
		{
			worksheetMethodList.FillHeaderMethoList();
			var methodListRowStart = 2;
			foreach (var method in methods)
			{
				
				if (!string.IsNullOrWhiteSpace(method.Name.Value))
				{
					worksheetMethodList.Cells[methodListRowStart, ExcelMethodListColumns.MethodName] = method.Name.Value;
				}
				if (!string.IsNullOrWhiteSpace(method.OutputType))
				{
					worksheetMethodList.Cells[methodListRowStart, ExcelMethodListColumns.OutputType] = method.OutputType;
				}
				if (!string.IsNullOrWhiteSpace(method.OutputDes))
				{
					worksheetMethodList.Cells[methodListRowStart, ExcelMethodListColumns.OutputDes] = method.OutputDes;
				}
				var InputTypeRowStart = methodListRowStart;
				if(method.Inputs != null && method.Inputs.ListInput.Count > 0)
				{
					foreach (var input in method.Inputs.ListInput )
					{
						if (!string.IsNullOrEmpty(input.Type) )
						{
							worksheetMethodList.Cells[InputTypeRowStart, ExcelMethodListColumns.InputType] = input.Type;
						}
						if (!string.IsNullOrEmpty(input.Name) )
						{
							worksheetMethodList.Cells[InputTypeRowStart, ExcelMethodListColumns.InputName] = input.Name;
						}
						InputTypeRowStart++;
					}
				}
				var AuditLogRowStart = methodListRowStart;
				if (method.AuditLogs != null && method.AuditLogs.ListAuditLog.Count > 0)
				{
					foreach(var auditlog in method.AuditLogs.ListAuditLog)
					{
						if (!string.IsNullOrEmpty(auditlog.Name))
						{
							worksheetMethodList.Cells[AuditLogRowStart, ExcelMethodListColumns.AuditLogName] = auditlog.Name;
						}
						if (!string.IsNullOrEmpty(auditlog.Type))
						{
							worksheetMethodList.Cells[AuditLogRowStart, ExcelMethodListColumns.AuditLogType] = auditlog.Type;
						}
						AuditLogRowStart++;
					}
				}
				var newC = Math.Max(methodListRowStart, Math.Max(InputTypeRowStart, AuditLogRowStart));

				if (methodListRowStart == newC)
				{
					methodListRowStart++;
				}
				else
				{
					methodListRowStart = newC;
				}
			}
			worksheetMethodList.Columns.AutoFit();
		}
		internal static void FillHeaderMethoList(this Excel._Worksheet worksheetMethodListHeader)
		{
			worksheetMethodListHeader.Cells[1, ExcelMethodListColumns.MethodName] = ExcelCells.MethodName;
			worksheetMethodListHeader.Cells[1,ExcelMethodListColumns.OutputType] = ExcelCells.OutputType;
			worksheetMethodListHeader.Cells[1, ExcelMethodListColumns.OutputDes] = ExcelCells.OutputDes;
			worksheetMethodListHeader.Cells[1, ExcelMethodListColumns.InputType] = ExcelCells.InputType;
			worksheetMethodListHeader.Cells[1, ExcelMethodListColumns.InputName] = ExcelCells.InputName;
			worksheetMethodListHeader.Cells[1, ExcelMethodListColumns.AuditLogName] = ExcelCells.AuditLogName;
			worksheetMethodListHeader.Cells[1, ExcelMethodListColumns.AuditLogType] = ExcelCells.AuditLogType;
			worksheetMethodListHeader.Columns.AutoFit();
		}
	}
}
