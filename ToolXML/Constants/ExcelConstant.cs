using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ToolXML.Constants
{
	public class ExcelSheet
	{
		internal const string MethodList = " MethodList";
	}

	internal static class ExcelMethodListColumns
	{
		
		internal const string MethodName = "A";
		internal const string OutputType = "B";
		internal const string OutputDes = "C";
		internal const string InputType = "D";
		internal const string InputName = "E";
		internal const string AuditLogName = "F";
		internal const string AuditLogType = "G";
		//internal const string InputName = "M";
		//internal const string InputDescription = "N";
		//internal const string AuditLogInfoElementName = "O";
		//internal const string AuditLogInfoValue = "P";
		//internal const string AuditLogInfoNoQuot = "Q";
		//internal const string SvcId = "R";
	}
	internal static class ExcelCells
	{
		internal const string MethodName = "MethodName";
		internal const string OutputType = "OutputType";
		internal const string OutputDes = "OutputDes";
		internal const string InputType = "InputType";
		internal const string InputName = "InputName";
		internal const string AuditLogName = "AuditLogName";
		internal const string AuditLogType = "AuditLogType";
	}
}
