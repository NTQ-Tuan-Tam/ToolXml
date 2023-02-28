using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ToolXML.Model
{
	public class Name
	{
		[XmlAttribute(AttributeName ="Value")]
		public string Value { get; set; }
	}
}
