using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ToolXML.Model
{
	public class MethodLists
	{
		[XmlElement(ElementName ="Method")]
		public List<Method> Method { get; set; }
	}
}
