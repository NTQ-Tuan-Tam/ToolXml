using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace ToolXML.Model
{
	public class Method
	{
		[XmlAttribute(AttributeName = "OutputType")]
		public string OutputType { get; set; }

		[XmlAttribute(AttributeName = "OutputDes")]
		public string OutputDes { get; set; }

		[XmlElement(ElementName = "Name")]
		public  Name Name { get; set; }

		[XmlElement(ElementName = "Inputs")]
		public Inputs Inputs { get; set; }
	}
}
