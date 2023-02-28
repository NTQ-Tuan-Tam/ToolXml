using System;
using System.Collections.Generic;
using System.IO;

using System.Xml.Serialization;

namespace ToolXML.Controll
{
	internal static  class XmlHelper
	{
		internal static T SerializerObject<T>(this TextReader textReader)
		{
			var serializer = new XmlSerializer(typeof(T));
			return (T)serializer.Deserialize(textReader);
		}
	}
}
