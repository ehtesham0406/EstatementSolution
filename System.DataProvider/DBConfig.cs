using System;
using System.Data;
using Microsoft.Win32;
using System.IO;

namespace System.DataProvider
{
		public sealed class DBConfig
 {
	 // Fields
	 public static readonly DataSet dsConnection = new DataSet();
	 private const string XMLFileName = "DBConfig";

	 // Methods
	 static DBConfig()
	 {
		 string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
		 try
		 {
			 if (!File.Exists(baseDirectory + "DBConfig.xsd"))
			 {
				 throw new Exception("A required file: DBConfig.xsd is not found is the location: " + baseDirectory);
			 }
			 dsConnection.ReadXmlSchema(baseDirectory + "DBConfig.xsd");
			 if (!File.Exists(baseDirectory + "DBConfig.xml"))
			 {
				 throw new Exception("A required file: DBConfig.xml is not found is the location: " + baseDirectory);
			 }
			 dsConnection.ReadXml(baseDirectory + "DBConfig.xml");
		 }
		 catch (Exception exception)
		 {
			 throw new Exception(exception.Message);
		 }
	 }
 }

 

}
