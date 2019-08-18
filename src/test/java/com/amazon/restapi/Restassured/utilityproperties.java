package com.amazon.restapi.Restassured;



import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.Properties;

public class utilityproperties {
	

		static Properties properties;
		
		public static void loaddata() throws IOException
		{
			properties= new Properties();
			File file = new File(System.getProperty("user.dir")+"/resources/browsertype.properties");
			FileReader obj = new FileReader(file);
			properties.load(obj);
		}
		
		public String getData(String Data) throws IOException
		{
		   loaddata();
		   String data = properties.getProperty(Data);
		   return data;
		   
		}

	


}
