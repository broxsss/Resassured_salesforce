package com.amazon.restapi.Restassured;

import static io.restassured.RestAssured.given;

import java.io.IOException;
import java.text.ParseException;

import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.HttpClientBuilder;
import org.apache.http.util.EntityUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.Test;

import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;

import io.restassured.http.ContentType;
import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;

public class crudrequests {
	
	private static utilityproperties ul = new utilityproperties();
	private static String accesstoken=null;

	private static String authenticateUser() throws IOException {

		//Getting the details from properties file
		String url= ul.getData("tokenurl");
		String password=ul.getData("password");
		String username=ul.getData("username");
		String client_secret=ul.getData("client_secret");
		String client_id=ul.getData("client_id");


		//Getting the access Token
		String response =
				given()
				.header("Content-Type", "application/x-www-form-urlencoded").formParam("grant_type", "password")
				.formParam("client_id", client_id)
				.formParam("client_secret", client_secret)
				.formParam("username", username)
				.formParam("password", password)
				.when()
				.post(url)
				.asString();

		//System.out.println(response);
		JsonPath jsonPath = new JsonPath(response);
		String accessToken = jsonPath.getString("access_token");

		return accessToken;
	}
	
	public static String createRecords(String token,String name) throws IOException {
		System.out.println("Insert");
		String id ="";
		String baseUri = ul.getData("testurl");
		String url = baseUri + "/services/data/v42.0/sobjects/Account/";
		try {
			
			Response res = given()
					.header("Content-Type", "application/x-www-form-urlencoded")
					.auth().oauth2(token)
					.contentType(ContentType.JSON)
					.accept(ContentType.JSON)
					.body("{\n" + 
							"  \"Name\" : \""+name+"\"\n" + 
							"}")
					.when()
					.post(url);
		            
			if(res.getStatusCode()==201)
			{
				System.out.println("Record is create with code :::"+res.getStatusCode());
				
			}
			System.out.println(res.asString());
			
			JsonParser jparser = new JsonParser();
			JsonElement jElement = jparser.parse(res.asString());
			id = jElement.getAsJsonObject().get("id").getAsString();
			
		}
		catch(Exception ex)
		{
			ex.printStackTrace();
		}
		
		return id;
	}
	
	
	public static void updateRecords(String token,String id,String place) throws IOException {
		System.out.println("Update");
		String baseUri = ul.getData("testurl");
		String url = baseUri + "/services/data/v42.0/sobjects/Account/";
		try {
			
			Response res = given()
					.header("Content-Type", "application/x-www-form-urlencoded")
					.auth().oauth2(token)
					.contentType(ContentType.JSON)
					.accept(ContentType.JSON)
					.body("{\n" + 
							"    \"BillingCity\" : \""+place+"\"\n" + 
							"}")
					.when()
					.patch(url+id);
		            
			if(res.getStatusCode()==204)
			{
				System.out.println("Record is update :::"+res.getStatusCode());
			}
			System.out.println(res.asString());
			
		}
		catch(Exception ex)
		{
			ex.printStackTrace();
		}	
	}
	
	public static void deleteRecords(String token,String id) throws IOException {
		System.out.println("delete");
		String baseUri = ul.getData("testurl");
		String url = baseUri + "/services/data/v42.0/sobjects/Account/";
		try {
			
			Response res = given()
					.header("Content-Type", "application/x-www-form-urlencoded")
					.auth().oauth2(token)
					.contentType(ContentType.JSON)
					.accept(ContentType.JSON)
					.when()
					.delete(url+id);
		            
			if(res.getStatusCode()==204)
			{
				System.out.println("Record is delete :::"+res.getStatusCode());
			}
			System.out.println(res.asString());
			
		}
		catch(Exception ex)
		{
			ex.printStackTrace();
		}	
	}
	
	@Test
	public static void check() throws IOException, EncryptedDocumentException, InvalidFormatException, ParseException
	{
		accesstoken =authenticateUser();
		System.out.println(accesstoken);
		String id1 = createRecords(accesstoken,"akshay saini 5");
		System.out.println();
		updateRecords(accesstoken,id1,"bijnore");
		System.out.println();
		String id2 = createRecords(accesstoken,"akshay saini 6");
		System.out.println();
		deleteRecords(accesstoken,id2);
		
	}
}
