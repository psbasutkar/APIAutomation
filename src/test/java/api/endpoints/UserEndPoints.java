package api.endpoints;

import static io.restassured.RestAssured.*;
import static io.restassured.matcher.RestAssuredMatchers.*;
import static org.hamcrest.Matchers.*;

import org.testng.annotations.Test;

import api.payload.User;
import io.restassured.RestAssured;
import io.restassured.http.ContentType;
import io.restassured.response.Response;
import io.restassured.response.ValidatableResponse;


public class UserEndPoints {

	public static Response createUser(User payload)
	{
		    Response response = given()
		    .contentType(ContentType.JSON)
		    .accept(ContentType.JSON)
		    .body(payload)		    
		    .when()
		      .post(Routes.base_url);
		    
		    return  response ;
		      
		   
	}
	

	public static Response readUser(String userName)
	{
		    Response response = given()
		    .pathParam("username", userName)
			    
		    .when()
		      .get(Routes.get_url);
		    //response.jsonPath().getString(userName)
		    return  response ;
	
}
	
	
	public static Response updateUser(String userName , User payload)
	{
		    Response response = given()
		    .contentType(ContentType.JSON)
		    .accept(ContentType.JSON)
		    .pathParam("username", userName)		    
		    .body(payload)		    
		    .when()
		      .put(Routes.update_url);
		    
		    return  response ;
		      
		   
	}
	
	
	public static Response deleteUser(String userName)
	{
		    Response response = given()
		    .contentType(ContentType.JSON)
		    .accept(ContentType.JSON)
		    .pathParam("username", userName)		    
		   
		    .when()
		      .delete(Routes.update_url);
		    
		    return  response ;
		      
		   
	}
	
	
}
