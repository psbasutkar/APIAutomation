package api.test;

import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import com.github.javafaker.Faker;

import api.endpoints.UserEndPoints;
import api.payload.User;
import io.restassured.response.Response;
import junit.framework.Assert;

public class UserTests {

	Faker faker; 
	User userPayload;
	
	
	@BeforeClass
	public void setupData()
	{
		faker = new Faker();
		userPayload = new User();
		userPayload.setId(faker.idNumber().hashCode());
		userPayload.setEmail(faker.internet().safeEmailAddress());
		userPayload.setFirstName(faker.name().firstName());
		userPayload.setLastName(faker.name().lastName());
		userPayload.setPassword(faker.internet().password());
		userPayload.setPhone(faker.phoneNumber().cellPhone());
		userPayload.setUsername(faker.name().username());
				
		
	}
	
	@Test(priority=1)
	void testPostUser()
	{
		Response res = UserEndPoints.createUser(userPayload);
		res.then().log().all();
		Assert.assertEquals(res.getStatusCode(), 200);
	}
	
	
	@Test(priority=2)
	void testGetUserByName()
	{
		Response res = UserEndPoints.readUser(userPayload.getUsername());
		res.then().log().all();
		Assert.assertEquals(res.getStatusCode(), 200);
	}
	

	@Test(priority=3)
	void testUpdateUserByName()
	
	{
		
		userPayload.setFirstName(faker.name().firstName());
		userPayload.setLastName(faker.name().lastName());
				Response res = UserEndPoints.updateUser(this.userPayload.getUsername(), userPayload);
		res.then().log().all();
		Assert.assertEquals(res.getStatusCode(), 200);
		Assert.assertEquals(res.body().jsonPath().getString("firstname"), 200);
		
	}
	
	
}
