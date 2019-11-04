using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq.Expressions;
using System.IO;
using System.Threading;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using System.Collections;
using Newtonsoft.Json.Linq;

namespace ApplicationLogic
{
    public class SheetsAccess
    {
	    public async Task Run()
	    {
		    //create credential
		    UserCredential credential;
		    credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets()
		    {
			    ClientId = "876944923425-6cdvrf8ls90fpnf0crb9kkmnedd2ojat.apps.googleusercontent.com",
			    ClientSecret = "6aS9xgysd8G1lXHfGEQo7XlP"
		    }, new[] { SheetsService.Scope.Spreadsheets }, "user", CancellationToken.None);

		    var service = new SheetsService(new BaseClientService.Initializer()
		    {
			    HttpClientInitializer = credential,
			    ApplicationName = "Sheets API Test"
		    });

			//List Responses
			//await ListSheets(service);
			await GetApplications(service);
	    }

	    private async Task ListSheets(SheetsService service)
	    {
		    var sheetdata = await service.Spreadsheets.Values.Get("1G1q3htVrxga_FEO-UvQ0BYFAli64WIBbonFym7VS0MA", "A1:QN")
			    .ExecuteAsync();

		    //get last row method
		    Console.WriteLine(sheetdata.Values.Count);

		    //write data
		    Console.WriteLine(JsonConvert.SerializeObject(sheetdata.Values,Formatting.Indented));

	    }

		private async Task GetApplications(SheetsService service)
		{
			//get sheet data
			var sheetData = await service.Spreadsheets.Values.Get("1gf8fqSEYNCfr28aU_fHTvO8IED1C77MwhVmSbpoBF6c", "A1:BD").ExecuteAsync();

			JArray data = JArray.FromObject(sheetData.Values);
			//JObject data = JObject.FromObject(sheetData.Values,new JsonSerializer());

			//Console.WriteLine(data.ToString());

			JArray result = new JArray();
			JArray headers = data[0].ToObject<JArray>();
			var cols = headers.Count();
			JArray row = new JArray();
			for (var i = 1;  i< data.Count(); i++)
			{
				row = data[i].ToObject<JArray>();
				JObject obj = new JObject();
				for (var col = 0; col < cols; col++)
				{
					obj[headers[col].ToString()] = row[col].ToString();
				}

				result.Add(obj);
			}

			//Console.WriteLine(result.ToString());



			//set up
			/*
			var result = new IList[0];
			var headers = sheetData.Values[0];
			var cols = headers.Count();
			IList row = new IList[0];

			//Console.WriteLine(cols);

			for(var i = 1; i < sheetData.Values.Count(); i++)
			{
				row = sheetData.Values[i];
				var obj = new object { };
				for(var col = 0; col < cols; col++)
				{
					obj[headers[col]] = row[col];
				}
			}
			*/
			//Console.WriteLine(sheetData.Values) ;

			//data to object
			//var serializedData = JsonConvert.SerializeObject(sheetData.Values,Formatting.Indented);

			//write data
			//Console.WriteLine(serializedData);



		}

	}
}
