using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Newtonsoft.Json;

namespace GoogleSheetsAPITesting
{
	class Program
	{
		static void Main(string[] args)
		{
			Console.WriteLine("Sheets API Sample: Get Test Results");
			Console.WriteLine("===================================");
			try
			{
				new Program().Run().Wait();

			}
			catch (AggregateException ex)
			{
				foreach (var e in ex.InnerExceptions)
				{
					Console.WriteLine("Error: " + e.Message);
				}
			}

			Console.WriteLine("Press any key to continue...");
			Console.ReadKey();
		}

		private async Task Run()
		{
			//create credential
			UserCredential credential;
			credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets()
			{
				ClientId = "876944923425-6cdvrf8ls90fpnf0crb9kkmnedd2ojat.apps.googleusercontent.com",
				ClientSecret = "6aS9xgysd8G1lXHfGEQo7XlP"
			}, new[] {SheetsService.Scope.Spreadsheets}, "user", CancellationToken.None);

			var service = new SheetsService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = "Sheets API Test"
			});

			//List Responses
			await ListSheets(service);

		}

		private async Task ListSheets(SheetsService service)
		{
			var sheetdata = await service.Spreadsheets.Values.Get("1G1q3htVrxga_FEO-UvQ0BYFAli64WIBbonFym7VS0MA", "A1:QN")
				.ExecuteAsync();
				
			//get last row method
				Console.WriteLine(sheetdata.Values.Count);

			//write data
				Console.WriteLine( JsonConvert.SerializeObject(sheetdata.Values.First()) );

		}

	}
}
