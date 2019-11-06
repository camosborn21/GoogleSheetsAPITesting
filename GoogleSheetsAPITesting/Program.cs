using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using ApplicationLogic;
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
			Console.WriteLine("Import Applicant Data and Test Results from Google Sheets");
			Console.WriteLine("===================================");
			try
			{
				new SheetsAccess().Run().Wait();

			}
			catch (AggregateException ex)
			{
				foreach (var e in ex.InnerExceptions)
				{
					Console.WriteLine("Error: " + e.Message);
				}
			}
			Console.WriteLine("Import complete");
			Console.WriteLine("Press any key to continue...");
			Console.ReadKey();
		}

		
	}
}
