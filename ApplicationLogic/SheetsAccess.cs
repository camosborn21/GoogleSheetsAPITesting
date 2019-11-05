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
		const string outputFilePath = "//files.biola.edu/LibraryCircAccess/Software/Applicants/data/imports/";
		const string ApplicantsSheetId = "1gf8fqSEYNCfr28aU_fHTvO8IED1C77MwhVmSbpoBF6c";
		const string TestingSheetId = "1G1q3htVrxga_FEO-UvQ0BYFAli64WIBbonFym7VS0MA";
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

			//TestAnswerKey();
			await GetTestData(service);
			//await GetApplications(service);
			
		}

		private void TestAnswerKey()
		{
			Console.WriteLine("Loading Answer Key");
			string answerkeyFile = outputFilePath + "answerkey.json";
			JObject answerKey = JObject.Parse(System.IO.File.ReadAllText(answerkeyFile));

			//answerKey
			Console.WriteLine(answerKey["answers"]["name_test"]["name_1"]);


		}


		private async Task GetTestData(SheetsService service)
		{


			Console.WriteLine("Getting test data from GoogleSheets");
			//A:QN for pre-grading
			var sheetdata = await service.Spreadsheets.Values.Get(TestingSheetId, "A1:QN").ExecuteAsync();
			JArray data = JArray.FromObject(sheetdata.Values);
			
			//Row 1 for headers
			//Row 2 for Answer key
			//		Column #s 0 indexed
			//	G:EZ Name Test (6:155)
			//	FA:KT Number Test (156:305)
			//  KU:QN Alphabetizing Test (306:455)
			//Row 3+ for Data

			string testdataFile = outputFilePath + "testdataimports.json";


			//Load Answer Key
			Console.WriteLine("Loading Answer Key");
			string answerkeyFile = outputFilePath + "answerkey.json";
			JObject answerKey = JObject.Parse(System.IO.File.ReadAllText(answerkeyFile));



			Console.WriteLine("Formatting JSON data");
			JArray result = new JArray();
			JArray headers = data[0].ToObject<JArray>();

			for (var i = 2; i < data.Count(); i++)
			{
				//get row data
				JArray row = data[i].ToObject<JArray>();
				JObject test = new JObject();

				test["id"] = row[0].ToString();


				//create metadata subobject and add to test
				JObject metadata = new JObject();
				for (var meta = 1; meta < 6; meta++)
				{
					metadata[NormalizeJsonHeader( headers[meta].ToString())] = row[meta].ToString();
				}
				test["metadata"] = metadata;

				JObject grades = new JObject();
				int namesGrade = 0;
				int numbersGrade = 0;
				int alphaGrade = 0;

				JObject answers = new JObject();

				//create name answers subobject and add to test
				JObject nameAnswers = new JObject();
				for(var names = 6; names < 156; names++)
				{
					string key = NormalizeJsonHeader(headers[names].ToString());
					nameAnswers[key] = row[names].ToString();
					if(answerKey["answers"]["name_test"][key].ToString() == row[names].ToString())
					{
						namesGrade += 1;
					}
				}
				answers["name_test"] = nameAnswers;
				grades["names"] = namesGrade;

				//create number answers suboject and add to test
				JObject numberAnswers = new JObject();
				for(var nums = 156; nums < 306; nums++)
				{
					string key = NormalizeJsonHeader(headers[nums].ToString());
					numberAnswers[key] = row[nums].ToString();
					if (answerKey["answers"]["number_test"][key].ToString() == row[nums].ToString())
					{
						numbersGrade += 1;
					}
				}
				answers["number_test"] = numberAnswers;
				grades["numbers"] = numbersGrade;


				//create alpha answers subobject and add to test
				JObject alphaAnswers = new JObject();
				for(var alphas = 306; alphas < 456; alphas++)
				{
					string key = NormalizeJsonHeader(headers[alphas].ToString());
					alphaAnswers[key] = row[alphas].ToString();
					if (answerKey["answers"]["alphabetizing_test"][key].ToString() == row[alphas].ToString())
					{
						alphaGrade += 1;
					}
				}
				answers["alphabetizing_test"] = alphaAnswers;
				grades["alphabetizing"] = alphaGrade;
				int total = namesGrade + numbersGrade + alphaGrade;
				grades["total"] = total;
				int missed = 450 - total;
				grades["missed"] = missed;
				double percent = ((50-missed) / 50);
				grades["percent"] = percent;

				test["answers"] = answers;
				test["grades"] = grades;


				//Add the formatted JSON for this test to the result array
				result.Add(test);
			}


			//write output file
			Console.WriteLine(String.Format("Writing test data to: {0}", testdataFile));
			System.IO.File.WriteAllText(testdataFile, result.ToString());

			//get last row method
			//Console.WriteLine(sheetdata.Values.Count);

			//write data
			//Console.WriteLine(JsonConvert.SerializeObject(sheetdata.Values, Formatting.Indented));

		}

		private async Task GetApplications(SheetsService service)
		{

			const string applicantsFile = outputFilePath + "applicantimports.json";
			//get sheet data
			Console.WriteLine("Getting Applicant data from GoogleSheets");
			var sheetData = await service.Spreadsheets.Values.Get(ApplicantsSheetId, "A1:BD").ExecuteAsync();

			JArray data = JArray.FromObject(sheetData.Values);
			//JObject data = JObject.FromObject(sheetData.Values,new JsonSerializer());

			//Console.WriteLine(data.ToString());
			Console.WriteLine("Formatting JSON data");
			JArray result = new JArray();
			JArray headers = data[0].ToObject<JArray>();
			var cols = headers.Count();
			JArray row = new JArray();
			for (var i = 1; i < data.Count(); i++)
			{
				row = data[i].ToObject<JArray>();
				JObject obj = new JObject();
				for (var col = 0; col < cols; col++)
				{
					obj[headers[col].ToString()] = row[col].ToString();
				}

				result.Add(obj);
			}

			//write output file
			Console.WriteLine(String.Format("Writing applicant data to: {0}", applicantsFile));
			System.IO.File.WriteAllText(applicantsFile, result.ToString());

		}

		private string NormalizeJsonHeader(string text)
		{
			return text.ToLower().Replace(" ", "_");
		}

	}
}
