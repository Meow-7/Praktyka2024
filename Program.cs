using System;
using System.Globalization;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;

class Program
{
    static async Task Main(string[] args)
    {
        string mainLink = "";
        string OaLink = "";
        string typeLink = "";

        Console.WriteLine("Type '1' to generate CSV sheets only ");
        Console.WriteLine("Type '2' to generate XLSX sheets only ");
        Console.WriteLine("Type '3' to generate both CSV and XLSX ");
        string generateNumber = Console.ReadLine();
        for (int i = 1; i <= 11; i++)
        {


            switch (i)
            {
                case 1:
                    {
                        mainLink = "https://api.openalex.org/sources/S4210230065";
                        OaLink = "https://api.openalex.org/works?group_by=open_access.is_oa&per_page=200&filter=primary_location.source.id:s4210230065";
                        typeLink = "https://api.openalex.org/works?group_by=type&per_page=200&filter=primary_location.source.id:s4210230065";
                    }
                    break;
                case 2:
                    {
                        mainLink = "https://api.openalex.org/sources/S4210189803";
                        OaLink = "https://api.openalex.org/works?group_by=open_access.is_oa&per_page=200&filter=primary_location.source.id:s4210189803";
                        typeLink = "https://api.openalex.org/works?group_by=type&per_page=200&filter=primary_location.source.id:s4210189803";
                    }
                    break;
                case 3:
                    {
                        mainLink = "https://api.openalex.org/sources/S4210198626";
                        OaLink = "https://api.openalex.org/works?group_by=open_access.is_oa&per_page=200&filter=primary_location.source.id:s4210198626";
                        typeLink = "https://api.openalex.org/works?group_by=type&per_page=200&filter=primary_location.source.id:s4210198626";
                    }
                    break;
                case 4:
                    {
                        mainLink = "https://api.openalex.org/sources/S4210184798";
                        OaLink = "https://api.openalex.org/works?group_by=open_access.is_oa&per_page=200&filter=primary_location.source.id:s4210184798";
                        typeLink = "https://api.openalex.org/works?group_by=type&per_page=200&filter=primary_location.source.id:s4210184798";
                    }
                    break;
                case 5:
                    {
                        mainLink = "https://api.openalex.org/sources/S4210220691";
                        OaLink = "https://api.openalex.org/works?group_by=open_access.is_oa&per_page=200&filter=primary_location.source.id:s4210220691";
                        typeLink = "https://api.openalex.org/works?group_by=type&per_page=200&filter=primary_location.source.id:s4210220691";
                    }
                    break;
                case 6:
                    {
                        mainLink = "https://api.openalex.org/sources/S4210174713";
                        OaLink = "https://api.openalex.org/works?group_by=open_access.is_oa&per_page=200&filter=primary_location.source.id:s4210174713";
                        typeLink = "https://api.openalex.org/works?group_by=type&per_page=200&filter=primary_location.source.id:s4210174713";
                    }
                    break;
                case 7:
                    {
                        mainLink = "https://api.openalex.org/sources/S4210235104";
                        OaLink = "https://api.openalex.org/works?group_by=open_access.is_oa&per_page=200&filter=primary_location.source.id:s4210235104";
                        typeLink = "https://api.openalex.org/works?group_by=type&per_page=200&filter=primary_location.source.id:s4210235104";
                    }
                    break;
                case 8:
                    {
                        mainLink = "https://api.openalex.org/sources/S4210184363";
                        OaLink = "https://api.openalex.org/works?group_by=open_access.is_oa&per_page=200&filter=primary_location.source.id:s4210184363";
                        typeLink = "https://api.openalex.org/works?group_by=type&per_page=200&filter=primary_location.source.id:s4210184363";
                    }
                    break;
                case 9:
                    {
                        mainLink = "https://api.openalex.org/sources/S4210219001";
                        OaLink = "https://api.openalex.org/works?group_by=open_access.is_oa&per_page=200&filter=primary_location.source.id:s4210219001";
                        typeLink = "https://api.openalex.org/works?group_by=type&per_page=200&filter=primary_location.source.id:s4210219001";
                    }
                    break;
                case 10:
                    {
                        mainLink = "https://api.openalex.org/sources/S4210231866";
                        OaLink = "https://api.openalex.org/works?group_by=open_access.is_oa&per_page=200&filter=primary_location.source.id:s4210231866";
                        typeLink = "https://api.openalex.org/works?group_by=type&per_page=200&filter=primary_location.source.id:s4210231866";
                    }
                    break;
                case 11:
                    {
                        mainLink = "https://api.openalex.org/sources/S4210229811";
                        OaLink = "https://api.openalex.org/works?group_by=open_access.is_oa&per_page=200&filter=primary_location.source.id:s4210229811";
                        typeLink = "https://api.openalex.org/works?group_by=type&per_page=200&filter=primary_location.source.id:s4210229811";
                    }
                    break;
                default:
                    Console.WriteLine("Nothing");
                    break;
            }

            // URL для конкретного джерела за ID

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    // Виконання запиту до API
                    HttpResponseMessage response1 = await client.GetAsync(mainLink);
                    HttpResponseMessage response2 = await client.GetAsync(OaLink);
                    HttpResponseMessage response3 = await client.GetAsync(typeLink);
                    response1.EnsureSuccessStatusCode();
                    response2.EnsureSuccessStatusCode();
                    response3.EnsureSuccessStatusCode();

                    string responseBody1 = await response1.Content.ReadAsStringAsync();
                    string responseBody2 = await response2.Content.ReadAsStringAsync();
                    string responseBody3 = await response3.Content.ReadAsStringAsync();

                    // Десеріалізація JSON у об'єкт 
                    var source = JsonSerializer.Deserialize<Source>(responseBody1);
                    var opeAccess = JsonSerializer.Deserialize<OpenAccess>(responseBody2);
                    var type = JsonSerializer.Deserialize<Type>(responseBody3);

                    //!!!!!! Change directories below, but do not change filename !!!!!!!

                    string csvFilePathMainBasic = $"C:\\Users\\Olexii\\Desktop\\jsons\\all_basic.csv";
                    string csvFilePathBasic = $"C:\\Users\\Olexii\\Desktop\\jsons\\{source.display_name}_basic.csv";
                    string csvFilePathYear = $"C:\\Users\\Olexii\\Desktop\\jsons\\{source.display_name}_year.csv";
                    string csvFilePathTopics = $"C:\\Users\\Olexii\\Desktop\\jsons\\{source.display_name}_topics.csv";
                    string csvFilePathTopicsShare = $"C:\\Users\\Olexii\\Desktop\\jsons\\{source.display_name}_topics_share.csv";
                    string csvFilePathXconcepts = $"C:\\Users\\Olexii\\Desktop\\jsons\\{source.display_name}_x_concepts.csv";
                    string csvFilePathOpenAccess = $"C:\\Users\\Olexii\\Desktop\\jsons\\{source.display_name}_open_access.csv";
                    string csvFilePathType = $"C:\\Users\\Olexii\\Desktop\\jsons\\{source.display_name}_type.csv";


                    //CSV
                    if (generateNumber == "1" || generateNumber == "3")
                    {
                        using (StreamWriter writer = new StreamWriter(csvFilePathMainBasic, append: true))
                        {
                            if (i == 1)
                            {
                                writer.WriteLine("ISSN, e-ISSN, Display Name, Host Organization Name, Works Count, Cited by Count," +
                                "2year Mean Citedness, H-index, Cited at Least 10 Times, Is OpenAccess, Is in DOAJ," +
                                " Is Core, Price USD, Country Code, Societies, Alternative Titles, Abbreviated Title, Type, Updated Date");
                            }


                            string societiesJoined = source.societies != null ? string.Join(";", source.societies) : "";
                            string titlesJoined = source.alternate_titles != null ? string.Join(";", source.alternate_titles) : "";



                            writer.WriteLine($"{source.issn[0]},{source.issn[1]},{source.display_name},{source.host_organization_name}," +
                                $"{source.works_count},{source.cited_by_count},{source.summary_stats.two_year_mean_citedness}" +
                                $",{source.summary_stats.h_index},{source.summary_stats.i10_index},{source.is_oa}" +
                                $",{source.is_in_doaj},{source.is_core},{source.apc_usd},{source.country_code},{societiesJoined}" +
                                $",{titlesJoined},{source.abbreviated_title},{source.type},{source.updated_date}");
                        } //1
                        using (StreamWriter writer = new StreamWriter(csvFilePathYear))
                        {
                            writer.WriteLine("Year, Works Count, Cited by Count");

                            foreach (YearStats stat in source.counts_by_year)
                            {
                                writer.WriteLine($"{stat.year}, {stat.works_count}, {stat.cited_by_count}");
                            }
                        }  //2          
                        using (StreamWriter writer = new StreamWriter(csvFilePathTopics))
                        {
                            using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
                            {
                                writer.WriteLine("Topic Name, Count, Subfield, Field, Domain");
                                foreach (Topics top in source.topics)
                                {
                                    csv.WriteField(top.display_name);
                                    csv.WriteField(top.count);
                                    csv.WriteField(top.subfield.display_name);
                                    csv.WriteField(top.field.display_name);
                                    csv.WriteField(top.domain.display_name);
                                    csv.NextRecord();
                                }


                                //csv.WriteRecords(source.topics); 
                            }
                        } //3
                        using (StreamWriter writer = new StreamWriter(csvFilePathTopicsShare))
                        {
                            using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
                            {
                                writer.WriteLine("Topic Name, Weight/Relevance, Subfield, Field, Domain");
                                foreach (Topic_share top in source.topic_share)
                                {
                                    csv.WriteField(top.display_name);
                                    csv.WriteField(top.value);
                                    csv.WriteField(top.subfield.display_name);
                                    csv.WriteField(top.field.display_name);
                                    csv.WriteField(top.domain.display_name);
                                    csv.NextRecord();
                                }


                                //csv.WriteRecords(source.topics); 
                            }
                        } //4
                        using (StreamWriter writer = new StreamWriter(csvFilePathXconcepts))
                        {
                            using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
                            {
                                writer.WriteLine("Concept Name, Level, Score");
                                foreach (x_Concept top in source.x_concepts)
                                {
                                    csv.WriteField(top.display_name);
                                    csv.WriteField(top.level);
                                    csv.WriteField(top.score);
                                    csv.NextRecord();
                                }


                                //csv.WriteRecords(source.topics); 
                            }
                        }   //5
                        using (StreamWriter writer = new StreamWriter(csvFilePathOpenAccess))
                        {
                            using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
                            {
                                writer.WriteLine("Is Open Access, Count");
                                foreach (IsOpen top in opeAccess.group_by)
                                {
                                    csv.WriteField(top.key_display_name);
                                    csv.WriteField(top.count);
                                    csv.NextRecord();
                                }


                                //csv.WriteRecords(source.topics); 
                            }
                        }  //6
                        using (StreamWriter writer = new StreamWriter(csvFilePathType))
                        {
                            using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)))
                            {
                                writer.WriteLine("Article type, Count");
                                foreach (ArticleType top in type.group_by)
                                {
                                    csv.WriteField(top.key_display_name);
                                    csv.WriteField(top.count);
                                    csv.NextRecord();
                                }


                                //csv.WriteRecords(source.topics); 
                            }
                        }  //7
                        using (StreamWriter writer = new StreamWriter(csvFilePathBasic))
                        {
                            writer.WriteLine("ISSN, e-ISSN, Display Name, Host Organization Name, Works Count, Cited by Count," +
                                "2year Mean Citedness, H-index, Cited at Least 10 Times, Is OpenAccess, Is in DOAJ," +
                                " Is Core, Price USD, Country Code, Societies, Alternative Titles, Abbreviated Title, Type, Updated Date");



                            string societiesJoined = source.societies != null ? string.Join(";", source.societies) : "";
                            string titlesJoined = source.alternate_titles != null ? string.Join(";", source.alternate_titles) : "";



                            writer.WriteLine($"{source.issn[0]},{source.issn[1]},{source.display_name},{source.host_organization_name}," +
                                $"{source.works_count},{source.cited_by_count},{source.summary_stats.two_year_mean_citedness}" +
                                $",{source.summary_stats.h_index},{source.summary_stats.i10_index},{source.is_oa}" +
                                $",{source.is_in_doaj},{source.is_core},{source.apc_usd},{source.country_code},{societiesJoined}" +
                                $",{titlesJoined},{source.abbreviated_title},{source.type},{source.updated_date}");
                        }
                    }
                    //CSV 


                    //Excel
                    if (generateNumber == "2" || generateNumber == "3")
                    {
                        using (var workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("Sheet1");

                            worksheet.Cell(1, 1).Value = "ISSN";
                            worksheet.Cell(1, 2).Value = "e-ISSN";
                            worksheet.Cell(1, 3).Value = "Display Name";
                            worksheet.Cell(1, 4).Value = "Host Organization Name";
                            worksheet.Cell(1, 5).Value = "Works Count";
                            worksheet.Cell(1, 6).Value = "Cited by Count";
                            worksheet.Cell(1, 7).Value = "2year Mean Citedness";
                            worksheet.Cell(1, 8).Value = "H-index";
                            worksheet.Cell(1, 9).Value = "Cited at Least 10 Times";
                            worksheet.Cell(1, 10).Value = "Is OpenAccess";
                            worksheet.Cell(1, 11).Value = "Is in DOAJ";
                            worksheet.Cell(1, 12).Value = "Is Core";
                            worksheet.Cell(1, 13).Value = "Price USD";
                            worksheet.Cell(1, 14).Value = "Country Code";
                            worksheet.Cell(1, 15).Value = "Societies";
                            worksheet.Cell(1, 16).Value = "Alternative Titles";
                            worksheet.Cell(1, 17).Value = "Abbreviated Title";
                            worksheet.Cell(1, 18).Value = "Type";
                            worksheet.Cell(1, 19).Value = "Updated Date";

                            string societiesJoined = source.societies != null ? string.Join(";", source.societies) : "";
                            string titlesJoined = source.alternate_titles != null ? string.Join(";", source.alternate_titles) : "";

                            worksheet.Cell(2, 1).Value = source.issn[0];
                            worksheet.Cell(2, 2).Value = source.issn[1];
                            worksheet.Cell(2, 3).Value = source.display_name;
                            worksheet.Cell(2, 4).Value = source.host_organization_name;
                            worksheet.Cell(2, 5).Value = source.works_count;
                            worksheet.Cell(2, 6).Value = source.cited_by_count;
                            worksheet.Cell(2, 7).Value = source.summary_stats.two_year_mean_citedness;
                            worksheet.Cell(2, 8).Value = source.summary_stats.h_index;
                            worksheet.Cell(2, 9).Value = source.summary_stats.i10_index;
                            worksheet.Cell(2, 10).Value = source.is_oa;
                            worksheet.Cell(2, 11).Value = source.is_in_doaj;
                            worksheet.Cell(2, 12).Value = source.is_core;
                            worksheet.Cell(2, 13).Value = source.apc_usd;
                            worksheet.Cell(2, 14).Value = source.country_code;
                            worksheet.Cell(2, 15).Value = societiesJoined;
                            worksheet.Cell(2, 16).Value = titlesJoined;
                            worksheet.Cell(2, 17).Value = source.abbreviated_title;
                            worksheet.Cell(2, 18).Value = source.type;
                            worksheet.Cell(2, 19).Value = source.updated_date;


                            workbook.SaveAs(csvFilePathBasic.Replace(".csv", ".xlsx"));
                        } //basic
                        using (var workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("Sheet1");
                            int k = 2;

                            worksheet.Cell(1, 1).Value = "Article type";
                            worksheet.Cell(1, 2).Value = "Count";

                            foreach (ArticleType top in type.group_by)
                            {
                                int l = 1;

                                worksheet.Cell(k, l).Value = top.key_display_name; l++;
                                worksheet.Cell(k, l).Value = top.count; l++;
                                k++;
                            }

                            workbook.SaveAs(csvFilePathType.Replace(".csv", ".xlsx"));
                        } //type
                        using (var workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("Sheet1");
                            int k = 2;

                            worksheet.Cell(1, 1).Value = "Is Open Access";
                            worksheet.Cell(1, 2).Value = "Count";

                            foreach (IsOpen top in opeAccess.group_by)
                            {
                                int l = 1;

                                worksheet.Cell(k, l).Value = top.key_display_name; l++;
                                worksheet.Cell(k, l).Value = top.count; l++;
                                k++;
                            }

                            workbook.SaveAs(csvFilePathOpenAccess.Replace(".csv", ".xlsx"));
                        } //oa
                        using (var workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("Sheet1");
                            int k = 2;

                            worksheet.Cell(1, 1).Value = "Concept Name";
                            worksheet.Cell(1, 2).Value = "Level";
                            worksheet.Cell(1, 3).Value = "Score";

                            foreach (x_Concept top in source.x_concepts)
                            {
                                int l = 1;

                                worksheet.Cell(k, l).Value = top.display_name; l++;
                                worksheet.Cell(k, l).Value = top.level; l++;
                                worksheet.Cell(k, l).Value = top.score; l++;
                                k++;
                            }

                            workbook.SaveAs(csvFilePathXconcepts.Replace(".csv", ".xlsx"));
                        } //x_concepts
                        using (var workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("Sheet1");
                            int k = 2;

                            worksheet.Cell(1, 1).Value = "Topic Name";
                            worksheet.Cell(1, 2).Value = "Weight/Relevance";
                            worksheet.Cell(1, 3).Value = "Subfield";
                            worksheet.Cell(1, 4).Value = "Field";
                            worksheet.Cell(1, 5).Value = "Domain";

                            foreach (Topic_share top in source.topic_share)
                            {
                                int l = 1;

                                worksheet.Cell(k, l).Value = top.display_name; l++;
                                worksheet.Cell(k, l).Value = top.value; l++;
                                worksheet.Cell(k, l).Value = top.subfield.display_name; l++;
                                worksheet.Cell(k, l).Value = top.field.display_name; l++;
                                worksheet.Cell(k, l).Value = top.domain.display_name; l++;
                                k++;
                            }

                            workbook.SaveAs(csvFilePathTopicsShare.Replace(".csv", ".xlsx"));
                        } //topic_share
                        using (var workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("Sheet1");
                            int k = 2;

                            worksheet.Cell(1, 1).Value = "Topic Name";
                            worksheet.Cell(1, 2).Value = "Count";
                            worksheet.Cell(1, 3).Value = "Subfield";
                            worksheet.Cell(1, 4).Value = "Field";
                            worksheet.Cell(1, 5).Value = "Domain";

                            foreach (Topics top in source.topics)
                            {
                                int l = 1;

                                worksheet.Cell(k, l).Value = top.display_name; l++;
                                worksheet.Cell(k, l).Value = top.count; l++;
                                worksheet.Cell(k, l).Value = top.subfield.display_name; l++;
                                worksheet.Cell(k, l).Value = top.field.display_name; l++;
                                worksheet.Cell(k, l).Value = top.domain.display_name; l++;
                                k++;
                            }

                            workbook.SaveAs(csvFilePathTopics.Replace(".csv", ".xlsx"));
                        } //topic
                        using (var workbook = new XLWorkbook())
                        {
                            var worksheet = workbook.Worksheets.Add("Sheet1");
                            int k = 2;

                            worksheet.Cell(1, 1).Value = "Year";
                            worksheet.Cell(1, 2).Value = "Works Count";
                            worksheet.Cell(1, 3).Value = "Cited by Count";

                            foreach (YearStats stat in source.counts_by_year)
                            {
                                int l = 1;

                                worksheet.Cell(k, l).Value = stat.year; l++;
                                worksheet.Cell(k, l).Value = stat.works_count; l++;
                                worksheet.Cell(k, l).Value = stat.cited_by_count; l++;
                                k++;
                            }

                            workbook.SaveAs(csvFilePathYear.Replace(".csv", ".xlsx"));
                        } //year
                        if (File.Exists(csvFilePathMainBasic.Replace(".csv", ".xlsx")))
                        {
                            using (var workbook = new XLWorkbook(csvFilePathMainBasic.Replace(".csv", ".xlsx")))
                            {
                                var worksheet = workbook.Worksheets.TryGetWorksheet("Sheet1", out var existingWorksheet)
                                ? existingWorksheet
                                : workbook.Worksheets.Add("Sheet1");

                                if (i == 1)
                                {


                                    worksheet.Cell(1, 1).Value = "ISSN";
                                    worksheet.Cell(1, 2).Value = "e-ISSN";
                                    worksheet.Cell(1, 3).Value = "Display Name";
                                    worksheet.Cell(1, 4).Value = "Host Organization Name";
                                    worksheet.Cell(1, 5).Value = "Works Count";
                                    worksheet.Cell(1, 6).Value = "Cited by Count";
                                    worksheet.Cell(1, 7).Value = "2year Mean Citedness";
                                    worksheet.Cell(1, 8).Value = "H-index";
                                    worksheet.Cell(1, 9).Value = "Cited at Least 10 Times";
                                    worksheet.Cell(1, 10).Value = "Is OpenAccess";
                                    worksheet.Cell(1, 11).Value = "Is in DOAJ";
                                    worksheet.Cell(1, 12).Value = "Is Core";
                                    worksheet.Cell(1, 13).Value = "Price USD";
                                    worksheet.Cell(1, 14).Value = "Country Code";
                                    worksheet.Cell(1, 15).Value = "Societies";
                                    worksheet.Cell(1, 16).Value = "Alternative Titles";
                                    worksheet.Cell(1, 17).Value = "Abbreviated Title";
                                    worksheet.Cell(1, 18).Value = "Type";
                                    worksheet.Cell(1, 19).Value = "Updated Date";
                                }

                                string societiesJoined = source.societies != null ? string.Join(";", source.societies) : "";
                                string titlesJoined = source.alternate_titles != null ? string.Join(";", source.alternate_titles) : "";


                                worksheet.Cell(i + 1, 1).Value = source.issn[0];
                                worksheet.Cell(i + 1, 2).Value = source.issn[1];
                                worksheet.Cell(i + 1, 3).Value = source.display_name;
                                worksheet.Cell(i + 1, 4).Value = source.host_organization_name;
                                worksheet.Cell(i + 1, 5).Value = source.works_count;
                                worksheet.Cell(i + 1, 6).Value = source.cited_by_count;
                                worksheet.Cell(i + 1, 7).Value = source.summary_stats.two_year_mean_citedness;
                                worksheet.Cell(i + 1, 8).Value = source.summary_stats.h_index;
                                worksheet.Cell(i + 1, 9).Value = source.summary_stats.i10_index;
                                worksheet.Cell(i + 1, 10).Value = source.is_oa;
                                worksheet.Cell(i + 1, 11).Value = source.is_in_doaj;
                                worksheet.Cell(i + 1, 12).Value = source.is_core;
                                worksheet.Cell(i + 1, 13).Value = source.apc_usd;
                                worksheet.Cell(i + 1, 14).Value = source.country_code;
                                worksheet.Cell(i + 1, 15).Value = societiesJoined;
                                worksheet.Cell(i + 1, 16).Value = titlesJoined;
                                worksheet.Cell(i + 1, 17).Value = source.abbreviated_title;
                                worksheet.Cell(i + 1, 18).Value = source.type;
                                worksheet.Cell(i + 1, 19).Value = source.updated_date;

                                workbook.SaveAs(csvFilePathMainBasic.Replace(".csv", ".xlsx"));

                            }
                            



                        } //all_basic
                        else
                        {
                            using (var workbook = new XLWorkbook())
                            {
                                var worksheet = workbook.Worksheets.TryGetWorksheet("Sheet1", out var existingWorksheet)
                                ? existingWorksheet
                                : workbook.Worksheets.Add("Sheet1");

                                if (i == 1)
                                {


                                    worksheet.Cell(1, 1).Value = "ISSN";
                                    worksheet.Cell(1, 2).Value = "e-ISSN";
                                    worksheet.Cell(1, 3).Value = "Display Name";
                                    worksheet.Cell(1, 4).Value = "Host Organization Name";
                                    worksheet.Cell(1, 5).Value = "Works Count";
                                    worksheet.Cell(1, 6).Value = "Cited by Count";
                                    worksheet.Cell(1, 7).Value = "2year Mean Citedness";
                                    worksheet.Cell(1, 8).Value = "H-index";
                                    worksheet.Cell(1, 9).Value = "Cited at Least 10 Times";
                                    worksheet.Cell(1, 10).Value = "Is OpenAccess";
                                    worksheet.Cell(1, 11).Value = "Is in DOAJ";
                                    worksheet.Cell(1, 12).Value = "Is Core";
                                    worksheet.Cell(1, 13).Value = "Price USD";
                                    worksheet.Cell(1, 14).Value = "Country Code";
                                    worksheet.Cell(1, 15).Value = "Societies";
                                    worksheet.Cell(1, 16).Value = "Alternative Titles";
                                    worksheet.Cell(1, 17).Value = "Abbreviated Title";
                                    worksheet.Cell(1, 18).Value = "Type";
                                    worksheet.Cell(1, 19).Value = "Updated Date";
                                }

                                string societiesJoined = source.societies != null ? string.Join(";", source.societies) : "";
                                string titlesJoined = source.alternate_titles != null ? string.Join(";", source.alternate_titles) : "";


                                worksheet.Cell(i + 1, 1).Value = source.issn[0];
                                worksheet.Cell(i + 1, 2).Value = source.issn[1];
                                worksheet.Cell(i + 1, 3).Value = source.display_name;
                                worksheet.Cell(i + 1, 4).Value = source.host_organization_name;
                                worksheet.Cell(i + 1, 5).Value = source.works_count;
                                worksheet.Cell(i + 1, 6).Value = source.cited_by_count;
                                worksheet.Cell(i + 1, 7).Value = source.summary_stats.two_year_mean_citedness;
                                worksheet.Cell(i + 1, 8).Value = source.summary_stats.h_index;
                                worksheet.Cell(i + 1, 9).Value = source.summary_stats.i10_index;
                                worksheet.Cell(i + 1, 10).Value = source.is_oa;
                                worksheet.Cell(i + 1, 11).Value = source.is_in_doaj;
                                worksheet.Cell(i + 1, 12).Value = source.is_core;
                                worksheet.Cell(i + 1, 13).Value = source.apc_usd;
                                worksheet.Cell(i + 1, 14).Value = source.country_code;
                                worksheet.Cell(i + 1, 15).Value = societiesJoined;
                                worksheet.Cell(i + 1, 16).Value = titlesJoined;
                                worksheet.Cell(i + 1, 17).Value = source.abbreviated_title;
                                worksheet.Cell(i + 1, 18).Value = source.type;
                                worksheet.Cell(i + 1, 19).Value = source.updated_date;

                                workbook.SaveAs(csvFilePathMainBasic.Replace(".csv", ".xlsx"));

                            }

                        }
                    }
                    //Excel
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Помилка: {ex.Message}");
            }
        }


    }

    
}

public class Source
{
    public string display_name { get; set; } //basic
    public int works_count { get; set; }
    public List<string> issn { get; set; }
    public string host_organization_name { get; set; }
    public int cited_by_count { get; set; }
    public SummaryStats summary_stats { get; set; }
    public bool is_oa { get; set; }
    public bool is_in_doaj { get; set; }
    public bool is_core { get; set; }
    public int? apc_usd { get; set; }
    public string country_code { get; set; }
    public List<string> societies { get; set; }
    public List<string> alternate_titles { get; set; }
    public string abbreviated_title { get; set; }
    public string type { get; set; }
    public string updated_date { get; set; }  
    public List<YearStats> counts_by_year { get; set; } //year
    public List<Topics> topics { get; set; }            //topics
    public List<Topic_share> topic_share { get; set; }  //topic_share
    public List<x_Concept> x_concepts { get; set; }     //x_concepts
}
public class OpenAccess
{
    public List<IsOpen> group_by { get; set; }
}
public class Type
{
    public List<ArticleType> group_by { get; set; }
}
public class SummaryStats
{
    public double two_year_mean_citedness { get; set; }
    public int h_index { get; set; }
    public int i10_index { get; set; }
}
public class YearStats
{
    public int year { get; set; }
    public int works_count { get; set; }
    public int cited_by_count { get; set; }
}
public class Topics
{
    public string display_name { get; set; }
    public int count { get; set; }
    public Subfield subfield { get; set; }
    public Field field { get; set; }
    public Domain domain { get; set; }
}
public class Topic_share
{
    public string display_name { get; set; }
    public double value { get; set; }
    public Subfield subfield { get; set; }
    public Field field { get; set; }
    public Domain domain { get; set; }
}
public class x_Concept
{
    public string display_name { get; set; }
    public int level { get; set; }
    public float score { get; set; }
}
public class Subfield
{
    public string display_name { get; set; }
}
public class Field
{
    public string display_name { get; set; }
}
public class Domain
{
    public string display_name { get; set; }
}
public class IsOpen
{
    public string key_display_name { get; set;}
    public int count { get; set; }
}
public class ArticleType
{
    public string key_display_name { get; set; }
    public int count { get; set; }
}