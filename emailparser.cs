using System;
using System.Text.RegularExpressions;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections.Generic;
using Microsoft.Azure;
using Microsoft.Azure.Storage;
using Microsoft.Azure.Storage.Blob;
using Microsoft.Azure.Storage.File;
using MimeKit;
using CsvHelper;


namespace EmailHeaderParser
{
    class Program
    {
        static void Main(string[] args)
        {
            // This is only required if emails are stored in as blobs in Storage Account
            // Azure Blob Storage Account keys and connection string declaration
            string storageAccount_connectionString = "DefaultEndpointsProtocol=https;AccountName=use-your-AzureStorageAccountName;AccountKey=use-your-AzureStorageAccountKey;EndpointSuffix=core.windows.net";
            var storageAccount = CloudStorageAccount.Parse(storageAccount_connectionString);
            var blobClient = storageAccount.CreateCloudBlobClient();
            var container = blobClient.GetContainerReference("{your-container-name}");
            container.CreateIfNotExists(BlobContainerPublicAccessType.Blob);

            // Downlaod all emails with .eml extension stored in Azure Blob one by one by passing the email file names.
            // Convert them to .txt format as your download
            string[] lines = System.IO.File.ReadAllLines(@"C:\Downloads\EmailHeaderParser\DownloadFromBlob.txt");
            foreach (string line in lines)
            {
                Console.WriteLine("\t" + line);
                CloudBlockBlob blockBlob = container.GetBlockBlobReference(line + ".eml");
                string path = (@"C:\Downloads\EmailHeaderParser\Emails\" + System.IO.Path.GetFileNameWithoutExtension(line) + ".txt");
                blockBlob.DownloadToFile(path, System.IO.FileMode.OpenOrCreate);
            }

            string[] value = new string[100];
            string[] receivedFrom = new string[50];
            string[] receivedBy = new string[50];
            string[] receivedAt = new string[50];

            string[] Servers = new string[] { "microsoftexchange" };

            var csv = new StringBuilder();
            var filePath = @"C:\Downloads\EmailHeaderParser\myOutput.csv";

            // Define attributes that you wish to parse
            void ParseEmail(int i, string fileName)
            {
                int firstStringPosition = value[i].IndexOf("from");
                int secondStringPosition = value[i].IndexOf("by");
                int thirdStringPosition = value[i].IndexOf("with");
                int fourthStringPosition = value[i].IndexOf(";");
                receivedFrom[i] = value[i].Substring(firstStringPosition + 5, (secondStringPosition - firstStringPosition) - 5);
                receivedBy[i] = value[i].Substring(secondStringPosition + 3, (thirdStringPosition - secondStringPosition) - 3);
                receivedAt[i] = value[i].Substring(fourthStringPosition + 2);

                DateTime convertedDate = DateTime.Parse(receivedAt[i].Replace("(EST)", ""));

                var firstColumn = receivedFrom[i].ToString();
                var secondColumn = receivedBy[i].ToString();
                var thirdColumn = convertedDate.ToString();
                var newLine = string.Format("{0},{1},{2},{3}", fileName, firstColumn, secondColumn, thirdColumn);
                csv.AppendLine(newLine);
            }

            string[] fileDirectory = Directory.GetFiles(@"C:\Downloads\EmailHeaderParser\Emails\");

            //Identify the headers needed to parse the content
            foreach (string files in fileDirectory)
            {
                int fileExtPosition = files.IndexOf("Emails");
                string fileName = files.Substring(fileExtPosition + 7);

                using (var stream = System.IO.File.OpenRead(files))
                {
                    var headers = HeaderList.Load(stream);

                    for (int i = 0; i < headers.Count; i++)
                    {
                        value[i] = headers[i].ToString();

                        if ((value[i].Contains("Received: from")) && (value[i].Contains("ABC")))
                        {
                            ParseEmail(i, fileName);
                        }

                        if ((value[i].Contains("Received")) && ((value[i].Contains("DEF")) || (value[i].Contains("ABC")) || (value[i].Contains("ASD")) || (value[i].Contains("XYZ"))))
                        {
                            ParseEmail(i, fileName);
                        }

                        if ((value[i].Contains("Received")) && ((value[i].Contains("abc@google.com")) || (value[i].Contains("xyz@google.com"))))
                        {
                            for (int j = 0; j < 8; j++)
                            {
                                if (value[i].Contains("by " + Servers[j]))
                                {
                                    ParseEmail(i, fileName);
                                }
                            }
                        }

                    }

                    string cGeneratedAt = headers[HeaderId.Date].ToString();
                    DateTime convertedCDate = DateTime.Parse(cGeneratedAt.Replace("(EST)", ""));
                    Console.WriteLine(convertedCDate);

                    var newLine1 = string.Format("{0},EmailGenerated,C,{1}", fileName, convertedCDate);
                    csv.AppendLine(newLine1);
                    File.WriteAllText(filePath, csv.ToString());

                }

            }

            using (var reader = new StreamReader(@"C:\Downloads\EmailHeaderParser\myOutput.csv"))
            {
                List<string> readListA = new List<string>();
                List<string> readListB = new List<string>();
                List<string> readListC = new List<string>();
                List<string> readListD = new List<string>();
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    readListA.Add(values[0]);
                    readListB.Add(values[1]);
                    readListC.Add(values[2]);
                    readListD.Add(values[3]);
                }

                string[] columnA = readListA.ToArray();
                string[] columnB = readListB.ToArray();
                string[] columnC = readListC.ToArray();
                string[] columnD = readListD.ToArray();

                var filePath1 = @"C:\Downloads\EmailHeaderParser\myOutput1.csv";
                using (StreamWriter csv1 = new StreamWriter(filePath1, true))
                {
                    csv1.WriteLine("ItemID, MileStone1ReceivedFrom, MileStone1ReceivedBy, MileStone1ReceivedAt, MileStone2ReceivedFrom, MileStone2ReceivedBy, MileStone2ReceivedAt, MileStone3ReceivedFrom, MileStone3ReceivedBy, MileStone3ReceivedAt, MileStone4ReceivedFrom, MileStone4ReceivedBy, MileStone4ReceivedAt");
                }

                for (int i = 0; i < readListA.Count; i++)
                {
                    if (columnB[i].Contains("EmailGenerated"))
                    {
                        using (StreamWriter csv1 = new StreamWriter(filePath1, true))
                        {
                            csv1.WriteLine("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12}", columnA[i], columnB[i], columnC[i], columnD[i], columnB[i-1], columnC[i-1], columnD[i-1], columnB[i-2], columnC[i-2], columnD[i-2], columnB[i-3], columnC[i-3], columnD[i-3]);
                        }
                        
                    }

                }
            }


            System.Console.ReadKey();

        }

    }
}