using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using HtmlAgilityPack;
using Amazon;
using Amazon.Polly;
using Amazon.Polly.Model;
using Amazon.Runtime;

namespace ConsoleApp1
{
    class Program
    {
        struct struct_General_Provision
        {
            public long number_result;
            public string name_search;
            public struct_General_Provision(long number_result, string name_search)
            {
                this.number_result = number_result;
                this.name_search = name_search;
            }

        }
        struct struct_Duties
        {
            public long number_result;
            public string name_search;
            public struct_Duties(long number_result, string name_search)
            {
                this.number_result = number_result;
                this.name_search = name_search;
            }

        }
        struct struct_Rights
        {
            public long number_result;
            public string name_search;
            public struct_Rights(long number_result, string name_search)
            {
                this.number_result = number_result;
                this.name_search = name_search;
            }

        }
        [STAThread]
        static void Main(string[] args)
        {
            Console.Write("Введите название профессии: ");
            string NameProf = Console.ReadLine();
            string path = CreateCatalog(NameProf) + "\\";
            string AccessKeyID = "AKIAIKFSWLHW3D6LO4YA";
            string SecretAccessKey = "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";


            List<string> general_Provisions = new List<string> { "Общие положения", "ОБЩИЕ ПОЛОЖЕНИЯ", "Общие правила", "ОБЩИЕ ПРАВИЛА" };
            List<string> duties = new List<string> { "Обязанности", "обязанности", "ОБЯЗАННОСТИ" };
            List<string> rights = new List<string> { "Права", "права", "ПРАВА" };
            List<string> general_Provisions_list = new List<string> { };
            List<string> duties_list = new List<string> { };
            List<string> rights_list = new List<string> { };
            List<int> number_result;
            int id_duties = 0, id_rights = 0, id_general_Provisions = 0;

            List<string> parText = new List<string> { };
            //Открытие файла и выгрузка его в список
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "Файлы MS Word |*.rtf",
                Multiselect = false
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                Word.Application app = new Word.Application();
                Object fileName = dialog.FileName;
                app.Documents.Open(ref fileName);
                Word.Document doc = app.ActiveDocument;
                // Нумерация параграфов начинается с одного
                for (int i = 1; i < doc.Paragraphs.Count; i++)
                {

                    parText.Add(doc.Paragraphs[i].Range.Text);
                    Console.WriteLine(parText[parText.Count - 1]);

                }

                app.Quit();

            }
            //Нахождение строк основных разделителей на группы
            for (int i = 0; i < parText.Count; i++)
            {
                foreach (string list in duties)
                {
                    if (parText[i].Contains(list)) id_duties = i;
                }
                foreach (string list in rights)
                {
                    if (parText[i].Contains(list)) id_rights = i;
                }
                foreach (string list in general_Provisions)
                {
                    if (parText[i].Contains(list)) id_general_Provisions = i;
                }
            }

            //Загрузка общих положений
            while (parText[id_general_Provisions + 2].ToString().Length > 5)
            {
                general_Provisions_list.Add(parText[id_general_Provisions + 2]);
                id_general_Provisions++;
            }
            //Звгрузка обязанностей
            while (parText[id_duties + 2].ToString().Length > 5)
            {
                duties_list.Add(parText[id_duties + 2]);
                id_duties++;
            }
            //Загрузка прав
            while (parText[id_rights + 2].ToString().Length > 5)
            {
                rights_list.Add(parText[id_rights + 2]);
                id_rights++;

            }

            List<struct_General_Provision> s_G_P = new List<struct_General_Provision>();
            List<struct_Duties> s_D = new List<struct_Duties>();
            List<struct_Rights> s_R = new List<struct_Rights>();
            foreach (string list in general_Provisions_list)
            {
                Random rnd = new Random();
                int value = rnd.Next(800, 1400);
                s_G_P.Add(new struct_General_Provision(pars(list), list));
                Thread.Sleep(value);
            }
            foreach (string list in duties_list)
            {
                Random rnd = new Random();
                int value = rnd.Next(800, 1400);
                s_D.Add(new struct_Duties(pars(list), list));
                Thread.Sleep(value);
            }
            foreach (string list in rights_list)
            {
                Random rnd = new Random();
                int value = rnd.Next(800, 1400);
                s_R.Add(new struct_Rights(pars(list), list));
                Thread.Sleep(value);
            }
            //Подсчёт среднего арифметического
            double sum = 0;
            double avg_general_Provisions = 0, avg_Duties = 0, avg_Rights = 0;
            for (int i = 0; i < s_G_P.Count; i++)
            {
                sum += s_G_P[i].number_result;
            }
            avg_general_Provisions = Convert.ToDouble(sum / s_G_P.Count) * 0.7;
            sum = 0;
            for (int i = 0; i < s_D.Count; i++)
            {
                sum += s_D[i].number_result;
            }
            avg_Duties = Convert.ToDouble(sum / s_D.Count) * 0.7;
            sum = 0;
            for (int i = 0; i < s_R.Count; i++)
            {
                sum += s_R[i].number_result;
            }
            avg_Rights = Convert.ToDouble(sum / s_R.Count) * 0.7;

            // Сохранение в файлы

            string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string audiotext  = "";


            using (StreamWriter w = new StreamWriter(path + "General Provisions.txt", false, Encoding.GetEncoding(1251)))
            {
                for (int i = 0; i < s_G_P.Count; i++)
                {
                    if (s_G_P[i].number_result > avg_general_Provisions)
                        continue;
                    w.WriteLine(s_G_P[i].name_search);
                    audiotext += s_G_P[i].name_search;
                }
            }

            using (StreamWriter w = new StreamWriter(path + "Duties.txt", false, Encoding.GetEncoding(1251)))
            {
                for (int i = 0; i < s_D.Count; i++)
                {
                    if (s_D[i].number_result > avg_Duties)
                        continue;
                    w.WriteLine(s_D[i].name_search);
                    audiotext += s_D[i].name_search;
                }
            }



            using (StreamWriter w = new StreamWriter(path + "Rights.txt", false, Encoding.GetEncoding(1251)))
            {
                for (int i = 0; i < s_R.Count; i++)
                {
                    if (s_R[i].number_result > avg_Rights)
                        continue;
                    w.WriteLine(s_R[i].name_search);
                    audiotext += s_R[i].name_search;
                }
            }

            BasicAWSCredentials awsCredentials =
               new BasicAWSCredentials(AccessKeyID, SecretAccessKey);

            // создаём объект класса AmazonPollyClient, 
            // передавая данные аккаунта и указывая используемый регион
            AmazonPollyClient amazonPollyClient = new AmazonPollyClient(awsCredentials, RegionEndpoint.EUCentral1);
            // создаём объект запроса
            string text_Speech = audiotext;
            SynthesizeSpeechRequest synthesizeSpeechRequest = MakeSynthesizeSpeechRequest(text_Speech);
            // получаем ответ от AWS Polly
            SynthesizeSpeechResponse synthesizeSpeechResponse = amazonPollyClient.SynthesizeSpeech(synthesizeSpeechRequest);

            CreateMp3File(synthesizeSpeechResponse.AudioStream, path, NameProf);
            MessageBox.Show("Формирование выжимки окончено");
            Console.ReadKey();

        }

        //Парс колличества результатов
        static long pars(string name)
        {
            string url = "https://google.com/search?q="; //Дополняемая ссылка 
            url = url + name;
			HtmlAgilityPack.HtmlWeb webDoc = new HtmlWeb();
			HtmlAgilityPack.HtmlDocument doc = webDoc.Load(url);
			string text = null;
			string res = null;
			foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//div[@id]"))
			{
				text = node.GetAttributeValue("id", String.Empty);
				if (text == "result-stats")
				{
					text = node.InnerText;
					break;
				}

			}
			foreach (string s in text.Split(new char[] { ' ' }))
				foreach (char c in s.ToCharArray())
				{
					if ((Convert.ToInt32(c) > 47 && Convert.ToInt32(c) < 58) || (Convert.ToInt32(c) == 160))
					{
						if (Convert.ToInt32(c) != 160) res += c;
					}
					else break;
				}
            return Convert.ToInt32(res);
		}
        private static SynthesizeSpeechRequest MakeSynthesizeSpeechRequest(string text)
        {
            // создаём объект запроса
            SynthesizeSpeechRequest synthesizeSpeechRequest = new SynthesizeSpeechRequest();
            // передаём необходимый текст
            synthesizeSpeechRequest.Text = text;
            // указываем код передаваемого языка
            synthesizeSpeechRequest.LanguageCode = LanguageCode.RuRU;
            // указываем выходной формат
            synthesizeSpeechRequest.OutputFormat = OutputFormat.Mp3;
            // указываем желаемый голос
            synthesizeSpeechRequest.VoiceId = VoiceId.Maxim;

            return synthesizeSpeechRequest;
        }

        // метод для создания mp3 файла
        private static void CreateMp3File(Stream audioStream, string path, string Name)
        {
            // указываем путь к сохраняемому mp3 файлу
            string pathToMp3 = path + Name + ".mp3";

            using (FileStream fileStream = File.Create(pathToMp3))
            {
                audioStream.CopyTo(fileStream);
                fileStream.Flush();
                fileStream.Close();
            }
        }

        public static string CreateCatalog(string NameFile)
        {
            string path = @"B:\File";
            string subpath = @""+NameFile;
            DirectoryInfo dirInfo = new DirectoryInfo(path);
            if (!dirInfo.Exists)
            {
                dirInfo.Create();
            }
            dirInfo.CreateSubdirectory(subpath);
            return path + "\\"+subpath;
        }
    }
   
}
