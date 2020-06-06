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

namespace ConsoleApp1
{
    class Program
    {
        struct struct_General_Provision
        {
            public int number_result;
            public string name_search;
            public struct_General_Provision(int number_result, string name_search)
            {
                this.number_result = number_result;
                this.name_search = name_search;
            }

        }
        struct struct_Duties
        {
            public int number_result;
            public string name_search;
            public struct_Duties(int number_result, string name_search)
            {
                this.number_result = number_result;
                this.name_search = name_search;
            }

        }
        struct struct_Rights
        {
            public int number_result;
            public string name_search;
            public struct_Rights(int number_result, string name_search)
            {
                this.number_result = number_result;
                this.name_search = name_search;
            }

        }
        [STAThread]
        static void Main(string[] args)
        {
            
            List<string>general_Provisions = new List<string> { "Общие положения", "ОБЩИЕ ПОЛОЖЕНИЯ"};
            List<string> duties = new List<string> { "Обязанности","обязанности","ОБЯЗАННОСТИ"};
            List<string> rights = new List<string> { "Права", "права", "ПРАВА" };
            List<string> general_Provisions_list = new List<string> { };
            List<string> duties_list = new List<string> { };
            List<string> rights_list = new List<string> { };
            List<int> number_result;
            int id_duties = 0,id_rights = 0, id_general_Provisions = 0;
           
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
                    Console.WriteLine(parText[parText.Count-1]);
                    
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
            while (parText[id_general_Provisions+2].ToString().Length > 5)
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
            foreach (string list in general_Provisions)
            {
                Random rnd = new Random();
                int value = rnd.Next(800,1400);
                s_G_P.Add(new struct_General_Provision(pars(list),list));
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
                s_R.Add(new struct_Rights(pars(list),list));
                Thread.Sleep(value);
            }

            
            Console.ReadKey();

        }
        //Парс колличества результатов
        static int pars(string name)
        {
            string url = "https://google.com/search?q=";
            url = url + name;
            HtmlAgilityPack.HtmlWeb webDoc = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc = webDoc.Load(url);
            string text = null;
            int num = 0;
            foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//div[@id]"))
            {
                text = node.GetAttributeValue("id", String.Empty);
                if (text == "result-stats")
                {
                    text = node.InnerHtml;
                    break;
                }

            }
            string[] split = text.Split(new char[] { '<', '>' });
            string[] res = null;
            foreach (string s in split)
            {
                res = s.Split(new char[] { ' ' });
                if (res[0] == "Результатов:")
                {
                    break;
                }
                else res = null;
            }
            text = String.Join(String.Empty, res);
            foreach (char c in text.ToCharArray())
            {
                if (Convert.ToInt32(c) > 47 && Convert.ToInt32(c) < 58)
                    text += c;

                num = Convert.ToInt32(text);
            }
            return num;
        }
    }
   
}
