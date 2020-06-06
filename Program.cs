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
            while (parText[id_general_Provisions+2].ToString().Length > 5)
            {
                general_Provisions_list.Add(parText[id_general_Provisions + 2]);
                id_general_Provisions++;
            }
            while (parText[id_duties + 2].ToString().Length > 5)
            {
                duties_list.Add(parText[id_duties + 2]);
                id_duties++;
            }
            while (parText[id_rights + 2].ToString().Length > 5)
            {
                rights_list.Add(parText[id_rights + 2]);
                id_rights++;
                
            }

            Console.ReadKey();

        }

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
