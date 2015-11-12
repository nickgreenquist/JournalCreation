using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.IO.Packaging;
using System.Xml;
using System.Text.RegularExpressions;

namespace WordCounter
{
    public partial class Form1 : Form
    {

        public class WordCount
        {
            public string Word { get; set; }
            public int Count { get; set; }
            public int Rank { get; set; }
        }

        public Form1()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            List<WordCount> wordCount = new List<WordCount>();
            string[] filesArray = (string[])e.Data.GetData(DataFormats.FileDrop);
            List<string> files = filesArray.ToList();
            files.Sort();

            foreach (string file in files)
            {
                Console.WriteLine(file);
                string journalEntry = TextFromWord(file);
                journalEntry = journalEntry.ToUpper();
                Regex rgx = new Regex("[^a-zA-Z0-9 ']");
                journalEntry = rgx.Replace(journalEntry, "");
                List<string> words = journalEntry.Split(' ').ToList<string>();
                words.Sort();
                Int32 index = 0;
                while (index < words.Count - 1)
                {
                    if (words[index].Equals(words[index + 1]))
                        words.RemoveAt(index);
                    else
                        index++;
                }
                for (int i = 0; i < words.Count; i++)
                {
                    int count = Regex.Matches(journalEntry, words[i]).Count;
                    WordCount newWord = new WordCount();
                    newWord.Word = words[i];
                    newWord.Count = count;
                    wordCount.Add(newWord);
                }
                wordCount = wordCount.OrderBy(w => w.Count).ToList();
                for(int i = 0; i < wordCount.Count; i++)
                {
                    wordCount[i].Rank = wordCount.Count - i;
                }
                dataGridView1.DataSource = wordCount;
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private string TextFromWord(string file)
        {
            const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            StringBuilder textBuilder = new StringBuilder();
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(file, false))
            {
                // Manage namespaces to perform XPath queries.  
                NameTable nt = new NameTable();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace("w", wordmlNamespace);

                // Get the document part from the package.  
                // Load the XML in the document part into an XmlDocument instance.  
                XmlDocument xdoc = new XmlDocument(nt);
                xdoc.Load(wdDoc.MainDocumentPart.GetStream());

                XmlNodeList paragraphNodes = xdoc.SelectNodes("//w:p", nsManager);
                foreach (XmlNode paragraphNode in paragraphNodes)
                {
                    XmlNodeList textNodes = paragraphNode.SelectNodes(".//w:t", nsManager);
                    foreach (System.Xml.XmlNode textNode in textNodes)
                    {
                        textBuilder.Append(textNode.InnerText);
                    }
                    textBuilder.Append(Environment.NewLine);
                }

            }
            return textBuilder.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                Console.WriteLine(row.Cells[0].Value);
                if (row.Cells[0].Value.ToString().Equals(textBox1.Text.ToUpper()))
                {
                    dataGridView1.CurrentCell = row.Cells[0];
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //nothing
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    Console.WriteLine(row.Cells[0].Value);
                    if (row.Cells[0].Value.ToString().Equals(textBox1.Text.ToUpper()))
                    {
                        dataGridView1.CurrentCell = row.Cells[0];
                    }
                }
            }
        }
    }
}
