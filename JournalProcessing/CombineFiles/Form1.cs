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
using RasterEdge.Imaging.Basic;
using RasterEdge.XDoc.Word;

namespace CombineFiles
{
    public partial class Form1 : Form
    {
        public static string AllText;
        public static List<string> OrderedListOfFiles;

        public Form1()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
            AllText = "";
            OrderedListOfFiles = new List<string>();
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] foldersArray = (string[])e.Data.GetData(DataFormats.FileDrop);
            List<string> folders = foldersArray.ToList();
            folders.Sort();

            //keep getting all subfolders until we hit files
            foreach(string folder in foldersArray)
            {
                Console.WriteLine(folder);
                GetSubFolders(folder);
            }

            DOCXDocument.CombineDocument(OrderedListOfFiles.ToArray(), folders[0] + "test.docx");

            //reset lists
            OrderedListOfFiles.Clear();
            AllText = "";
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }
        
        private void GetSubFolders(string path)
        {
            string[] foldersArray = Directory.GetDirectories(path);
            List<string> folders = foldersArray.ToList();
            folders.Sort();
            foreach (string folder in folders)
            {
                GetSubFolders(folder);
            }

            //once we have dealt with any subfolders, we add the files we find in the folder
            string[] filesInfolder = Directory.GetFiles(path);
            List<string> files = filesInfolder.ToList();
            files.Sort();
            foreach (string file in files)
            {
                string ext = Path.GetExtension(file);
                if (ext.Equals(".docx") || ext.Equals(".doc"))
                {
                    try
                    {
                        TryToOpenWord(file);
                        OrderedListOfFiles.Add(file);
                    }
                    catch
                    {
                        //if this fails, this is not a valid word doc
                    }
                }
            }
        }

        private void TryToOpenWord(string file)
        {
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(file, false))
            {
            }
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
    }
}
