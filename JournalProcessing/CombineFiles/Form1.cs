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

            bool isDir = false;

            //keep getting all subfolders of the folder we dragged in until we hit files
            foreach (string folder in foldersArray)
            {
                //if anything we dropped is a file, we want to add those
                try {
                    isDir = (File.GetAttributes(folder) & FileAttributes.Directory)
                     == FileAttributes.Directory;

                    string ext = Path.GetExtension(folder);
                    if (!isDir && ext.Equals(".docx"))
                    {
                        OrderedListOfFiles.Add(folder);
                    }
                    else
                    {
                        GetSubFolders(folder);
                    }
                }
                catch
                {
                    //this should never hit
                    isDir = true;
                }
            }

            foreach(string file in OrderedListOfFiles)
            {
                Console.WriteLine(file);
            }

            DOCXDocument.CombineDocument(OrderedListOfFiles.ToArray(), folders[0] + ".docx");

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
            Console.WriteLine(path);
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
                if (ext.Equals(".docx"))
                {
                    try
                    {
                        //this will check if the file is clean
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
    }
}
