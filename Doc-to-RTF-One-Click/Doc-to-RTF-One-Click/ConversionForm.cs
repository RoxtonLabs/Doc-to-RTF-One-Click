using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;    //For the directory handling
using Word = Microsoft.Office.Interop.Word; //Right-click solution, Add>Rference>Microsoft Word 14.0 Object Library

namespace Doc_to_RTF_One_Click
{
    public partial class ConversionForm : Form
    {
        public ConversionForm()
        {
            InitializeComponent();
        }

        private void convertFiles()
        {   //Converts all the .doc and .docxf files in the folder into .rtf
            //UI faffery

            //First, create the output directory if it doesn't exist
            string curDir = Application.StartupPath;
            if (!Directory.Exists(curDir + @"\Output"))
            {
                statusLabel.Text = "Creating output directory...";
                Directory.CreateDirectory(curDir + @"\Output");
            }

            //Next, get a list of all the eligible files in the current folder
            statusLabel.Text = "Collating files...";
            string[] fileNames = Directory.GetFiles(curDir, "*.doc*");
            List<string> converted = new List<string>();    //This list will hold all the files we have successfully converted

            //UI faffery
            progressBar.Minimum = 0;
            progressBar.Maximum = fileNames.Length;
            progressBar.Value = 0;
            progressBar.Step = 1;

            //Create our Word objects
            try
            {
                statusLabel.Text = "Initialising MS Word...";
                Word.Application wordApp = new Word.Application();
                wordApp.Visible = true; //For development purposes at least; turn this off for production
                Word.Document wordDoc;
                object unknown = Type.Missing;  //Object needed for the save method
                object format = Word.WdSaveFormat.wdFormatRTF;
                //Start Word and get transforming
                foreach (string fileName in fileNames)
                {   //Open the file
                    progressBar.PerformStep();
                    object newName = curDir + @"\Output\" + Path.GetFileNameWithoutExtension(fileName) + ".rtf";
                    try
                    {   //Often an error is returned even if the file works perfectly
                        statusLabel.Text = "Opening " + Path.GetFileName(fileName).ToString() + "...";
                        wordDoc = wordApp.Documents.Open(fileName);
                        wordDoc.SaveAs2(newName, format);
                        wordDoc.Close();
                        converted.Add(fileName);    //Add the converted file to the list
                    }
                    catch (Exception) { }
                }

                //Close Word and release our objects
                //Not entirely sure if releasing them is necessary but better safe etc.
                statusLabel.Text = "Closing MS Word...";
                wordApp.Quit();
                releaseObject(wordApp);
                wordDoc = null; //We have to do this because of that damn try-loop
                releaseObject(wordDoc);

                //Delete the files we've converted (if we can)
                statusLabel.Text = "Deleting input files...";
                string[] convertedFiles = converted.ToArray();
                foreach(string fileName in convertedFiles)
                {
                    try
                    {
                        File.Delete(fileName);
                    }
                    catch (Exception) { }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Could not open Microsoft Word.");
                Environment.Exit(0);
            }
        }

        //Release interop objects
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception x)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();   //Call in the binman
            }
        }

        private void ConversionForm_Shown(object sender, EventArgs e)
        {
            convertFiles();
            Environment.Exit(1);
        }
    }
}
