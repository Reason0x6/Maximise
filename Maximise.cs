using System;
using System.IO;
using System.Linq;
using System.Printing;
using System.Threading;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Collections.Generic;
using MaterialSkin;
using MaterialSkin.Controls;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Content;

namespace Maximise
{
    public partial class Form1 : MaterialForm
    {

        static string OutputPath = @"C:\Support\MaximoOutput";
        static string GscriptPath = @"C:\Support\_gscript\bin\gswin64.exe";
        static string GSPrintPath = @"C:\Support\_gsprint\gsprint.exe";
        string SelectedPrinter = "\\\\<Print Server>\\KITMFD3";

        public Form1()
        {
            InitializeComponent();
            // Create a material theme manager and add the form to manage (this)
            MaterialSkinManager materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;

            // Configure color schema
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.Blue400, Primary.Blue500, Primary.Blue500, Accent.LightBlue200, TextShade.WHITE);


            // Enable form timer to check printer queue
            System.Windows.Forms.Timer printerTimer = new System.Windows.Forms.Timer();
            printerTimer.Tick += new EventHandler(printerTimer_Tick);
            printerTimer.Interval = 2000; // in miliseconds
            printerTimer.Start();

        }

        private void printerTimer_Tick(object sender, EventArgs e)
        {
            GetRemainingJobs(SelectedPrinter);
        }

        private void
        Browse_Button_Click(object sender, EventArgs e)
        {
            // Show the dialog that allows user to select a file, the
            // call will result a value from the DialogResult enum
            // when the dialog is dismissed.
            DialogResult ResultDialog = this.openFileDialog1.ShowDialog();
            // if a file is selected
            if (ResultDialog == DialogResult.OK)
            {
                // Set the selected file URL to the textbox
                this.lbl_FilePath.Text = this.openFileDialog1.FileName;
                this.btn_BrowseFiles.Location = new System.Drawing.Point(lbl_FilePath.Width + 30, 142);
                this.btn_CreateTestFile.Visible = true;
                this.btn_PrintFiles.Visible = true;
            }

            string InputPDFPath = "C:\\";
            string FileName = "Temp.pdf";
            PdfSharp.Pdf.PdfDocument inputPDFFile;
            try
            {
                // Input PDF Info
                InputPDFPath = this.openFileDialog1.FileName;
                FileName = System.IO.Path.GetFileNameWithoutExtension(this.openFileDialog1.FileName);

                inputPDFFile = PdfReader.Open(InputPDFPath, PdfDocumentOpenMode.Import);

            }
            catch
            {
                return;
            }

            // Get the total pages in the PDF
            var totalPagesInInputPDFFile = inputPDFFile.PageCount;

            // Page counters
            PdfSharp.Pdf.PdfDocument cut = new PdfSharp.Pdf.PdfDocument();
            int pnum = 0;
            bool IsTitlePage = false;
            int TitlePageCount = 0;

            // Loops through complete input pdf
            while (pnum < totalPagesInInputPDFFile)
            {
                // Current page context
                PdfPage p = inputPDFFile.Pages[pnum];
                IsTitlePage = false;

                // Text on current page
                PdfDictionary.PdfStream stream = p.Contents.Elements.GetDictionary(0).Stream;
                var content = ContentReader.ReadContent(p);

                // Extracted text array to string for ease of comparison
                string ExtractedText = string.Join("", PdfSharpExtensions.ExtractText(content).ToArray());

                IsTitlePage = ExtractedText.Contains("Job Plan")
                    && ExtractedText.Contains("Location")
                    && ExtractedText.Contains("WORK ORDER");

                if (IsTitlePage) { TitlePageCount++; }

                // Step Page number
                pnum++;

            }
            // Alters the label to show workorder count
            this.lbl_Validated.Text = TitlePageCount + " workorders | " + totalPagesInInputPDFFile + " pages";
        }

        private void
        testfile_button_Click(object sender, EventArgs e)
        {

            this.btn_CreateTestFile.ForeColor = System.Drawing.Color.LightSeaGreen;
            string InputPDFFilePath = this.openFileDialog1.FileName;
            string InputFileName = System.IO.Path.GetFileNameWithoutExtension(this.openFileDialog1.FileName);
            PdfSharp.Pdf.PdfDocument inputPDFFile
                = PdfReader.Open(InputPDFFilePath, PdfDocumentOpenMode.Import);

            // Get the total pages in the PDF
            var totalPagesInInputPDFFile = inputPDFFile.PageCount;

            PdfSharp.Pdf.PdfDocument cut = new PdfSharp.Pdf.PdfDocument();
            int CurrentPageNumber = 0;
            bool IsTitlePage = false;
            int PrintedPageCount = 1;

            while (CurrentPageNumber < totalPagesInInputPDFFile)
            {
                // CUrrent page import
                PdfPage CurrentPage = inputPDFFile.Pages[CurrentPageNumber];
                IsTitlePage = false;


                // Page content extracted as array
                PdfDictionary.PdfStream stream = CurrentPage.Contents.Elements.GetDictionary(0).Stream;
                var content = ContentReader.ReadContent(CurrentPage);

                // Converstion to string for comparison
                string ExtractedText = string.Join("", PdfSharpExtensions.ExtractText(content).ToArray());

                IsTitlePage = ExtractedText.Contains("Job Plan") 
                    && ExtractedText.Contains("Location")
                    && ExtractedText.Contains("WORK ORDER");


                if (CurrentPageNumber != 0 && IsTitlePage)
                {
                    // Pdf output function
                    SaveOutputPDF(cut, InputFileName + "_WORun_" + PrintedPageCount++);
                }

                // storing pdf title
                string titles = InputFileName + "_WORun_" + PrintedPageCount;
                cut = new PdfSharp.Pdf.PdfDocument(titles);

                // a3 or horizontal page to a4 vertical converstion
                if (HasContent(inputPDFFile.Pages[CurrentPageNumber]))
                {

                    // creates tamp page to rotate
                    PdfPage TempPage = inputPDFFile.Pages[CurrentPageNumber];
                    TempPage.Rotate = 0;
                    TempPage.Orientation = PdfSharp.PageOrientation.Portrait;

                    // Adds page to the PdfDocument instance
                    cut.AddPage(TempPage);
                }


                // Step page number
                CurrentPageNumber++;
            }

            // Saves last pdf of run
            SaveOutputPDF(cut, InputFileName + "_WORun_" + PrintedPageCount++);

        }

        private static void
        SaveOutputPDF(PdfSharp.Pdf.PdfDocument outputPDFDocument, String title)
        {
            // Output file path
            string outputPDFFilePath = Path.Combine(OutputPath, title + ".pdf");

            // Save the document
            outputPDFDocument.Save(outputPDFFilePath);
        }


        // reads the page contents and rules if it is blank or has content
        public bool
        HasContent(PdfPage page)
        {
            for (var i = 0; i < page.Contents.Elements.Count; i++)
            {
                if (page.Contents.Elements.GetDictionary(i).Stream.Length > 76)
                {
                    return true;
                }
            }
            return false;
        }

        private void
        print_button_Click(object sender, EventArgs e)
        {

            // Creates temp storage of pdf titles
            LinkedList<String> files = new LinkedList<string>();

            // Path for input PDF
            string inputPDFFilePath = this.openFileDialog1.FileName;

            // removes '.pdf' from title & removes spaces, before creating pdf object
            string fileName = System.IO.Path.GetFileNameWithoutExtension(this.openFileDialog1.FileName);
            fileName = fileName.Replace(" ", "_");
            PdfSharp.Pdf.PdfDocument inputPDFFile
                = PdfReader.Open(inputPDFFilePath, PdfDocumentOpenMode.Import);

            // Get the total pages in the PDF
            var totalPagesInInputPDFFile = inputPDFFile.PageCount;

            // Shows button has been pressed
            this.btn_PrintFiles.ForeColor = System.Drawing.Color.LightSeaGreen;

            // Init document variables
            PdfSharp.Pdf.PdfDocument cut = new PdfSharp.Pdf.PdfDocument();
            int CurrentPageNumber = 0;
            bool titlepage = false;
            int printed = 1;

            // Pdf reading Loop
            while (CurrentPageNumber < totalPagesInInputPDFFile)
            {
                // Current page object & contents checking
                PdfPage p = inputPDFFile.Pages[CurrentPageNumber];
                titlepage = false;

                PdfDictionary.PdfStream stream = p.Contents.Elements.GetDictionary(0).Stream;
                var content = ContentReader.ReadContent(p);

                string testString = string.Join("", PdfSharpExtensions.ExtractText(content).ToArray());

                titlepage = testString.Contains("Job Plan") && testString.Contains("Location")
                            && testString.Contains("WORK ORDER");

                // Title page based logic
                if (titlepage)
                {
                    // Saves pdf for workorder
                    if (CurrentPageNumber != 0)
                    {
                        SaveOutputPDF(cut, fileName + "_WORun_" + printed++);
                    }

                    // Creates new file & pdf object for next workorder
                    string titles = fileName + "_WORun_" + printed;
                    cut = new PdfSharp.Pdf.PdfDocument(titles);
                    files.AddLast(titles);
                }


                // A3 and Horizontal page rotation
                if (HasContent(inputPDFFile.Pages[CurrentPageNumber]))
                {

                    PdfPage TempPage = inputPDFFile.Pages[CurrentPageNumber];
                    TempPage.Rotate = 0;
                    TempPage.Orientation = PdfSharp.PageOrientation.Portrait;
                    // Add a specific page to the PdfDocument instance
                    cut.AddPage(TempPage);

                }

                // Step Current page 
                CurrentPageNumber++;
            }

            // Saves last pdf
            SaveOutputPDF(cut, fileName + "_WORun_" + printed++);

            // Start of printing queue
            this.lbl_Validated.Text = "Sending Work Orders To Printer";

            // Open Print Queue
            ProcessStartInfo Queue = new ProcessStartInfo();
            Queue.FileName = "rundll32.exe";
            Queue.Arguments = "printui.dll,PrintUIEntry /o /n " + SelectedPrinter;
            Process.Start(Queue);

            // Sends each pdf to the selected printer with a wait time to allow in-order printing
            foreach (string forprint in files)
            {
                // Printer Objects
                PrinterSettings settings = new PrinterSettings();
                ProcessStartInfo startInfo = new ProcessStartInfo();

                // Starts the printer process
                startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                startInfo.CreateNoWindow = true;
                startInfo.FileName = (GSPrintPath);

                startInfo.Arguments = ("-ghostscript "
                       + GscriptPath + " -sPAPERSIZE=a4 -dFIXEDMEDIA -dPDFFitPage -dBATCH -dNOPAUSE"
                       + " -printer " + SelectedPrinter + " -all C:\\Support\\MaximoOutput\\"
                       + forprint + ".pdf");

                // Debugging of arguments
                Debug.WriteLine(startInfo.Arguments);

                // Send to printer
                startInfo.UseShellExecute = false;
                Process.Start(startInfo);

                // Wait to allow pdf's to load in order
                Thread.Sleep(3000);
            }

        }

        // Printer selections for each of the in app options
        private void PrinterSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.cbo_PrinterSelect.SelectedIndex == 0)
            {
                SelectedPrinter = "\\\\<Print Server>\\KMNTMFD2_CLR";
            }
            else if (this.cbo_PrinterSelect.SelectedIndex == 1)
            {
                SelectedPrinter = "\\\\<Print Server>\\CMNTMFD2_CLR";
            }
            else if (this.cbo_PrinterSelect.SelectedIndex == 2)
            {
                SelectedPrinter = "\\\\<Print Server>\\CMNTMFD4_CLR";
            }
            else
            {
                SelectedPrinter = "\\\\<Print Server>\\KITMFD3";
            }
        }

        // Testing of print queue checking
        public static bool GetRemainingJobs(string printer)
        {

            String Server = "\\\\<Print Server>\\";
            //Get Printer name
            printer = printer.Replace(Server, "");


            // Print Server objects
            PrintServer server = new PrintServer(@"\\<Print Server>");
            PrintQueue queues = server.GetPrintQueue(printer);

            foreach (PrintSystemJobInfo curr in queues.GetPrintJobInfoCollection())
            {
                Debug.WriteLine(curr.JobName);
            }

            // Yeeting that print server object
            server.Dispose();

            return true;
        }
    }
}

