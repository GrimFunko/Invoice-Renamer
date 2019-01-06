using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Text.RegularExpressions;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;


namespace Invoice_Renamer
{

    class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("This console application renames Deliveroo invoices to \'Invoice yyyy-mm-dd\'.\n\n" +
                "Please ensure that all your invoices are located in one folder,\nthis is to ensure all invoices are " +
                "renamed in one go, although this \nprogram can be used as many times as required.\n");

            Console.WriteLine("Please input your invoice folder's location: \n(e.g \'C:\\Users\\JohnSmith\\Documents\\Invoices\')");

            //string pathInput = Console.ReadLine();
            string pathInput = "C:\\Users\\luke\\Desktop\\roo inv 2";
            string folderPath = pathInput + "\\"; 

            Console.WriteLine("Are you sure you wish to proceed? \n \'y / n\'");
            ConsoleKeyInfo answer;

            do
            {
                answer = Console.ReadKey();
                if (answer.Key != ConsoleKey.Y || answer.Key != ConsoleKey.N)
                    Console.WriteLine("\nSorry I don't understand.. \nPress \'y\' to proceed with file renaming or \'n\' to exit.");
     
                if (answer.Key == ConsoleKey.N)
                    Environment.Exit(0);

            } while (answer.Key != ConsoleKey.Y);

            Console.WriteLine();
            Console.WriteLine("Thanks, your files shall be renamed momentarily =]\n");
            Console.WriteLine("renaming...\n");
         
            List<string> targetFiles = TargetFileNames(GetFolderContents(folderPath));
            if (targetFiles.Count == 0)
            {
                Console.WriteLine("Sorry, there don't seem to be any invoices in this folder. " +
                      "\nPlease try again with the correct folder path.");
                Console.ReadKey();
                Environment.Exit(0);
            }

            foreach(string str in targetFiles)
            {
                string filePath = folderPath + str;
                string invoiceDate = GetInvoiceDate(filePath);
                ChangeInvoiceName(folderPath, str, invoiceDate, 0);
            }

            Console.WriteLine("Completed. Press any key to exit.");
            Console.ReadKey();
        }

        /// <summary>
        /// Searches a directory for files of type .pdf, then removes the directory from the filePath to give an array
        /// of fileNames.
        /// </summary>
        /// <param name="folderPath">Directory name e.g 'C:\Users\'</param>
        /// <returns></returns>
        static string[] GetFolderContents(string folderPath)
        {
            string[] pathFiles = Directory.GetFiles(folderPath, "*.pdf", SearchOption.TopDirectoryOnly);
            for (int i = 0; i < pathFiles.Length; i++)
            {
                string fileName = pathFiles[i].Replace(folderPath, "");
                pathFiles[i] = fileName;
            }
            return pathFiles;
        }

        /// <summary>
        /// Narrows an array of files to a list of list of filenames that need updating through the use of Regex 
        /// matching.
        /// </summary>
        /// <param name="files">An array of filenames.pdf</param>
        /// <returns></returns>
        static List<string> TargetFileNames(string[] files)
        {
            List<string> targetFiles = new List<string>();
            Regex pattern = new Regex(@"^Invoice\s\d{4}\-\d{2}\-\d{2}\.pdf");
            foreach (string file in files)
            {
                if (!pattern.IsMatch(file))
                {
                    targetFiles.Add(file);
                }
            }
            return targetFiles;
        }

        /// <summary>
        /// Readies a pdf file for reading, and then calls DateParser to extract the date
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        static string GetInvoiceDate(string filePath)
        {
            string invoiceDate = "";
            using (PdfReader read = new PdfReader(filePath))
            {
                using (PdfDocument pdfDocument = new PdfDocument(read))
                {
                    string content = PdfTextExtractor.GetTextFromPage(pdfDocument.GetFirstPage());

                    invoiceDate = DateParser(content);
                }
            }
            return invoiceDate;
        }

        /// <summary>
        /// Searches for the invoice date in a string, and formats it into 'yyyy-mm-dd'
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        static string DateParser(string content)
        {
            string date = "";

            Regex invoiceDatePattern = new Regex(@"Invoice\sDate\:\s\d{2}\s\w+\s\d{4}");
            Match match = invoiceDatePattern.Match(content);
            if (match.Success)
            {
                string invoiceDate = match.Value.Replace("Invoice Date: ", ""); 
                string dateSplit = @"\s";
                string[] dateComponents = Regex.Split(invoiceDate, dateSplit);
                date = dateComponents[2] + "-" + GetDictionary()[dateComponents[1]] + "-" + dateComponents[0];
            }
            return date;
        }

        /// <summary>
        /// A dictionary that keys month names with their value in mm numeric format, e.g K = January, V = 01;
        /// </summary>
        /// <returns></returns>
        static Dictionary<string,string> GetDictionary()
        {
            Dictionary<string, string> Months = new Dictionary<string, string>();
            Months.Add("January", "01");
            Months.Add("February", "02");
            Months.Add("March", "03");
            Months.Add("April", "04");
            Months.Add("May", "05");
            Months.Add("June", "06");
            Months.Add("July", "07");
            Months.Add("August", "08");
            Months.Add("September", "09");
            Months.Add("October", "10");
            Months.Add("November", "11");
            Months.Add("December", "12");

            return Months;
        }

        /// <summary>
        /// Renames a file 'Invoice yyyy-mm-dd'
        /// </summary>
        /// <param name="folderPath"></param>
        /// <param name="fileName"></param>
        /// <param name="invoiceDate"></param>
        static void ChangeInvoiceName(string folderPath, string fileName, string invoiceDate, int attempt)
        {   
            string targetFile = attempt != 0 ? 
                folderPath + "Invoice " + invoiceDate + " (" + attempt + ")" + ".pdf" : folderPath + "Invoice " + invoiceDate + ".pdf";

            attempt += 1;
            try
            {
                File.Move(folderPath + fileName, targetFile);
            }
            catch (IOException)
            {
                ChangeInvoiceName(folderPath, fileName, invoiceDate, attempt);
            }
            Console.WriteLine("renamed " + targetFile);
        }
    }

}
