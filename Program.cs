using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace wordpdf
{
    class Program
    {
        static void Main(string[] args)
        {
            Application word = new Application();
            DirectoryInfo dirInfo;
            if (args.Count() != 0)
            {
                dirInfo = new DirectoryInfo(args[0]);
            }
            else
            {
                dirInfo = new DirectoryInfo(Directory.GetCurrentDirectory());
            }
            try
            {
                FileInfo[] files = dirInfo.GetFiles();
                //wordFiles.

                foreach (FileInfo file in files)
                {
                    if (file.Extension.Contains("doc"))
                    {

                        // Cast as Object for word Open method 
                        Object filename = (Object)file.FullName;
                        Console.Write(string.Format("Conviertiendo archivo {0} a ", file.Name));
                        // Use the dummy value as a placeholder for optional arguments 
                        Document doc = word.Documents.Open(ref filename);//, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing); 
                        doc.Activate();
                        object outputFileName = file.FullName.Replace(file.Extension, ".pdf");
                        object fileFormat = WdSaveFormat.wdFormatPDF;
                        // Save document into PDF Format 
                        doc.SaveAs(ref outputFileName, ref fileFormat);//, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing); 
                                                                       // Close the Word document, but leave the Word application open. 
                                                                       // doc has to be cast to type _Document so that it will find the 
                                                                       // correct Close method. 
                        object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                        doc.Close(ref saveChanges);// ref oMissing, ref oMissing); 
                        Console.WriteLine(file.Name.Replace(file.Extension, ".pdf"));
                        Console.WriteLine("-------------------------");

                        doc = null;
                    }
                }
                Console.WriteLine("FIN.");
                Console.ReadLine();
            }
            catch(Exception e)
            {
                Console.WriteLine("error al convertir el documento");
                Console.ReadLine();
            }

        }
    }
}
