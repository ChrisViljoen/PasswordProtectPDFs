using System;
using System.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace PasswordProtectPDFs
{
    class Program
    {
        private static bool status = true;
        private static readonly ConsoleColor defaultColor = Console.ForegroundColor;
        private static object padLock = new object();

        [STAThread]
        static void Main(string[] args)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Please select the folder containing the benefit statements";
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                var start = DateTime.Now;
                Parallel.ForEach(Directory.GetFiles(fbd.SelectedPath), (filePath) =>
                {
                    ProtectAndRenameFile(filePath);
                });
                if (status)
                    Console.WriteLine("Process Complete");
                else
                    Console.WriteLine("Process Completed with errors. See red lines.");
                var end = DateTime.Now;
                var diff = end - start;
                Console.WriteLine($"Time taken: {diff}");
                Console.WriteLine("Press any key to exit");
                Console.ReadKey();
            }
        }

        private static void ProtectAndRenameFile(string filePath)
        {
            var fileName = Path.GetFileName(filePath);
            
            try
            {
                PdfDocument document = PdfReader.Open(filePath);
                var fileNameWithoutExtention = Path.GetFileNameWithoutExtension(filePath);
                var split = fileNameWithoutExtention.Split('_');
                var idOrPassportNumber = split[2];
                if (idOrPassportNumber.Length < 1)
                    throw new Exception("ID or Passport number not found in file name");
                document.SecuritySettings.UserPassword = idOrPassportNumber;
                document.SecuritySettings.OwnerPassword = idOrPassportNumber;
                document.Save(filePath);
                document.Close();
                var newFileName = filePath.Replace($"_{idOrPassportNumber}", string.Empty);
                File.Move(filePath, newFileName);
                Console.WriteLine(fileName + " has been password protected and renamed to " + Path.GetFileName(newFileName) + Environment.NewLine);
            }
            catch (Exception e)
            {
                lock (padLock)
                {
                    status = false;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(fileName + " failed to execute");
                    Console.WriteLine(e.Message);
                    Console.ForegroundColor = defaultColor;
                    Console.WriteLine();
                }                
            }
        }
    }
}
