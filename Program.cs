
using System.Data.SqlClient;
using System.Text;
using OfficeOpenXml;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.Reflection;
using System.IO;

namespace Agritec
{
    class Program{
        static async Task<int> Main(string[] args)
        {
            DirectoryInfo successDir = Directory.CreateDirectory(@".\Success");
            
            DirectoryInfo errorDir = Directory.CreateDirectory(@".\Error");
            var fileOption = new Option<FileInfo?>(
            name: "--file",
            description: "The file to read and process.");
            
            var silentMode = new Option<bool>(
            name: "--silent",
            description: "No info is displayed in console.");

            var autoCommit = new Option<bool>(
            name: "--autocommit",
            description: "Do not ask for confirmation to commit changes.");

            var rootCommand = new RootCommand("Reads the excel file and updates logdata table with new batchIdExt and client");
            
            rootCommand.AddOption(fileOption);
            rootCommand.AddOption(silentMode);
            rootCommand.AddOption(autoCommit);

            rootCommand.SetHandler((file, silent, auto) => 
            { 
                if (file == null)
                {
                    DirectoryInfo d = new DirectoryInfo(@".\");
                    FileInfo[] Files = d.GetFiles("*.xlsx"); //Getting XLSX files
                    
                    foreach(FileInfo file2 in Files )
                    {
                        int result = ProcessAging(file2.FullName, silent);
                        string destination = "";
                        if (result == 0)
                        {
                            destination = successDir.FullName + Path.DirectorySeparatorChar.ToString() + file2.Name; 
                        }
                        else
                        {
                            destination = errorDir.FullName + Path.DirectorySeparatorChar.ToString() + file2.Name;
                        }
                        File.Move(file2.FullName, destination);
                    }
                }
                else
                {
                    int result = ProcessAging(file.FullName, silent); 
                    if (result == 0)
                    {
                        file.MoveTo(successDir.FullName);
                    }
                    else
                    {
                        file.MoveTo(errorDir.FullName);
                    }
                }
            },
            fileOption, silentMode, autoCommit);

            return await rootCommand.InvokeAsync(args);
        }
        public static int ProcessAging(string fileLocation, bool silentMode, bool autoCommit = true)
        {
            if (!silentMode) Console.WriteLine(fileLocation);
            if (!silentMode) Console.WriteLine(silentMode);
            Dictionary<string, int>  columnIndex = new Dictionary<string, int>();
            columnIndex.Add("clientId", 1);
            columnIndex.Add("clientName", 2);
            
            columnIndex.Add("fechaEmision", 3);
            columnIndex.Add("expiredAmmount", 8);
            columnIndex.Add("serie", 6);
            columnIndex.Add("nroDocumento", 7);
            int incrementReport = 10;
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(fileLocation)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var myWorksheet = xlPackage.Workbook.Worksheets["Aging"]; //select sheet here
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;
                var firstCol = myWorksheet.Dimension.Start.Column;

                if (!silentMode) Console.WriteLine("Total rows read: " + (totalRows -1));

                
                int currentPercentage = 0;
                try 
                {
                    int totalRowsAffected = 0;
                    for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                    {
                        if (myWorksheet.Cells[rowNum, 1].Value == null)
                        {
                            continue;
                        }
                        var value = myWorksheet.Cells[rowNum, columnIndex["expiredAmmount"]].Value ?? "0.0";
                        double ammount = 0;
                        Double.TryParse(value.ToString(), out ammount);
                        if (ammount > 0.0)
                        {
                            string? client = myWorksheet.Cells[rowNum, columnIndex["clientName"]].Value.ToString();
                            string? clientId = myWorksheet.Cells[rowNum, columnIndex["clientId"]].Value.ToString();
                            string? fechaEmision = myWorksheet.Cells[rowNum, columnIndex["fechaEmision"]].Value.ToString();
                            string? serie = myWorksheet.Cells[rowNum, columnIndex["serie"]].Value.ToString();
                            string? nroDoc = myWorksheet.Cells[rowNum, columnIndex["nroDocumento"]].Value.ToString();
                            
                            Console.WriteLine(client + " con ID:" + clientId + " debe $" + ammount + " por la factura " + serie + " " + nroDoc + " emitida el: " + fechaEmision);
                        }
                        //pilaParam.Value = myWorksheet.Cells[rowNum, columnIndex["pila"]].Value.ToString();
                        //clienteOrigenParam.Value = Convert.ToInt32(myWorksheet.Cells[rowNum, columnIndex["cliente origen"]].Value);
                        //clienteFinalParam.Value = Convert.ToInt32(myWorksheet.Cells[rowNum, columnIndex["cliente final"]].Value ?? clienteOrigenParam.Value);
                        
                        //totalRowsAffected += cmd.ExecuteNonQuery();
                        if (!silentMode && rowNum/(double)totalRows >= (currentPercentage + incrementReport)/100.0)
                        {
                            Console.WriteLine("{0}% done", currentPercentage);
                            Console.WriteLine("Total rows affected: " + totalRowsAffected);
                            currentPercentage += incrementReport;
                        }
                    }
                    /* if (totalRowsAffected<= totalRows-1)
                    {
                        string choice = "y";                        
                        if (!silentMode) 
                        {
                            Console.WriteLine("100% done. Total rows affected: " + totalRowsAffected);
                            Console.WriteLine("Commit changes? y/n");
                            if (!autoCommit) choice = Console.ReadLine().ToLower();
                        } 

                        if (choice == "y" || silentMode || autoCommit)
                        {
                            transaction.Commit();
                            if (!silentMode) Console.WriteLine("Changes commited. Total rows affected: " + totalRowsAffected);
                        }
                        else
                        {
                            transaction.Rollback();
                            if (!silentMode) Console.WriteLine("Changes rolledback. No rows were affected ");
                        }
                    }
                    else
                    {
                        transaction.Rollback();
                        if (!silentMode) Console.WriteLine("Error Total rows: "+ totalRows +" and total affected rows: " + totalRowsAffected);
                    }
                    con.Close(); */
                    xlPackage.Dispose();
                    return 0;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    return -1;
                }
            }
        }
    }
}