using System;
using System.Diagnostics;
using System.Data.OleDb;
using System.IO;
using System.Data;
using System.Text;
using System.Linq;
using System.Collections.Generic;

namespace PlastiQ
{
    class Program
    {
        #region Initializing global variables
        const string workDirectory = "D:/delta_plast/";
        const string plastikaDirectory = "X:/PLASTIKA/2022/";
        const string LoadDirectory = "X:/PLASTIKA/Для загрузки/";
        static OleDbConnection connection = new OleDbConnection();
        static Microsoft.Office.Interop.Access.Application application = new Microsoft.Office.Interop.Access.Application();
        static StringBuilder sb = new StringBuilder();
        #endregion

        /// <summary>
        /// Main method
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            try
            {
                application.Visible = false;
                //application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                var now = DateTime.Now.DayOfWeek;
                try
                {
                    Ost4monitoring();
                }
                catch(Exception e) { Log(e.Message); };
            }
            catch (Exception e) { Log(e.Message); }
            finally
            {
                File.AppendAllText($"{workDirectory}tools/log.txt", sb.ToString());
            }
        }
        /// <summary>
        /// Ost4Monitoring execution method
        /// </summary>
        static void Ost4monitoring()
        {
            try
            {
                string[] commands = new string[] { "[MODULE]", "[make_final]", "[make_final_vbal]", "[make_final_vbal2]", "[replace_comma-dot_vbal2]", "[make_monitoring]", "[make_monitoring_vbal]", "[replace_comma-dot]", "[replace_comma-dot_vbal]" };

                Log("-----------------------------------Start OST4Monitoring.-----------------------------------");
                var files = Directory.GetFiles($"{plastikaDirectory}upd").Where(x => x.ToLower().Contains("txt")).Where(x => x.ToLower().Contains(DateTime.Now.AddDays(-1).ToString("yyyyMMdd"))).ToArray();
                var compareFile = Directory.GetFiles($"{plastikaDirectory}upd/compare").Where(x => x.ToLower().Contains("txt")).Where(x => x.ToLower().Contains("compare_"+DateTime.Now.AddDays(-1).ToString("dd.MM.yyyy"))).ToArray();

                List<string> compareFileList = new List<string>();

                if (compareFile.Length == 1)
                    compareFileList = ReadFileData(compareFile[0]);

                Log("Compare file was being read");

                if (files.Length == 5)
                {
                    foreach(var f in files)
                    {
                        File.Copy($"{plastikaDirectory}upd/{Path.GetFileName(f)}", $"{workDirectory}{Path.GetFileName(f)}", true);
                        List<string> selectedAccntInVbal = SelectCompareVbals(f, compareFileList);
                        Log($"{Path.GetFileName(f)} was compared");
                        File.WriteAllLines($"{workDirectory}{Path.GetFileName(f)}", selectedAccntInVbal);
                        Log($"{Path.GetFileName(f)} was rewritten");

                        if (!File.Exists($"{workDirectory}/bak/originalInput/{Path.GetFileName(f)}"))
                        {
                            File.Move(f, $"{workDirectory}/bak/originalInput/{Path.GetFileName(f)}");
                            Log($"{Path.GetFileName(f)} was moved to {workDirectory}bak/originalInput/");
                        }
                        else
                        {
                            File.Delete(f);
                            Log($"{Path.GetFileName(f)} from {plastikaDirectory}  was deleted due to existence of the file {workDirectory}bak/originalInput/");
                        }

                    }
                }
                else
                {
                    throw new Exception($"Not all 4monitoring files found 5({files.Length})");
                }

                connection.ConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={workDirectory}ost4monitoring.accdb;";
                StartBatFile($"{workDirectory}tools/ost_join.bat");
                Log("ost_join.bat done.");
                connection.Open();
                new OleDbCommand(SQLCOMMANDS.DeleteOSTFinal, connection).ExecuteNonQuery();
                new OleDbCommand(SQLCOMMANDS.DeleteOSTFinal_VBAL, connection).ExecuteNonQuery();
                new OleDbCommand(SQLCOMMANDS.DeleteOSTMonitoring, connection).ExecuteNonQuery();
                new OleDbCommand(SQLCOMMANDS.DeleteOSTMonitoring_VBAL, connection).ExecuteNonQuery();
                new OleDbCommand(SQLCOMMANDS.DeleteOSTNew, connection).ExecuteNonQuery();
                new OleDbCommand(SQLCOMMANDS.DeleteOSTNew_VBAL1, connection).ExecuteNonQuery();
                new OleDbCommand(SQLCOMMANDS.DeleteOSTNew_VBAL2, connection).ExecuteNonQuery();
                connection.Close();
                Log("ost_final,ost_final_vbal,ost_monitoring,ost_monitoring_vbal,ost_new,ost_new_vbal,ost_new_vbal2 cleared.");
                try
                {
                    CompactAndRepair($"{workDirectory}ost4monitoring.accdb", $"{workDirectory}temp_ost4monitoring.accdb");
                    File.Delete($"{workDirectory}ost4monitoring.accdb");
                    File.Move($"{workDirectory}temp_ost4monitoring.accdb", $"{workDirectory}ost4monitoring.accdb");
                    Log("ost4monitoring compact and repair.");
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                application.OpenCurrentDatabase($"{workDirectory}ost4monitoring.accdb");

                application.DoCmd.TransferText(Microsoft.Office.Interop.Access.AcTextTransferType.acImportDelim, "ost_vbal", "ost_new_vbal", $"{workDirectory}ost_vbal4import.txt", true);
                Log("ost_vbal4import.txt imported.");
                application.DoCmd.TransferText(Microsoft.Office.Interop.Access.AcTextTransferType.acImportDelim, "ost_new2", "ost_new", $"{workDirectory}ost_bal4import.txt", true);
                Log("ost_bal4import.txt imported.");
                //application.CloseCurrentDatabase();
               // application.Quit();

                connection.Open();
                Log("Start queries.");
                foreach (var c in commands)
                {
                    OleDbCommand command = new OleDbCommand(c, connection);
                    command.CommandType = CommandType.StoredProcedure;
                    command.ExecuteNonQuery();
                    command.Dispose();
                    Log($"Query {c} done.");
                }
                connection.Close();

                application.DoCmd.TransferText(Microsoft.Office.Interop.Access.AcTextTransferType.acExportDelim, "export_ost_monitoring", "ost_monitoring", $"{workDirectory}ost {DateTime.Now.AddDays(-1).ToString("dd-MM_yyyy")}.txt", false);
                Log($"bal {DateTime.Now.AddDays(-1).ToString("dd.MM.yyyy")}.txt exported.");
                application.DoCmd.TransferText(Microsoft.Office.Interop.Access.AcTextTransferType.acExportDelim, "export_ost_monitoring_vbal", "ost_monitoring_vbal", $"{workDirectory}vbal {DateTime.Now.AddDays(-1).ToString("dd-MM_yyyy")}.txt", false);
                Log($"vbal {DateTime.Now.AddDays(-1).ToString("dd.MM.yyyy")}.txt exported.");

                StartBatFile($"{workDirectory}tools/ost_clean_after.bat");
                Log("ost_clean_after.bat done.");

                if (!Directory.Exists($"{LoadDirectory}{DateTime.Now.AddDays(-1).ToString("ddMMyyyy")}"))
                {
                    Directory.CreateDirectory($"{LoadDirectory}{DateTime.Now.AddDays(-1).ToString("ddMMyyyy")}");
                    Log($"Folder {DateTime.Now.AddDays(-1).ToString("ddMMyyyy")} created");
                }

                var formove = Directory.GetFiles($"{workDirectory}")
                    .Where(x => x.ToLower().Contains("_") && x.ToLower().Contains("-") && (x.ToLower().Contains("bal") || x.ToLower().Contains("ost"))).ToArray();
                foreach (var f in formove)
                {
                    File.Copy($"{workDirectory}{Path.GetFileName(f)}", $"{LoadDirectory}{DateTime.Now.AddDays(-1).ToString("ddMMyyyy")}/{Path.GetFileName(f).Replace('_', '.').Replace('-', '.')}", true);
                    Log($"{Path.GetFileName(f)} moved to {LoadDirectory}{DateTime.Now.AddDays(-1).ToString("ddMMyyyy")}");
                    if (f.Contains("ost"))
                    {
                        if (!Directory.Exists(@"S:\Exchange\Северин"))
                        {
                            Directory.CreateDirectory(@"S:\Exchange\Северин");
                        }
                        File.Copy($"{workDirectory}{Path.GetFileName(f)}", $"S:/Exchange/Северин/{Path.GetFileName(f).Replace('_', '.').Replace('-', '.')}", true);
                        Log($"{Path.GetFileName(f)} moved to S/Exchange/Северин");

                    }
                    File.Delete(f);
                }


            }
            catch (Exception e)
            {
                Log($"ERROR: {e.Message}");
                StartBatFile($"{workDirectory}tools/errors.bat", $"4Monitoring {e.Message}");
            }
            finally
            {
                application.CloseCurrentDatabase();
            }
        }
        /// <summary>
        /// Method which read data from file to list of strings
        /// </summary>
        /// <param name="file"></param>
        static List<string> ReadFileData(string file)
        {
            FileInfo fileInfo = new FileInfo(file);
            List<string> fileData = new List<string>();
            using (StreamReader reader = new StreamReader(fileInfo.FullName, Encoding.Default))
            {
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    fileData.Add(line);
                }
            }
            return fileData;
        }

        /// <summary>
        /// Method which writes existing accnt to processing file
        /// </summary>
        /// <param name="fileVbal"></param>
        /// param name="compareFile"></param>
        static List<string> SelectCompareVbals(string fileVbal, List<string> compareFile)
        {
            List<string> comparedVbal = new List<string>();
            List<string> fileVbalList = ReadFileData(fileVbal);
            for (int lineCounter = 0; lineCounter < fileVbalList.Count; lineCounter++)
            {
                string splittedString = fileVbalList[lineCounter].ToString().Split(';').Skip(2).FirstOrDefault();

                if (searchAccntInCompare(splittedString, compareFile) == true)
                {
                    comparedVbal.Add(fileVbalList[lineCounter].ToString());
                }
            }
            return comparedVbal;
        }

        /// <summary>
        /// Method which starts bat file without arguments
        /// </summary>
        /// <param name="path"></param>
        static void StartBatFile(string path)
        {
            Process.Start(path).WaitForExit();
        }
        /// <summary>
        /// Method which finds out if accnt was found in the comparable file
        /// </summary>
        /// <param name="vbalAccnt"></param>
        /// <param name="compareFile"></param>
        static bool searchAccntInCompare(string vbalAccnt, List<string> compareFile)
        {
            for (int compareFileLineCounter = 0; compareFileLineCounter < compareFile.Count; compareFileLineCounter++)
            {
                if (vbalAccnt == compareFile[compareFileLineCounter])
                    return true;

            }
            return false;
        }

        /// <summary>
        /// Method which starts bat file with arguments
        /// </summary>
        /// <param name="path"></param>
        /// <param name="arguments"></param>
        static void StartBatFile(string path,string arguments)
        {
            Process.Start(path,arguments).WaitForExit();
        }
        /// <summary>
        /// Method to compact and repair access file
        /// </summary>
        /// <param name="input"></param>
        /// <param name="output"></param>
        static void CompactAndRepair(string input, string output)
        {
            var dbe = new Microsoft.Office.Interop.Access.Dao.DBEngine();
            try
            {
                dbe.CompactDatabase(input, output);
            }
            catch
            {
                throw;
            }
        }
        /// <summary>
        /// Log Method
        /// </summary>
        /// <param name="data"></param>
        static void Log(string data)
        {
            data = $"{DateTime.Now} || {data}";
            sb.Append($"{data} \n");
            Console.WriteLine(data);
        }
    }
}