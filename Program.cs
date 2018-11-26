using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;   //IMPORTANT: Excel Library
using System.Net.Mail;
using System.Net;

namespace readExcelData
{
    class Program
    {
        public static bool exit = false;

        static void Main(string[] args)
        {
            //initializing default paths to read src and dst path later
            initializeDefaultpaths();

            int choice = 0;
            while (!exit)
            {

                #region Menu
                Console.WriteLine("CollectExcelData from Fabian Denk");
                Console.WriteLine();
                Console.WriteLine("After reading the Excel Files and re-writing them, ");
                Console.WriteLine("you need to press 'save' when the standard Windows prompt opens up");
                Console.WriteLine("Remove all filters on the Destination Excel file and ");
                Console.WriteLine("close all of the Excel Files the Program uses before collecting Data.");
                Console.WriteLine();
                Console.WriteLine("1.) collect Data");
                Console.WriteLine("2.) change Sourcepath");
                Console.WriteLine("3.) change Destinationpath");
                Console.WriteLine("4.) show Information");
                Console.WriteLine("5.) report Bug");
                Console.WriteLine("6.) Exit");
                Console.WriteLine();
                #endregion

                try
                {
                    choice = int.Parse(Console.ReadLine());

                    switch (choice)
                    {
                        //collect and gather data
                        case 1:
                            Console.Clear();
                            if(readAndWriteExcelFiles(getSrcPath(), getDstPath()))
                            {
                                Console.WriteLine("Data has been read and rewritten in file!");
                            }
                            else
                            {
                                Console.WriteLine("Try entering new Source and Destination path");
                            }
                            Console.ReadLine();
                            Console.WriteLine("Case 1 finished");
                            break;

                        //change the source path
                        case 2:
                            Console.Clear();
                            if (!ChangeSourcepath())
                            {
                                Console.WriteLine();
                                Console.WriteLine("Error trying to change Sourcepath!");
                            }
                            else
                            {
                                Console.WriteLine();
                                Console.WriteLine("Sourcepath successfully changed!");
                            }
                            Console.ReadLine();
                            break;
                        
                        //change the destination path
                        case 3:
                            Console.Clear();
                            if (!ChangeDestinationpath())
                            {
                                Console.WriteLine();
                                Console.WriteLine("Error trying to change Destinationpath!");
                            }
                            else
                            {
                                Console.WriteLine();
                                Console.WriteLine("Destinationpath successfully changed!");
                            }
                            Console.ReadLine();
                            break;
                        
                        //show interesting data
                        case 4:
                            Console.Clear();
                            if (!showData())
                            {
                                Console.WriteLine("Error while trying to show Data, try again later");
                            }
                            Console.ReadLine();
                            break;
                        case 5:
                            Console.Clear();
                            if(!reportBug())
                            {
                                Console.WriteLine("Error while trying to report a Bug");
                            }
                            Console.ReadLine();
                            break;
                        //exit the program
                        case 6:
                            Console.Clear();
                            Console.WriteLine("Exiting Program, press enter key..");
                            exit = true;
                            Console.ReadLine();
                            break;

                    }
                    Console.Clear();
                }
                catch(Exception ex)
                {
                    //catching parsing error
                    Console.Clear();
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("Enter a number between 1 and 4 only!");
                    Console.ReadLine();
                }
            }
        }

        #region report Bug

        public static bool reportBug()
        {
            //activate the right settings in gmx:
            //click on email
            //click on settings (bottom left)
            //click on POP3/IMAP Abruf
            //click on POP3 und IMAP Zugriff erlauben
            //you can also see SMTP sever here
            try
            {
                //reading in users email
                Console.WriteLine("Enter your email adress: ");

                string emailAdress;

                emailAdress = Console.ReadLine();

                //simple validation of email format
                if(!(emailAdress.Contains("@") && emailAdress.Contains(".")))
                {
                    Console.WriteLine("Wrong mail adress format");
                    return false;
                }

                //seperate email -> everything before @ is the username and will be sent too
                string[] seperator = new string[1];
                seperator[0] = "@";
                //seperating mail adress
                string[] trimmed = emailAdress.Split(seperator, StringSplitOptions.RemoveEmptyEntries);
                //building greetings strings
                string mfg = "Mit freundlichen Grüßen \r\n" + trimmed[0];

                Console.WriteLine();

                //reading main message with bug report
                Console.WriteLine("Enter Message: ");
                Console.WriteLine("(hit enter to send email)");

                string msg;

                msg = Console.ReadLine();

                //setting up mail
                MailMessage mail = new MailMessage();
                //sender mail
                mail.From = new MailAddress("testdenk@gmx.at");
                //receiver mail
                mail.To.Add("fabian-denk@gmx.at");

                //Setting up smtp client with gmx settings
                SmtpClient client = new SmtpClient();
                client.Port = 25;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                //gmx host
                client.Host = "mail.gmx.net";
                //building subject with neccessary information
                mail.Subject = "Bug Report " + DateTime.Now.ToShortDateString() + ", " + DateTime.Now.ToShortTimeString();
                //building main message with neccessary information
                mail.Body = "Report sent from: " + emailAdress + "\r\n\r\n" + msg + "\r\n\r\n" +mfg;
                //enabling ssl secure connection, important for most hosts
                client.EnableSsl = true;
                //setting email and password from sender email
                client.Credentials = new NetworkCredential("testdenk@gmx.at", "Denk)71)test");

                //send mail
                client.Send(mail);

                Console.WriteLine();
                Console.WriteLine("Email has been sent!");
                Console.ReadLine();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
            return true;
        }

        #endregion

        #region initialize defaultpaths
        public static string sourcepathSave;
        public static string destinationpathSave;

        public static void initializeDefaultpaths()
        {
            //Defaultpath for  Sourcepathsave
            sourcepathSave = Directory.GetCurrentDirectory();
            sourcepathSave = Path.Combine(sourcepathSave, "srcPathSave.txt");
            //Console.WriteLine(sourcepathSave);

            //Defaultpath for  Destinationpathsave
            destinationpathSave = Directory.GetCurrentDirectory();
            destinationpathSave = Path.Combine(destinationpathSave, "dstpathSave.txt");
            //Console.WriteLine(destinationpathSave);
        }
        #endregion

        #region ChangePaths
        public static bool ChangeSourcepath()
        {
            Console.WriteLine("Change Sourcepath");
            Console.WriteLine(@"Example: C:\Test\TestfolderWithExcel\");
            Console.WriteLine();

            #region Streamvariables
            FileStream fs1 = null;
            StreamReader sr = null;
            FileStream fs2 = null;
            StreamWriter sw = null;
            #endregion

            try
            {
                //open existing file
                fs1 = new FileStream(sourcepathSave, FileMode.Open);
                sr = new StreamReader(fs1);
                string oldPath = sr.ReadLine();
                //printing old path
                Console.WriteLine("Old path: " + oldPath);
            }
            catch(Exception ex)
            {
                //if no file exists (first startup)
                Console.WriteLine("Old path: none, file will be created");
            }
            Console.WriteLine();

            //closing old stream so that the file isnt used anymore
            if (sr != null)
            {
                sr.Close();
            }

            //read new path
            Console.Write("new path: ");
            string newPath = Console.ReadLine();
                        
            try
            {
                //open existing file 
                //create new file (old one doesnt matter)
                fs2 = new FileStream(sourcepathSave, FileMode.Create);
                sw = new StreamWriter(fs2);
                //write new path into file
                sw.WriteLine(newPath);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }

            if(sw != null)
            {
                sw.Close();
            }

            return true;
        }

        public static bool ChangeDestinationpath()
        {
            Console.WriteLine("Change Sourcepath");
            Console.WriteLine(@"Example: C:\Test\TestfolderWithExcel\collectedData.xlsx");
            Console.WriteLine();

            #region Streamvariables
            FileStream fs1 = null;
            StreamReader sr = null;
            FileStream fs2 = null;
            StreamWriter sw = null;
            #endregion

            try
            {
                //open existing file
                fs1 = new FileStream(destinationpathSave, FileMode.Open);
                sr = new StreamReader(fs1);
                string oldPath = sr.ReadLine();
                //printing old path
                Console.WriteLine("Old path: " + oldPath);
            }
            catch (Exception ex)
            {
                //if no file exists (first startup)
                Console.WriteLine("Old path: none, file will be created");
            }
            Console.WriteLine();

            //closing old stream so that the file isnt used anymore
            if (sr != null)
            {
                sr.Close();
            }

            //read new path
            Console.Write("new path: ");
            string newPath = Console.ReadLine();

            try
            {
                //open existing file 
                //create new file (old one doesnt matter)
                fs2 = new FileStream(destinationpathSave, FileMode.Create);
                sw = new StreamWriter(fs2);
                //write new path into file
                sw.WriteLine(newPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }

            if (sw != null)
            {
                sw.Close();
            }

            return true;
        }
        #endregion

        #region getPaths
        public static string getSrcPath()
        {
            string path = null;

            FileStream fs = null;
            StreamReader sr = null;

            try
            {
                //open stream
                 fs = new FileStream(sourcepathSave, FileMode.Open);
                 sr = new StreamReader(fs);

                //read path from file
                path = sr.ReadLine();
            }
            catch(Exception ex)
            {
                //if -> no path has been found, file probably doesnt exist
                Console.WriteLine("No Sourcepath found!");
                return null;
            }

            //closing streams, freeing resources
            finally
            {
                if (sr != null) {
                    sr.Close();
                }
            }

            return path;
        }

        public static string getDstPath()
        {
            string path = null;

            FileStream fs = null;
            StreamReader sr = null;

            try
            {
                //open stream
                fs = new FileStream(destinationpathSave, FileMode.Open);
                sr = new StreamReader(fs);

                //read path from file
                path = sr.ReadLine();
            }
            catch (Exception ex)
            {
                //if -> no path has been found, file probably doesnt exist
                Console.WriteLine("No Destinationpath found!");
                return null;
            }

            //closing streams, freeing resources
            finally
            {
                if (sr != null)
                {
                    sr.Close();
                }
            }

            return path;
        }

        #endregion

        #region readExcel

        public static int successful = 0;
        public static int count = 0;
        public static string[] errorFiles = null;
        public static int errorCount = 0;

        //src: file path where the path for the source excel files is saved
        //dst: file path where the path for the destination excel file is saved
        public static bool readAndWriteExcelFiles(string src, string dst)
        {
            successful = 0;
            count = 0;
            ExcelObj dstExcel = null;
            errorFiles = new string[10];
            try
            {

                //checking if both paths were entered
                if (src == null || dst == null)
                {
                    return false;
                }

                //get all the .xlsx files in the folder
                DirectoryInfo d;
                FileInfo[] files;
                try
                {
                    //getting every excel file needed
                    d = new DirectoryInfo(src);
                    files = d.GetFiles("*.csv");
                    errorFiles = new string[files.Count<FileInfo>()];
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("error while scanning direcotry");
                    return false;
                }

                count = files.Count<FileInfo>();

                //open Destination excel file
                dstExcel = new ExcelObj(getDstPath(), true);
                int oldRow = dstExcel.row;

                foreach (FileInfo file in files)
                {
                    //send each file to reader function
                    if (readExcelFile(file, dst, dstExcel, count))
                    {
                        successful++;
                    }
                }

                //release every process needed for the destination excel
                dstExcel.deleteCells(oldRow);

                Console.WriteLine();
                Console.WriteLine("Error trying to read these Files: ");
                foreach (string errorfile in errorFiles)
                {
                    //writing every failed excel file to the console
                    if (errorfile != null)
                    {
                        Console.WriteLine(errorfile);
                    }
                }
                Console.WriteLine();
                Console.WriteLine(successful + " of " + count + " have been read and re-saved");
                Console.WriteLine("Hit save on the Windows Save Prompt to continue. (if you dont see it use alt+tab to look for it)");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(errorFiles.Count<string>());
                Console.WriteLine("Big error");
            }
            finally
            {
                if (dstExcel != null)
                {
                    Console.WriteLine("trying to free data collection excel");
                    dstExcel.Free(true);
                }
            }
            return true;
        }

        //fi: fileinfo of a single file (path, and so on..)
        public static bool readExcelFile(FileInfo fi, string dstpath, ExcelObj dst, int count)
        {

            string filePath = "not initialized";
            int internalCount;
            Dictionary<string, string> sysInfo;
            ExcelObj src = null;
            try
            {
                //full path of a xlsx file
                filePath = fi.FullName;
                internalCount = 1;
                Console.WriteLine(filePath);

                sysInfo = new Dictionary<string, string>();

                src = new ExcelObj(filePath, false);
                Console.WriteLine("Rows: " + src.row + ", Column: " + src.column);

                #region readingSysInfo

                //reading everything that is on these fixed cells (most of the time)

                sysInfo.Add("IP-Name", src.getValue(1, 1));

                sysInfo.Add("Benutzer", src.getValue(8, 1));

                sysInfo.Add("CPU", src.getValue(3, 1));

                sysInfo.Add("Typenbezeichner", src.getValue(5, 1));

                sysInfo.Add("Seriennummer", src.getValue(10, 1));
                
                sysInfo.Add("Massenspeicher", src.getValue(12, 1) + " " + src.getValue(14, 1));

                
                /*
                sysInfo.Add("IP-Adresse", src.getValue(28, 1));
                */

                //filtering ip adress -> not in the same cell every time
                string value = "empty";
                for(int i = 20; i <= 40; i++)
                {
                    
                    value = src.getValue(i, 1);
                    
                    if (value != null)
                    {
                        //looking for possible ip adresses
                        if (value.Contains("IPv4-Adresse") && (value.Contains("192.") || value.Contains("169.") || value.Contains("10.")))
                        {
                            //trimming ip adress, neccessary because a lot of useless information is attached here
                            value = src.getValue(i, 1);
                            string[] seperator = new string[1];
                            seperator[0] = ":";
                            string[] trimmed = value.Split(seperator, StringSplitOptions.RemoveEmptyEntries);
                            
                            value = trimmed[1];
                            break;
                        }
                    }
                    else
                    {
                        value = null;
                    }
                }
                sysInfo.Add("IP-Adresse", value);

                #endregion

                //writing to the console every read info, debug purposes
                foreach (KeyValuePair<string, string> kvp in sysInfo)
                {
                    if (kvp.Key != null && kvp.Value != null)
                    {
                        Console.WriteLine(kvp.Key + ", " + kvp.Value);
                    }

                }
                //free excel and free resources
                //src.Free(true);

                #region saving in Summary
                //save to big summary
                //IP-Name
                //validating everything and writing it into the summary excel
                Console.WriteLine(dst.row + 1);
                if (sysInfo["IP-Name"] != null)
                {
                    //writing to excel
                    Console.WriteLine("Saved IP-name");
                    dst.saveValue(dst.row + 1, 9, sysInfo["IP-Name"]);
                }
                else
                {
                    //writing none to excel
                    dst.saveValue(dst.row + 1, 9, "none");
                }
                //CPU
                if (sysInfo["CPU"] != null)
                {
                    Console.WriteLine("Saved CPU");
                    dst.saveValue(dst.row + 1, 10, sysInfo["CPU"]);
                }
                else
                {
                    dst.saveValue(dst.row + 1, 10, "none");
                }
                //Typenbezeichner
                if (sysInfo["Typenbezeichner"] != null)
                {
                    Console.WriteLine("Saved typenbezeichner");
                    dst.saveValue(dst.row + 1, 7, sysInfo["Typenbezeichner"]);
                }
                else
                {
                    dst.saveValue(dst.row + 1, 7, "none");
                }
                //Benutzer
                if (sysInfo["Benutzer"] != null)
                {
                    Console.WriteLine("Saved Benutzer");
                    dst.saveValue(dst.row + 1, 4, sysInfo["Benutzer"]);
                }
                else
                {
                    dst.saveValue(dst.row + 1, 4, "none");
                }
                //Seriennummer
                if (sysInfo["Seriennummer"] != null)
                {
                    Console.WriteLine("Saved Seriennummer");
                    dst.saveValue(dst.row + 1, 8, sysInfo["Seriennummer"]);
                }
                else
                {
                    dst.saveValue(dst.row + 1, 8, "none");
                }
                //Massenspeicher
                if (sysInfo["Massenspeicher"] != null)
                {
                    Console.WriteLine("Saved massenspeicher");
                    dst.saveValue(dst.row + 1, 12, sysInfo["Massenspeicher"]);
                }
                else
                {
                    Console.WriteLine("no mass storage");
                    dst.saveValue(dst.row + 1, 12, "none");
                }
                //IP-Adresse
                if (sysInfo["IP-Adresse"] != null)
                {
                    dst.saveValue(dst.row + 1, 13, sysInfo["IP-Adresse"]);
                }
                else
                {
                    Console.WriteLine("no ip adress");
                    dst.saveValue(dst.row + 1, 13, "none");
                }

                dst.IncrementRow();
                #endregion

                internalCount++;
                return true;
            }
            catch(Exception ex)
            {
                errorCount++;
                errorFiles[errorCount] = filePath;

                /*
                if(src != null)
                {
                    Console.WriteLine("tried to error free");
                    src.Free(true);
                }
                */
                Console.WriteLine("Error Single Read and Save, " + ex.Message);
                Console.WriteLine("error file: " + filePath);
            }
            finally
            {
                if(src != null)
                {
                    Console.WriteLine("trying to free single read");
                    src.Free(true);
                }
            }
            return false;
        }

        public class ExcelObj
        {
            public Excel.Application xApp { get; set; }
            public Excel.Workbook xWorkbook { get; set; }
            public Excel._Worksheet xWorksheet { get; set; }
            public Excel.Range xRange { get; set; }

            public int row { get; set; }
            public int column { get; set; }

            //constructor
            public ExcelObj(string dst, bool isDst)
            {
                xApp = new Excel.Application();
                xWorkbook = xApp.Workbooks.Open(dst);
                xWorksheet = xWorkbook.Sheets[1];
                xRange = xWorksheet.UsedRange;

                //row = xWorksheet.UsedRange.Rows.Count;
                //column = xWorksheet.UsedRange.Columns.Count;

                //row = xWorksheet.Cells.Find("*", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, false, Type.Missing).Row;
                //if the excel file is the destination excel -> find last index (library returns shit sometimes)
                if (isDst)
                {
                    row = this.getRow();
                    Console.WriteLine("getRow used");
                }
                //else use the default value library returns
                else
                {
                    row = xWorksheet.UsedRange.Rows.Count;
                    Console.WriteLine("normal rows used");
                }
                
                //initialized but never used i think
                column = xWorksheet.Cells.Find("*", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, false, Type.Missing).Column;

                Console.WriteLine("Row: " + row);
            }

            //finding last index of the destination excel file
            private int getRow()
            {
                int columnIndex = 9;
                bool found = false;
                int iterator = 4;
                string name = "empty";
                int starterRow = 4;
                int checksum = 0;
                //search over IP-Name
                Console.WriteLine("Enter search row loop");
                while (!found)
                {
                    //if cell a and cell a+1 are empty -> last index (where the cell has no value) is found
                    name = this.getValue(iterator, columnIndex);
                    if (name == null)
                    {
                        //check if this one cell has no value
                        Console.WriteLine("first if");
                        checksum++;
                    }
                    name = this.getValue(iterator + 1, columnIndex);
                    if (name == null)
                    {
                        //checking if next one has no value
                        Console.WriteLine("second if");
                        checksum++;
                    }
                    if(checksum == 2)
                    {
                        //if both had no value -> last index found
                        Console.WriteLine("row found, exiting loop");
                        //setting flag to exit loop
                        found = true;
                        //saving index of last row
                        starterRow = iterator;
                    }
                    else
                    {
                        //if only one has no value or both have actual value 
                        //checksum gets reset and search continues
                        Console.WriteLine("Checksum reset");
                        checksum = 0;
                    }
                    //iterator is the row counter
                    iterator++;
                }

                Console.WriteLine("StarterRow = " + starterRow);
                return starterRow;
            }

            //deleting range in a certain range
            public void deleteCells(int r)
            {
                try
                {
                    Console.WriteLine("trying to delete correct range");
                    //building range for delete
                    string rangeToDel = "A4:";
                    if(r < 4)
                    {
                        Console.WriteLine("index too small!");
                        return;
                    }
                    //building string
                    string lastRow = r.ToString();
                    rangeToDel = rangeToDel + "R" + lastRow;
                    Console.WriteLine("Range: " + rangeToDel);

                    //get the correct range
                    Excel.Range toDel = xWorksheet.get_Range(rangeToDel);
                    Console.WriteLine("got range object");

                    //delete range and shift cells up
                    toDel.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    Console.WriteLine("Range deleted");
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.WriteLine("Error trying to delete range");
                }
            }

            public void IncrementRow()
            {
                //well, incrementing row
                row++;
            }

            public void IncrementColumn()
            {
                //well, incrementing column, never used i think
                column++;
            }

            //getting a value in a certain cell
            public string getValue(int r, int c)
            {
                return xRange.Cells[r, c].Value2;
            }

            //saving a value to a certain cell
            public void saveValue(int r, int c, string value)
            {
                xRange.Cells[r, c] = value;
            }

            //releasing every resource needed earlier
            public void Free(bool last)
            {
                //releasing everything
                //trying to call free in every case so no corpse is left in the background
                //if the console windows is terminated during collecting data by using the "X" button (dont do that)
                //non excpected exit -> corpses will be in the background and must be freed through task manager
                try
                {
                    //cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    //release com objects to fully kill excel process from running in the background
                    Marshal.ReleaseComObject(xRange);
                    Marshal.ReleaseComObject(xWorksheet);

                    //close and release
                    xWorkbook.Close();
                    Marshal.ReleaseComObject(xWorkbook);
                    //quit and release
                    xApp.Quit();
                    Marshal.ReleaseComObject(xApp);
                    Console.WriteLine("Excel freed");
                }
                catch(Exception ex)
                {
                    //yeah use task manager your only hope
                    Console.WriteLine("Error trying to free!! Try manually in TaskManager!");
                }
            }
        }

        #endregion

        #region show Data

        public static bool showData()
        {
            //get the neccessary paths
            string src = getSrcPath();
            string dst = getDstPath();

            //checking if both paths are available
            if(src == null || dst == null)
            {
                //probably no files with paths
                return false;
            }

            //writing everything to the user
            Console.WriteLine("Data: ");
            Console.WriteLine();
            Console.WriteLine("Source Folder: " + src);
            Console.WriteLine("Destination File: " + dst);
            return true;
        }

        #endregion
    }
}