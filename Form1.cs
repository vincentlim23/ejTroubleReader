using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Runtime.InteropServices;


namespace EJTroubleReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
              

        private void btnEnter_Click(object sender, EventArgs e)
        {
            //string input;
            string browsepath;
            browsepath = path_dir.Text.Trim();
            if (browsepath == "Browse your directory.")
            {
                MessageBox.Show("Please set your path correctly.");

            }
            //input = inputBox.Text.Trim();
            else
            {
                // disbale button
                btnEnter.Enabled = false;
                btnExcel.Enabled = false;
                excel_dir.Enabled = false;
                path_dir.Enabled = false;
                

                string[] lines;
                string[] total;
                string[] filteredTotal;
                var list = new List<string>();
                var listDate = new List<string>();
                var listTotal = new List<string>();
                var filterTotal = new List<string>();
                // System.IO.File.Delete(@"C:\Users\sw_sychan\Desktop\EJTroubleReader\EJ November\READ\T"+input+".txt");

                // Check if the /READ folder exists, else create it
                System.IO.Directory.CreateDirectory(@browsepath + "/READ");
                
                // Delete all files inside before create/append txt.
                System.IO.DirectoryInfo di = new DirectoryInfo(@browsepath + "/READ");
                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }               
                String sourceDirectory = browsepath;
               //var txtFiles = Directory.EnumerateFiles(sourceDirectory, "*Test-*");
               var txtFiles = Directory.EnumerateFiles(sourceDirectory, "*.txt*");

                foreach (var currentfile in txtFiles)
                {
                    list.Clear(); // to clear the list to avoid repeating 
                    listTotal.Clear();
                    filterTotal.Clear();

                    string[] filename = null;
                    if (currentfile.Contains("-"))
                    {
                        filename = currentfile.Replace(@browsepath + "\\", "").Split('-');
                    }
                    else
                    {
                        filename = currentfile.Replace(@browsepath + "\\", "").Split('_');
                    }
                        var fileStream = new FileStream(currentfile, FileMode.Open, FileAccess.Read);
                        using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
                        {
                            string line = "";
                            while ((line = streamReader.ReadLine()) != null)
                            {
                                list.Add(line);
                            }
                        }
                        lines = list.ToArray();

                        for (int i = 0; i < lines.Length; i++)
                        {                                        
                            //
                            // Print the date and time
                            // 
                            string strDate = "";
                            if (lines[i].Contains("----------201"))
                            {
                                strDate = lines[i];
                                listTotal.Add(strDate);
                            }
                  
                            //
                            // **POWER ON/OFF
                            //
                            string strPower = "";
                            if (lines[i].Contains("**POWER"))
                            {
                                strPower = lines[i];
                                listTotal.Add(strPower);
                            }
                            //
                            // find **SUPERVISOR 
                            //
                            string strSuper = "";
                            if (lines[i].Contains("**SUPERVISOR "))
                            {
                                strSuper = lines[i];
                                listTotal.Add(strSuper);
                            }
                            //
                            // find    IN-SERVICE OK / OUT-OF-SERVICE                         
                            //
                            string strINS = "";
                            if (lines[i].Contains("IN-SERVICE OK"))
                            {
                                strINS = lines[i - 5]  + lines[i];
                                listTotal.Add(strINS);
                            }
                            string strOOS = "";
                            if (lines[i].Contains("OUT-OF-SERVICE"))
                            {
                                strOOS = lines[i];
                                listTotal.Add(strOOS);                            
                            }

                            /*
                            // find INIT CASH:
                            // TEMPORARY REMOVE
                           string strInit = "";
                            if (lines[i].Contains("INIT CASH:"))
                            {
                                strInit = lines[i] + "\n" + lines[i + 1] + "\n" + lines[i + 2];
                                listTotal.Add(strInit);
                            }
                           */

                            // skip exceptional log
                            if (lines[i].Contains("EXCEPTIONAL LOG"))
                            {
                                do
                                {
                                    i++;
                                } while (lines[i] != " ");                            
                            
                            }
                       
                            // 
                            //find *W
                            string strW = "";
                            //if (lines[i].Contains("*W") && !lines[i].Contains("++") && !lines[i].Contains("+0"))
                            if (lines[i].Contains("*W"))
                            {
                                if (lines[i].Contains("*W2-CASSETTE FULL") && lines[i + 1].Contains("----------201"))
                                { 
                                    // take the time only.
                                    string[] time = lines[i+1].Replace("----------","").Split(' ');
                                    strW = time[1].ToString() + "         " + lines[i]; 
                                }
                                else if (i != 0 && lines[i - 1].Contains("["))
                                {
                                    strW = lines[i - 1] + "\n" + lines[i];
                                }
                                else if (lines[i].Contains("*W2-CASSETTE FULL") && lines[i - 1].Contains("**RESET") && lines[i + 1].Contains("**CONFIGURATION ID"))
                                {
                                    strW = lines[i - 1] + "\n" + lines[i] + "\n" + lines[i + 1];
                                }
                                else
                                {
                                    strW = lines[i];
                                }
                                listTotal.Add(strW);
                            }
                            //
                            // find **RETRACT
                            string strRET = "";
                            if (lines[i].Contains("**RETRACT"))
                            {
                                strRET = lines[i].Replace("**", "");
                                listTotal.Add(strRET);
                            }

                            //
                            //find *E4
                            string strE = "";
                            //if (lines[i].Contains("*E4") && !lines[i].Contains("++") && !lines[i].Contains("+0"))
                            if(lines[i].Contains("*E4"))
                            {
                                if (lines[i - 2].Contains("["))
                                {                                       
                                    strE = lines[i - 2].Split(']')[0].Replace("[","") + "   " + lines[i];
                                }
                                else if (lines[i - 1].Contains("**RESET"))
                                {
                                    string[] gettime = lines[i - 1].Split(' ');
                                    strE = lines[i - 1] + "\n" +gettime[20].Trim()+ lines[i] ;                                
                                }
                                else
                                {
                                    strE = lines[i];
                                }
                                // listE.Add(strE);
                                listTotal.Add(strE);
                            }

                            //
                            // find **TROUBLE
                            // not State table configuration error
                            //
                            string strTrouble = "";
                            if (lines[i].Contains("**TROUBLE"))
                            {
                                if (!lines[i+1].Contains("State table configuration error ")) 
                                {
                                    if ( lines[i+1].Contains("SAM module reset failure") && lines[i - 2].Contains("**RESET") && i !=1 )
                                    {
                                         strTrouble =lines[i-2]+ lines[i - 1] + "\n" + lines[i] + lines[i + 1];
                                    }
                                    else if (lines[i + 1].Contains("The specified cash kind doesn't exist in"))
                                    {
                                        strTrouble = lines[i] + lines[i + 1] + lines[i + 2];                                    
                                    }
                                    else
                                    {
                                        strTrouble = lines[i] + lines[i + 1];
                                    }
                                   
                                    listTotal.Add(strTrouble);
                                }
                            }

                            //
                            //COMMS LINE DOWN
                            //
                            string strCOMM = "";
                            if (lines[i].Contains("COMMS LINE DOWN"))
                            {
                                strCOMM = lines[i];
                                listTotal.Add(strCOMM);
                            }
                            //
                            //find *AUTO RESET FAILURE
                            //if (lines[i].Contains("*AUTO RESET FAILURE**") && !lines[i].Contains("++") && !lines[i].Contains("+0"))
                            if (lines[i].Contains("*AUTO RESET FAILURE**"))
                            {
                                string strAuto = lines[i];
                                // listAutoReset.Add(strAuto);
                                listTotal.Add(strAuto);
                            }

                            //commencing sw update
                            string strUpdate = "";
                            if (lines[i].Contains("**COMMENCING SW UPDATE"))
                            {
                                if ((i + 1) != lines.Length) 
                                {
                                    strUpdate = lines[i] + "\n" + lines[i + 1] + "\n" + lines[i + 2] + "\n" + lines[i + 3];
                                    listTotal.Add(strUpdate);
                                }
                                
                            }

                            if (i == (lines.Length - 1) && lines[i].Contains("*W3-"))
                            {
                                listTotal.Add("23:59:59.END OF DAY");  // To indicate the end of a Day
                            }
        
                        }
                        // end of searchandmatch for loop                       
                        total = listTotal.ToArray();
                       
                        // remove the unwanted dates
                        for (int i = 0; i < total.Length; i++)
                        {   
                            string filterStr= "";
                            if (i ==0)
                            {
                                filterStr = total[i];
                                filterTotal.Add(filterStr);
                            }                     
                           else
                            {
                               
                                // remove business program detection Error 
                               if (total[i].Contains("BUSINESS PROGRAM DETECTION ERROR") || total[i].Contains("Browser display error"))
                                {
                                    // do nth
                                }

                                else if (total[i - 1].Contains("**TROUBLE") && total[i].Contains("BUSINESS PROGRAM DETECTION ERROR"))
                                {
                                    // do nth
                                }

                                else if (total[i - 1].Contains("----------201") && total[i].Contains("----------201"))
                                {
                                    // do nth
                                }
                                else
                                {
                                    filterStr = total[i];
                                    filterTotal.Add(filterStr);
                                }
                            }
                        
                        } // end of remove unwanted data
                        filteredTotal = filterTotal.ToArray();

                         //   if (total.Length.ToString() != "0")
                             if (filteredTotal.Length.ToString()!= "0")
                            {
                                System.IO.File.AppendAllLines(@browsepath + "/READ/" + filename[0].ToString() + ".txt", filteredTotal);                  
                                //System.IO.File.AppendAllLines(@browsepath + "/READ/" + filename[0].ToString() + ".txt", total);
                            }

                    } // end of  foreach loop                               

                // enable button 
                btnEnter.Enabled = true;
                btnExcel.Enabled = true;
                excel_dir.Enabled = true;
                path_dir.Enabled = true;

            }// end else 
        }
        //
        //  Display Excel Sheet
        //
        private void btnExcel_Click(object sender, EventArgs e)
        {
            string browsepath;
            browsepath = excel_dir.Text.Trim();
            if (browsepath == "Browse your directory.")
            {
                MessageBox.Show("Please set your path correctly.");
            }
            else
            {
                // disable button after click
                btnExcel.Enabled = false;
                btnEnter.Enabled = false;
                excel_dir.Enabled = false;
                path_dir.Enabled = false;

                string[] lines;
                var list = new List<string>();
                string start = DateTime.Now.ToString("h:mm:ss tt");

               // String sourceDirectory = @"C:\Users\sw_sychan\Desktop\EJTroubleReader\EJ November\Read";
                String sourceDirectory = browsepath;
                
                // delete the old excel file 
                //System.IO.File.Delete(@"C:\Users\sw_wllim\Desktop\EJTroubleReader\Result.xlsx"); 
                
               //var txtFiles = Directory.EnumerateFiles(sourceDirectory, "*T347.txt");
                var txtFiles = Directory.EnumerateFiles(sourceDirectory, "*.txt*"); 
                            
                var excelApp = new Excel.Application();
               // excelApp.Visible = true;             
                
                excelApp.Workbooks.Add();
                Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;               
                excelApp.Visible = false;
                                               
                // Excel.Worksheet workSheet2 = (Excel.Worksheet)excelApp.ActiveSheet;

                foreach (string currentfile in txtFiles)
                {
                    list.Clear(); // to clear the list to avoid repeating
                    var fileStream = new FileStream(currentfile, FileMode.Open, FileAccess.Read);
                    using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
                    {
                        string line;
                        while ((line = streamReader.ReadLine()) != null)
                        {
                            list.Add(line);
                        }
                    }

                    lines = list.ToArray();

                    var row = 2;
                   // var rowsSheet2 = 2

                    //add to new sheet
                    workSheet = (Excel.Worksheet)excelApp.Worksheets.Add();
                  
                  // workSheet2 = (Excel.Worksheet)excelApp.Worksheets.Add();
                                        

                    //change sheet name
                    string[] sheetname = currentfile.Split(Path.DirectorySeparatorChar);
                    workSheet.Name = sheetname[sheetname.Length - 1].Replace(".txt", "");
             //       workSheet2.Name = sheetname[sheetname.Length - 1].Replace(".txt", "Denote");

                    workSheet.Cells[1, "A"] = "Device ID";
                    workSheet.Cells[1, "B"] = "Date and Time";
                    workSheet.Cells[1, "C"] = "Events";
                    workSheet.Cells[1, "D"] = "Time of Occurence:";
                    workSheet.Cells[1, "E"] = "SOP Mode";
                    workSheet.Cells[1, "F"] = "W3/W7/W9-Error";
                    workSheet.Cells[1, "G"] = "Dispenser Faulty";
                    workSheet.Cells[1, "H"] = "Receipt Error";
                    workSheet.Cells[1, "I"] = "Card Reader Error";
                    workSheet.Cells[1, "J"] = "SAM module/chip error";

       /*
                    // workSheet2
                    workSheet2.Cells[1, "A"] = "Device ID";
                    workSheet2.Cells[1, "B"] = "CDM-RM10";
                    workSheet2.Cells[1, "C"] = "CDM-RM20";
                    workSheet2.Cells[1, "D"] = "CDM-RM50";
                    workSheet2.Cells[1, "E"] = "CDM-RM100";
                    workSheet2.Cells[1, "F"] = "ATM-RM50";
                    workSheet2.Cells[1, "G"] = "ATM-RM100";
       */
                    workSheet.Cells[2, "A"] = sheetname[sheetname.Length - 1].Replace(".txt", "");
       //             workSheet2.Cells[2, "A"] = sheetname[sheetname.Length - 1].Replace(".txt", "");                   
                   
                    
                    for (int i = 1; i < lines.Length; i++)
                    {
                        //power on/off                 
                        if (lines[i].Contains("**POWER"))
                        {

                            if (lines[i - 1].Contains("----------201"))
                            {
                                workSheet.Cells[row, "B"] = lines[i - 1].Replace("-", "");
                            }
                            if (lines[i].Contains("**POWER ON"))
                            {
                                workSheet.Cells[row, "C"] = "POWER ON";
                            }
                            else
                            {
                                workSheet.Cells[row, "C"] = "POWER OFF";
                            }
                            // GET TIME
                            string check = lines[i].Trim().Replace(" ", "");
                            workSheet.Cells[row, "D"] = check.Substring(check.Length - 8);
                            row++;
                        }

                       // trouble          
                       // cashjam = E                       
                       // F=Dispenser Faulty
                       // G=Casette Full
                       // SAM module/chip error= H 

                        else if (lines[i].Contains("**TROUBLE"))
                        {

                            if (lines[i - 1].Contains("----------201"))
                            {
                                workSheet.Cells[row, "B"] = lines[i - 1].Replace("-", "");
                            }
                            //record first trouble downtime,                   
                            //except State table configuration error && cash handler unit
                            //and line before is not Trouble and *W and *E4 and **Power OFf and in-service-ok

                            if (!lines[i].Contains("CASH HANDLER UNIT;") && !lines[i].Contains("Card length abnormal"))
                            {
                                //if (!lines[i - 1].Contains("**TROUBLE") && !lines[i - 1].Contains("*W") && !lines[i - 1].Contains("*E4") && !lines[i - 1].Contains("**POWER OFF"))
                                if ( !lines[i - 1].Contains("*E4") && !lines[i - 1].Contains("**POWER OFF"))
                                {
                                    // Get time
                                    string[] downtime = lines[i].Replace("**TROUBLE", "").Split(' ');
                                    if (downtime[10] == "" && downtime[11] != "" )
                                    {
                                        workSheet.Cells[row, "D"] = downtime[11];
                                    }
                                    else {
                                        workSheet.Cells[row, "D"] = downtime[10];
                                    }                                    
                                }
                            }
                            workSheet.Cells[row, "C"] = lines[i].Remove(0, 35).TrimStart();
                            row++;
                        }

                        //in-service ok
                        else if (lines[i].Contains("DATE:"))
                        {
                               string dateParse = "";
                                string[] getDate = lines[i].Split(' ');
                                dateParse = getDate[0].Replace("DATE:", "");
                                DateTime dt = DateTime.ParseExact(dateParse, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                workSheet.Cells[row, "B"] = dt.ToShortDateString() + "  " + getDate[7].Replace("TIME:", "");                                                       

                                workSheet.Cells[row, "C"] = "IN-SERVICE OK";
                                // get time 
                                workSheet.Cells[row, "D"] = getDate[7].Replace("TIME:", "");
                                row++;                      
                                                  
                        }

                        // out of service 
                        else if (lines[i].Contains("OUT-OF-SERVICE"))
                        {
                            if (lines[i - 1].Contains("----------201"))
                             {
                                 workSheet.Cells[row, "B"] = lines[i - 1].Replace("-", "");
                             }
                            workSheet.Cells[row, "C"] = lines[i].Trim();

                            //GET TIME
                            workSheet.Cells[row,"D"] = lines[i-1].Split(' ')[1].Replace("-","");
                            row++;
                        }

                        //supervisor mode
                        else if (lines[i].Contains("**SUPERVISOR"))
                        {
                            if (lines[i - 1].Contains("----------201"))
                            {
                                workSheet.Cells[row, "B"] = lines[i - 1].Replace("-", "");
                            }
                            else if (!lines[i - 1].Contains("----------2016/11") && lines[i + 1].Contains("----------2016/11"))
                            {
                                workSheet.Cells[row, "B"] = lines[i + 1].Replace("-", "");
                            }                           

                            // record supervisor downtime
                            string[] downtime = lines[i].Split(' ');
                            workSheet.Cells[row, "C"] = downtime[0].Replace("**", "") + " " + downtime[1] + " " + downtime[2];
                            if (lines[i].Contains("**SUPERVISOR MODE EXIT"))
                            {
                                workSheet.Cells[row, "D"] = downtime[7];
                            }
                            else
                            {
                                workSheet.Cells[row, "D"] = downtime[6];
                            }
                            row++;
                        }

        /*                    //init cash
                        else if (lines[i].Contains("INIT CASH:"))
                        {
                            if (lines[i - 1].Contains("----------2016/11"))
                            {
                                workSheet.Cells[row, "B"] = lines[i - 1].Replace("-", "");
                            }
                            //workSheet.Cells[row, "G"] = lines[i + 1] + lines[i + 2];
                            workSheet.Cells[row, "C"] = lines[i + 1] + " " + lines[i + 2].Trim();
                            row++;
                        }
         */
                        //*W
                        else if (lines[i].Contains("*W"))
                        {
                            if (lines[i - 1].Contains("----------201"))
                            {
                                workSheet.Cells[row, "B"] = lines[i - 1].Replace("-", "");
                            }                                                    
                            string[] split = lines[i].Replace("+", "").Split('*');
                            workSheet.Cells[row, "C"] = split[split.Length - 2];
                            row++;
                        }

                        //* E4
                        else if (lines[i].Contains("*E4"))
                        {
                            if (lines[i - 1].Contains("----------201"))
                            {
                                workSheet.Cells[row, "B"] = lines[i - 1].Replace("-", "");
                            }
                            
                           // if (lines[i - 1].Contains("["))
                           // {   string[] stre4 = lines[i - 1].Split(']');
                           //     workSheet.Cells[row, "D"] = stre4[0].Replace("[", "");}
                            
                            string[] getTime = lines[i].Split(' '); 
                            if(getTime[0].Contains("*E4"))
                            {
                                workSheet.Cells[row, "D"] = getTime[22];
                            }
                            else 
                            {
                                workSheet.Cells[row, "D"] = getTime[0];
                            }
                            
                            string[] split = lines[i].Replace("+", "").Split('*');
                            workSheet.Cells[row, "C"] = split[split.Length - 2];
                            row++;
                        }

                        // **RETRACT
                        else if (lines[i].Contains("RETRACT"))
                        {
                            if (lines[i - 1].Contains("----------201"))
                            {
                                workSheet.Cells[row, "B"] = lines[i - 1].Replace("-", "");
                            }                          
                            string check = lines[i].Trim().Replace(" ", "");
                            string[] retractStr = lines[i].Split(' ');
                            workSheet.Cells[row, "C"] = "RETRACT";
                            
                            if (lines[i - 1].Contains("**SUPERVISOR"))
                            {
                                workSheet.Cells[row, "D"] = check.Substring(check.Length - 8);
                            }
                            else if(retractStr[retractStr.Length-8] != "")
                            {
                                workSheet.Cells[row, "D"] = retractStr[retractStr.Length- 8];                               
                            }
                            row++;
                        }

                        // *Auto Reset Failure
                        else if (lines[i].Contains("*AUTO RESET FAILURE**"))
                        {
                            if (lines[i - 1].Contains("----------201"))
                            {
                                workSheet.Cells[row, "B"] = lines[i - 1].Replace("-", "");
                            }                                    
                            workSheet.Cells[row, "C"] = "AUTO RESET FAILURE";
                            row++;
                        }

                        // COMMS LINE DOWN
                        else if (lines[i].Contains("COMMS LINE DOWN"))
                        {
                            if (lines[i - 1].Contains("----------201"))
                            {
                                workSheet.Cells[row, "B"] = lines[i - 1].Replace("-", "");
                            }                          
                            workSheet.Cells[row, "C"] = lines[i];
                            row++;
                        }               
                        
                        // END OF DAY 
                        else if (lines[i].Contains("END OF DAY"))
                        {
                            string[] strEOD = lines[i].Split('.');
                            workSheet.Cells[row, "C"] = strEOD[1];
                            workSheet.Cells[row, "D"] = strEOD[0];
                            row++;
                        
                        }

            
                    }// end of for loop   

                    //calculate  
                    int lastrow = workSheet.UsedRange.Row + workSheet.UsedRange.Rows.Count - 1;
                    string firstOccur = "";
                    string secondOccur = "";
                   // string EOD = "";
                    string totalE = "";                    
                    TimeSpan totalDurE = new TimeSpan(0, 0, 0);
                    string totalF = "";
                    TimeSpan totalDurF = new TimeSpan(0, 0, 0);
                    string totalG = "";
                    TimeSpan totalDurG = new TimeSpan(0, 0, 0);
                    string totalH = "";
                    TimeSpan totalDurH = new TimeSpan(0, 0, 0);
                    string totalI = "";
                    TimeSpan totalDurI = new TimeSpan(0, 0, 0);
                    string totalJ = "";
                    TimeSpan totalDurJ = new TimeSpan(0, 0, 0);
                    
                    
                    for (int i = 2; i <= lastrow; i++)
                    {
                        // SOP entry-SOP exit - downtime calculation
                        // Column 'E'

                        if (workSheet.Cells[i, "C"].Text.ToString().Contains("SUPERVISOR MODE ENTRY") && firstOccur =="")
                        {
                            firstOccur = workSheet.Cells[i, "D"].Text.ToString();                            
                        }
                        else if (workSheet.Cells[i, "C"].Text.ToString().Contains("SUPERVISOR MODE EXIT") || workSheet.Cells[i,"C"].Text.ToString().Contains("POWER OFF"))
                        {
                            secondOccur = workSheet.Cells[i, "D"].Text.ToString();
                            if (firstOccur != null && firstOccur != "" && secondOccur != null && secondOccur != "")
                            {
                                int[] getTimeStart = StringToIntArray(firstOccur);
                                int[] getTimeEnd = StringToIntArray(secondOccur);

                                if (getTimeStart[0] >= 12 && getTimeEnd[0] < 12)// eg, 13:09 - 5:05
                                {
                                    getTimeEnd[0] += 24; // more than 1 day
                                }
                                if ((getTimeStart[0] >= 12 && getTimeEnd[0] >= 12) && (getTimeEnd[0] <= getTimeStart[0])) // e.g day1 18:00 - day 2 16:00 
                                {
                                    // temporary ignore if more than 24 hours system down, eg. day 1 14:00 - day 2 18:00                                   
                                    getTimeEnd[0] += 24;
                                }       

                                TimeSpan dtStart = new TimeSpan(getTimeStart[0], getTimeStart[1], getTimeStart[2]);
                                TimeSpan dtEnd = new TimeSpan(getTimeEnd[0], getTimeEnd[1], getTimeEnd[2]);

                                TimeSpan difTime = dtEnd - dtStart;
                                string duration = string.Format("{0:d} hrs {1:d} mins {2:d} secs", difTime.Hours, difTime.Minutes, difTime.Seconds);
                               
                                workSheet.Cells[i, "E"] = duration.Replace("-", "");
                                firstOccur = "";
                                secondOccur = "";                              
                                
                            }   
                        }                    
                   } // end of SOP - downtime for-loop                    
                

                        for (int i = 2; i <= lastrow; i++)
                        {
                            // Cash Jammed(W3/W7/W9) - downtime calculation
                            // Column 'F'
                            if (workSheet.Cells[i, "C"].Text.ToString().Contains("W7-CASH JAMMED") && firstOccur == "")
                            {
                                if (workSheet.Cells[i, "D"].Text.ToString() == "")
                                {
                                    firstOccur = workSheet.Cells[i + 1, "D"].Text.ToString();
                                }
                            }

                            if (workSheet.Cells[i, "C"].Text.ToString().Contains("W9-UNABLE TO ENCASH") && firstOccur == "")
                            {
                                if (workSheet.Cells[i, "D"].Text.ToString() == "")
                                {
                                    firstOccur = workSheet.Cells[i + 1, "D"].Text.ToString();
                                }

                            }

                            if (workSheet.Cells[i, "C"].Text.ToString().Contains(""))

                                if (workSheet.Cells[i, "C"].Text.ToString().Contains("W3-CASH UNIT FAILED") && workSheet.Cells[i - 1, "C"].Text.ToString() != ("W3-CASH UNIT FAILED") && workSheet.Cells[i + 1, "C"].Text.ToString().Contains("W3-CASH UNIT FAILED") && firstOccur == "")
                                {
                                    if (workSheet.Cells[i, "D"].Text.ToString() == "")
                                    {
                                        firstOccur = workSheet.Cells[i - 1, "D"].Text.ToString();
                                    }
                                }                         

                                // line 1 is  end of day, line 2 is not SOP entry

                              else if (workSheet.Cells[i, "C"].Text.ToString().Contains("END OF DAY") && !workSheet.Cells[i + 1, "C"].Text.ToString().Contains("SUPERVISOR MODE ENTRY"))
                                {
                                    secondOccur = workSheet.Cells[i, "D"].Text.ToString();  
                                  
                                    if (firstOccur != null && firstOccur != "" && secondOccur != null && secondOccur != "")
                                   {
                                    int[] getTimeStart = StringToIntArray(firstOccur);
                                    int[] getTimeEnd = StringToIntArray(secondOccur);

                                    if (getTimeStart[0] >= 12 && getTimeEnd[0] < 12)// eg, 13:09 - 5:05
                                    {
                                        getTimeEnd[0] += 24; // more than 1 day
                                    }
                                    if ((getTimeStart[0] >= 12 && getTimeEnd[0] >= 12) && (getTimeEnd[0] <= getTimeStart[0])) // e.g day1 18:00 - day 2 16:00 
                                    {
                                        // temporary ignore if more than 24 hours system down, eg. day 1 14:00 - day 2 18:00                                   
                                        getTimeEnd[0] += 24;
                                    }

                                    TimeSpan dtStart = new TimeSpan(getTimeStart[0], getTimeStart[1], getTimeStart[2]);
                                    TimeSpan dtEnd = new TimeSpan(getTimeEnd[0], getTimeEnd[1], getTimeEnd[2]);

                                    TimeSpan difTime = dtEnd - dtStart;
                                    string duration = string.Format("{0:d} hrs {1:d} mins {2:d} secs", difTime.Hours, difTime.Minutes, difTime.Seconds);

                                    workSheet.Cells[i, "F"] = duration.Replace("-", "");
                                    firstOccur = "";
                                    secondOccur = "";
                                  }

                            }
                            // if line is  "END OF DAY" and next line is SOP-Entry, TIME = EOD - FirstOccur
                            else if (workSheet.Cells[i, "C"].Text.ToString().Contains("END OF DAY") && workSheet.Cells[i + 1, "C"].Text.ToString().Contains("SUPERVISOR MODE ENTRY"))
                            {
                                secondOccur = workSheet.Cells[i, "D"].Text.ToString();

                                // calculate firstOccur up to EOD, then EOD-> firstOccur , SOP -> secondOccur
                                if (firstOccur != null && firstOccur != "" && secondOccur != null && secondOccur != "")
                                {
                                    int[] getTimeStart = StringToIntArray(firstOccur);
                                    int[] getTimeEnd = StringToIntArray(secondOccur);

                                    if (getTimeStart[0] >= 12 && getTimeEnd[0] < 12)// eg, 13:09 - 5:05
                                    {
                                        getTimeEnd[0] += 24; // more than 1 day
                                    }
                                    if ((getTimeStart[0] >= 12 && getTimeEnd[0] >= 12) && (getTimeEnd[0] <= getTimeStart[0])) // e.g day1 18:00 - day 2 16:00 
                                    {
                                        // temporary ignore if more than 24 hours system down, eg. day 1 14:00 - day 2 18:00                                   
                                        getTimeEnd[0] += 24;
                                    }

                                    TimeSpan dtStart = new TimeSpan(getTimeStart[0], getTimeStart[1], getTimeStart[2]);
                                    TimeSpan dtEnd = new TimeSpan(getTimeEnd[0], getTimeEnd[1], getTimeEnd[2]);

                                    TimeSpan difTime = dtEnd - dtStart;
                                    string duration = string.Format("{0:d} hrs {1:d} mins {2:d} secs", difTime.Hours, difTime.Minutes, difTime.Seconds);

                                    workSheet.Cells[i, "F"] = duration.Replace("-", "");

                                    // Calculate (SOP entry - End Of Day)
                                    firstOccur = workSheet.Cells[i, "D"].Text.ToString(); //  EOD  time
                                    secondOccur = workSheet.Cells[i + 1, "D"].Text.ToString();  // SOP-ENTRY time

                                    int[] getTimeEOD = StringToIntArray(firstOccur);
                                    int[] getTimeSOP = StringToIntArray(secondOccur);

                                    if (getTimeEOD[0] >= 12 && getTimeSOP[0] < 12)// eg, 13:09 - 5:05
                                    {
                                        getTimeSOP[0] += 24; // more than 1 day
                                    }
                                    if ((getTimeEOD[0] >= 12 && getTimeSOP[0] >= 12) && (getTimeSOP[0] <= getTimeEOD[0])) // e.g day1 18:00 - day 2 16:00 
                                    {
                                        // temporary ignore if more than 24 hours system down, eg. day 1 14:00 - day 2 18:00                                   
                                        getTimeSOP[0] += 24;
                                    }

                                    TimeSpan dtEODStart = new TimeSpan(getTimeEOD[0], getTimeEOD[1], getTimeEOD[2]);
                                    TimeSpan dtSOPEnd = new TimeSpan(getTimeSOP[0], getTimeSOP[1], getTimeSOP[2]);

                                    TimeSpan dt = dtSOPEnd - dtEODStart;
                                    string durationdt = string.Format("{0:d} hrs {1:d} mins {2:d} secs", dt.Hours, dt.Minutes, dt.Seconds);

                                    workSheet.Cells[i + 1, "F"] = durationdt.Replace("-", "");
                                    firstOccur = "";
                                    secondOccur = "";
                                }
                            }

                                // if line is  "END OF DAY" , line +1 = W3  , and line +2 != SOP entry
                                else if (workSheet.Cells[i,"C"].Text.ToString().Contains("W3-CASH UNIT FAILED") && workSheet.Cells[i-1, "C"].Text.ToString().Contains("END OF DAY") && !workSheet.Cells[i + 1, "C"].Text.ToString().Contains("SUPERVISOR MODE ENTRY"))
                                {
                                    if( workSheet.Cells[i, "D"].Text.ToString() =="")
                                    {
                                      secondOccur = workSheet.Cells[i+1, "D"].Text.ToString();
                                    }
                                      int[] getTimeothers = StringToIntArray(secondOccur);
                                      TimeSpan dtOthers = new TimeSpan(getTimeothers[0], getTimeothers[1], getTimeothers[2]);
                                       string durationdOthers = string.Format("{0:d} hrs {1:d} mins {2:d} secs", dtOthers.Hours, dtOthers.Minutes, dtOthers.Seconds);
                                       workSheet.Cells[i + 1, "F"] = durationdOthers.Replace("-", "");                                      
                                        secondOccur = "";                                  
                                }
                               


                            // current line is SOP-entry, previous is "W3-CASH UNIT FAILED"
                            else if (workSheet.Cells[i, "C"].Text.ToString().Contains("SUPERVISOR MODE ENTRY") && workSheet.Cells[i - 1, "C"].Text.ToString().Contains("W3-CASH UNIT FAILED"))
                            {
                                secondOccur = workSheet.Cells[i, "D"].Text.ToString();

                                if (firstOccur != null && firstOccur != "" && secondOccur != null && secondOccur != "")
                                {
                                    int[] getTimeStart = StringToIntArray(firstOccur);
                                    int[] getTimeEnd = StringToIntArray(secondOccur);

                                    if (getTimeStart[0] >= 12 && getTimeEnd[0] < 12)// eg, 13:09 - 5:05
                                    {
                                        getTimeEnd[0] += 24; // more than 1 day
                                    }
                                    if ((getTimeStart[0] >= 12 && getTimeEnd[0] >= 12) && (getTimeEnd[0] <= getTimeStart[0])) // e.g day1 18:00 - day 2 16:00 
                                    {
                                        // temporary ignore if more than 24 hours system down, eg. day 1 14:00 - day 2 18:00                                   
                                        getTimeEnd[0] += 24;
                                    }

                                    TimeSpan dtStart = new TimeSpan(getTimeStart[0], getTimeStart[1], getTimeStart[2]);
                                    TimeSpan dtEnd = new TimeSpan(getTimeEnd[0], getTimeEnd[1], getTimeEnd[2]);

                                    TimeSpan difTime = dtEnd - dtStart;
                                    string duration = string.Format("{0:d} hrs {1:d} mins {2:d} secs", difTime.Hours, difTime.Minutes, difTime.Seconds);

                                    workSheet.Cells[i, "F"] = duration.Replace("-", "");
                                    firstOccur = "";
                                    secondOccur = "";
                                }

                            }                                       

                        }// end of w3 for loop                       

                        for (int i = 2; i <= lastrow; i++)
                        {
                            // E4 -Dispenser Faulty
                            // Column 'G'

                            if (workSheet.Cells[i, "C"].Text.ToString().Contains("E4-DISPENSER FAULTY") && !workSheet.Cells[i - 1, "C"].Text.ToString().Contains("SUPERVISOR MODE ENTRY") && firstOccur =="")
                            {
                                firstOccur = workSheet.Cells[i, "D"].Text.ToString();
                            }

                            // if Inside the E4 , has another downtime calculation  --------------------------05/04/2017
                            if (firstOccur != "" && workSheet.Cells[i, "C"].Text.ToString().Contains("SUPERVISOR MODE ENTRY") && workSheet.Cells[i, "F"].Text.ToString() != "") 
                            {
                                workSheet.Cells[i, "F"] = "";   
                            }

                            else if (workSheet.Cells[i, "C"].Text.ToString().Contains("IN-SERVICE OK") && !workSheet.Cells[i - 1, "C"].Text.ToString().Contains("SUPERVISOR MODE EXIT"))
                            {

                                secondOccur = workSheet.Cells[i, "D"].Text.ToString();

                                if (firstOccur != null && firstOccur != "" && secondOccur != null && secondOccur != "")
                                {
                                    int[] getTimeStart = StringToIntArray(firstOccur);
                                    int[] getTimeEnd = StringToIntArray(secondOccur);

                                    if (getTimeStart[0] >= 12 && getTimeEnd[0] < 12)// eg, 13:09 - 5:05
                                    {
                                        getTimeEnd[0] += 24; // more than 1 day
                                    }
                                    if ((getTimeStart[0] >= 12 && getTimeEnd[0] >= 12) && (getTimeEnd[0] <= getTimeStart[0])) // e.g day1 18:00 - day 2 16:00 
                                    {
                                        // temporary ignore if more than 24 hours system down, eg. day 1 14:00 - day 2 18:00                                   
                                        getTimeEnd[0] += 24;
                                    }

                                    TimeSpan dtStart = new TimeSpan(getTimeStart[0], getTimeStart[1], getTimeStart[2]);
                                    TimeSpan dtEnd = new TimeSpan(getTimeEnd[0], getTimeEnd[1], getTimeEnd[2]);

                                    TimeSpan difTime = dtEnd - dtStart;
                                    string duration = string.Format("{0:d} hrs {1:d} mins {2:d} secs", difTime.Hours, difTime.Minutes, difTime.Seconds);

                                    workSheet.Cells[i, "G"] = duration.Replace("-", "");
                                    firstOccur = "";
                                    secondOccur = "";
                                }
                            }                                                
                        
                        } // end of E4- downtime calculation for-loop                       

                        for (int i = 2; i <= lastrow; i++)
                        {
                            // receipt error - IN-SERVICE-OK
                            // Column 'H'

                            if (workSheet.Cells[i, "C"].Text.ToString().Contains("RECEIPT PRINTER") && !workSheet.Cells[i - 1, "C"].Text.ToString().Contains("SUPERVISOR MODE ENTRY") && firstOccur == "")
                            {
                                firstOccur = workSheet.Cells[i, "D"].Text.ToString();

                            }
                            else if (workSheet.Cells[i, "C"].Text.ToString().Contains("IN-SERVICE OK") && !workSheet.Cells[i - 1, "C"].Text.ToString().Contains("SUPERVISOR MODE EXIT"))
                            {
                                secondOccur = workSheet.Cells[i, "D"].Text.ToString();

                                if (firstOccur != null && firstOccur != "" && secondOccur != null && secondOccur != "")
                                {
                                    int[] getTimeStart = StringToIntArray(firstOccur);
                                    int[] getTimeEnd = StringToIntArray(secondOccur);

                                    if (getTimeStart[0] >= 12 && getTimeEnd[0] < 12)// eg, 13:09 - 5:05
                                    {
                                        getTimeEnd[0] += 24; // more than 1 day
                                    }
                                    if ((getTimeStart[0] >= 12 && getTimeEnd[0] >= 12) && (getTimeEnd[0] <= getTimeStart[0])) // e.g day1 18:00 - day 2 16:00 
                                    {
                                        // temporary ignore if more than 24 hours system down, eg. day 1 14:00 - day 2 18:00                                   
                                        getTimeEnd[0] += 24;
                                    }

                                    TimeSpan dtStart = new TimeSpan(getTimeStart[0], getTimeStart[1], getTimeStart[2]);
                                    TimeSpan dtEnd = new TimeSpan(getTimeEnd[0], getTimeEnd[1], getTimeEnd[2]);

                                    TimeSpan difTime = dtEnd - dtStart;
                                    string duration = string.Format("{0:d} hrs {1:d} mins {2:d} secs", difTime.Hours, difTime.Minutes, difTime.Seconds);
                                    
          
                                    workSheet.Cells[i, "H"] = duration.Replace("-", "");
                                    firstOccur = "";
                                    secondOccur = "";
                                }
                            }
                        } // end of receipt error calculate downtime for-loop                      

                        for (int i = 2; i <= lastrow; i++)
                        {
                            // Card reader error - IN-SERVICE-OK
                            // Column 'I'

                            if (workSheet.Cells[i, "C"].Text.ToString().Contains("CARD READER") && !workSheet.Cells[i - 1, "C"].Text.ToString().Contains("SUPERVISOR MODE ENTRY") && firstOccur == "")
                            {
                                firstOccur = workSheet.Cells[i, "D"].Text.ToString();
                            }
                            else if (workSheet.Cells[i, "C"].Text.ToString().Contains("IN-SERVICE OK") && !workSheet.Cells[i - 1, "C"].Text.ToString().Contains("SUPERVISOR MODE EXIT"))
                            {
                                secondOccur = workSheet.Cells[i, "D"].Text.ToString();

                                if (firstOccur != null && firstOccur != "" && secondOccur != null && secondOccur != "")
                                {

                                    int[] getTimeStart = StringToIntArray(firstOccur);
                                    int[] getTimeEnd = StringToIntArray(secondOccur);

                                    if (getTimeStart[0] >= 12 && getTimeEnd[0] < 12)// eg, 13:09 - 5:05
                                    {
                                        getTimeEnd[0] += 24; // more than 1 day
                                    }
                                    if ((getTimeStart[0] >= 12 && getTimeEnd[0] >= 12) && (getTimeEnd[0] <= getTimeStart[0])) // e.g day1 18:00 - day 2 16:00 
                                    {
                                        // temporary ignore if more than 24 hours system down, eg. day 1 14:00 - day 2 18:00                                   
                                        getTimeEnd[0] += 24;
                                    }

                                    TimeSpan dtStart = new TimeSpan(getTimeStart[0], getTimeStart[1], getTimeStart[2]);
                                    TimeSpan dtEnd = new TimeSpan(getTimeEnd[0], getTimeEnd[1], getTimeEnd[2]);

                                    TimeSpan difTime = dtEnd - dtStart;
                                    string duration = string.Format("{0:d} hrs {1:d} mins {2:d} secs", difTime.Hours, difTime.Minutes, difTime.Seconds);
                                  
                                    workSheet.Cells[i, "I"] = duration.Replace("-", "");
                                    firstOccur = "";
                                    secondOccur = "";
                                }
                            }

                        } // end of CARD READER error calculate downtime for-loop
                        for (int i = 2; i <= lastrow; i++)
                        {
                            // SAM module Error - IN-SERVICE-OK
                            // Column 'J'

                            if (workSheet.Cells[i, "C"].Text.ToString().Contains("SAM module") && !workSheet.Cells[i - 1, "C"].Text.ToString().Contains("SUPERVISOR MODE ENTRY") && firstOccur == "")
                            {
                                firstOccur = workSheet.Cells[i, "D"].Text.ToString();
                            }
                            else if (workSheet.Cells[i, "C"].Text.ToString().Contains("IN-SERVICE OK") && !workSheet.Cells[i - 1, "C"].Text.ToString().Contains("SUPERVISOR MODE EXIT"))
                            {
                                secondOccur = workSheet.Cells[i, "D"].Text.ToString();

                                if (firstOccur != null && firstOccur != "" && secondOccur != null && secondOccur != "")
                                {
                                    int[] getTimeStart = StringToIntArray(firstOccur);
                                    int[] getTimeEnd = StringToIntArray(secondOccur);

                                    if (getTimeStart[0] >= 12 && getTimeEnd[0] < 12)// eg, 13:09 - 5:05
                                    {
                                        getTimeEnd[0] += 24; // more than 1 day
                                    }
                                    if ((getTimeStart[0] >= 12 && getTimeEnd[0] >= 12) && (getTimeEnd[0] <= getTimeStart[0])) // e.g day1 18:00 - day 2 16:00 
                                    {
                                        // temporary ignore if more than 24 hours system down, eg. day 1 14:00 - day 2 18:00                                   
                                        getTimeEnd[0] += 24;
                                    }

                                    TimeSpan dtStart = new TimeSpan(getTimeStart[0], getTimeStart[1], getTimeStart[2]);
                                    TimeSpan dtEnd = new TimeSpan(getTimeEnd[0], getTimeEnd[1], getTimeEnd[2]);

                                    TimeSpan difTime = dtEnd - dtStart;
                                    string duration = string.Format("{0:d} hrs {1:d} mins {2:d} secs", difTime.Hours, difTime.Minutes, difTime.Seconds);

                                    workSheet.Cells[i, "J"] = duration.Replace("-", "");
                                    firstOccur = "";
                                    secondOccur = "";
                                }
                            }

                        } // end of SAM module error calculate downtime for-loop     
/*
                        // check if there machine Down for more than 1 day..
                        // replace the time at column "F"
                        string RemoveTime = "";
                        TimeSpan OverlappedTime = new TimeSpan(0, 0, 0);
                        string overlappedTimeString = "";
                        for (int i = 2; i <= lastrow; i++)
                        {
                            
                            if (workSheet.Cells[i, "C"].Text.ToString().Contains("END OF DAY") && EOD == "")
                            {
                                EOD = "first EOD";
                            }

                            else if (workSheet.Cells[i, "C"].Text.ToString().Contains("IN-SERVICE OK") && EOD == "first EOD")
                            {
                                EOD = "";
                            }

                            else if (workSheet.Cells[i, "C"].Text.ToString().Contains("SUPERVISOR MODE ENTRY") && workSheet.Cells[i,"F"].Text.ToString() !="" && EOD == "first EOD")
                            {
                                RemoveTime = workSheet.Cells[i, "F"].Text.ToString();
                                RemoveTime = RemoveTime.Replace("hrs", ":").Replace("mins", ":").Replace("secs", "");
                                int[] getTime = StringToIntArray(RemoveTime);
                                TimeSpan convrtTime = new TimeSpan(getTime[0],getTime[1],getTime[2]);
                                OverlappedTime += convrtTime;
                                overlappedTimeString = string.Format("{0:d} hrs {1:d} mins {2:d} secs", OverlappedTime.Hours, OverlappedTime.Minutes, OverlappedTime.Seconds);
                                RemoveTime = "";

                            }
                            else if (workSheet.Cells[i, "C"].Text.ToString().Contains("END OF DAY") && EOD == "first EOD" && overlappedTimeString != null)
                            {
                                // continue here...... minus the extra time
                                // EOD - all SOP entry time 
                                string EODtime = "23:59:59";
                                overlappedTimeString = overlappedTimeString.Replace("hrs", ":").Replace("mins", ":").Replace("secs", "");
                                int[] getEODTime = StringToIntArray(EODtime);
                                int[] getOverTime = StringToIntArray(overlappedTimeString);
                                TimeSpan convrtEODTime = new TimeSpan(getEODTime[0], getEODTime[1], getEODTime[2]);
                                TimeSpan convrtOverTime = new TimeSpan(getOverTime[0], getOverTime[1], getOverTime[2]);
                                TimeSpan difEOD = convrtEODTime - convrtOverTime;
                                string durationEOD = string.Format("{0:d} hrs {1:d} mins {2:d} secs", difEOD.Hours, difEOD.Minutes, difEOD.Seconds);
                                workSheet.Cells[i, "F"] = durationEOD.Replace("-", "");
                                EOD = "";
                                overlappedTimeString = "";
                            }                        
                            else
                            {

                            }

                        }
                   
 */  

  // TOTAL DOWNTIME

                     // SOP Mode ,E
                       Excel.Range xlRngE = workSheet.Range["E:E"];
                        TimeSpan totalTimeE = new TimeSpan(0, 0, 0);
                        foreach (string value in xlRngE.Value)
                        {
                            if (value != "" && value != null && ! value.Contains("SOP Mode"))
                            {
                                string strResultE = value.Replace("hrs", ":").Replace("mins", ":").Replace("secs", ":");
                                int[] getStrtoArray = StringToIntArray(strResultE);
                                TimeSpan timeE = new TimeSpan(getStrtoArray[0], getStrtoArray[1], getStrtoArray[2]);
                                totalTimeE += timeE;
                            }
                        }                      
                        totalE = string.Format("{0} hrs {1} mins {2} secs", ((int)totalTimeE.TotalHours), totalTimeE.Minutes, totalTimeE.Seconds);                        
                        workSheet.Cells[5, "K"] = "Total SOP downtime:";
                        workSheet.Cells[6, "K"] = totalE;

                        //W3/W7/W9-Error , F
                       Excel.Range xlRngF = workSheet.Range["F:F"];
                        TimeSpan totalTimeF = new TimeSpan(0, 0, 0);
                        foreach (string value in xlRngF.Value)
                        {
                            if (value != "" && value != null && !value.Contains("W3/W7/W9-Error"))
                            {
                                string strResultF = value.Replace("hrs", ":").Replace("mins", ":").Replace("secs", ":");
                                int[] getStrtoArray = StringToIntArray(strResultF);
                                TimeSpan timeF = new TimeSpan(getStrtoArray[0], getStrtoArray[1], getStrtoArray[2]);
                                totalTimeF += timeF;
                            }
                        }
                        totalF = string.Format("{0} hrs {1} mins {2} secs", ((int)totalTimeF.TotalHours), totalTimeF.Minutes, totalTimeF.Seconds);
                        workSheet.Cells[5, "L"] = "Total CashJam downtime:";
                        workSheet.Cells[6, "L"] = totalF;

                        // Dispenser Faulty, G
                        Excel.Range xlRngG = workSheet.Range["G:G"];
                        TimeSpan totalTimeG = new TimeSpan(0, 0, 0);
                        foreach (string value in xlRngG.Value)
                        {
                            if (value != "" && value != null && !value.Contains("Dispenser Faulty"))
                            {
                                string strResultG = value.Replace("hrs", ":").Replace("mins", ":").Replace("secs", ":");
                                int[] getStrtoArray = StringToIntArray(strResultG);
                                TimeSpan timeG = new TimeSpan(getStrtoArray[0], getStrtoArray[1], getStrtoArray[2]);
                                totalTimeG += timeG;
                            }
                        }
                        totalG = string.Format("{0} hrs {1} mins {2} secs", ((int)totalTimeG.TotalHours), totalTimeG.Minutes, totalTimeG.Seconds);
                        workSheet.Cells[5, "M"] = "Total E4-Dispenser Faulty downtime:";
                        workSheet.Cells[6, "M"] = totalG;

                        //Receipt Error , H
                        Excel.Range xlRngH = workSheet.Range["H:H"];
                        TimeSpan totalTimeH = new TimeSpan(0, 0, 0);
                        foreach (string value in xlRngH.Value)
                        {
                            if (value != "" && value != null && !value.Contains("Receipt Error"))
                            {
                                string strResultH = value.Replace("hrs", ":").Replace("mins", ":").Replace("secs", ":");
                                int[] getStrtoArray = StringToIntArray(strResultH);
                                TimeSpan timeH = new TimeSpan(getStrtoArray[0], getStrtoArray[1], getStrtoArray[2]);
                                totalTimeH += timeH;
                            }
                        }
                        totalH = string.Format("{0} hrs {1} mins {2} secs", ((int)totalTimeH.TotalHours), totalTimeH.Minutes, totalTimeH.Seconds);
                        workSheet.Cells[5, "N"] = "Total Receipt Error downtime:";
                        workSheet.Cells[6, "N"] = totalH;

                        //Card Reader Error , I
                        Excel.Range xlRngI = workSheet.Range["I:I"];
                        TimeSpan totalTimeI = new TimeSpan(0, 0, 0);
                        foreach (string value in xlRngI.Value)
                        {
                            if (value != "" && value != null && !value.Contains("Card Reader Error"))
                            {
                                string strResultI = value.Replace("hrs", ":").Replace("mins", ":").Replace("secs", ":");
                                int[] getStrtoArray = StringToIntArray(strResultI);
                                TimeSpan timeI = new TimeSpan(getStrtoArray[0], getStrtoArray[1], getStrtoArray[2]);
                                totalTimeI += timeI;
                            }
                        }
                        totalI = string.Format("{0} hrs {1} mins {2} secs", ((int)totalTimeI.TotalHours), totalTimeI.Minutes, totalTimeI.Seconds);
                        workSheet.Cells[5, "O"] = "Total Card Reader Error downtime:";
                        workSheet.Cells[6, "O"] = totalI;

                        //SAM module/chip error, J
                        Excel.Range xlRngJ = workSheet.Range["J:J"];
                        TimeSpan totalTimeJ = new TimeSpan(0, 0, 0);
                        foreach (string value in xlRngJ.Value)
                        {
                            if (value != "" && value != null && !value.Contains("SAM module/chip error"))
                            {
                                string strResultJ = value.Replace("hrs", ":").Replace("mins", ":").Replace("secs", ":");
                                int[] getStrtoArray = StringToIntArray(strResultJ);
                                TimeSpan timeJ = new TimeSpan(getStrtoArray[0], getStrtoArray[1], getStrtoArray[2]);
                                totalTimeJ += timeJ;
                            }
                        }
                        totalJ = string.Format("{0} hrs {1} mins {2} secs", ((int)totalTimeJ.TotalHours), totalTimeJ.Minutes, totalTimeJ.Seconds);
                        workSheet.Cells[5, "P"] = "Total SAM module/chip error downtime:";
                        workSheet.Cells[6, "P"] = totalJ;


                        /*   TEMPORARY REMOVE DOWNTIME
                        // calculate downtime
                        int lastrow = workSheet.UsedRange.Row + workSheet.UsedRange.Rows.Count - 1;

                        string firstnum = "";
                        string secondnum = "";
                        string thirdnum = "";
                        string total = "";
                        TimeSpan totalDur = new TimeSpan(0,0,0);

                        for (int i = 2; i <= lastrow; i++)
                        {
                            if (workSheet.Cells[i, "D"].Value != null)
                            {
                                if (workSheet.Cells[i, "C"].Text.ToString().Contains("SUPERVISOR MODE ENTRY"))
                                {
                                    // condition: if Entry -> secondnum
                                    if (firstnum != "")
                                    {
                                        secondnum = workSheet.Cells[i, "D"].Text.ToString();
                                        thirdnum = secondnum;
                                    }
                                    else
                                    {
                                        //condition: if Entry -> firstnum
                                        firstnum = workSheet.Cells[i, "D"].Text.ToString();
                                    }

                                }
                                else if (workSheet.Cells[i, "C"].Text.ToString().Contains("SUPERVISOR MODE EXIT"))
                                {
                                    secondnum = workSheet.Cells[i, "D"].Text.ToString();
                                    //condition: if Entry from thirdnum ->firstnum
                                    if (thirdnum != "")
                                    {
                                        firstnum = thirdnum;
                                        thirdnum = "";
                                    }
                                }
                                else if (workSheet.Cells[i, "C"].Text.ToString().Contains("POWER OFF"))
                                {
                                    secondnum = workSheet.Cells[i, "D"].Text.ToString();
                                }
                                else if (workSheet.Cells[i, "C"].Text.ToString().Contains("E4-"))
                                {
                                    firstnum = workSheet.Cells[i, "D"].Text.ToString();
                                }
                                else
                                {
                                    firstnum = workSheet.Cells[i, "D"].Text.ToString();
                                }
                            }
                            //
                            if (firstnum != null && firstnum != "" && secondnum != null && secondnum != "")
                            {
                                int[] GetTimeStart;
                                int[] GetTimeEnd;
                                GetTimeStart = StringToIntArray(firstnum);

                                //TempTimeEnd = TroubleTimeEnd.Replace(":", " ");
                                GetTimeEnd = StringToIntArray(secondnum);

                                if (GetTimeStart[0] > 12 && GetTimeEnd[0] < 12)
                                {
                                    GetTimeEnd[0] += 24;
                                }
                                TimeSpan DownTimeStart = new TimeSpan(GetTimeStart[0], GetTimeStart[1], GetTimeStart[2]);

                                TimeSpan DownTimeEnd = new TimeSpan(GetTimeEnd[0], GetTimeEnd[1], GetTimeEnd[2]);

                                TimeSpan DifferenceTime = DownTimeEnd - DownTimeStart;
                                string duration = string.Format("{0:d} hrs {1:d} mins {2:d} secs", DifferenceTime.Hours, DifferenceTime.Minutes, DifferenceTime.Seconds);

                                totalDur += DifferenceTime;

                                // MessageBox.Show(duration);
                                workSheet.Cells[i, "I"] = duration.Replace("-", "");
                                firstnum = "";
                                secondnum = "";
                            }
                        }

                        // Sum the downtime and display it
                        total = string.Format("{0} hrs {1} mins {2} secs",((int)totalDur.TotalHours) , totalDur.Minutes, totalDur.Seconds);
                        workSheet.Cells[5, "K"] = "Total Duration:" + total.Replace("-", "");
    */           
                        // autofit columns
                        for (int i = 1; i <= 20; i++)
                        {
                            workSheet.Columns[i].AutoFit();
                        }
/*
                    for (int i = 1; i <= 14; i++)
                    {
                        workSheet2.Columns[i].AutoFit();
                    }
*/
                }// end of foreach loop

                // save this result          
                
              // excelApp.ActiveWorkbook.SaveCopyAs(@"C:\Users\sw_wllim\Desktop\EJTroubleReader\Result.xlsx");
               var desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
               var fullpath = Path.Combine(desktopFolder, "Result.xlsx");
               excelApp.ActiveWorkbook.SaveCopyAs(fullpath);
         
               excelApp.DisplayAlerts = false;
               excelApp.ActiveWorkbook.Close(0);
               excelApp.Quit();
               
                //check time taken for the program
                string end = DateTime.Now.ToString("h:mm:ss tt");
                MessageBox.Show("Completed!" +"\n" +"\nStart = " + start + "\n" + "End = " + end);

                // enable button back
                btnExcel.Enabled = true;
                btnEnter.Enabled = true;
                excel_dir.Enabled = true;
                path_dir.Enabled = true;

                
                // release excel from background process
                Marshal.ReleaseComObject(workSheet);
  //              Marshal.ReleaseComObject(workSheet2);
                Marshal.ReleaseComObject(excelApp);

            }// end of else        
    }          
        //
        // delete empty rows
        //
        private static void DeleteEmptyRowsCols(Excel.Worksheet workSheet)
        {
            Excel.Range targetCells = workSheet.UsedRange;
            object[,] allValues = (object[,])targetCells.Cells.Value;
            int totalRows = targetCells.Rows.Count;
            int totalCols = targetCells.Columns.Count;

            List<int> emptyRows = GetEmptyRows(allValues, totalRows, totalCols);
            DeleteRows(emptyRows, workSheet);        
        
        }        
        private static void DeleteRows(List<int> rowsToDelete, Excel.Worksheet worksheet)
        {
            // the rows are sorted high to low - so index's wont shift
            foreach (int rowIndex in rowsToDelete)
            {
                worksheet.Rows[rowIndex].Delete();
            }
        }
        private static List<int> GetEmptyRows(object[,] allValues, int totalRows, int totalCols)
        {
            List<int> emptyRows = new List<int>();

            for (int i = 1; i < totalRows; i++)
            {
                if (IsRowEmpty(allValues, i, totalCols))
                {
                    emptyRows.Add(i);
                }
            }
            // sort the list from high to low
            return emptyRows.OrderByDescending(x => x).ToList();
        }
        private static bool IsRowEmpty(object[,] allValues, int rowIndex, int totalCols)
        {
            for (int i = 1; i <= totalCols; i++)
            {
                if (allValues[rowIndex, i] != null)
                {
                    return false;
                }
            }
            return true;
        }

        private static int[] StringToIntArray(string myNumbers)
        {
            List<int> myIntegers = new List<int>();
            Array.ForEach(myNumbers.Split(":".ToCharArray()), s =>
            {
                int currentInt;
                if (Int32.TryParse(s, out currentInt))
                    myIntegers.Add(currentInt);
            });
            return myIntegers.ToArray();
        }

        private void path_dir_Click(object sender, EventArgs e) 
        {
            DialogResult result = this.folderBrowserDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                string path = this.folderBrowserDialog1.SelectedPath;
                path_dir.Text = path;
            }
        }


        private void excel_dir_Click(object sender, EventArgs e)
        {

            DialogResult result = this.folderBrowserDialog2.ShowDialog();
            if (result == DialogResult.OK)
            {
                string path = this.folderBrowserDialog2.SelectedPath;
                excel_dir.Text = path;

            }       
            
        }
    }
}
