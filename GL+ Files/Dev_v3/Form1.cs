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
using System.Threading;
using teemtalk;
using System.Globalization;
using System.Diagnostics;
using NLog;
using Excel = Microsoft.Office.Interop.Excel;


namespace Status_changer
{
    
    public partial class Form1 : Form
    {

        private static Logger logger = LogManager.GetCurrentClassLogger(); // Nlog

        public Form1()
        {
            InitializeComponent();
        }



        static teemtalk. Application teemApp;

        public string EventDepot { get; private set; }

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox_login.Text == "")
                {
                    MessageBox.Show("Вы не ввели логин", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (textBox_pw.Text == "")
                {
                    MessageBox.Show("Вы не ввели пароль", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }


                //поиск файла Excel
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Multiselect = false;
                ofd.DefaultExt = "*.xls;*.xlsx";
                ofd.Filter = "Microsoft Excel (*.xls*)|*.xls*";
                ofd.Title = "Выберите документ Excel";

                if (ofd.ShowDialog() != DialogResult.OK)
                {
                    MessageBox.Show("Вы не выбрали файл для открытия", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                                                      
                            
                //Login into Mainframe
                var login = textBox_login.Text;
                var password = textBox_pw.Text;

                //var login = Properties.Settings.Default.loginMF;
                //var password = Properties.Settings.Default.pwdMF;
                //var consData = DBContext.GetConsStatus();

                teemApp = new teemtalk.Application();
            
                teemApp.CurrentSession.Name = "Mainframe";

                teemApp.CurrentSession.Network.Protocol = ttNetworkProtocol.ProtocolWinsock;
                teemApp.CurrentSession.Network.Hostname = "mainframe.gb.tntpost.com";
                teemApp.CurrentSession.Network.Telnet.Port = 23;
                teemApp.CurrentSession.Network.Telnet.Name = "IBM-3278-2-E";
                teemApp.CurrentSession.Emulation = ttEmulations.IBM3270Emul;

                teemApp.CurrentSession.Network.Connect();

                teemApp.Visible = Properties.Settings.Default.isVisible;
            

                var host = teemApp.CurrentSession.Host;
                var disp = teemApp.CurrentSession.Display;

                ForAwait(35, 16, "INTERNATIONAL");

                host.Send("SM");
                host.Send("<ENTER>");

                ForAwait(13, 23, "USER ID");
                Thread.Sleep(2000);
                host.Send(login);
                host.Send("<TAB>");
                host.Send(password);
                host.Send("<ENTER>");

                Thread.Sleep(2000);
                if (teemApp.CurrentSession.Display.CursorCol == 40)
                {
                    TeemTalkClose();
                    MessageBox.Show("Вы ввели неверный логин или пароль. Введите правильные данные и нажмите кнопку START", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;

                }
                else if (teemApp.CurrentSession.Display.CursorCol == 35)
                {
                    TeemTalkClose();
                    MessageBox.Show("Ваш пароль устарел. Измените пароль в Mainframe, введите правильные данные и нажмите кнопку START", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;

                }                
                

                //if (!ForAwait(2, 2, "Command")) goto StartMaimframe;
                ForAwait(2, 2, "Command");
                host.Send("2");
                host.Send("<ENTER>");

                Thread.Sleep(2000);
                if (teemApp.CurrentSession.Display.CursorCol == 01)
                {
                    TeemTalkClose();
                    MessageBox.Show("Пользователь "+login+" уже авторизован в Terminal I. Выйдете из сессии Terminal I и нажмите кнопку START", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;

                }

                logger.Debug("User:"+login, this.Text); //LOG 

                ForAwait(20, 7, "Job Description");
                host.Send("<F12>");
                Thread.Sleep(500);
                if (disp.CursorRow != 2)


                host.Send("PK01");
                logger.Debug("PK01", this.Text); //LOG
                host.Send("<ENTER>");
                Thread.Sleep(2200);

                if (teemApp.CurrentSession.Display.CursorCol == 25)// Пункт 7.1 Окно, которое нужно закрыть
                {
                    host.Send("<F12>");
                }

                //Проверка на возможность доступа в PK01
                //var PK01 = "";
                //PK01= disp.ScreenData[1, 73, 4];

                if (disp.ScreenData[73, 1, 4] != "PK01")
                {
                    TeemTalkClose();
                    MessageBox.Show("Пользователь " + login + " не имеет доступа в Special Service Order Search Criteria", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                //Открываем excel файл
                string xlFileName = ofd.FileName; //имя нашего Excel файла
                Excel.Application ObjWorkExcel = new Excel.Application(); //создаём приложение Excel
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(xlFileName); //открываем наш файл          
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист               
                var last = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);//1 ячейку                          
                int lastUsedRow = last.Row;
                logger.Debug("lastUsedRow:"+ lastUsedRow, this.Text); //LOG



                //Копируем открываемый excel файл в папку User Logs
                string UserLogsPath = @"User Logs";
                if (!Directory.Exists(UserLogsPath)) //Если папки нет...
                    Directory.CreateDirectory(UserLogsPath); //...создадим ее
                string xlFileExt = xlFileName.Split(new[] { '.' }).Last();
                string xlFileNameWithExt = xlFileName.Split(new[] { '\\' }).Last();
                string xlLogFileName = "Log_" + DateTime.Now.ToString("HH.mm.ss_ddMMMyy", CultureInfo.GetCultureInfo("en-us")) + "." + xlFileExt;
                string destxlLogFileName = Path.Combine(UserLogsPath, xlLogFileName);
                File.Copy(xlFileName, destxlLogFileName);


                // Создаем пользовательский LOG файл
                string UserLogName = "Log_" + DateTime.Now.ToString("HH.mm.ss_ddMMMyy", CultureInfo.GetCultureInfo("en-us")) + ".txt";
                string destUserLog = Path.Combine(UserLogsPath, UserLogName);                
                StreamWriter UserLog = new StreamWriter(destUserLog, true);
                UserLog.WriteLine("#User: "+login);
                UserLog.Close();




                for (int i = 2; i <= lastUsedRow; i++) 
                {
                    // Colnum
                    string colnum = Convert.ToString(i);
                                      

                    logger.Debug(colnum, this.Text); //LOG


                    #region variables /////Объявление переменных//////////
                    ////////////////////////////////////////////////////////////////Объявление переменных/////////////////////////////////////////////////////////////////////////

                    //CName - Caller Name
                    //var CName = "0";
                    //TelNo - Telephone No
                    //var TelNo = "0";


                    //Select - Selection
                    //var Select = "41";


                    //Con - Connote из колонки CN Number
                    var excelcon = ObjWorkSheet.get_Range("B" + colnum, Type.Missing).Value2;
                    if (excelcon == null)
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine("#Row " + colnum + " - No CN Number");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }
                    string Con = excelcon.ToString();
                    if (Con.Length != 9)
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con +" - Bad CN Number Lenght");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }

                    //Acc - Account No из колонки acc (A)
                    var excelAcc = ObjWorkSheet.get_Range("A" + colnum, Type.Missing).Value2;
                    if (excelAcc == null)
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - No Acc data");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }
                    
                    string CheckAcc = excelAcc.ToString();// Добавление 0 к Acc No
                    string Acc = "";
                    if(CheckAcc.Length < 9)
                    {
                        if (CheckAcc.Length == 8) { Acc = "0" + CheckAcc; }
                        else if (CheckAcc.Length == 7) { Acc = "00" + CheckAcc; }
                        else if (CheckAcc.Length == 6) { Acc = "000" + CheckAcc; }
                        else if (CheckAcc.Length == 5) { Acc = "0000" + CheckAcc; }
                        else if (CheckAcc.Length == 4) { Acc = "00000" + CheckAcc; }
                        //short z = 1;
                        //do
                        //{
                        //    string zero = "0";
                        //    Acc = zero * z + CheckAcc;

                        //    z++

                        //} while (Acc.Length != 9);
                    }
                    else if (CheckAcc.Length == 9)
                    {
                        CheckAcc = Acc;
                    }
                    else
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - Acc has to much symbols");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }

                    //SSstat - SS status
                    //var SSstat = "BK";
                    var excelSSstat = ObjWorkSheet.get_Range("AB" + colnum, Type.Missing).Value2;
                    if (excelSSstat == null)
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - No Status data");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }
                    string SSstat = excelSSstat.ToString();

                    if (SSstat != "bk+ve"
                        && SSstat != "bk"
                        && SSstat == "BK"
                        && SSstat != "BK+VE"
                        && SSstat != "bk+VE"
                        && SSstat != "BK+ve")
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - Wrong Status");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }


                    ////RecName - Receaver Name
                    //var excelRecName = ObjWorkSheet.get_Range("C" + colnum, Type.Missing).Value2;
                    //if (excelRecName == null)
                    //{
                    //    UserLog = new StreamWriter(destUserLog, true);
                    //    UserLog.WriteLine(Con + " - No Receaver Name");
                    //    UserLog.Close();
                    //    continue; //переход к следующей итерации FOR
                    //}
                    //string RecName = excelRecName.ToString();


                    ////RecAddr - Receaver Address
                    //var excelRecAddr = ObjWorkSheet.get_Range("D" + colnum, Type.Missing).Value2;
                    //if (excelRecAddr == null)
                    //{
                    //    UserLog = new StreamWriter(destUserLog, true);
                    //    UserLog.WriteLine(Con + " - No Receaver Address");
                    //    UserLog.Close();
                    //    continue; //переход к следующей итерации FOR
                    //}
                    //string RecAddr = excelRecAddr.ToString();
                    //if (RecAddr.Length > 20)
                    //{
                    //    RecAddr = RecAddr.Substring(0, 30);// ограничивание накладной по количеству знаков 30
                    //}

                    ////RecTown - Receaver Town
                    //var excelRecTown = ObjWorkSheet.get_Range("E" + colnum, Type.Missing).Value2;
                    //if (excelRecTown == null)
                    //{
                    //    UserLog = new StreamWriter(destUserLog, true);
                    //    UserLog.WriteLine(Con + " - No Receaver Town");
                    //    UserLog.Close();
                    //    continue; //переход к следующей итерации FOR
                    //}
                    //string RecTown = excelRecTown.ToString();

                    ////RecPost - Receaver Postcode
                    //var excelRecPost = ObjWorkSheet.get_Range("F" + colnum, Type.Missing).Value2;
                    //if (excelRecPost == null)
                    //{
                    //    UserLog = new StreamWriter(destUserLog, true);
                    //    UserLog.WriteLine(Con + " - No Receaver Postcode");
                    //    UserLog.Close();
                    //    continue; //переход к следующей итерации FOR
                    //}
                    //string RecPost = excelRecPost.ToString();


                    ////GdsDesk - GDS Desc  
                    //var excelGdsDesk = ObjWorkSheet.get_Range("G" + colnum, Type.Missing).Value2;
                    //if (excelGdsDesk == null)
                    //{
                    //    UserLog = new StreamWriter(destUserLog, true);
                    //    UserLog.WriteLine(Con + " - No GDS Description");
                    //    UserLog.Close();
                    //    continue; //переход к следующей итерации FOR
                    //}
                    //string GdsDesk = excelGdsDesk.ToString();

                    ////WeightKG - Weight KG
                    //var excelWeightKG = ObjWorkSheet.get_Range("I" + colnum, Type.Missing).Value2;// Изменил колонку на I
                    //if (excelWeightKG == null)
                    //{
                    //    UserLog = new StreamWriter(destUserLog, true);
                    //    UserLog.WriteLine(Con + " - No Weight data");
                    //    UserLog.Close();
                    //    continue; //переход к следующей итерации FOR
                    //}
                    //string WeightKG = excelWeightKG.ToString();


                    //Weight - Weight 
                    var excelWeight = ObjWorkSheet.get_Range("H" + colnum, Type.Missing).Value2;
                    if (excelWeight == null)
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - No Weight data");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }
                    string Weight = excelWeight.ToString();

                    if (Weight.Contains(","))
                    {
                        Weight = Weight.Replace(",", ".");
                    }




                    //Items
                    var excelItems = ObjWorkSheet.get_Range("J" + colnum, Type.Missing).Value2;
                    if (excelItems == null)
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - No Items data");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }                    
                    string Items = excelItems.ToString();


                    ////Length - 10
                    //var Length = "10";
                    ////Widht - 10
                    //var Widht = "10";
                    ////Height - 10
                    //var Height = "10";


                    //Экспорт даты забора CollDate - Collection Date
                    Excel.Range excelCollDate = ObjWorkSheet.get_Range("K" + colnum);
                    object CollDate_v = excelCollDate.Value2;

                    if (CollDate_v == null)
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - No Collection Date");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }

                    DateTime dCt = DateTime.FromOADate((double)CollDate_v);
                    string CollDate = dCt.ToString("ddMMMyy", CultureInfo.GetCultureInfo("en-us"));

                    if (CollDate_v is double)
                    {
                        dCt = DateTime.FromOADate((double)CollDate_v);
                    }
                    else
                    {
                        DateTime.TryParse((string)CollDate_v, out dCt);
                    }




                    //CollTime - Collection Time = 0900
                    var CollTime = "0900";

                    //CollTimeTo - Collection Time To = 1800
                    var CollTimeTo = "1800";



                    //Экспорт даты доставки DelDate - Delivery Date From
                    Excel.Range excelDelDate = ObjWorkSheet.get_Range("L" + colnum);
                    object DelDate_v = excelDelDate.Value2;

                    if (DelDate_v == null)
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - No Delivery Date");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }

                    DateTime dDt = DateTime.FromOADate((double)DelDate_v);
                    string DelDate = dDt.ToString("ddMMMyy", CultureInfo.GetCultureInfo("en-us"));

                    if (DelDate_v is double)
                    {
                        dDt = DateTime.FromOADate((double)DelDate_v);
                    }
                    else
                    {
                        DateTime.TryParse((string)DelDate_v, out dDt);
                    }


                    //Триггер на даты забора и доставки
                    if (dCt.Date > dDt.Date)
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - The Collection date cannot be earlier than the Delivery Date");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }


                    //DelTime - Delivery Time = 0900
                    var Deltime = "0900";

                    //DelTimeTo - Delivery Time To = 2359
                    var DeltimeTo = "2359";


                    //Dev - Dev = s
                    var excelDiv = ObjWorkSheet.get_Range("M" + colnum, Type.Missing).Value2;                    
                    if (excelDiv == null)
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - No Div");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }
                    string Div = excelDiv.ToString();
                    if (Div != "s" && Div != "S")// Div всегда должен быть s, нужна ли проверка или можно дать значение s?
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - Bad Div");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }


                    //Prod - Prod
                    var excelProd = ObjWorkSheet.get_Range("N" + colnum, Type.Missing).Value2;
                    if (excelProd == null)
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - No Prod date");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }
                    string Prod = excelProd.ToString();
                    



                    //Payer - Payer = S || R
                    var excelPayer = ObjWorkSheet.get_Range("O" + colnum, Type.Missing).Value2;
                    if (excelPayer == null)
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - No Payer data");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }
                    string Payer = excelPayer.ToString();
                    if (Payer != "s" && Payer != "S"
                        && Payer != "r" && Payer != "R")// Payer всегда должен быть s или r
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - Bad Payer");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }

                    ///Stack - Stackable = y
                    //var Stack = "y";


                    //CMair
                    var CMair = "4";

                    //CMairVendor - Vendor
                    var excel_CMairVendor = ObjWorkSheet.get_Range("Q" + colnum, Type.Missing).Value2;
                    string CMairVendor;
                    if (excel_CMairVendor == null)
                    {
                        CMairVendor = null;
                    }
                    else
                    {
                        CMairVendor = excel_CMairVendor.ToString();
                    }


                    //CMairQt - Quote Amount
                    var excel_CMairQtDouble = ObjWorkSheet.get_Range("R" + colnum, Type.Missing).Value2;
                    string CMairQt;
                    if (excel_CMairQtDouble == null)
                    {
                        CMairQt = null;
                    }
                    else
                    {
                        float excel_CMairQt = Convert.ToSingle(excel_CMairQtDouble);// Fix для не дробных значений
                        CMairQt = excel_CMairQt.ToString();

                        if (CMairQt.Contains(","))
                        {
                            CMairQt = CMairQt.Replace(",", ".");
                        }
                    }

                    



                    //LCLPU
                    var LCLPU = "24";

                    //LCLPU_Vendor - Vendor
                    var excel_LCLPU_Vendor = ObjWorkSheet.get_Range("T" + colnum, Type.Missing).Value2;
                    string LCLPU_Vendor;
                    if (excel_LCLPU_Vendor == null)
                    {
                        LCLPU_Vendor = null;
                    }
                    else
                    {
                        LCLPU_Vendor = excel_LCLPU_Vendor.ToString();
                    }


                    //LCLPU_Qt - Quote Amount
                    var excel_LCLPU_QtDouble = ObjWorkSheet.get_Range("U" + colnum, Type.Missing).Value2;
                    string LCLPU_Qt;
                    if (excel_LCLPU_QtDouble == null)
                    {
                        LCLPU_Qt = null;
                    }
                    else
                    {
                        float excel_LCLPU_Qt = Convert.ToSingle(excel_LCLPU_QtDouble);// Fix для не дробных значений
                        LCLPU_Qt = excel_LCLPU_Qt.ToString();

                        if (LCLPU_Qt.Contains(","))
                        {
                            LCLPU_Qt = LCLPU_Qt.Replace(",", ".");
                        }
                    }

                    


                    //LCDl - Local Deilvery

                    var LCDL = "23";

                    //LCDL_Vendor - Vendor
                    var excel_LCDL_Vendor = ObjWorkSheet.get_Range("W" + colnum, Type.Missing).Value2;
                    string LCDL_Vendor;
                    if (excel_LCDL_Vendor == null)
                    {
                        LCDL_Vendor = null;
                    }
                    else
                    {
                        LCDL_Vendor = excel_LCDL_Vendor.ToString();
                    }


                    //LCDL_Qt - Quote Amount
                    var excel_LCDL_QtDouble = ObjWorkSheet.get_Range("X" + colnum, Type.Missing).Value2;
                    string LCDL_Qt;
                    if (excel_LCDL_QtDouble == null)
                    {
                        LCDL_Qt = null;
                    }
                    else
                    {
                        float excel_LCDL_Qt = Convert.ToSingle(excel_LCDL_QtDouble);// Fix для не дробных значений
                        LCDL_Qt = excel_LCDL_Qt.ToString();

                        if (LCDL_Qt.Contains(","))
                        {
                            LCDL_Qt = LCDL_Qt.Replace(",", ".");
                        }
                    }

                    


                    //Revao
                    var Revao_n = "43";

                    var excelRevaoDouble = ObjWorkSheet.get_Range("Y" + colnum, Type.Missing).Value2;
                    string Revao;
                    if (excelRevaoDouble == null)// Если ячейки пустые, то ничего не вводить                 
                    {
                        Revao = null;
                    }
                    else
                    {
                        float excelRevao = Convert.ToSingle(excelRevaoDouble);// Fix для не дробных значений
                        Revao = excelRevao.ToString();

                        if (Revao.Contains(","))
                        {
                            Revao = Revao.Replace(",", ".");
                        }
                    }

                    



                    //Disc
                    var Disc_n = "7";

                    var excelDiscDouble = ObjWorkSheet.get_Range("Z" + colnum, Type.Missing).Value2;
                    
                    string Disc;
                    if (excelDiscDouble == null)// Если ячейки пустые, то ничего не вводить                 
                    {
                        Disc = null;
                    }
                    else
                    {
                        float excelDisc = Convert.ToSingle(excelDiscDouble);// Fix для не дробных значений
                        Disc = excelDisc.ToString();

                        if (Disc.Contains(","))
                        {
                            Disc = Disc.Replace(",", ".");
                        }

                        if (Disc.Contains("-"))
                        {
                            Disc = Disc.Replace("-", "");
                        }
                    }


                    //var excelDisc = ObjWorkSheet.get_Range("Z" + colnum, Type.Missing).Value2;
                    //string DiscString;
                    //double Disc;
                    //if (excelDisc == null)// Если ячейки пустые, то ничего не вводить                 
                    //{
                    //    DiscString = null;
                    //}
                    //else
                    //{
                    //    DiscString = excelDisc.ToString();

                    //    if (DiscString.Contains(","))
                    //    {
                    //        DiscString = DiscString.Replace(",", ".");
                    //    }

                    //    if (DiscString.Contains("-"))
                    //    {
                    //        DiscString = DiscString.Replace("-", "");
                    //    }

                    //    Disc = Convert.ToDouble(DiscString);

                    //}



                    //Sale
                    var excelSaleDouble = ObjWorkSheet.get_Range("AA" + colnum, Type.Missing).Value2;
                    string Sale = "";
                    if (excelSaleDouble == null)
                    {
                        if (SSstat != "bk" || SSstat == "BK")
                        {
                            UserLog = new StreamWriter(destUserLog, true);
                            UserLog.WriteLine(Con + " - No Sale data");
                            UserLog.Close();
                            continue; //переход к следующей итерации FOR
                        }
                        
                    }
                    else
                    {
                        float excelSale = Convert.ToSingle(excelSaleDouble);// Fix для не дробных значений
                        string PreSale = excelSale.ToString();
                        
                        //Sale_MF.length = 14
                        if (PreSale.Contains(",") || PreSale.Contains("."))
                        {
                            if (PreSale.Contains(","))
                            {
                                PreSale = PreSale.Replace(",", ".");
                            }

                            if (PreSale.Length == 12) { Sale = "  " + PreSale; }
                            else if (PreSale.Length == 11) { Sale = "   " + PreSale; }
                            else if (PreSale.Length == 10) { Sale = "    " + PreSale; }
                            else if (PreSale.Length == 9) { Sale = "     " + PreSale; }
                            else if (PreSale.Length == 8) { Sale = "      " + PreSale; }
                            else if (PreSale.Length == 7) { Sale = "       " + PreSale; }
                            else if (PreSale.Length == 6) { Sale = "        " + PreSale; }
                            else if (PreSale.Length == 5) { Sale = "         " + PreSale; }
                            else if (PreSale.Length == 4) { Sale = "          " + PreSale; }
                            else if (PreSale.Length == 3) { Sale = "           " + PreSale; }
                            else
                            {
                                UserLog = new StreamWriter(destUserLog, true);
                                UserLog.WriteLine(Con + " - Sale too low");
                                UserLog.Close();
                                continue; //переход к следующей итерации FOR
                            }
                        }
                        else if (PreSale.Length <= 10)
                        {

                            if (PreSale.Length == 10) { Sale = " " + PreSale + ".00"; }
                            else if (PreSale.Length == 9) { Sale = "  " + PreSale + ".00"; }
                            else if (PreSale.Length == 8) { Sale = "   " + PreSale + ".00"; }
                            else if (PreSale.Length == 7) { Sale = "    " + PreSale + ".00"; }
                            else if (PreSale.Length == 6) { Sale = "     " + PreSale + ".00"; }
                            else if (PreSale.Length == 5) { Sale = "      " + PreSale + ".00"; }
                            else if (PreSale.Length == 4) { Sale = "       " + PreSale + ".00"; }
                            else if (PreSale.Length == 3) { Sale = "        " + PreSale + ".00"; }
                            else
                            {
                                UserLog = new StreamWriter(destUserLog, true);
                                UserLog.WriteLine(Con + " - Sale too low");
                                UserLog.Close();
                                continue; //переход к следующей итерации FOR
                            }

                        }
                        else
                        {
                            UserLog = new StreamWriter(destUserLog, true);
                            UserLog.WriteLine(Con + " - Sale has to much symbols");
                            UserLog.Close();
                            continue; //переход к следующей итерации FOR
                        }

                    }

                    #endregion//

                    //string CheckAcc = excelAcc.ToString();// Добавление 0 к Acc No
                    //string Acc = "";
                    //if (CheckAcc.Length < 9)
                    //{

                    //    do
                    //    {

                    //        Acc = "0" + CheckAcc;

                    //    } while (CheckAcc.Length == 9);
                    //}
                    //else
                    //{
                    //    CheckAcc = Acc;
                    //}





                    /////////////////////////////////////////////////////////////Data Entry//////////////////////////////////////////////////////////////////////////

                    

                    ForAwaitCol(21);
                    Thread.Sleep(600);
                    host.Send(Con);// Вводим CN Number
                    host.Send("<ENTER>");


                    //Провеорка на ошибку PK0105 -  NO ORDERS FOUND MATCHING THE SEARCH CRITERIA
                    //var PK0105 = "";
                    //PK0105 = disp.ScreenData[2, 24, 6];
                    if (disp.ScreenData[2, 24, 6] == "PK0105")
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - NO ORDERS FOUND MATCHING THE SEARCH CRITERIA");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR
                    }                    
                    logger.Debug("Con:" + Con, this.Text); //LOG


                    ForAwaitCol(15);// Проверка соответствия Div и Prod  + определения значения Revenue and Cost                 
                    var CheckDiv = "";
                    var CheckProd = "";
                    var CheckRevCost = "";
                    short j = 1;

                    do
                    {
                        short col = (Int16)(9 + j);
                        CheckDiv = disp.ScreenData[41, col, 1];
                        CheckRevCost = disp.ScreenData[74, col, 6];

                        if (Prod.Length == 3 )
                        {
                            CheckProd = disp.ScreenData[45, col, 3];
                        }
                        else if (Prod.Length == 2)
                        {
                            CheckProd = disp.ScreenData[45, col, 2];
                        }                    

                        Thread.Sleep(600);

                        if (CheckDiv == Div && CheckProd == Prod)
                        {
                            host.Send(j.ToString());
                            host.Send("<ENTER>");
                            break;//Прерывание цикла Do While
                        }                            
                        j++;

                    } while (CheckDiv.Trim() != "");

                    if (CheckDiv != Div && CheckProd != Prod)
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - Such Con with specified Div and Prod not found ");
                        UserLog.Close();

                        host.Send("<F12>");
                        Thread.Sleep(600);
                        host.Send("<F12>");
                        Thread.Sleep(2000);

                        if (teemApp.CurrentSession.Display.CursorCol == 25)// Пункт 7.1 Окно, которое нужно закрыть
                        {
                            host.Send("<F12>");
                            Thread.Sleep(600);
                        }

                        continue; //переход к следующей итерации FOR
                    }


                    if (CheckRevCost == "      ")// Проверка значения Revao & Cost
                    {
                        //////////////////////////////////////Функция Прохода по меню
                        ForAwaitCol(7);
                        host.Send("S");
                        host.Send("<ENTER>");
                        Thread.Sleep(600);

                        ForAwaitCol(21);
                        host.Send("<TAB>");

                        ForAwaitCol(39);
                        host.Send("<TAB>");

                        ForAwaitCol(52);
                        host.Send("<TAB>");

                        ForAwaitCol(75);
                        host.Send("<TAB>");

                        ForAwaitCol(8);
                        host.Send("<TAB>");

                        ForAwaitCol(47);
                        host.Send("<TAB>");

                        ForAwaitCol(69);
                        host.Send("<TAB>");

                        ForAwaitCol(20);
                        host.Send("<TAB>");

                        ForAwaitCol(31);
                        host.Send("<TAB>");

                        ForAwaitCol(40);
                        host.Send("<TAB>");

                        ForAwaitCol(51);
                        host.Send("<TAB>");

                        ForAwaitCol(20);
                        host.Send("<TAB>");

                        ForAwaitCol(20);
                        host.Send("<TAB>");

                        ForAwaitCol(75);
                        host.Send("<TAB>");

                        ForAwaitCol(20);
                        host.Send("<TAB>");

                        ForAwaitCol(27);
                        host.Send("<TAB>");

                        ForAwaitCol(34);
                        host.Send("<TAB>");
                    }
                    else if (CheckRevCost == "100.00")
                    {
                        //////////////////////////////////////Функция Прохода по меню СОКРАЩЕННАЯ
                        ForAwaitCol(7);
                        host.Send("S");
                        host.Send("<ENTER>");
                        Thread.Sleep(600);

                        ForAwaitCol(21);
                        host.Send("<TAB>");

                        ForAwaitCol(39);
                        host.Send("<TAB>");

                        ForAwaitCol(52);
                        host.Send("<TAB>");

                        ForAwaitCol(75);
                        host.Send("<TAB>");

                        ForAwaitCol(8);
                        host.Send("<TAB>");

                        ForAwaitCol(47);
                        host.Send("<TAB>");

                        ForAwaitCol(69);
                        host.Send("<TAB>");

                        ForAwaitCol(20);
                        host.Send("<TAB>");
                        Thread.Sleep(600);



                        ForAwaitCol(75);
                        Thread.Sleep(1000);
                        host.Send("<TAB>");

                        ForAwaitCol(20);
                        host.Send("<TAB>");

                        ForAwaitCol(27);
                        host.Send("<TAB>");

                        ForAwaitCol(34);
                        host.Send("<TAB>");
                    }
                    else
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - Con have already entered non-compliant Revenue and Cost value ");
                        UserLog.Close();

                        host.Send("<F12>");
                        Thread.Sleep(600);
                        host.Send("<F12>");
                        Thread.Sleep(2000);

                        if (teemApp.CurrentSession.Display.CursorCol == 25)// Пункт 7.1 Окно, которое нужно закрыть
                        {
                            host.Send("<F12>");
                            Thread.Sleep(600);
                        }

                        continue; //переход к следующей итерации FOR
                    }



                    ////////////////////////////////////////Функция Прохода по меню
                    //ForAwaitCol(7);
                    //host.Send("S");
                    //host.Send("<ENTER>");
                    //Thread.Sleep(600);

                    //ForAwaitCol(21);
                    //host.Send("<TAB>");

                    //ForAwaitCol(39);
                    //host.Send("<TAB>");

                    //ForAwaitCol(52);
                    //host.Send("<TAB>");

                    //ForAwaitCol(75);
                    //host.Send("<TAB>");

                    //ForAwaitCol(8);
                    //host.Send("<TAB>");

                    //ForAwaitCol(47);
                    //host.Send("<TAB>");

                    //ForAwaitCol(69);
                    //host.Send("<TAB>");

                    //ForAwaitCol(20);
                    //host.Send("<TAB>");

                    //ForAwaitCol(31);
                    //host.Send("<TAB>");

                    //ForAwaitCol(40);
                    //host.Send("<TAB>");

                    //ForAwaitCol(51);
                    //host.Send("<TAB>");

                    //ForAwaitCol(20);
                    //host.Send("<TAB>");

                    //ForAwaitCol(20);
                    //host.Send("<TAB>");

                    //ForAwaitCol(75);
                    //host.Send("<TAB>");

                    //ForAwaitCol(20);
                    //host.Send("<TAB>");

                    //ForAwaitCol(27);
                    //host.Send("<TAB>");

                    //ForAwaitCol(34);
                    //host.Send("<TAB>");


                    ForAwaitCol(24);
                    host.Send("<F4>");//Переход в меню отправки
                    ///////////////////////////////////////////////////////////////////Конец Функции
                    ForAwaitCol(43);
                    host.Send("<TAB>");

                    ForAwaitCol(62);
                    host.Send("<TAB>");

                    ForAwaitCol(7);
                    host.Send("<F4>");//Переход к сверке по аккаунту
                    Thread.Sleep(1000);

                    // Сверка по Account No
                    var MfAcc = disp.ScreenData[13, 8, 9];
                    if(Acc == MfAcc )
                    {                        
                        host.Send("<F12>");
                        Thread.Sleep(1000);
                    }
                    else
                    {
                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - Excel Acc № " + Acc + " does not match with Mainframe Acc № " + MfAcc);
                        UserLog.Close();

                        host.Send("<F12>");
                        Thread.Sleep(2000);
                        host.Send("<F12>");
                        Thread.Sleep(600);
                        host.Send("<F12>");
                        Thread.Sleep(2000);

                        if (teemApp.CurrentSession.Display.CursorCol == 25)// Пункт 7.1 Окно, которое нужно закрыть
                        {
                            host.Send("<F12>");
                            Thread.Sleep(600);
                        }

                        continue; //переход к следующей итерации FOR
                    }

                    ForAwaitCol(7);
                    host.Send("<Shift+TAB>");// Первое Введение Shift + TAB
                    Thread.Sleep(600);

                    ForAwaitCol(62);
                    host.Send("<Shift+TAB>");// Первое Введение Shift + TAB
                    Thread.Sleep(600);



                    ForAwaitCol(43);//Проверка на то какой статус проставлять - BK или BK+VE                    


                    if (SSstat == "bk" || SSstat == "BK")
                    {
                        /////////////////////////////////////////////////Начало функции BK
                        host.Send("bk"); // SS Status
                        Thread.Sleep(600);

                        ForAwaitCol(62);
                        host.Send("<TAB>");

                        ForAwaitCol(7);
                        host.Send("<TAB>");

                        ForAwaitCol(46);
                        host.Send("<TAB>");

                        ForAwaitCol(7);
                        host.Send("<TAB>");

                        ForAwaitCol(46);
                        host.Send("<TAB>");

                        ForAwaitCol(13);
                        host.Send("<TAB>");

                        ForAwaitCol(50);
                        host.Send("<TAB>");

                        ForAwaitCol(14);
                        host.Send("<END>");
                        Thread.Sleep(600);
                        host.Send(Weight); // Weight Только КГ!!!!
                        host.Send("<TAB>");

                        ForAwaitCol(28);
                        host.Send("<TAB>");

                        ForAwaitCol(55);
                        host.Send("<TAB>");

                        ForAwaitCol(73);
                        host.Send("<TAB>");

                        ForAwaitCol(8);
                        host.Send("<TAB>");

                        ForAwaitCol(20);
                        host.Send("<TAB>");

                        ForAwaitCol(37);
                        host.Send("<TAB>");

                        ForAwaitCol(53);
                        host.Send("<TAB>");

                        ForAwaitCol(70);
                        host.Send("<TAB>");

                        if (CheckRevCost == "      ") // Проверка ChecRevCost - надо ли вводить дату и время пикапа и REvao
                        {
                            ForAwaitCol(20);
                            host.Send("<END>");
                            Thread.Sleep(600);
                            host.Send(CollDate); // Collection Date
                            Thread.Sleep(600);
                            host.Send("<TAB>");

                            ForAwaitRow(16);
                            host.Send(CollTime); // =0900  
                            Thread.Sleep(600);
                            ForAwaitCol(31);
                            host.Send(CollTimeTo); // =1800
                            host.Send("<TAB>");

                            ForAwaitCol(51);
                            host.Send("<TAB>");

                            ForAwaitCol(22);
                            host.Send(DelDate); // Delivery Date
                            Thread.Sleep(600);
                            host.Send("<TAB>");

                            ForAwaitCol(40);
                            host.Send(Deltime); // =0900  
                            Thread.Sleep(600);

                            ForAwaitCol(56);
                            host.Send(DelDate); // Delivery Date
                            Thread.Sleep(600);
                            host.Send("<TAB>");

                            ForAwaitCol(74);
                            host.Send(DeltimeTo); // =2359  
                            Thread.Sleep(600);

                            //Вход в Tariff No
                            ForAwaitCol(7);
                            host.Send("<F10>");

                            ForAwaitCol(15);
                            host.Send("<F5>");
                            
                            


                            ForAwaitCol(7);
                            host.Send("<F4>");

                            ForAwaitCol(23);
                            host.Send("43");
                            host.Send("<ENTER>");

                            ForAwaitCol(7);
                            host.Send("<TAB>");

                            ForAwaitCol(15);
                            host.Send("<TAB>");

                            ForAwaitCol(28);
                            host.Send("100"); // Revao = 100
                            Thread.Sleep(600);
                            host.Send("<ENTER>");

                            host.Send("<F12>");
                            Thread.Sleep(2000);
                            host.Send("<F12>");
                            Thread.Sleep(600);
                            host.Send("<F12>");
                            Thread.Sleep(2000);

                            ForAwaitCol(73);
                        }

                        
                        

                        //ForAwaitCol(73); - Больше не нужно
                        host.Send("<ENTER>"); // 1й раз
                        Thread.Sleep(1000);
                        host.Send("<ENTER>");// 2й раз
                        Thread.Sleep(1000);
                        host.Send("<ENTER>");// 3й раз
                        Thread.Sleep(1000);
                        host.Send("<ENTER>");// 4й раз
                        Thread.Sleep(1000);
                        host.Send("<ENTER>");// 5й раз
                        Thread.Sleep(1000);
                        host.Send("<ENTER>");// 6й раз

                        ForAwaitCol(62);
                        host.Send("y");
                        Thread.Sleep(600);
                        host.Send("<ENTER>");

                        Thread.Sleep(600);
                        host.Send("<F12>");
                        Thread.Sleep(1500);

                        if (teemApp.CurrentSession.Display.CursorCol == 25)// Пункт 7.1 Окно, которое нужно закрыть
                        {
                            host.Send("<F12>");
                            Thread.Sleep(600);
                        }

                        UserLog = new StreamWriter(destUserLog, true);
                        UserLog.WriteLine(Con + " - Done");
                        UserLog.Close();
                        continue; //переход к следующей итерации FOR

                    }

                    else if(SSstat != "BK+VE" || SSstat != "bk+VE" || SSstat != "BK+ve")
                    {

                        /////////////////////////////////////////////////Начало функции BK
                        host.Send("bk"); // SS Status
                        Thread.Sleep(600);

                        ForAwaitCol(62);
                        host.Send("<TAB>");

                        ForAwaitCol(7);
                        host.Send("<TAB>");

                        ForAwaitCol(46);
                        host.Send("<TAB>");

                        ForAwaitCol(7);
                        host.Send("<TAB>");

                        ForAwaitCol(46);
                        host.Send("<TAB>");

                        ForAwaitCol(13);
                        host.Send("<TAB>");

                        ForAwaitCol(50);
                        host.Send("<TAB>");

                        ForAwaitCol(14);
                        host.Send("<END>");
                        Thread.Sleep(600);
                        host.Send(Weight); // Weight Только КГ!!!!
                        host.Send("<TAB>");

                        ForAwaitCol(28);
                        host.Send("<TAB>");

                        ForAwaitCol(55);
                        host.Send("<TAB>");

                        ForAwaitCol(73);
                        host.Send("<TAB>");

                        ForAwaitCol(8);
                        host.Send("<TAB>");

                        ForAwaitCol(20);
                        host.Send("<TAB>");

                        ForAwaitCol(37);
                        host.Send("<TAB>");

                        ForAwaitCol(53);
                        host.Send("<TAB>");

                        ForAwaitCol(70);
                        host.Send("<TAB>");

                        if (CheckRevCost == "      ") // Проверка ChecRevCost - надо ли вводить дату и время пикапа
                        {
                            ForAwaitCol(20);
                            host.Send("<END>");
                            Thread.Sleep(600);
                            host.Send(CollDate); // Collection Date
                            Thread.Sleep(600);
                            host.Send("<TAB>");

                            ForAwaitRow(16);
                            host.Send(CollTime); // =0900  
                            Thread.Sleep(600);
                            ForAwaitCol(31);
                            host.Send(CollTimeTo); // =1800
                            host.Send("<TAB>");

                            ForAwaitCol(51);
                            host.Send("<TAB>");

                            ForAwaitCol(22);
                            host.Send(DelDate); // Delivery Date
                            Thread.Sleep(600);
                            host.Send("<TAB>");

                            ForAwaitCol(40);
                            host.Send(Deltime); // =0900  
                            Thread.Sleep(600);

                            ForAwaitCol(56);
                            host.Send(DelDate); // Delivery Date
                            Thread.Sleep(600);
                            host.Send("<TAB>");

                            ForAwaitCol(74);
                            host.Send(DeltimeTo); // =2359  
                            Thread.Sleep(600);
                        }


                        //Вход в Tariff No
                        ForAwaitCol(7);
                        host.Send("<F10>");
                        Thread.Sleep(1000);

                        if (disp.ScreenData[7, 10, 5] == "REVAO")// Удаление Revao = 100
                        {
                            ForAwaitCol(15);
                            host.Send("1");
                            Thread.Sleep(600);
                            host.Send("<ENTER>");

                            ForAwaitCol(7);
                            host.Send("<TAB>");

                            ForAwaitCol(15);
                            host.Send("<TAB>");

                            ForAwaitCol(28);
                            host.Send("<TAB>");

                            ForAwaitCol(48);
                            host.Send("<TAB>");

                            ForAwaitCol(67);
                            host.Send("<TAB>");

                            ForAwaitCol(77);
                            host.Send("Y");
                            Thread.Sleep(600);
                            host.Send("<ENTER>");
                        }

                        ForAwaitCol(15);
                        host.Send("<F5>");

                        ////////////////////////////////////////////////////////Конец Функции BK

                        
                        if (CMairVendor == null && CMairQt == null)// Если ячейки пустые, то ничего не вводить                 
                        {
                                                                                  
                        }
                        else if (CMairVendor == null)
                        {
                            UserLog = new StreamWriter(destUserLog, true);
                            UserLog.WriteLine(Con + " - No CMair Vendor data");
                            UserLog.Close();
                        }
                        else if (CMairQt == null)
                        {
                            UserLog = new StreamWriter(destUserLog, true);
                            UserLog.WriteLine(Con + " - No CMair Quote Amount");
                            UserLog.Close();
                        }
                        else
                        {
                            Thread.Sleep(1000);
                            if (disp.CursorCol == 7)
                            {
                                host.Send("<F4>");
                                Thread.Sleep(600);
                            }


                            ForAwaitCol(23);
                            host.Send(CMair);// =4
                            host.Send("<ENTER>");

                            ForAwaitCol(7);
                            host.Send("<TAB>");

                            ForAwaitCol(15);
                            host.Send(CMairVendor); // CMairVendor
                            Thread.Sleep(600);
                            host.Send("<TAB>");

                            ForAwaitCol(28);
                            host.Send(CMairQt); // CMairQt
                            Thread.Sleep(600);
                            host.Send("<ENTER>");
                        }

                        if (LCLPU_Vendor == null && LCLPU_Qt == null)// Если ячейки пустые, то ничего не вводить                 
                        {
                            
                        }
                        else if (LCLPU_Vendor == null)
                        {
                            UserLog = new StreamWriter(destUserLog, true);
                            UserLog.WriteLine(Con + " - No LCLPU Vendor data");
                            UserLog.Close();
                        }
                        else if (LCLPU_Qt == null)
                        {
                            UserLog = new StreamWriter(destUserLog, true);
                            UserLog.WriteLine(Con + " - No LCLPU Quote Amount");
                            UserLog.Close();
                        }
                        else
                        {
                            //ForAwaitCol(7);
                            //host.Send("<F4>");
                            Thread.Sleep(1000);
                            if (disp.CursorCol == 7)
                            {
                                host.Send("<F4>");
                                Thread.Sleep(600);
                            }

                            ForAwaitCol(23);
                            Thread.Sleep(600);
                            host.Send(LCLPU); // =24
                            host.Send("<ENTER>");

                            ForAwaitCol(7);
                            host.Send("<TAB>");

                            ForAwaitCol(15);
                            host.Send(LCLPU_Vendor); // LCLPU_Vendor
                            Thread.Sleep(600);
                            host.Send("<TAB>");

                            ForAwaitCol(28);
                            host.Send(LCLPU_Qt); // LCLPU_Qt
                            Thread.Sleep(600);
                            host.Send("<ENTER>");
                        }

                        if (LCDL_Vendor == null && LCDL_Qt == null)// Если ячейки пустые, то ничего не вводить                 
                        {
                            
                        }
                        else if (LCDL_Vendor == null)
                        {
                            UserLog = new StreamWriter(destUserLog, true);
                            UserLog.WriteLine(Con + " - No LCDL Vendor data");
                            UserLog.Close();
                        }
                        else if (LCDL_Qt == null)
                        {
                            UserLog = new StreamWriter(destUserLog, true);
                            UserLog.WriteLine(Con + " - No LCDL Quote Amount");
                            UserLog.Close();
                        }
                        else
                        {
                            Thread.Sleep(1000);
                            if (disp.CursorCol == 7)
                            {
                                host.Send("<F4>");
                                Thread.Sleep(600);
                            }

                            ForAwaitCol(23);
                            host.Send(LCDL); // =23
                            host.Send("<ENTER>");

                            ForAwaitCol(7);
                            host.Send("<TAB>");

                            ForAwaitCol(15);
                            host.Send(LCDL_Vendor); // LCDL_Vendor
                            Thread.Sleep(600);
                            host.Send("<TAB>");

                            ForAwaitCol(28);
                            host.Send(LCDL_Qt); // LCDL_Qt
                            Thread.Sleep(600);
                            host.Send("<ENTER>");
                        }

                        if (Revao == null)// Если ячейки пустые, то ничего не вводить                 
                        {
                            UserLog = new StreamWriter(destUserLog, true);
                            UserLog.WriteLine(Con + " - No Revao data");
                            UserLog.Close();
                        }
                        else
                        {
                            Thread.Sleep(1000);
                            if (disp.CursorCol == 7)
                            {
                                host.Send("<F4>");
                                Thread.Sleep(600);
                            }

                            ForAwaitCol(23);
                            host.Send(Revao_n); // =43
                            host.Send("<ENTER>");

                            ForAwaitCol(7);
                            host.Send("<TAB>");

                            ForAwaitCol(15);
                            host.Send("<TAB>");

                            ForAwaitCol(28);
                            host.Send(Revao); // Revao
                            Thread.Sleep(600);
                            host.Send("<ENTER>");
                        }


                        if (Disc == null)// Если ячейки пустые, то ничего не вводить                 
                        {
                            UserLog = new StreamWriter(destUserLog, true);
                            UserLog.WriteLine(Con + " - No Disc data");
                            UserLog.Close();
                        }
                        else
                        {
                            //UserLog = new StreamWriter(destUserLog, true);
                            //UserLog.WriteLine(Con + " - Was added additional Disc = 100, because already was entered temporary revao value");
                            //UserLog.Close();

                            Thread.Sleep(1000);
                            if (disp.CursorCol == 7)
                            {
                                host.Send("<F4>");
                                Thread.Sleep(600);
                            }

                            ForAwaitCol(23);
                            host.Send(Disc_n); // =7
                            host.Send("<ENTER>");

                            ForAwaitCol(7);
                            host.Send("<TAB>");

                            ForAwaitCol(15);
                            host.Send("<TAB>");

                            ForAwaitCol(28);
                            host.Send(Disc); // Disc
                            Thread.Sleep(600);
                            host.Send("<ENTER>");
                        }


                        //if(CheckRevCost == "100.00")// Дополнительная скидка, чтобы убрать ранее введеный ревао
                        //{
                        //    Thread.Sleep(1000);
                        //    if (disp.CursorCol == 7)
                        //    {
                        //        host.Send("<F4>");
                        //        Thread.Sleep(600);
                        //    }

                        //    ForAwaitCol(23);
                        //    host.Send(Disc_n); // =7
                        //    host.Send("<ENTER>");

                        //    ForAwaitCol(7);
                        //    host.Send("<TAB>");

                        //    ForAwaitCol(15);
                        //    host.Send("<TAB>");

                        //    ForAwaitCol(28);
                        //    host.Send("100.00"); // Disc
                        //    Thread.Sleep(600);
                        //    host.Send("<ENTER>");
                        //}


                        ForAwaitCol(7);
                        string NetSale = disp.ScreenData[9, 21, 14];
                        if (NetSale == Sale)/////////////////////////////////// Нужны тесты!!!
                        {
                            ForAwaitCol(7);
                            host.Send("<F12>");
                            ForAwaitCol(15);
                            host.Send("<F12>");
                            Thread.Sleep(1000);
                            host.Send("<F12>");

                            ForAwaitCol(73);
                            host.Send("<ENTER>"); // 1й раз
                            Thread.Sleep(1000);
                            host.Send("<ENTER>");// 2й раз
                            Thread.Sleep(1000);
                            host.Send("<ENTER>");// 3й раз
                            Thread.Sleep(1000);
                            host.Send("<ENTER>");// 4й раз
                            Thread.Sleep(1000);
                            host.Send("<ENTER>");// 5й раз
                            Thread.Sleep(1000);
                            host.Send("<ENTER>");// 6й раз

                            ForAwaitCol(62);
                            host.Send("y");
                            Thread.Sleep(1000);
                            host.Send("<ENTER>");


                            ForAwaitCol(15);                            
                            host.Send("1");
                            Thread.Sleep(600);
                            host.Send("<ENTER>");

                            //////////////////////////////////////Функция Прохода по меню СОКРАЩЕННАЯ
                            ForAwaitCol(7);
                            host.Send("S");
                            host.Send("<ENTER>");
                            Thread.Sleep(600);

                            ForAwaitCol(21);
                            host.Send("<TAB>");

                            ForAwaitCol(39);
                            host.Send("<TAB>");

                            ForAwaitCol(52);
                            host.Send("<TAB>");

                            ForAwaitCol(75);
                            host.Send("<TAB>");

                            ForAwaitCol(8);
                            host.Send("<TAB>");

                            ForAwaitCol(47);
                            host.Send("<TAB>");

                            ForAwaitCol(69);
                            host.Send("<TAB>");

                            ForAwaitCol(20);
                            host.Send("<TAB>");
                            Thread.Sleep(600);

                            

                            ForAwaitCol(75);
                            Thread.Sleep(1000);
                            host.Send("<TAB>");

                            ForAwaitCol(20);
                            host.Send("<TAB>");

                            ForAwaitCol(27);
                            host.Send("<TAB>");
                             
                            ForAwaitCol(34);
                            host.Send("<TAB>");


                            ForAwaitCol(24);
                            host.Send("<F4>");//Переход в меню отправки
                            ///////////////////////////////////////////////////////////////////Конец Функции


                            ForAwaitCol(43);//Ввод статуса VE
                            host.Send("ve");

                            host.Send("<ENTER>"); // 1й раз
                            Thread.Sleep(1000);
                            host.Send("<ENTER>");// 2й раз
                            Thread.Sleep(1000);
                            host.Send("<ENTER>");// 3й раз
                            Thread.Sleep(1000);
                            host.Send("<ENTER>");// 4й раз
                            Thread.Sleep(1000);
                            host.Send("<ENTER>");// 5й раз
                            Thread.Sleep(1000);
                            host.Send("<ENTER>");// 6й раз

                            ForAwaitCol(62);
                            host.Send("y");
                            Thread.Sleep(1000);
                            host.Send("<ENTER>");

                            Thread.Sleep(1000);
                            host.Send("<F12>");
                            Thread.Sleep(1500);

                            if (teemApp.CurrentSession.Display.CursorCol == 25)// Пункт 7.1 Окно, которое нужно закрыть
                            {
                                host.Send("<F12>");
                                Thread.Sleep(600);
                            }

                            UserLog = new StreamWriter(destUserLog, true);
                            UserLog.WriteLine(Con + " - Done");
                            UserLog.Close();
                            continue; //переход к следующей итерации FOR

                        }

                        else//Выход в исходную точку
                        {
                            ForAwaitCol(7);
                            host.Send("<F12>");
                            ForAwaitCol(15);
                            host.Send("<F12>");
                            Thread.Sleep(600);
                            host.Send("<F12>");
                            Thread.Sleep(600);
                            host.Send("<F12>");
                            Thread.Sleep(600);
                            host.Send("<F12>");


                            UserLog = new StreamWriter(destUserLog, true);
                            UserLog.WriteLine(Con + " - total cost: "+NetSale+ " is not equal to Sale: "+Sale);
                            UserLog.Close();
                            continue; //переход к следующей итерации FOR
                        }

                    }





                    //ForAwaitRow(9);
                    //Thread.Sleep(600);
                    //host.Send(RecName); // Receaver Name
                    //Thread.Sleep(600);
                    //host.Send("<TAB>");

                    //ForAwaitRow(10);
                    //host.Send(RecAddr); // Receaver Address
                    //Thread.Sleep(600);
                    //if(RecAddr.Length<30)
                    //{
                    //    host.Send("<TAB>");
                    //}
                    
                    //ForAwaitCol(46);
                    //host.Send("<TAB>");
                    //ForAwaitRow(11);
                    //host.Send("<TAB>");

                    //ForAwaitRow(12);
                    //host.Send(RecTown); // Receaver Town
                    //Thread.Sleep(600);
                    //host.Send("<TAB>");

                    //ForAwaitCol(57);
                    //host.Send(RecPost); // Receaver Postcode
                    //host.Send("<ENTER>");
                    //Thread.Sleep(1000);
                    //host.Send("<TAB>");

                    //ForAwaitCol(7);
                    //host.Send("<TAB>");

                    //ForAwaitCol(46);
                    //host.Send("<TAB>");

                    //ForAwaitCol(13);
                    //host.Send("<TAB>");

                    //ForAwaitCol(50);
                    //host.Send(GdsDesk); // GDS Desc 
                    //host.Send("<TAB>");

                    //ForAwaitCol(14);
                    //host.Send(Weight); // Weight Только КГ!!!!
                    //host.Send("<TAB>");
                    //ForAwaitCol(28);
                    ////host.Send(WeightG);
                    //host.Send("<TAB>");

                    //ForAwaitCol(55);
                    //host.Send("<TAB>");

                    //ForAwaitCol(73);
                    //host.Send("<TAB>");

                    //ForAwaitCol(8);
                    //host.Send("<TAB>");

                    //ForAwaitCol(20);
                    //host.Send(Items); // Items
                    //host.Send("<TAB>");

                    //ForAwaitCol(37);
                    //host.Send(Length); // =10 
                    //host.Send("<TAB>");

                    //ForAwaitCol(53);
                    //host.Send(Widht); // =10 
                    //host.Send("<TAB>");

                    //ForAwaitCol(70);
                    //host.Send(Height); // =10  
                    //host.Send("<TAB>");

                    //ForAwaitCol(20);
                    //host.Send(CollDate); // Collection Date
                    //Thread.Sleep(600);
                    //host.Send("<TAB>");

                    //ForAwaitRow(16);
                    //host.Send(CollTime); // =1100  
                    //Thread.Sleep(600);
                    //ForAwaitCol(31);
                    //host.Send(CollTimeTo); // =1400
                    //host.Send("<TAB>");

                    ////ForAwaitCol(40);
                    ////host.Send("<TAB>");

                    //ForAwaitCol(51);
                    //host.Send("<TAB>");

                    //ForAwaitCol(22);
                    //host.Send(DelDate); // Delivery Date
                    //Thread.Sleep(600);
                    //host.Send("<TAB>");

                    //ForAwaitCol(40);
                    //host.Send(Deltime); // =2359  
                    //Thread.Sleep(600);

                    //ForAwaitCol(56);
                    //host.Send(DelDate); // Delivery Date
                    //Thread.Sleep(600);
                    //host.Send("<TAB>");

                    //ForAwaitCol(74);
                    //host.Send(DeltimeTo); // =2359  
                    //Thread.Sleep(600);

                    //ForAwaitCol(7);
                    //host.Send(Div); // =s                        
                    //host.Send("<TAB>");

                    //ForAwaitCol(20);
                    //host.Send(Prod); // Prod
                    //Thread.Sleep(600);
                    //host.Send("<TAB>");

                    //ForAwaitCol(37);
                    //host.Send("<TAB>");
                    //ForAwaitCol(45);
                    //host.Send("<TAB>");
                    //ForAwaitCol(53);
                    //host.Send("<TAB>");
                    //ForAwaitCol(61);
                    //host.Send("<TAB>");

                    //ForAwaitCol(76);
                    //host.Send(Payer); // Payer
                    //Thread.Sleep(600);
                    //host.Send("<TAB>");


                    //ForAwaitCol(38);
                    //host.Send("<TAB>");
                    //ForAwaitCol(76);
                    //host.Send("<TAB>");
                    //ForAwaitCol(17);
                    //host.Send("<TAB>");
                    //ForAwaitCol(47);
                    //host.Send("<TAB>");

                    //ForAwaitCol(13);
                    //host.Send(Stack); // =y

                    ////Вход в Tariff No
                    //ForAwaitCol(48);
                    //host.Send("<F10>");

                    //ForAwaitCol(15);
                    //host.Send("<F5>");

                    

                    //if (CMairVendor == null || CMairQt == null)// Если ячейки пустые, то ничего не вводить                 
                    //{

                    //}
                    //else
                    //{
                    //    Thread.Sleep(1000);
                    //    if (disp.CursorCol == 7)
                    //    {
                    //        host.Send("<F4>");
                    //        Thread.Sleep(600);
                    //    }


                    //    ForAwaitCol(23);
                    //    host.Send(CMair);// =4
                    //    host.Send("<ENTER>");

                    //    ForAwaitCol(7);
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(15);
                    //    host.Send(CMairVendor); // CMairVendor
                    //    Thread.Sleep(600);
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(28);
                    //    host.Send(CMairQt); // CMairQt
                    //    Thread.Sleep(600);
                    //    host.Send("<ENTER>");
                    //}

                    //if (LCLPU_Vendor == null || LCLPU_Qt == null)// Если ячейки пустые, то ничего не вводить                 
                    //{

                    //}
                    //else
                    //{
                    //    //ForAwaitCol(7);
                    //    //host.Send("<F4>");
                    //    Thread.Sleep(1000);
                    //    if (disp.CursorCol == 7)
                    //    {                            
                    //        host.Send("<F4>");
                    //        Thread.Sleep(600);
                    //    }

                    //    ForAwaitCol(23);
                    //    Thread.Sleep(600);
                    //    host.Send(LCLPU); // =24
                    //    host.Send("<ENTER>");

                    //    ForAwaitCol(7);
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(15);
                    //    host.Send(LCLPU_Vendor); // LCLPU_Vendor
                    //    Thread.Sleep(600);
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(28);
                    //    host.Send(LCLPU_Qt); // LCLPU_Qt
                    //    Thread.Sleep(600);
                    //    host.Send("<ENTER>");
                    //}

                    //if (LCDL_Vendor == null || LCDL_Qt == null)// Если ячейки пустые, то ничего не вводить                 
                    //{

                    //}
                    //else
                    //{
                    //    Thread.Sleep(1000);
                    //    if (disp.CursorCol == 7)
                    //    {
                    //        host.Send("<F4>");
                    //        Thread.Sleep(600);
                    //    }

                    //    ForAwaitCol(23);
                    //    host.Send(LCDL); // =23
                    //    host.Send("<ENTER>");

                    //    ForAwaitCol(7);
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(15);
                    //    host.Send(LCLPU_Vendor); // LCDL_Vendor
                    //    Thread.Sleep(600);
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(28);
                    //    host.Send(LCDL_Qt); // LCDL_Qt
                    //    Thread.Sleep(600);
                    //    host.Send("<ENTER>");
                    //}

                    //if (Revao == null)// Если ячейки пустые, то ничего не вводить                 
                    //{

                    //}
                    //else
                    //{
                    //    Thread.Sleep(1000);
                    //    if (disp.CursorCol == 7)
                    //    {
                    //        host.Send("<F4>");
                    //        Thread.Sleep(600);
                    //    }

                    //    ForAwaitCol(23);
                    //    host.Send(Revao_n); // =43
                    //    host.Send("<ENTER>");

                    //    ForAwaitCol(7);
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(15);
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(28);
                    //    host.Send(Revao); // Revao
                    //    Thread.Sleep(600);
                    //    host.Send("<ENTER>");
                    //}


                    //if (Disc == null)// Если ячейки пустые, то ничего не вводить                 
                    //{

                    //}
                    //else
                    //{
                    //    Thread.Sleep(1000);
                    //    if (disp.CursorCol == 7)
                    //    {
                    //        host.Send("<F4>");
                    //        Thread.Sleep(600);
                    //    }

                    //    ForAwaitCol(23);
                    //    host.Send(Disc_n); // =7
                    //    host.Send("<ENTER>");

                    //    ForAwaitCol(7);
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(15);
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(28);
                    //    host.Send(Disc); // Disc
                    //    Thread.Sleep(600);
                    //    host.Send("<ENTER>");
                    //}


                    //ForAwaitCol(7);
                    //host.Send("<F12>");
                    //ForAwaitCol(15);
                    //host.Send("<F12>");
                    //Thread.Sleep(600);
                    //host.Send("<F12>");

                    //ForAwaitCol(73);
                    //host.Send("<ENTER>"); // 1й раз
                    //Thread.Sleep(1000);
                    //host.Send("<ENTER>");// 2й раз
                    //Thread.Sleep(1000);
                    //host.Send("<ENTER>");// 3й раз
                    //Thread.Sleep(1000);
                    //host.Send("<ENTER>");// 4й раз
                    //Thread.Sleep(1000);
                    //host.Send("<ENTER>");// 5й раз
                    //Thread.Sleep(1000);
                    //host.Send("<ENTER>");// 6й раз

                    //ForAwaitCol(62);
                    //host.Send("y");
                    //Thread.Sleep(600);
                    //host.Send("<ENTER>");


                }


                // Закрываем TeemTalk
                TeemTalkClose();
                logger.Debug("TeemTalkNormalClose", this.Text); //LOG


                //teemApp.Close();
                //foreach (Process proc in Process.GetProcessesByName("teem2k"))
                //{
                //proc.Kill();
                //}
                //teemApp.Application.Close();
                //Thread.Sleep(1000);
                //host.Send("<ENTER>");

                //Thread.Sleep(2000);
                //MessageBox.Show("Готово!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //return;
                Process.Start(destUserLog);

                //Закрываем Excel
                ObjWorkBook.Close(true);
                ObjWorkExcel.Quit();
                foreach (Process proc in Process.GetProcessesByName("excel"))
                {
                    proc.Kill();
                }

                              

            }
            catch (Exception ex)
            {
                // Вывод сообщения об ошибке
                logger.Debug(ex.ToString());
            }




        }

        static void TeemTalkClose()// Закрываем TeemTalk
        {

            teemApp.CurrentSession.Network.Close();
            Thread.Sleep(500);
            teemApp.Close();
        }


        static bool ForAwait(short col, short row, string keyword)
        {
            byte count = 0;
            
                do
                {
                    count++;
                    
                    if (count > 70)
                    {
                        teemApp.CurrentSession.Network.Close();
                        Thread.Sleep(1000);
                        teemApp.Close();

                        System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("teem2k");

                        foreach (System.Diagnostics.Process p in process)
                        {
                            if (!string.IsNullOrEmpty(p.ProcessName))
                            {
                                try
                                {
                                    p.Kill();
                                }

                                catch (Exception ex)
                                {
                                    // Вывод сообщения об ошибке
                                    logger.Debug(ex.ToString());
                                }
                            }
                        }

                        return false;
                    }

                    Thread.Sleep(100);

                } while ((teemApp.CurrentSession.Display.ScreenData[col, row, (short)keyword.Length] != keyword));
            return true;
        }

        static bool ForAwaitRow(short keyword)
        {
            byte count = 0;

            do
            {
                count++;

                if (count > 70)
                {
                    teemApp.CurrentSession.Network.Close();
                    Thread.Sleep(1000);
                    teemApp.Close();

                    System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("teem2k");

                    foreach (System.Diagnostics.Process p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch (Exception ex)
                            {
                                // Вывод сообщения об ошибке
                                logger.Debug(ex.ToString());
                            }
                        }
                    }

                    return false;
                }

                Thread.Sleep(100);

            } while ((teemApp.CurrentSession.Display.CursorRow != keyword));
            return true;
        }
        static bool ForAwaitCol(short keyword)
        {
            byte count = 0;

            do
            {
                count++;

                if (count > 70)
                {
                    teemApp.CurrentSession.Network.Close();
                    Thread.Sleep(1000);
                    teemApp.Close();

                    System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("teem2k");

                    foreach (System.Diagnostics.Process p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch (Exception ex)
                            {
                                // Вывод сообщения об ошибке
                                logger.Debug(ex.ToString());
                            }
                        }
                    }

                    return false;
                }

                Thread.Sleep(100);

            } while ((teemApp.CurrentSession.Display.CursorCol != keyword));
            return true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
