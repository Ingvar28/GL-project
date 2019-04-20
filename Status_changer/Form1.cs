using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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

            

                //поиск файла Excel
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Multiselect = false;
                ofd.DefaultExt = "*.xls;*.xlsx";
                ofd.Filter = "Microsoft Excel (*.xls*)|*.xls*";
                ofd.Title = "Выберите документ Excel";
                if (ofd.ShowDialog() != DialogResult.OK)
                {
                    MessageBox.Show("Вы не выбрали файл для открытия", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string xlFileName = ofd.FileName; //имя нашего Excel файла

                Excel.Application ObjWorkExcel = new Excel.Application(); //создаём приложение Excel
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(xlFileName); //открываем наш файл 
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист





                //Login into Mainframe
                //var login = textBox1.Text;
                //var password = textBox2.Text;
                var login = Properties.Settings.Default.loginMF;
                var password = Properties.Settings.Default.pwdMF;
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

                //if (!ForAwait(2, 2, "Command")) goto StartMaimframe;
                ForAwait(2, 2, "Command");
                host.Send("2");
                host.Send("<ENTER>");

                ForAwait(20, 7, "Job Description");
                host.Send("<F12>");
                Thread.Sleep(500);
                if (disp.CursorRow != 2)

                host.Send("JK04");
                logger.Debug("JK04", this.Text); //LOG
                host.Send("<ENTER>");



                



                var last = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);//1 ячейку                          
                int lastUsedRow = last.Row;

                //Дата, введенная пользователем в DateSform1
                //var DateSform = this.DateSform1.Text;
                ////DateTime DateS = DateTime.Parse((string) DateSform);
                //logger.Debug(DateSform, this.Text); //LOG


                // foreach (DataRow row in consData.Rows) // Старое
                //for (int i = 2; i <= lastUsedRow; i++)
                for (int i = 2; i <= lastUsedRow; i++) 
                {
                    // Colnum
                    string colnum = Convert.ToString(i);
                    //string colnum = "1";



                    // DataS эскпорт                          
                    //Excel.Range exceldateS = ObjWorkSheet.get_Range("S" + colnum);             
                    //object dateS_v = exceldateS.Value2;

                    //if (dateS_v == null || dateS_v is string)
                    //{
                    //    continue; //переход к следующей итерации FOR
                    //}

                    //DateTime dSt = DateTime.FromOADate((double)dateS_v);

                    //string dateS = dSt.ToString("dd.MM.yyyy", CultureInfo.GetCultureInfo("RU-ru"));

                    //if (dateS_v is double)
                    //{
                    //    dSt = DateTime.FromOADate((double)dateS_v);
                    //}
                    //else
                    //{
                    //    DateTime.TryParse((string)dateS_v, out dSt);
                    //}


                    // Done
                    string done = "DONE";


                    //Если статус Введенная дата есть в ячейке, то  цикл продолжается, если нет, то перескакивает к следующему i
                    //if (DateSform == dateS) 
                    //{
                        logger.Debug(colnum, this.Text); //LOG

                        //Объявление переменных

                        //CName - Caller Name
                        var CName = "0";
                        //TelNo - Telephone No
                        var TelNo = "0";

                        //SrCrit - Search Criteria из колонки acc (A)
                        var excelSrCrit = ObjWorkSheet.get_Range("A" + colnum, Type.Missing).Value2;
                        string SrCrit = excelSrCrit.ToString();

                        //Select - Selection
                        var Select = "41";
                        //SSstat - SS status
                        var SSstat = "BK";

                        //Con - Connote из колонки CN Number
                        var excelcon = ObjWorkSheet.get_Range("B" + colnum, Type.Missing).Value2;
                        string Con = excelcon.ToString();


                        //RecName - Receaver Name
                        var excelRecName = ObjWorkSheet.get_Range("C" + colnum, Type.Missing).Value2;
                        string RecName = excelRecName.ToString();

                        //RecAddr - Receaver Address
                        var excelRecAddr = ObjWorkSheet.get_Range("D" + colnum, Type.Missing).Value2;
                        string RecAddr = excelRecAddr.ToString();

                        //RecTown - Receaver Town
                        var excelRecTown = ObjWorkSheet.get_Range("E" + colnum, Type.Missing).Value2;
                        string RecTown = excelRecTown.ToString();

                        //RecPost - Receaver Postcode
                        var excelRecPost = ObjWorkSheet.get_Range("F" + colnum, Type.Missing).Value2;
                        string RecPost = excelRecPost.ToString();


                        //GdsDesk - GDS Desc  
                        var excelGdsDesk = ObjWorkSheet.get_Range("G" + colnum, Type.Missing).Value2;
                        string GdsDesk = excelGdsDesk.ToString();

                        //WeightKG - Weight KG
                        var excelWeightKG = ObjWorkSheet.get_Range("H" + colnum, Type.Missing).Value2;
                        string WeightKG = excelWeightKG.ToString();
                        //WeightG - Weight G


                        //Items
                        var excelItems = ObjWorkSheet.get_Range("J" + colnum, Type.Missing).Value2;
                        string Items = excelItems.ToString();


                        //Length - 10
                        var Length = "10";
                        //Widht - 10
                        var Widht = "10";
                        //Height - 10
                        var Height = "10";


                        //Экспорт даты забора CollDate - Collection Date
                        Excel.Range excelCollDate = ObjWorkSheet.get_Range("K" + colnum);
                        object CollDate_v = excelCollDate.Value2;
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


                        //CollTime - Collection Time = 1100
                        var CollTime = "1100";

                        //CollTimeTo - Collection Time To = 1400
                        var CollTimeTo = "1400";

                        //Экспорт даты доставки DelDate - Delivery Date From
                        Excel.Range excelDelDate = ObjWorkSheet.get_Range("L" + colnum);
                        object DelDate_v = excelDelDate.Value2;
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

                        //DelTime - Delivery Time = 2359
                        var Deltime = "2359";

                        //DelTimeTo - Delivery Time To = 2359
                        var DeltimeTo = "2359";


                        //Dev - Dev = s
                        var Div = "s";

                        //Prod - Prod
                        var excelProd = ObjWorkSheet.get_Range("N" + colnum, Type.Missing).Value2;
                        string Prod = excelProd.ToString();

                        //Payer - Payer
                        var excelPayer = ObjWorkSheet.get_Range("O" + colnum, Type.Missing).Value2;
                        string Payer = excelPayer.ToString();

                        //Stack - Stackable = y
                        var Stack = "y";


                        //CMair
                        var CMair = "4";

                            //CMairVendor - Vendor
                            var excel_CMairVendor = ObjWorkSheet.get_Range("Q" + colnum, Type.Missing).Value2;
                            string CMairVendor = excel_CMairVendor.ToString();

                            //CMairQt - Quote Amount
                            var excel_CMairQt = ObjWorkSheet.get_Range("R" + colnum, Type.Missing).Value2;
                            string CMairQt = excel_CMairQt.ToString();


                        //LCLPU
                        var LCLPU = "24";

                            //LCLPU_Vendor - Vendor
                            var excel_LCLPU_Vendor = ObjWorkSheet.get_Range("T" + colnum, Type.Missing).Value2;
                            string LCLPU_Vendor = excel_LCLPU_Vendor.ToString();

                            //LCLPU_Qt - Quote Amount
                            var excel_LCLPU_Qt = ObjWorkSheet.get_Range("U" + colnum, Type.Missing).Value2;
                            string LCLPU_Qt = excel_LCLPU_Qt.ToString();


                        //Insur - Insurance         ???????????????               
                        //LocDel - Local Deilvery   ???????????????
                        //Handling                  ???????????????
                        
                        //Revao
                        var excelRevao = ObjWorkSheet.get_Range("W" + colnum, Type.Missing).Value2;
                        string Revao;
                        if (excelRevao == null)// Если ячейки пустые, то ничего не вводить                 
                        {
                            Revao = null;
                        }
                        else
                        {
                            Revao = excelRevao.ToString();
                        }                    
                        var Revao_n = "43";

                        //Disc
                        var excelDisc = ObjWorkSheet.get_Range("X" + colnum, Type.Missing).Value2;
                        string Disc = excelDisc.ToString();
                        var Disc_n = "7";


                        ForAwait(20, 1, "Customer Service System");
                        Thread.Sleep(600);                        
                        host.Send(CName);// 0                        
                        Thread.Sleep(100); //костыль
                        host.Send("<TAB>");

                        host.Send(TelNo);// 0 
                        Thread.Sleep(100); //костыль
                        host.Send("<TAB>");
                        host.Send(TelNo);// 0 
                        Thread.Sleep(100); //костыль
                        host.Send("<TAB>");

                        ForAwaitCol(19);
                        Thread.Sleep(600);
                        host.Send(SrCrit);
                        Thread.Sleep(500);
                        host.Send("<ENTER>");
                        Thread.Sleep(1000);
                        host.Send("<F12>");

                        ForAwaitCol(13);
                        Thread.Sleep(600);
                        host.Send(Select);// =41                        
                        host.Send("<ENTER>");


                        ForAwait(31, 8, "SS Status");
                        Thread.Sleep(600);
                        host.Send(SSstat);// = BK                        
                        Thread.Sleep(100);

                        ForAwaitCol(62);
                        Thread.Sleep(600);
                        host.Send(Con);// CN Number                        
                        host.Send("<TAB>");
                        ForAwaitCol(7);
                        host.Send("<TAB>");

                        ForAwaitCol(46);
                        host.Send("<F4>");

                        ForAwaitRow(9);
                        Thread.Sleep(600);
                        host.Send(RecName); // Receaver Name
                        Thread.Sleep(600);
                        host.Send("<TAB>");

                        ForAwaitRow(10);
                        host.Send(RecAddr); // Receaver Address
                        Thread.Sleep(600);
                        host.Send("<TAB>");
                        ForAwaitCol(46);
                        host.Send("<TAB>");
                        ForAwaitRow(11);                        
                        host.Send("<TAB>");

                        ForAwaitRow(12);
                        host.Send(RecTown); // Receaver Town
                        Thread.Sleep(600);
                        host.Send("<TAB>");

                        ForAwaitCol(57);
                        host.Send(RecPost); // Receaver Postcode
                        host.Send("<ENTER>");
                        Thread.Sleep(1000);
                        host.Send("<TAB>");

                        ForAwaitCol(7);
                        host.Send("<TAB>");

                        ForAwaitCol(46);
                        host.Send("<TAB>");

                        ForAwaitCol(13);
                        host.Send("<TAB>");

                        ForAwaitCol(50);
                        host.Send(GdsDesk); // GDS Desc 
                        host.Send("<TAB>");

                        ForAwaitCol(14);
                        host.Send(WeightKG); // Weight Только КГ!!!!
                        host.Send("<TAB>");
                        ForAwaitCol(28);
                        //host.Send(WeightG);
                        host.Send("<TAB>");

                        ForAwaitCol(55);
                        host.Send("<TAB>");

                        ForAwaitCol(73);
                        host.Send("<TAB>");

                        ForAwaitCol(8);
                        host.Send("<TAB>");

                        ForAwaitCol(20);
                        host.Send(Items); // Items
                        host.Send("<TAB>");

                        ForAwaitCol(37);
                        host.Send(Length); // =10 
                        host.Send("<TAB>");

                        ForAwaitCol(53);
                        host.Send(Widht); // =10 
                        host.Send("<TAB>");

                        ForAwaitCol(70);
                        host.Send(Height); // =10  
                        host.Send("<TAB>");

                        ForAwaitCol(20);
                        host.Send(CollDate); // Collection Date
                        Thread.Sleep(600);
                        host.Send("<TAB>");

                        ForAwaitRow(16);
                        host.Send(CollTime); // =1100  
                        Thread.Sleep(600);
                        ForAwaitCol(31);
                        host.Send(CollTimeTo); // =1400
                        host.Send("<TAB>");

                        //ForAwaitCol(40);
                        //host.Send("<TAB>");

                        ForAwaitCol(51);
                        host.Send("<TAB>");

                        ForAwaitCol(22);
                        host.Send(DelDate); // Delivery Date
                        Thread.Sleep(600);
                        host.Send("<TAB>");

                        ForAwaitCol(40);
                        host.Send(Deltime); // =2359  
                        Thread.Sleep(600);

                        ForAwaitCol(56);
                        host.Send(DelDate); // Delivery Date
                        Thread.Sleep(600);
                        host.Send("<TAB>");

                        ForAwaitCol(74);
                        host.Send(DeltimeTo); // =2359  
                        Thread.Sleep(600);

                        ForAwaitCol(7);
                        host.Send(Div); // =s                        
                        host.Send("<TAB>");

                        ForAwaitCol(20);
                        host.Send(Prod); // Prod
                        Thread.Sleep(600);
                        host.Send("<TAB>");

                        ForAwaitCol(37);
                        host.Send("<TAB>");
                        ForAwaitCol(45);
                        host.Send("<TAB>");
                        ForAwaitCol(53);
                        host.Send("<TAB>");
                        ForAwaitCol(61);
                        host.Send("<TAB>");

                        ForAwaitCol(76);
                        host.Send(Payer); // Payer
                        Thread.Sleep(600);
                        host.Send("<TAB>");
                    
                    
                        ForAwaitCol(38);
                        host.Send("<TAB>");
                        ForAwaitCol(76);
                        host.Send("<TAB>");
                        ForAwaitCol(17);
                        host.Send("<TAB>");
                        ForAwaitCol(47);
                        host.Send("<TAB>");
                        
                        ForAwaitCol(13);
                        host.Send(Stack); // =y

                        //Вход в Tariff No
                        ForAwaitCol(48);
                        host.Send("<F10>");

                        ForAwaitCol(15);
                        host.Send("<F5>");

                        ForAwaitCol(7);
                        host.Send("<F4>");

                        if (CMairVendor == null || CMairQt == null)// Если ячейки пустые, то ничего не вводить                 
                        {
                            
                        }
                        else
                        {
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

                        if (LCLPU_Vendor == null || LCLPU_Qt == null)// Если ячейки пустые, то ничего не вводить                 
                        {

                        }
                        else
                        {
                            ForAwaitCol(7);
                            host.Send("<F4>");

                            ForAwaitCol(23);
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


                        if (Revao == null)// Если ячейки пустые, то ничего не вводить                 
                        {

                        }
                        else
                        {
                            ForAwaitCol(7);
                            host.Send("<F4>");

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

                        }
                        else
                        {
                            ForAwaitCol(7);
                            host.Send("<F4>");

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


                        ForAwaitCol(7);
                        host.Send("<F12>");
                        ForAwaitCol(15);
                        host.Send("<F12>");
                        Thread.Sleep(600);
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
                        Thread.Sleep(600);
                        host.Send("<ENTER>");


                    //    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    //    // Status
                    //    var excelstatus = ObjWorkSheet.get_Range("R" + colnum, Type.Missing).Value2;
                    //    string status = excelstatus.ToString();                                  
                    
                    //    // Экспорт даты доставки dateD                
                    //    Excel.Range exceldateD = ObjWorkSheet.get_Range("Q" + colnum);
                    //    object dateD_v = exceldateD.Value2;
                                           
                    //    // Экспорт даты забора dateZ                
                    //    Excel.Range exceldateZ = ObjWorkSheet.get_Range("P" + colnum);
                    //    object dateZ_v = exceldateZ.Value2;
                    
                    //    //Time - Можно по умолчанию вводить "1000"
                    //    var time = "1000";

                    //    //Depo - EVENTDEPOT                     
                    //    var eventdepot = "MOW";

                    //    // Consigment номер накладной
                    //    var excelcon = ObjWorkSheet.get_Range("N" + colnum, Type.Missing).Value2;
                    //    string con_wocheck = excelcon.ToString();
                    //    string con = con_wocheck.Substring(0, 9);// ограничивание накладной по количеству знаков

                    //    // количество обрабатываемых накладных за раз (по умолчанию 1)
                    //    var qty = "1";

                    //    // Delv zone - по умолчанию "b"
                    //    var delvz = "B";

                    //    ForAwait(15, 2, "Consignment Status Entry");
                    //    Thread.Sleep(600);                        
                    //    host.Send(status);//Вводим статус
                    //    logger.Debug(status, this.Text); //LOG
                    //    Thread.Sleep(600); //костыль
                    //    if (disp.CursorCol != 28 && disp.CursorCol != 10)
                    //        host.Send("<TAB>");

                    //    ForAwaitCol(28);//Вводим дату доставки, если ОК или дату забора, если OF
                    //    if (status == "OK")
                    //    {
                    //        if (dateD_v == null)// Если даты нет, то переходим к вводу статуса и след. строке
                    //        {
                    //            host.Send("<F12>");
                    //            continue; //переход к следующей итерации FOR
                    //        }

                    //        DateTime dDt = DateTime.FromOADate((double)dateD_v);
                            
                    //        string dateD = dDt.ToString("ddMMMyy", CultureInfo.GetCultureInfo("en-us"));

                    //        if (dateD_v is double)
                    //        {
                    //            dDt = DateTime.FromOADate((double)dateD_v);
                    //        }
                    //        else
                    //        {
                    //            DateTime.TryParse((string)dateD_v, out dDt);
                    //        }

                    //        host.Send(dateD);
                    //        Thread.Sleep(100);
                    //        logger.Debug(dateD, this.Text);  //LOG                                                        
                    //        host.Send("<TAB>");
                    //    }
                    //    else if(status == "OF")
                    //    {
                    //        if (dateZ_v == null)// Если даты нет, то переходим к вводу статуса и след. строке                    
                    //        {
                    //            host.Send("<F12>");
                    //            continue; //переход к следующей итерации FOR
                    //        }

                    //        DateTime dZt = DateTime.FromOADate((double)dateZ_v);

                    //        string dateZ = dZt.ToString("ddMMMyy", CultureInfo.GetCultureInfo("en-us"));

                    //        if (dateZ_v is double)
                    //        {
                    //            dZt = DateTime.FromOADate((double)dateZ_v);
                    //        }
                    //        else
                    //        {
                    //            DateTime.TryParse((string)dateZ_v, out dZt);
                    //        }

                    //        host.Send(dateZ);
                    //        Thread.Sleep(100);
                    //        logger.Debug(dateZ, this.Text);  //LOG                                                       
                    //        host.Send("<TAB>");
                    //    }

                    //    ForAwaitCol(46);//Вводим время
                    //    host.Send(time);
                    //    Thread.Sleep(100);
                    //    if (disp.CursorCol != 70 && disp.CursorCol != 46) host.Send("<TAB>");
                                                
                    //    ForAwaitCol(70);//Вводим депо
                    //    Thread.Sleep(3000);
                    //    host.Send(eventdepot);
                    //    Thread.Sleep(3000);
                    //    host.Send("<TAB>");
                    //    Thread.Sleep(100);

                    //    ForAwaitCol(13);// Signatory - пропускаем
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(57);// REV Date - пропускаем                    
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(77);//Rems + Если статус OF, то делаем и вводим коммент = статусу OF
                    //    if (status == "OF")
                    //    {
                    //        host.Send("<F4>");
                    //        ForAwait(5, 5, "Seq Remarks");                            
                    //        host.Send(status);
                    //        Thread.Sleep(500);
                    //        host.Send("<ENTER>");

                    //        ForAwaitCol(9); // вторая строка seq remarks
                    //        host.Send("<F12>");

                    //        ForAwaitCol(18);// mode: add - пропускаем
                    //        host.Send("<F12>");//возвращаемся в общее меню на позицию REMS+ COL(77)
                    //        ForAwait(15, 2, "Consignment Status Entry");// проверяем                    
                    //    }                               
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(12);//Runsheet - пропускаем
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(33);//Round no - пропускаем
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(54);// Delv zone -  по умолчанию "b"
                    //    host.Send(delvz); 
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(73);// Delv area - пропускаем
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(24);//No of status Entries = 1
                    //    host.Send(qty);
                    //    host.Send("<ENTER>");
                    //    ForAwait(1, 10, "01");

                    //    host.Send(con);  // Con number        
                    //    logger.Debug(con, this.Text); //LOG
                    //    ForAwaitCol(26);//Позиция после ввода 9 символов номера накладной    
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(37);// Статус (повторный вывод) - пропускаем
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(48);// Time - пропускаем
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(58);// Solved - пропускаем
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(64);// Rev date (повторный вывод) - пропускаем
                    //    host.Send("<TAB>");

                    //    ForAwaitCol(17); // Signatory Если статус OK = OK, если OF = ""
                    //    if (status == "OK")
                    //    {
                    //        host.Send(status);
                    //    }
                    //    else
                    //    {
                    //        host.Send("");
                    //    }

                    //    host.Send("<ENTER>");//концовка и переход обратно к вводу статуса
                    //    host.Send("<F12>");
                    //    host.Send("<ENTER>");
                    //    Thread.Sleep(2500);

                                                                   

                    //    ForAwait(15, 2, "Consignment Status Entry");
                    //    // DBContext.ChangeRecordStatus(id); 

                    //     // Запись в ячейку даты внесения статуса отметки DONE
                    //    //ObjWorkSheet.Cells[18, i] = done;
                    //    //ObjWorkExcel.Interactive = false;
                    //    //ObjWorkBook.Save();
                    //    //ObjWorkExcel.Interactive = true;
                    //    logger.Debug(done, this.Text);  //LOG
                    ////}
                    ////else
                    ////{
                    ////    continue; //переход к следующей итерации FOR
                    ////} 
                    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                
                }


                // Закрываем TeemTalk
                teemApp.Close();
                foreach (Process proc in Process.GetProcessesByName("teem2k"))
                {
                    proc.Kill();
                }
                //teemApp.Application.Close();
                Thread.Sleep(1000);
                //host.Send("<ENTER>");



                //Закрываем Excel
                ObjWorkBook.Close(true);
                ObjWorkExcel.Quit();
                foreach (Process proc in Process.GetProcessesByName("excel"))
                {
                    proc.Kill();
                }

                MessageBox.Show("Данные внесены", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;

            }
            catch (Exception ex)
            {
                // Вывод сообщения об ошибке
                logger.Debug(ex.ToString());
            }




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
