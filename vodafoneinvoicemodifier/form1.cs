using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Threading.Tasks;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace VodafoneInvoiceModifier
{
    public partial class Form1 : Form
    {
        private System.Diagnostics.FileVersionInfo myFileVersionInfo;
        private ContextMenu contextMenu1;
        private string myRegKey = @"SOFTWARE\RYIK\VodafoneInvoiceModifier";
        string pathToIni = Application.StartupPath + @"\VodafoneInvoiceModifier.ini"; //path to ini of tools

        private string pStop = @"ЗАГАЛОМ ЗА ВСІМА КОНТРАКТАМИ";
        private string pDiscount = @"ЗНИЖКА**";
        private string pFull = @"ЩОМІСЯЧНА ВАРТІСТЬ ПАКЕТА ТА ПОСЛУГИ, НАДАНІ ЗА МЕЖАМИ ПАКЕТА";

        private string about = "";
        private string dataStart = ""; // дата начала периода счета
        private string dataEnd = "";  // дата конца периода счета
        private string periodInvoice = ""; //Период
        private bool checkRahunok = false;
        private bool checkNomerRahunku = false;
        private bool checkPeriod = false;

        private string sConnection = ""; //string connection to MS SQL DB
        private string sConnectionServer = ""; //string connection to MS SQL DB
        private string sConnectionUserName = ""; //string connection to MS SQL DB
        private string sConnectionUserPasswords = ""; //string connection to MS SQL DB

        private string[] pListParseStrings = new string[]
        {
            // со счета
            @"Владелец",                                        //0
            @"Контракт №",                                      //1
            @"Моб.номер",                                       //2
            @"Ціновий Пакет",                                   //3
            @"ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА",                 //4
            @"ПОСЛУГИ МІЖНАРОДНОГО РОУМІНГУ",                   //5
            @"ЗНИЖКИ",                                          //6
            @"ЗАГАЛОМ ЗА КОНТРАКТОМ (БЕЗ ПДВ ТА ПФ)",           //7
            @"ПДВ",                                             //8
            @"ПФ",                                              //9
            @"Загалом з податками",                             //10
            @"GPRS/CDMA з'єд.  Роумінг",                        //11
            @"Передача даних - вартість пакету послуг",         //12
            @"Вихідні дзвінки  Міські номери",                  //13
            @"ПОСЛУГИ, НАДАНІ ЗА МЕЖАМИ ПАКЕТА",                //14
            @"НАДАНІ КОНТЕНТ-ПОСЛУГИ",                          //15
            @"Дата счета",                                      //16
            @"Дата кінця періоду",                              //17
            // из базы
            @"Таб. номер",                                      //18
            @"Отдел",                                           //19
            @"Действует c",                                     //20
            @"Модель",                                          //21
            @"Оплата владельцем",                               //22
            // со счета
            @"ПОСЛУГИ ЗА МЕЖАМИ ПАКЕТА",                        //23
            // анализ
            @"Контракт использовался",                          //24
            @"Контракт не заблокирован",                        //25            
            // доп.признаки строк
            "Вх",                                           //26
            "Вих",                                         //27
            "Переадр",                                         //28
            "GPRS",                                        //29
            "CDMA"                                        //30
        };
        private string[] pTranslate = new string[]
        {
            // со счета
            @"ФИО сотрудника",
            @"Контракт",
            @"Номер телефона абонента",
            @"Ціновий Пакет",
            @"ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА",
            @"Общая сумма в роуминге, грн",
            @"Скидка",
            @"Затраты по номеру, грн",
            @"НДС, грн",
            @"ПФ, грн",
            @"Итого по контракту, грн",
            @"Интернет в роуминге",
            @"Интернет за пределами пакета",
            @"Звонки на городские номера",
            @"ПОСЛУГИ, НАДАНІ ЗА МЕЖАМИ ПАКЕТА",
            @"КОНТЕНТ-ПОСЛУГИ",
            @"Дата счета",
            @"Дата окончания периода",
            // из базы
            @"Табельный номер",
            @"Подразделение",
            @"Действует c",
            @"ТАРИФНАЯ МОДЕЛЬ",
            @"К оплате владельцем номера, грн",
            // со счета
            @"ЗАМОВЛЕНІ ДОДАТКОВІ ПОСЛУГИ ЗА МЕЖАМИ ПАКЕТА",
            // анализ
            @"Контракт использовался",
            @"Контракт не заблокирован",
            @"Вхідні",     //26
            @"Вихідні",     //27
            @"Переадр",     //28
            @"GPRS",     //29
            @"CDMA"     //29
        };
        private string[] pToAccount = new string[]
        {
            // для бухгалтерии
            @"Дата счета",
            @"Номер телефона абонента",
            @"ФИО сотрудника",
            @"Затраты по номеру, грн",
            @"НДС, грн",
            @"ПФ, грн",
            @"Итого по контракту, грн",
            @"Общая сумма в роуминге, грн",
            @"Подразделение",
            @"Табельный номер",
            @"ТАРИФНАЯ МОДЕЛЬ",
            @"К оплате владельцем номера, грн",
            @"Контракт использовался",   //Test
            @"Контракт не заблокирован"  //Test
        };
        private int contractNumberFound = 0;
        private StringBuilder sbError = new StringBuilder();
        private DataTable dtMobile = new DataTable("MobileData");

        private DataColumn[] dcMobile ={
                                  // со счета
                                  new DataColumn("ФИО сотрудника",typeof(string)),
                                  new DataColumn("Контракт",typeof(string)),
                                  new DataColumn("Номер телефона абонента",typeof(string)),
                                  new DataColumn("Ціновий Пакет",typeof(string)),
                                  new DataColumn("ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА",typeof(double)),
                                  new DataColumn("Общая сумма в роуминге, грн",typeof(double)),
                                  new DataColumn("ЗНИЖКИ",typeof(double)),
                                  new DataColumn("Затраты по номеру, грн",typeof(double)),
                                  new DataColumn("НДС, грн",typeof(double)),
                                  new DataColumn("ПФ, грн",typeof(double)),
                                  new DataColumn("Итого по контракту, грн",typeof(double)),
                                  new DataColumn("GPRS/CDMA з'єд.  Роумінг",typeof(double)),
                                  new DataColumn("Передача даних - вартість пакету послуг",typeof(double)),
                                  new DataColumn("Вихідні дзвінки  Міські номери",typeof(double)),
                                  new DataColumn("ПОСЛУГИ, НАДАНІ ЗА МЕЖАМИ ПАКЕТА",typeof(double)),
                                  new DataColumn("КОНТЕНТ-ПОСЛУГИ",typeof(double)),
                                  new DataColumn("Дата счета",typeof(string)),
                                  new DataColumn("Дата кінця періоду",typeof(string)),
                                  // из базы
                                  new DataColumn("Табельный номер",typeof(string)),
                                  new DataColumn("Подразделение",typeof(string)),
                                  new DataColumn("Действует c",typeof(string)),
                                  new DataColumn("ТАРИФНАЯ МОДЕЛЬ",typeof(string)),
                                  new DataColumn("К оплате владельцем номера, грн",typeof(double)),
                                  // со счета
                                  new DataColumn("ЗАМОВЛЕНІ ДОДАТКОВІ ПОСЛУГИ ЗА МЕЖАМИ ПАКЕТА",typeof(double)),

                                  // анализ состояния контракта
                                  new DataColumn("NumberUsed",typeof(bool)),
                                  new DataColumn("NumberNoBlock",typeof(bool))
                              };

        private List<string> listTempContract = new List<string>();

        // из базы
        //  private List<string> lTarifData = new List<string>();
        private DataTable dtTarif = new DataTable("TarifListData");
        private DataColumn[] dcTarif ={
                                  new DataColumn("Номер телефона",typeof(string)),
                                  new DataColumn("ФИО",typeof(string)),
                                  new DataColumn("NAV",typeof(string)),
                                  new DataColumn("Подразделение",typeof(string)),
                                  new DataColumn("Основной",typeof(string)),
                                  new DataColumn("Действует c",typeof(string)),
                                  new DataColumn("Модель компенсации",typeof(string)),
                                  new DataColumn("Тарифный пакет",typeof(string))
                              };
        //search in DataTable:
        /*private static void ShowTable(DataTable table) {
          foreach (DataColumn col in table.Columns) {
            Console.Write("{0,-14}", col.ColumnName);
            }
           Console.WriteLine();

         foreach (DataRow row in table.Rows) {
           foreach (DataColumn col in table.Columns) {
            if (col.DataType.Equals(typeof(DateTime)))
               Console.Write("{0,-14:d}", row[col]);
            else if (col.DataType.Equals(typeof(Decimal)))
               Console.Write("{0,-14:C}", row[col]);
            else
               Console.Write("{0,-14}", row[col]);           
           }
           Console.WriteLine();
         }
         Console.WriteLine();
         }*/

        private List<string> lTarif = new List<string>();  //тарифные модели компенсаций
        private HashSet<string> listTarifData = new HashSet<string>(); //will write models in modelToPayment()

        private string[] arrayTarif = new string[] {
            @"L100% корпорация",                 //0
            @"L100% сотрудник",                  //1
            @"L100%,R80%",                       //2
            @"L50,R0%",                          //3
            @"L80,R0%",                          //4
            @"L100,R0%",                         //5
            @"L160,R0%",                         //6
            @"L250,R0%",                         //7
            @"L50%,R0%",                         //8
            @"L50%,R80%",                        //9
            @"L50%,R100%",                       //10
            @"L90%,R100%",                       //11
            @"Lpack100%,R0%,Paid0%",             //12
            @"Lмоб200,R0%,Paid0%",               //13
            @"L200,R0%"                          //14
        };

        private string infoStatusBar = "";
        private bool newModels = false; //stop calculating data
        private string strNewModels = "";

        private string filePathTxt; //path to the selected bill

        private List<string> listNumbers = new List<string>(); //list of numbers for the marketing report
        private List<string> listServices = new List<string>();//list of services for the marketing report

        private string parametrStart = "Контракт";
        private string parametrEnd = "ЗАГАЛОМ ЗА ВСІМА КОНТРАКТАМИ";

        double discountDouble = 1;
        double beforeDiscountDouble = 1;
         string discount = null;
        string beforeDiscount = null;
        private double discountValue = 0.70; //скидка в текущем счете (1-discountValue/100) - 30% 

        private bool loadedBill = false;
        private bool selectedServices = false;
        private bool selectedNumbers = false;
        private DataTable dtFullBill = new DataTable("FullListDataBill");
        private DataColumn[] dcFullBill ={
                                  new DataColumn("Контракт",typeof(string)),
                                  new DataColumn("Номер телефона",typeof(string)),
                                  new DataColumn("ФИО",typeof(string)),
                                  new DataColumn("NAV",typeof(string)),
                                  new DataColumn("Подразделение",typeof(string)),
                                  new DataColumn("Имя сервиса",typeof(string)),
                                  new DataColumn("Номер В",typeof(string)),
                                  new DataColumn("Дата",typeof(string)),
                                  new DataColumn("Время",typeof(string)),
                                  new DataColumn("Длительность А",typeof(string)),
                                  new DataColumn("Длительность В",typeof(string)),
                                  new DataColumn("Стоимость",typeof(string))
                              };
        private DataTable dtMarket = new DataTable("MarketReport");

        private List<string> listSavedServices = new List<string>();
        private List<string> listSavedNumbers = new List<string>();
        private string filepathLoadedData = "";  //current path to invoice
        private string strSavedPathToInvoice = "";  //previous session path to invoice
        private bool foundSavedData = false;


        public Form1()
        { InitializeComponent(); }

        private void Form1_Load(object sender, EventArgs e)
        {
            myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);
            about = myFileVersionInfo.Comments + " ver." + myFileVersionInfo.FileVersion + " " + myFileVersionInfo.LegalCopyright;
            StatusLabel1.Text = myFileVersionInfo.ProductName + " ver." + myFileVersionInfo.FileVersion + " " + myFileVersionInfo.LegalCopyright;
            StatusLabel1.Alignment = ToolStripItemAlignment.Right;

            notifyIcon1.Text = myFileVersionInfo.ProductName + " " + myFileVersionInfo.LegalCopyright;
            notifyIcon1.BalloonTipText = about;

            contextMenu1 = new ContextMenu();  //Context Menu on notify Icon
            notifyIcon1.ContextMenu = contextMenu1;
            contextMenu1.MenuItems.Add("About", AboutSoft);
            contextMenu1.MenuItems.Add("Exit", ApplicationExit);
            notifyIcon1.Text = myFileVersionInfo.ProductName + "\nv." + myFileVersionInfo.FileVersion + "\n" + myFileVersionInfo.CompanyName;
            this.Text = myFileVersionInfo.Comments;

            groupBox1.BackColor = System.Drawing.Color.Ivory;

            labelAccount.Visible = false;
            labelPeriod.Visible = false;
            labelBill.Visible = false;
            labelContracts.Visible = false;
            readinifile();

            makeReportAccountantItem.Enabled = false;
            makeFullReportItem.Enabled = false;
            makeReportMarketingItem.Enabled = false;
            prepareBillItem.Enabled = false;


            openBillItem.ToolTipText = "Открыть счет Voodafon в текстовом формате.\nMax количество строк - 500 000";
            makeFullReportItem.ToolTipText = "Подготовить полный отчет в Excel-файле.\nФайл будет сохранен в папке с программой";
            makeReportAccountantItem.ToolTipText = "Подготовить отчет для бух. в Excel-файле.\nФайл будет сохранен в папке с программой";
            useSavedDataItem.ToolTipText = "Использовать сохраненный список файлов и сервисов из предыдущей сессии";
            labelDiscount.Text = "";
            clearTextboxItem.ToolTipText = "Убрать весь текст из окна просмотра";
            aboutItem.ToolTipText = "О программе";
            exitItem.ToolTipText = "Выйти из программы и сохранить настройки и парсеры счета";

            /*buttonReport2.FlatAppearance.MouseOverBackColor = System.Drawing.Color.PaleGreen;
            buttonExit.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SandyBrown;
            */
            dtMobile.Columns.AddRange(dcMobile);
            dtTarif.Columns.AddRange(dcTarif);
            dtFullBill.Columns.AddRange(dcFullBill);
            dtMarket = dtFullBill.Clone();
            ListsRegistryDataCheck();
            useSavedDataItem.Enabled = foundSavedData;

            ProgressBar1.Value = 0;
        }


        private void AboutSoft()
        {
            string strVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

            MessageBox.Show(
                myFileVersionInfo.Comments + "\n\nВерсия: " + myFileVersionInfo.FileVersion + "\nBuild: " +
                strVersion + "\n" + myFileVersionInfo.LegalCopyright,
                "Информация о программе",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1);
        }

        private void ApplicationExit()
        {
            writeinitofile();
            Application.Exit();
        }

        private void openBillItem_Click(object sender, EventArgs e)//Menu "Open"
        { OpenBill(); }

        private void makeFullReportItem_Click(object sender, EventArgs e)
        { MakeExcelReport(ExportFullDataTableToExcel); }

        private void makeReportAccountantToolItem_Click(object sender, EventArgs e)
        { MakeExcelReport(ExportDataTableToExcelForAccount); }

        private void clearTextBoxItem_Click(object sender, EventArgs e)
        { textBoxLog.Clear(); }

        private void AboutSoft(object sender, EventArgs e)
        { AboutSoft(); }

        private void ApplicationExit(object sender, EventArgs e)
        { ApplicationExit(); }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        { ApplicationExit(); }

        private void selectListNumbersItem_Click(object sender, EventArgs e)
        { selectListNumbers(); }

        //limit of numbers <500
        private void selectListNumbers() //Prepare list of numbers for the marketing report - listNumbers
        {
            selectedNumbers = false;
            makeReportMarketingItem.Enabled = false;
            string strTemp = "";
            List<string> listWrongString = new List<string>();
            List<string> tempListString = LoadDataIntoList();
            int limitWrongNumber = 300;

            //clear target 
            listNumbers.Clear();
            textBoxLog.Clear();

            if (tempListString.Count > 0)
            {
                foreach (string s in tempListString)
                {
                    strTemp = MakeCommonViewPhone(s);
                    if (strTemp.Length == 13)  //check Length of a number == 13
                    { listNumbers.Add(strTemp); }
                    else
                    { listWrongString.Add(strTemp); }
                }

                if (0 < listWrongString.Count)
                {
                    textBoxLog.AppendText("List of first 300 wrong rows in the selected list:\n");
                    textBoxLog.AppendText("--------------------------------------------\n\n");
                    int wrongRow = 0;
                    foreach (string s in listWrongString)
                    {
                        textBoxLog.AppendText(s + "\n"); wrongRow++;
                        if (wrongRow > limitWrongNumber) { break; }

                    }
                    textBoxLog.AppendText("\n\n");
                }

                if (0 < listNumbers.Count)
                {
                    if (500 < listNumbers.Count)
                    {
                        textBoxLog.AppendText("The selected list contains so many Numbers!\nWill check the selected file!\n");
                        textBoxLog.AppendText("--------------------------------------------\n\n");
                    }
                    else
                    {
                        selectedNumbers = true;
                        ListNumbersRegistrySave();
                        _ControlSetItsText(labelContracts, listNumbers.Count.ToString() + " шт.");
                        _ControlVisibleEnabled(labelContracts, true);
                    }

                    textBoxLog.AppendText("List of numbers:\n");
                    textBoxLog.AppendText("--------------------------------------------\n\n");
                    foreach (string s in listNumbers)
                    { textBoxLog.AppendText(s + "\n"); }
                }
                else { textBoxLog.AppendText("Check the list of numbers\n"); }


                //clean temporary elements
                strTemp = null;
                listWrongString = null;
                tempListString = null;
            }
            CheckConditionEnableMarketingReport();
        }

        private void selectListServicesItem_Click(object sender, EventArgs e)
        { selectListServices(); }

        //limit of services <100
        private void selectListServices() //Prepare list of services for the marketing report - listServices
        {
            selectedServices = false;
            makeReportMarketingItem.Enabled = false;
            textBoxLog.Clear();

            listServices.Clear();
            listServices = LoadDataIntoList();

            if (0 < listServices.Count && listServices.Count < 100)
            {
                textBoxLog.AppendText("\nList of services:\n");
                foreach (string s in listServices)
                { textBoxLog.AppendText(s + "\n"); }
                selectedServices = true;

                ListServicesRegistrySave();
            }
            else
            {
                textBoxLog.AppendText("\n--------------------------------------------\n");
                textBoxLog.AppendText("The selected list is wrong!\nWill check the file!\nIt has to contain from 1 to 100 services.");
                textBoxLog.AppendText("\n--------------------------------------------\n");
            }
            CheckConditionEnableMarketingReport();
        }

        private async void prepareBillItem_Click(object sender, EventArgs e)
        { await Task.Run(() => LoadBillIntoMemory()); }

        private void LoadBillIntoMemory()
        {
            _ProgressBar1Start();
            _TextboxClear(textBoxLog);
            _ToolStripMenuItemEnabled(fileMenuItem, false);
            _ControlVisibleEnabled(labelPeriod, true);

            loadedBill = false;

            GetDataWithModel();
            string kontrakt = "";
            string numberMobile = "";
            string fio = "";
            string nav = "";
            string department = "";
            string tempRow = "";
            string serviceName = "";
            string numberB = "";
            string date = "";
            string time = "";
            string durationA = "";
            string durationB = "";
            string cost = "";

            pListParseStrings[1] = _ControlReturnItsText(textBoxP1);
            pListParseStrings[2] = _ControlReturnItsText(textBoxP2);

            List<string> filterBill = new List<string>();
            filterBill.Add(pListParseStrings[1]);
            filterBill.Add(pListParseStrings[2]);

            if (listServices.Count == 0)
            { listServices = listSavedServices; }
            if (listNumbers.Count == 0)
            { listNumbers = listSavedNumbers; }

            _ProgressWork1Step();

            foreach (string service in listServices)
            { filterBill.Add(service); }

            _ProgressWork1Step();

            List<string> loadedBillWithServicesFilter = LoadDataUsingParameters(filterBill, parametrStart, parametrEnd);
            filterBill = null;

            int allRow = loadedBillWithServicesFilter.Count * listServices.Count * (dtTarif.Rows.Count + listNumbers.Count); //всего строк для борабоки

            int counterstep = 10000;
            int countStepProgressBar = counterstep;
            int counted = 0;
            if (loadedBillWithServicesFilter.Count > 0)
            {
                dtFullBill.Rows.Clear();
                StringBuilder sb = new StringBuilder();

                //todo parsing strings of the filtered bill
                foreach (string sRowBill in loadedBillWithServicesFilter)
                {
                    countStepProgressBar--;
                    if (countStepProgressBar == 0)
                    {
                        _ProgressWork1Step();
                        countStepProgressBar = counterstep;
                    }

                    if (sRowBill.StartsWith(pListParseStrings[1]))
                    {
                        try
                        {
                            kontrakt = Regex.Split(sRowBill.Substring(sRowBill.IndexOf('№') + 1).Trim(), " ")[0].Trim();
                            tempRow = sRowBill.Substring(sRowBill.IndexOf(':') + 1).Trim();

                            if (tempRow.StartsWith("+"))
                            { numberMobile = tempRow; }
                            else { numberMobile = "+" + tempRow; } //set format number like '+380...'

                            tempRow = "";
                        }
                        catch
                        {
                            MessageBox.Show("Проверьте правильность выбора детализации разговоров!\n" +
                        "Возможно поменялся формат.\n" +
                        "Правильный формат:\n" +
                        "Контракт № 000000000  _номер_: 380000000000");
                        }
                    }
                    else
                    {
                        foreach (string service in listServices)
                        {
                            countStepProgressBar--;
                            if (countStepProgressBar == 0)
                            {
                                _ProgressWork1Step();
                                countStepProgressBar = counterstep;
                            }                            //"service" parsing. start position of a cell
                            /*
                            1-39	наименование услуги
                            40-52	номер(целевой)
                            53-63	дата
                            66-74	время
                            75-84	длительность
                            85-95	учтенная длительность оператором (для биллинга)
                            96-106	стоимость
                            */

                            try
                            {
                                serviceName = sRowBill.Substring(0, 38).Trim();
                                numberB = sRowBill.Substring(38, 13).Trim();
                                date = sRowBill.Substring(52, 10).Trim();
                                time = sRowBill.Substring(65, 8).Trim();
                                durationA = sRowBill.Substring(74, 9).Trim();
                                durationB = sRowBill.Substring(84, 9).Trim();
                                cost = sRowBill.Substring(95).Trim();

                                DataRow row = dtFullBill.NewRow(); //Full Table

                                row["Контракт"] = kontrakt;
                                row["Номер телефона"] = numberMobile;
                                row["Имя сервиса"] = serviceName;
                                row["Номер В"] = numberB;
                                row["Дата"] = date;
                                row["Время"] = time;
                                row["Длительность А"] = durationA;
                                row["Длительность В"] = durationB;
                                row["Стоимость"] = cost;

                                DataRow rowMarket = dtMarket.NewRow(); //for Market
                                rowMarket["Контракт"] = kontrakt;
                                rowMarket["Номер телефона"] = numberMobile;
                                rowMarket["Имя сервиса"] = serviceName;
                                rowMarket["Номер В"] = numberB;
                                rowMarket["Дата"] = date;
                                rowMarket["Время"] = time;
                                rowMarket["Длительность А"] = durationA;
                                rowMarket["Длительность В"] = durationB;
                                rowMarket["Стоимость"] = cost;

                                if (!time.Contains('.')) //except a common service with ". . ."
                                {
                                    foreach (DataRow rowTarif in dtTarif.Rows)
                                    {
                                        if (rowTarif["Номер телефона"].ToString().Contains(numberMobile))
                                        {
                                            fio = rowTarif["ФИО"].ToString();
                                            nav = rowTarif["NAV"].ToString();
                                            department = rowTarif["Подразделение"].ToString();

                                            row["ФИО"] = fio;
                                            row["NAV"] = nav;
                                            row["Подразделение"] = department;

                                            rowMarket["ФИО"] = fio;
                                            rowMarket["NAV"] = nav;
                                            rowMarket["Подразделение"] = department;
                                        }

                                        countStepProgressBar--;
                                        if (countStepProgressBar == 0)
                                        {
                                            _ProgressWork1Step("Обработано " + (++counted) + " строк из " + allRow / counterstep);
                                            countStepProgressBar = counterstep;
                                        }
                                    }

                                    tempRow = numberMobile + "\t" + fio + "\t" + nav + "\t" + department + "\t" + serviceName + "\t" + numberB + "\t" + date + "\t" + time + "\t" + durationA + "\t" + durationB + "\t" + cost;
                                    dtFullBill.Rows.Add(row);

                                    foreach (string sNumber in listNumbers)
                                    {
                                        if (tempRow.StartsWith(sNumber))
                                        {
                                            dtMarket.Rows.Add(rowMarket);
                                            sb.AppendLine(tempRow);
                                        }

                                        countStepProgressBar--;
                                        if (countStepProgressBar == 0)
                                        {
                                            _ProgressWork1Step("Обработано " + (++counted) + " строк из " + allRow / counterstep);
                                            countStepProgressBar = counterstep;
                                        }
                                    }
                                }
                                break;
                            }
                            catch (Exception expt) { MessageBox.Show(sRowBill + "\n" + expt.ToString()); }
                        }
                    }
                }
                loadedBill = true;
                File.WriteAllText(Path.GetDirectoryName(filepathLoadedData) + @"\listMarketingCollectRows.csv", sb.ToString(), Encoding.GetEncoding(1251));
                sb = null;
            }
            else
            { _TextboxAppendText(textBoxLog, "Нет в выборке ничего для указанных номеров!\n"); }

            CheckConditionEnableMarketingReport();
            _ToolStripStatusLabelSetItsText(StatusLabel1, "Файл сохранен в папку: " + Path.GetDirectoryName(filepathLoadedData));

            _ToolStripMenuItemEnabled(fileMenuItem, true);
            _ProgressBar1Stop();
        }

        private void makeReportMarketingItem_Click(object sender, EventArgs e)
        { MakeExcelReport(ExportMarketReport); }

        private void ExportMarketReport()
        { ExportDatatableToExcel(dtMarket, "_Marketing.xlsx"); }     //Заполнение таблицы в Excel  данными

        private void CheckConditionEnableMarketingReport() //enableing Marketing report if load data is correct
        {
            if (selectedServices && selectedNumbers && loadedBill)
            {
                _ToolStripMenuItemEnabled(prepareBillItem, true);
                _ToolStripMenuItemEnabled(makeReportMarketingItem, true);
            }
            else if (selectedServices && selectedNumbers)
            {
                _ToolStripMenuItemEnabled(prepareBillItem, true);
            }
        }

        private List<string> LoadDataIntoList() //max received List's length = 500 000 rows
        {
            int listMaxLength = 500000;
            List<string> listValue = new List<string>(listMaxLength);
            string s = "";
            int i = 0; // it is not empty's rows in the selected file

            string filepathLoadedData = _OpenFileDialogReturnPath(openFileDialog1);
            if (filepathLoadedData == null || filepathLoadedData.Length < 1)
            { MessageBox.Show("Не выбран файл."); }
            else
            {
                try
                {
                    var Coder = Encoding.GetEncoding(1251);
                    using (StreamReader Reader = new StreamReader(filepathLoadedData, Coder))
                    {
                        _ToolStripStatusLabelSetItsText(StatusLabel1, "Обрабатываю файл:  " + filepathLoadedData);
                        while ((s = Reader.ReadLine()) != null && i < listMaxLength)
                        {
                            if (s.Trim().Length > 0)
                            {
                                listValue.Add(s.Trim());
                                i++;
                            }
                        }
                    }
                }
                catch (Exception expt) { MessageBox.Show("Ошибка произошла на " + i + " строке:\n\n" + expt.ToString()); }
                if (i > listMaxLength - 10 || i == 0)
                { MessageBox.Show("Error was happened on " + i + " row\n You've been chosen the long file!"); }
            }
            return listValue;
        }

        private List<string> LoadDataUsingParameters(List<string> listParameters, string startStringLoad, string endStringLoad) //max List length = 500 000 rows
        {
            checkRahunok = false;
            checkNomerRahunku = false;
            checkPeriod = false;
            int countParameters = listParameters.Count;
            int countStepProgressBar = 500;
            int listMaxLength = 500000;
            List<string> listRows = new List<string>(listMaxLength);
            string s = "";
            string loadedString = "";
            bool newInvoice = true;
            try
            {
                if (strSavedPathToInvoice.Length > 1)
                {
                    DialogResult result = MessageBox.Show(
                          "Использовать предыдущий выбор файла?\n" + strSavedPathToInvoice,
                          "Внимание!",
                          MessageBoxButtons.YesNo,
                          MessageBoxIcon.Exclamation,
                          MessageBoxDefaultButton.Button1);
                    if (result == DialogResult.Yes)
                    {
                        newInvoice = false;
                    }
                }

                filepathLoadedData = "";
                bool startLoadData = false;
                bool endLoadData = false;
                var Coder = Encoding.GetEncoding(1251);
                if (countParameters > 0)
                {
                    if (newInvoice == true)
                    {
                        filepathLoadedData = _OpenFileDialogReturnPath(openFileDialog1);
                    }
                    else
                    {
                        filepathLoadedData = strSavedPathToInvoice;
                    }

                    if (filepathLoadedData == null || filepathLoadedData.Length < 1)
                    { MessageBox.Show("Did not select File!"); }
                    else
                    {
                        _ToolStripStatusLabelSetItsText(StatusLabel1, "Обрабатываю файл:  " + filepathLoadedData);
                        try
                        {
                            using (StreamReader Reader = new StreamReader(filepathLoadedData, Coder))
                            {
                                while ((s = Reader.ReadLine()) != null && !endLoadData && listRows.Count < listMaxLength)
                                {
                                    loadedString = s.Trim();

                                    //Set label Date
                                    if (s.Contains("Особовий рахунок")) { checkRahunok = true; }
                                    if (s.Contains("Номер рахунку")) { checkNomerRahunku = true; }
                                    if (s.Contains("Розрахунковий період"))
                                    {
                                        string[] substrings = Regex.Split(s, ": ");
                                        periodInvoice = substrings[substrings.Length - 1].Trim();
                                        checkPeriod = true;
                                    }

                                    if (loadedString.StartsWith(startStringLoad))
                                    { startLoadData = true; }
                                    else if (loadedString.StartsWith(endStringLoad))
                                    { endLoadData = true; }

                                    if (startLoadData)
                                    {
                                        foreach (string parameterString in listParameters)
                                        {
                                            if (loadedString.StartsWith(parameterString))
                                            {
                                                listRows.Add(loadedString);
                                                break;
                                            }
                                        }
                                    }
                                    countStepProgressBar--;
                                    if (countStepProgressBar == 0)
                                    {
                                        _ProgressWork1Step();
                                        countStepProgressBar = 500;
                                    }
                                }
                            }

                            if (checkPeriod && checkRahunok && checkNomerRahunku)
                            {
                                _ControlSetItsText(labelPeriod, periodInvoice);
                            }
                            ParameterLastInvoiceRegistrySave();
                        }
                        catch (Exception expt) { MessageBox.Show("Error was happened on " + listRows.Count + " row\n" + expt.ToString()); }
                        if (listMaxLength - 2 < listRows.Count || listRows.Count == 0)
                        { MessageBox.Show("Error was happened on " + (listRows.Count) + " row\n You've been chosen the long file!"); }
                    }
                }
            }
            catch (Exception expt) { MessageBox.Show(expt.ToString()); }
            return listRows;
        }

        private void useSavedDataItem_Click(object sender, EventArgs e)
        {
            if (strSavedPathToInvoice != null && strSavedPathToInvoice.Length > 1)
            { filepathLoadedData = strSavedPathToInvoice; }
            else { strSavedPathToInvoice = ""; }

            if (listSavedNumbers.Count > 0)
            { listNumbers = listSavedNumbers; }

            if (listSavedServices.Count > 0)
            { listServices = listSavedServices; }

            if (listSavedNumbers.Count > 0 && listSavedServices.Count > 0)
            { prepareBillItem.Enabled = true; }
        }

        /*
        public void ControlHoverChangeColorPale(Control control)
        { control.BackColor = System.Drawing.Color.PaleGreen; }

        public void ControlLeaveChangeColorSandy(Control control)
        { control.BackColor = System.Drawing.Color.SandyBrown; }

        public void ControlLeaveChangeColorNormal(Control control)
        { control.BackColor = System.Drawing.SystemColors.Control; }*/

        private async void OpenBill()
        {
            filePathTxt = null;
            sbError = new StringBuilder();
            StatusLabel1.BackColor = System.Drawing.SystemColors.Control;

            textBoxLog.Visible = false;
            newModels = false;
            makeReportAccountantItem.Enabled = false;
            makeFullReportItem.Enabled = false;
            openBillItem.Enabled = false;

            infoStatusBar = "";
            //Чтение параметров парсинга с textbox`es
            pListParseStrings[1] = textBoxP1.Text;
            pListParseStrings[2] = textBoxP2.Text;
            pListParseStrings[3] = textBoxP3.Text;
            pListParseStrings[4] = textBoxP4.Text;
            pListParseStrings[5] = textBoxP5.Text;
            pListParseStrings[6] = textBoxP6.Text;
            pListParseStrings[7] = textBoxP7.Text;
            pStop = textBoxP8.Text;

            StatusLabel1.Text = "Обрабатываю исходные данные...";
            bool billCorrect = ReadTxtAndWiteToMyTmp();

            if (billCorrect)
            {
                StatusLabel1.Text = "Получаю данные с базы Tfactura...";

                await Task.Run(() => GetDataWithModel());
                if (contractNumberFound <= 1)
                {
                    MessageBox.Show("Выбранный счет в базу данных Tfactura еще не импортирован!\nПеред обработкой счета, предварительно необходимо импортировать счет в базу!");
                    StatusLabel1.Text = "Обработка счета прекращена! Предварительно импортируйте счет в Tfactura!";
                    StatusLabel1.BackColor = System.Drawing.Color.SandyBrown;
                }
                else
                {
                    await Task.Run(() => CheckNewTarif());

                    if (!newModels)
                    {
                        MyTmpToMyArray();
                        DataRow[] results;
                        string columnName1 = dtMobile.Columns[0].ColumnName.Remove(3);
                        string columnName2 = dtMobile.Columns[2].ColumnName.Remove(14);
                        string columnName3 = dtMobile.Columns[3].ColumnName;
                        string columnName4 = dtMobile.Columns[10].ColumnName.Remove(6);
                        string columnName5 = dtMobile.Columns[21].ColumnName;
                        string columnName6 = "Роуминг";                     //dtMobile.Columns[5].ColumnName;
                        string columnName10 = dtMobile.Columns[24].ColumnName;
                        string columnName11 = dtMobile.Columns[25].ColumnName;

                        string sortOrder = dtMobile.Columns[0].ColumnName + " ASC";

                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("Дата счета:  " + dtMobile.Rows[1][16].ToString()); //Дата счета
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("====================================================\n");
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("\n");


                        //////////////////////////////
                        textBoxLog.AppendText("---= Список тарифных схем, не существующих в программе =---\n");
                        textBoxLog.AppendText(
                                                    String.Format("{0,-40}", columnName1) +
                                                    String.Format("{0,-15}", columnName2) +
                                                    String.Format("{0,-30}", columnName3) +
                                                    String.Format("{0,-10}", columnName4) +
                                                    String.Format("{0,-30}", columnName5) +
                                                    "\n");

                        foreach (string str in listTarifData)
                        {
                            textBoxLog.AppendText("\"" + str + "\"\n");
                            results = dtMobile.Select("'" + dtMobile.Columns[21].ColumnName.Length + "' LIKE '" + str + "'", sortOrder, DataViewRowState.Added);

                            for (int i = 0; i < results.Length; i++)
                            {

                                textBoxLog.AppendText(
                                 String.Format("{0,-40}", results[i][0].ToString()) +
                                 String.Format("{0,-15}", results[i][2].ToString()) +
                                 String.Format("{0,-30}", results[i][3].ToString()) +
                                 String.Format("{0,-10}", results[i][10].ToString()) +
                                 String.Format("{0,-30}", results[i][21].ToString()) +
                                 "\n"
                                  );
                            }
                        }
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("----------------------------------------------------\n");


                        /////////////////
                        textBoxLog.AppendText("---= Список контрактов, по которым не велась работа =---\n");
                        results = dtMobile.Select("NumberUsed='False' AND NumberNoBlock='True'", sortOrder, DataViewRowState.Added);
                        textBoxLog.AppendText(
                             String.Format("{0,-40}", columnName1) +
                             String.Format("{0,-15}", columnName2) +
                             String.Format("{0,-30}", columnName3) +
                             String.Format("{0,-10}", columnName4) +
                             String.Format("{0,-30}", columnName5) +
                             "\n");
                        for (int i = 0; i < results.Length; i++)
                        {

                            textBoxLog.AppendText(
                             String.Format("{0,-40}", results[i][0].ToString()) +
                             String.Format("{0,-15}", results[i][2].ToString()) +
                             String.Format("{0,-30}", results[i][3].ToString()) +
                             String.Format("{0,-10}", results[i][10].ToString()) +
                             String.Format("{0,-30}", results[i][21].ToString()) +
                             "\n"
                              );
                        }
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("----------------------------------------------------\n");


                        /////////////////
                        textBoxLog.AppendText("---= Список заблокированных контрактов =---\n");
                        results = dtMobile.Select("NumberNoBlock='False'", sortOrder, DataViewRowState.Added);
                        textBoxLog.AppendText(
                             String.Format("{0,-40}", columnName1) +
                             String.Format("{0,-15}", columnName2) +
                             String.Format("{0,-30}", columnName3) +
                             String.Format("{0,-10}", columnName4) +
                             String.Format("{0,-30}", columnName5) +
                             "\n");
                        for (int i = 0; i < results.Length; i++)
                        {
                            textBoxLog.AppendText(
                             String.Format("{0,-40}", results[i][0].ToString()) +
                             String.Format("{0,-15}", results[i][2].ToString()) +
                             String.Format("{0,-30}", results[i][3].ToString()) +
                             String.Format("{0,-10}", results[i][10].ToString()) +
                             String.Format("{0,-30}", results[i][21].ToString()) +
                             "\n"
                              );
                        }
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("----------------------------------------------------\n");


                        /////////////////
                        textBoxLog.AppendText("---= Все =---\n");
                        results = dtMobile.Select(dtMobile.Columns[0].ColumnName.Length + " > 0", sortOrder, DataViewRowState.Added);
                        textBoxLog.AppendText(
                             String.Format("{0,-40}", columnName1) +
                             String.Format("{0,-15}", columnName2) +
                             String.Format("{0,-30}", columnName3) +
                             String.Format("{0,-10}", columnName4) +
                             String.Format("{0,-10}", columnName6) +

                             String.Format("{0,-30}", columnName5) +
                             String.Format("{0,-12}", columnName10) +
                             String.Format("{0,-12}", columnName11) +
                             "\n");
                        for (int i = 0; i < results.Length; i++)
                        {

                            textBoxLog.AppendText(
                             String.Format("{0,-40}", results[i][0].ToString().Trim()) +
                             String.Format("{0,-15}", results[i][2].ToString()) +
                             String.Format("{0,-30}", results[i][3].ToString()) +
                             String.Format("{0,-10}", results[i][10].ToString()) +
                             String.Format("{0,-10}", results[i][5].ToString()) +

                             String.Format("{0,-30}", results[i][21].ToString()) +
                             String.Format("{0,-12}", results[i][24].ToString()) +
                             String.Format("{0,-12}", results[i][25].ToString()) +
                             "\n"
                              );
                        }
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("====================================================\n");
                        /////////////////


                        makeReportAccountantItem.Enabled = true;
                        makeFullReportItem.Enabled = true;

                        StatusLabel1.Text = "Обработка счета завершена!";
                    }
                    else
                    {
                        textBoxLog.AppendText("В базе найдены новые, не настроенные в данной программе на обработку,\n");
                        textBoxLog.AppendText("модели тарификации компенсации затрат сотрудников:\n");
                        textBoxLog.AppendText("\n");
                        int i = 0;
                        foreach (string str in listTarifData)
                        {
                            textBoxLog.AppendText(++i + ". \"" + str + "\"\n");
                        }
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("====================================================\n");
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText(sbError.ToString());
                    }

                    if (infoStatusBar.Length > 1)
                    {
                        StatusLabel1.Text = infoStatusBar;
                        StatusLabel1.BackColor = System.Drawing.Color.SandyBrown;
                    }
                    makeReportAccountantItem.Enabled = true;
                    makeFullReportItem.Enabled = true;
                }

                filepathLoadedData = filePathTxt;

                if (listSavedNumbers.Count > 0)
                { listNumbers = listSavedNumbers; }
                if (listSavedServices.Count > 0)
                { listServices = listSavedServices; }
                if (listSavedNumbers.Count > 0 && listSavedServices.Count > 0)
                { prepareBillItem.Enabled = true; }
            }
            else { StatusLabel1.Text = "Файл с детализацией не выбран!  "; }

            openBillItem.Enabled = true;
            textBoxLog.Visible = true;
            // перейти в конец текстового файла
            // textBox1.SelectionStart = textBox1.Text.Length;
            // textBox1.ScrollToCaret();
        }

        private async void MakeExcelReport(Action action)
        {
            StatusLabel1.Text = "Обрабатываю полученные данные и формирую отчет...";

            makeReportAccountantItem.Enabled = false;
            makeFullReportItem.Enabled = false;
            openBillItem.Enabled = false;
            makeReportMarketingMenuItem.Enabled = false;

            await Task.Run(() => action());

            makeReportAccountantItem.Enabled = true;
            makeFullReportItem.Enabled = true;
            openBillItem.Enabled = true;
            makeReportMarketingMenuItem.Enabled = true;

            StatusLabel1.Text = @"Формирование отчета завершено. Файл сохранен в папку:  " + Path.GetDirectoryName(filePathTxt);

            GC.Collect();
        }

        /*
        private void Save_Click(object sender, EventArgs e) //Кнопка "Save"
        {
            saveFileDialog1.FileName = openFileDialog1.FileName + ".csv";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    TextWriter Writer = new StreamWriter(saveFileDialog1.FileName, false, Encoding.GetEncoding(1251));
                    Writer.Write(textBox1.Text);
                    Writer.Flush();
                    Writer.Close();
                }
                catch (Exception Expt)
                { // Отчет обо всех возможных ошибках
                    MessageBox.Show(Expt.ToString(), "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
                }
            }
        }*/

        private void readinifile() //Чтение парсеров из ini файла
        {
            string s = "";
            string tempConnectionServer = "";
            string tempConnectionUserName = "";
            string tempConnectionUserPasswords = "";

            bool b1 = false, b2 = false;
            toolTip1.SetToolTip(this.groupBox1, "Использованы настройки программы");

            if (File.Exists(pathToIni))
            {
                var Coder = Encoding.GetEncoding(1251);
                using (StreamReader Reader = new StreamReader(pathToIni, Coder))
                {
                    while ((s = Reader.ReadLine()) != null)
                    {
                        //Проверка ini файла на наличие строк с авторством
                        if (s.Contains(myFileVersionInfo.ProductName))
                        { b1 = true; }
                        else if (s.Contains(@"Author " + myFileVersionInfo.LegalCopyright))
                        { b2 = true; }

                        if (b1 && b2)
                        {
                            if (s.StartsWith("pConnectionServer="))
                            {
                                tempConnectionServer = Regex.Split(s, "pConnectionServer=")[1].Trim();
                                if (tempConnectionServer.Length > 1)
                                { sConnectionServer = tempConnectionServer; }
                            }
                            else if (s.StartsWith("pConnectionUserName="))
                            {
                                tempConnectionUserName = Regex.Split(s, "pConnectionUserName=")[1].Trim();
                                if (tempConnectionUserName.Length > 1)
                                { sConnectionUserName = tempConnectionUserName; }
                            }
                            else if (s.StartsWith("pConnectionUserPasswords="))
                            {
                                tempConnectionUserPasswords = Regex.Split(s, "pConnectionUserPasswords=")[1].Trim();
                                if (tempConnectionUserPasswords.Length > 1)
                                { sConnectionUserPasswords = tempConnectionUserPasswords; }
                            }
                            else if (s.StartsWith("parametrStart="))
                            { parametrStart = Regex.Split(s, "parametrStart=")[1].Trim(); }
                            else if (s.StartsWith("parametrEnd="))
                            { parametrEnd = Regex.Split(s, "parametrEnd=")[1].Trim(); }
                            else if (s.StartsWith("pStop="))
                            { pStop = Regex.Split(s, "pStop=")[1].Trim(); }
                            else if (s.StartsWith("pDiscount=")) //Срока с величиной скидки
                            { pDiscount = Regex.Split(s, "pDiscount=")[1].Trim(); }
                            else if (s.StartsWith("pFull=")) ///Строка с суммой счета до скидки
                            { pFull = Regex.Split(s, "pFull=")[1].Trim(); }

                            //Далее - обработка ini файла только с наличием авторства
                            for (int i = 0; i < pListParseStrings.Length; i++)
                            {
                                if (s.StartsWith("p" + i + "="))
                                { pListParseStrings[i] = Regex.Split(s, "p" + i + "=")[1].Trim(); }
                            }
                        }
                    }
                }

                if ((b1 && b2 == false) || (b2 && b1 == false))
                { toolTip1.SetToolTip(this.groupBox1, "Настройки из VodafoneInvoiceModifier.ini проигнорированы. Изменен формат файла"); }
                else
                {
                    groupBox1.BackColor = System.Drawing.Color.Tan;
                    toolTip1.SetToolTip(this.groupBox1, "Парсинг модифицирован настройками из VodafoneInvoiceModifier.ini");
                }
            }

            if (sConnectionServer.Length > 1)
            {
                sConnection = "Data Source=" + sConnectionServer +
                "; Initial Catalog=EBP;Type System Version=SQL Server 2005;Persist Security Info =True;User ID=" +
                sConnectionUserName + "; Password=" + sConnectionUserPasswords + "; Connect Timeout=180";
            }

            textBoxP1.Text = pListParseStrings[1];
            textBoxP2.Text = pListParseStrings[2];
            textBoxP3.Text = pListParseStrings[3];
            textBoxP4.Text = pListParseStrings[4];
            textBoxP5.Text = pListParseStrings[5];
            textBoxP6.Text = pListParseStrings[6];
            textBoxP7.Text = pListParseStrings[7];
            textBoxP8.Text = pStop;
            if (sConnection.Length < 15)
            {
                infoStatusBar = "Строка подключения к базе Tfactura  слишком короткая:\npConnection=" + sConnection;
                MessageBox.Show(infoStatusBar + "\nДобавьте в " + pathToIni + " строку подключения к базе данных", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                StatusLabel1.Text = infoStatusBar;

                StatusLabel1.BackColor = System.Drawing.Color.SandyBrown;
            }

            s = null;
        }

        private void writeinitofile() //Запись всех рабочих парсеров в ini файл
        {
            StringBuilder sb = new StringBuilder(String.Empty);
            DateTime localDate = DateTime.Now;

            try
            {
                sb.AppendLine(@"; This VodafoneInvoiceModifier.ini for " + myFileVersionInfo.ProductName);
                sb.AppendLine(@"; " + @"Author " + myFileVersionInfo.LegalCopyright);
                sb.AppendLine(@"");

                for (int i = 0; i < pListParseStrings.Length; i++)
                {
                    if (pListParseStrings[i]?.Length > 0)
                    { sb.AppendLine("p" + i + "=" + pListParseStrings[i]); }
                    else { sb.AppendLine("p" + i + "="); }
                }

                if (sConnectionServer?.Length > 1)
                { sb.AppendLine(@"pConnectionServer=" + sConnectionServer); }
                else { sb.AppendLine(@"pConnectionServer="); }

                if (sConnectionUserName?.Length > 1)
                { sb.AppendLine(@"pConnectionUserName=" + sConnectionUserName); }
                else { sb.AppendLine(@"pConnectionUserName="); }

                if (sConnectionUserPasswords?.Length > 1)
                { sb.AppendLine(@"pConnectionUserPasswords=" + sConnectionUserPasswords); }
                else { sb.AppendLine(@"pConnectionUserPasswords="); }

                if (sConnection?.Length > 15)
                { sb.AppendLine(@"pConnection=" + sConnection); }
                else { sb.AppendLine(@"pConnection="); }

                if (parametrStart?.Length > 0)
                { sb.AppendLine(@"parametrStart=" + parametrStart); }
                else { sb.AppendLine(@"parametrStart="); }

                if (pDiscount?.Length > 0)
                { sb.AppendLine(@"pDiscount=" + pDiscount); }
                else { sb.AppendLine(@"pDiscount="); }

                if (pFull?.Length > 0)
                { sb.AppendLine(@"pFull=" + pFull); }
                else { sb.AppendLine(@"pFull="); }

                if (parametrEnd?.Length > 0)
                { sb.AppendLine(@"parametrEnd=" + parametrEnd); }
                else { sb.AppendLine(@"parametrEnd="); }

                if (pStop?.Length > 0)
                { sb.AppendLine(@"pStop=" + pStop); }
                else { sb.AppendLine(@"pStop="); }

                sb.AppendLine(@"");
                sb.AppendLine(@"; Дата обновления файла:  " + localDate.ToString());

                File.WriteAllText(pathToIni, sb.ToString(), Encoding.GetEncoding(1251));
            }
            catch (Exception Expt)
            { MessageBox.Show(Expt.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            sb = null;
        }

        private bool ReadTxtAndWiteToMyTmp() //Чтение исходного файл, и первичный разбор счета (удаление ненужных данных)
        {
            bool ChosenFile = false;
            listTempContract.Clear();

            filePathTxt = _OpenFileDialogReturnPath(openFileDialog1);

            if (filePathTxt == null || filePathTxt.Length < 1) { return false; }
            else
            {
                try
                {
                    var Coder = Encoding.GetEncoding(1251);
                     discount = null;
                     beforeDiscount = null;
                    string test = null;
                    using (StreamReader Reader = new StreamReader(filePathTxt, Coder))
                    {
                        string s; int i = 0;
                        bool mystatusbegin = false;
                        int lenghtData=0;
                        _ToolStripStatusLabelSetItsText(StatusLabel1, "Обрабатываю файл:  " + filePathTxt);
                        while ((s = Reader.ReadLine()) != null)
                        {
                            if (s.Contains("Особовий рахунок"))
                            {
                                string[] substrings = Regex.Split(s, ":| ");
                                labelAccount.Visible = true;
                                labelAccount.Text = substrings[substrings.Length - 1].Trim();
                            }

                            else if (s.Contains("Номер рахунку"))
                            {
                                string[] substrings = Regex.Split(s, ":| ");
                                labelBill.Visible = true;
                                labelBill.Text = substrings[substrings.Length - 3].Trim();
                            }

                            else if (s.Contains(pDiscount))
                            {
                                lenghtData = (s.Split(':')[1].Trim()).Split(' ').Length;
                                discount = (s.Split(':')[1].Trim()).Split(' ')[lenghtData-1];
                            }
                            else if (s.Contains(pFull))
                            {
                                lenghtData = (s.Split(':')[1].Trim()).Split(' ').Length;
                                beforeDiscount = (s.Split(':')[1].Trim()).Split(' ')[lenghtData - 1];
                            }
                            else if (s.Contains("Розрахунковий період"))
                            {
                                string[] substrings = Regex.Split(s, ": ");
                                periodInvoice = substrings[substrings.Length - 1].Trim();
                                labelPeriod.Visible = true;
                                labelPeriod.Text = periodInvoice;
                            }
                            
                            if (s.Contains(pListParseStrings[1]))
                            {
                                mystatusbegin = true;
                                i += 1;
                            }

                            foreach (string pParseString in pListParseStrings)
                            {
                                if ((s.Contains(pParseString) || s.Contains(pStop)) && mystatusbegin)
                                { listTempContract.Add(s.Trim()); break; }
                            }
                        }
                        labelContracts.Visible = true;
                        labelContracts.Text = " " + i + " шт.";

                         discountDouble = 1;
                         beforeDiscountDouble = 1;
                        if (double.TryParse(discount, out discountDouble)&& double.TryParse(beforeDiscount, out beforeDiscountDouble)) //calculate current discount in the biil
                        {
                            discountValue =1 -(Math.Abs(Math.Round((discountDouble / beforeDiscountDouble), 2, MidpointRounding.AwayFromZero) * 100))/100;
                            labelDiscount.Text = ((1-discountValue)*100).ToString();
                            labelDiscount.Visible = true;
                        }
                        //todo
                       // add switch with discount caltulating
                    }


                    //----- Test module The Start. Dump onto a  local disk -----
                    /*
                    StringBuilder sb = new StringBuilder(String.Empty);
                    try
                    {
                        foreach (string str in listTempContract.ToArray())
                        { sb.AppendLine(str); }

                        File.WriteAllText(Path.GetDirectoryName(filePathTxt) + @"\listTempContract.txt", sb.ToString(), Encoding.GetEncoding(1251));
                    }
                    catch (Exception Expt)
                    { MessageBox.Show(Expt.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                    finally { sb = null; }
                    */
                    //Test module The End -----
                    
                    ChosenFile = true;
                }
                catch (Exception Expt)
                {
                    ChosenFile = false;
                    MessageBox.Show(Expt.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return ChosenFile;
            }
        }

        private double modelToPayment(MobileContractPerson mobileContractPerson)
        {
            double result = 0;

            for (int i = 0; i < arrayTarif.Length; i++)
            {
                if (mobileContractPerson.modelCompensation.Contains(arrayTarif[i]))
                {
                    switch (i)
                    {
                        case (0):     // "L100% корпорация",     //0
                            result = mobileContractPerson.content + mobileContractPerson.romingData;
                            break;

                        case (1):     // "L100% сотрудник",      //1
                            result = mobileContractPerson.totalCostWithTax;
                            break;

                        case (2):     // "L100%,R80%",           //2
                            result = mobileContractPerson.content + (mobileContractPerson.roming - mobileContractPerson.romingData) * 0.2 + mobileContractPerson.romingData;
                            break;

                        case (3):      // "L50,R0%",              //3
                            result = (mobileContractPerson.totalCostWithTax - mobileContractPerson.roming - 50 - mobileContractPerson.content) < 0 ?
                                      mobileContractPerson.roming + mobileContractPerson.content :
                                     (mobileContractPerson.totalCostWithTax - 50);
                            break;

                        case (4):     // "L80,R0%",              //4
                            result = (mobileContractPerson.totalCostWithTax - mobileContractPerson.roming - 80 - mobileContractPerson.content) < 0 ?
                                      mobileContractPerson.roming + mobileContractPerson.content :
                                     (mobileContractPerson.totalCostWithTax - 80);
                            break;

                        case (5):     // "L100,R0%",             //5
                            result = (mobileContractPerson.totalCostWithTax - mobileContractPerson.roming - 100 - mobileContractPerson.content) < 0 ?
                                      mobileContractPerson.roming + mobileContractPerson.content :
                                     (mobileContractPerson.totalCostWithTax - 100);
                            break;

                        case (6):     // "L160,R0%",             //6
                            result = (mobileContractPerson.totalCostWithTax - mobileContractPerson.roming - 160 - mobileContractPerson.content) < 0 ?
                                      mobileContractPerson.roming + mobileContractPerson.content :
                                     (mobileContractPerson.totalCostWithTax - 160);
                            break;

                        case (7):     // "L250,R0%",             //7
                            result = (mobileContractPerson.totalCostWithTax - mobileContractPerson.roming - 250 - mobileContractPerson.content) < 0 ?
                                      mobileContractPerson.roming + mobileContractPerson.content :
                                     (mobileContractPerson.totalCostWithTax - 250);
                            break;

                        case (8):      // "L50%,R0%",             //8
                            result = (mobileContractPerson.totalCostWithTax - mobileContractPerson.roming - mobileContractPerson.content) * 0.5 +
                                      mobileContractPerson.roming + mobileContractPerson.content;
                            break;

                        case (9):     // "L50%,R80%",            //9
                            result = (mobileContractPerson.totalCostWithTax - mobileContractPerson.roming - mobileContractPerson.content) * 0.5 +
                                     (mobileContractPerson.roming - mobileContractPerson.romingData) * 0.2 + mobileContractPerson.romingData +
                                      mobileContractPerson.content;
                            break;

                        case (10):    // "L50%,R100%",           //10
                            result = (mobileContractPerson.totalCostWithTax - mobileContractPerson.roming - mobileContractPerson.content) * 0.5 +
                                      mobileContractPerson.romingData + mobileContractPerson.content;
                            break;

                        case (11):    // "L90%,R100%",           //11
                            result = (mobileContractPerson.totalCostWithTax - mobileContractPerson.roming - mobileContractPerson.content) * 0.1 +
                                      mobileContractPerson.romingData + mobileContractPerson.content;
                            break;

                        case (12):     // "Lpack100%,R0%,Paid0%", //12
                            result = (mobileContractPerson.totalCostWithTax - mobileContractPerson.monthCost - mobileContractPerson.roming - mobileContractPerson.content - mobileContractPerson.extraServiceOrdered) < 0 ?
                                      mobileContractPerson.roming + mobileContractPerson.content + mobileContractPerson.extraServiceOrdered :
                                     (mobileContractPerson.totalCostWithTax - mobileContractPerson.monthCost);
                            break;

                        case (13):     // "Lмоб200,R0%,Paid0%"    //13
                            result = (mobileContractPerson.totalCostWithTax - mobileContractPerson.outToCity - mobileContractPerson.roming - mobileContractPerson.content - mobileContractPerson.extraInternetOrdered - 200) < 0 ?
                                      mobileContractPerson.outToCity + mobileContractPerson.roming + mobileContractPerson.content + mobileContractPerson.extraInternetOrdered :
                                     (mobileContractPerson.totalCostWithTax - 200);
                            break;

                        case (14):    // "L200,R0%",             //14
                            result = (mobileContractPerson.totalCostWithTax - mobileContractPerson.roming - 200 - mobileContractPerson.content) < 0 ?
                                      mobileContractPerson.roming + mobileContractPerson.content :
                                     (mobileContractPerson.totalCostWithTax - 200);
                            break;


                        default:
                            result = 0;
                            break;
                    }
                }
            }
            return result;
        }

        private double tax(double beforePayTax)
        { return beforePayTax * 0.2; }

        private double pF(double beforePayToPF)
        { return beforePayToPF * 0.075; }

        private void MyTmpToMyArray() //Парсинг строк и передача результата текстовый редактор
        {
            dtMobile.Rows.Clear();
            DataRow row = dtMobile.NewRow();
            bool isUsedCurrent = false;
            bool isCheckFinishedTitles = false;

            dataStart = labelPeriod.Text.Split('-')[0].Trim(); // дата начала периода счета
            dataEnd = labelPeriod.Text.Split('-')[1].Trim();  // дата конца периода счета
            _ToolStripStatusLabelSetItsText(StatusLabel1, "Обрабатываю полученные данные...");
            string n = "", searchNumber;
            string[] substrings = new string[1];

            strNewModels = "";

            /*    // Test only
            StringBuilder sb = new StringBuilder(String.Empty);
            */

            MobileContractPerson mcpCurrent = new MobileContractPerson();
            try
            {
                foreach (string s in listTempContract.ToArray())
                {
                    if (s.Contains(pListParseStrings[1]) || s.Contains(pStop))
                    {
                        isCheckFinishedTitles = false;
                        if (mcpCurrent.contractName.Length > 1)
                        {
                            mcpCurrent.dateBillStart = dataStart;
                            mcpCurrent.dateBillEnd = dataEnd;
                            mcpCurrent.tax = tax(mcpCurrent.totalCost);
                            mcpCurrent.pF = pF(mcpCurrent.totalCost);
                            mcpCurrent.totalCostWithTax = mcpCurrent.totalCost * 1.275;  //number spend+НДС+ПФ

                            searchNumber = mcpCurrent.mobNumberName;
                            foreach (DataRow dr in dtTarif.Rows)
                            {
                                if (dr.ItemArray[0].ToString().Contains(searchNumber))
                                {
                                    mcpCurrent.ownerName = dr.ItemArray[1].ToString();
                                    mcpCurrent.NAV = dr.ItemArray[2].ToString();
                                    mcpCurrent.orgUnit = dr.ItemArray[3].ToString();
                                    mcpCurrent.startDate = dr.ItemArray[5].ToString();
                                    mcpCurrent.modelCompensation = dr.ItemArray[6].ToString();
                                }
                            }
                            mcpCurrent.payOwner = modelToPayment(mcpCurrent);

                            mcpCurrent.isUsed = isUsedCurrent;
                            if (mcpCurrent.totalCostWithTax > 0)
                            { mcpCurrent.isUnblocked = true; }

                            row = dtMobile.NewRow();
                            row[0] = mcpCurrent.ownerName;
                            row[1] = mcpCurrent.contractName;
                            row[2] = mcpCurrent.mobNumberName;
                            row[3] = mcpCurrent.tarifPackageName;
                            row[4] = Math.Round(mcpCurrent.monthCost, 2);
                            row[5] = Math.Round(mcpCurrent.roming, 2);
                            row[6] = Math.Round(mcpCurrent.discount, 2);
                            row[7] = Math.Round(mcpCurrent.totalCost, 2);
                            row[8] = Math.Round(mcpCurrent.tax, 2);
                            row[9] = Math.Round(mcpCurrent.pF, 2);
                            row[10] = Math.Round(mcpCurrent.totalCostWithTax, 2);
                            row[11] = Math.Round(mcpCurrent.romingData, 2);
                            row[12] = Math.Round(mcpCurrent.extraInternetOrdered, 2);
                            row[13] = Math.Round(mcpCurrent.outToCity, 2);
                            row[14] = Math.Round(mcpCurrent.extraService, 2);
                            row[15] = Math.Round(mcpCurrent.content, 2);
                            row[16] = mcpCurrent.dateBillStart;
                            row[17] = mcpCurrent.dateBillEnd;
                            row[18] = mcpCurrent.NAV;
                            row[19] = mcpCurrent.orgUnit;
                            row[20] = mcpCurrent.startDate;
                            row[21] = mcpCurrent.modelCompensation;
                            row[22] = Math.Round(mcpCurrent.payOwner, 2);
                            row[23] = Math.Round(mcpCurrent.extraServiceOrdered, 2);
                            //проверки контракта
                            row[24] = mcpCurrent.isUsed;
                            row[25] = mcpCurrent.isUnblocked;
                            //запись сформированной строки в таблицу
                            dtMobile.Rows.Add(row);

                            //запись дубля в список
                            //Test only
                            // sb.AppendLine(mcpCurrent.mobNumberName + " - " + mcpCurrent.totalCost * 1.275 + "(with tax) - " + mcpCurrent.totalCost + "(without tax) - ");
                        }

                        mcpCurrent = new MobileContractPerson();
                        substrings = s.Split('№')[s.Split('№').Length - 1].Trim().Split(' ');
                        mcpCurrent.contractName = substrings[0].Trim();

                        if (s.Contains(pListParseStrings[2]))
                        {
                            substrings = s.Split(':')[s.Split(':').Length - 1].Trim().Split(' ');
                            mcpCurrent.mobNumberName = substrings[substrings.Length - 1].Trim();
                        }
                    }

                    else if (s.Contains(pListParseStrings[3]))
                    {
                        substrings = s.Split(':');
                        mcpCurrent.tarifPackageName = substrings[substrings.Length - 1].Trim();
                    }

                    else if (s.Contains(pListParseStrings[4]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.monthCost = Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * discountValue * 1.275;
                    }

                    else if (s.Contains(pListParseStrings[5]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.roming = Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275;
                    }

                    else if (s.Contains(pListParseStrings[6]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.discount = Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
                    }

                    else if (s.Contains(pListParseStrings[7]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.totalCost = Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
                        isCheckFinishedTitles = true;
                        isUsedCurrent = false;
                    }

                    else if (s.Contains(pListParseStrings[11]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.romingData += Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275;
                    }

                    else if (s.Contains(pListParseStrings[12]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.extraInternetOrdered += Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275 * discountValue;
                    }

                    else if (s.Contains(pListParseStrings[13]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.outToCity += Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275 * discountValue;
                    }

                    else if (s.Contains(pListParseStrings[14]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.extraService += Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
                    }

                    else if (s.Contains(pListParseStrings[15]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.content += Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275;
                    }

                    else if (s.Contains(pListParseStrings[23]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.extraServiceOrdered += Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275 * discountValue;
                    }
                    else if (s.Equals(pDiscount))
                    {
                        discount = s.Split(':')[1].Trim();
                    }
                    else if (s.Equals(pFull))
                    {
                        beforeDiscount = s.Split(':')[1].Trim();
                    }
                    else if (isCheckFinishedTitles)
                    { isUsedCurrent = true; }
                }

                //additional payment for detalisation (the end row)
                mcpCurrent = new MobileContractPerson();
                row = dtMobile.NewRow();

                row[0] = "за детализацию счета, коррекция суммы";
                row[1] = mcpCurrent.contractName;
                row[2] = mcpCurrent.mobNumberName;
                row[3] = mcpCurrent.tarifPackageName;
                row[4] = mcpCurrent.monthCost;
                row[5] = mcpCurrent.roming;
                row[6] = Math.Round(-24.9999, 2);
                row[7] = Math.Round(83.3333, 2);
                row[8] = Math.Round(11.67, 2);
                row[9] = Math.Round(4.39, 2);
                row[10] = Math.Round(74.375, 2);
                row[11] = mcpCurrent.romingData;
                row[12] = mcpCurrent.extraInternetOrdered;
                row[13] = mcpCurrent.outToCity;
                row[14] = mcpCurrent.extraService;
                row[15] = mcpCurrent.content;
                row[16] = dataStart;
                row[17] = mcpCurrent.dateBillEnd;

                row[18] = "E22";
                row[19] = "IT-дирекция";
                row[20] = mcpCurrent.startDate;
                row[21] = "T[6] L100% корпорация";
                row[22] = mcpCurrent.payOwner;

                row[23] = mcpCurrent.extraServiceOrdered;

                dtMobile.Rows.Add(row);
            }
            catch (Exception Expt) { MessageBox.Show(Expt.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }

            //Test only
            //File.WriteAllText(Application.StartupPath + @"\VodafoneCollector.txt", sb.ToString(), Encoding.GetEncoding(1251));
            //sb = null;

            row = null;
            mcpCurrent = null;
            listTempContract.Clear();
        }

        /*
        [System.Runtime.InteropServices.DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);
        private static void KillExcel(Microsoft.Office.Interop.Excel.Application theApp)
        {
            int id = 0;
            IntPtr intptr = new IntPtr(theApp.Hwnd);
            System.Diagnostics.Process p = null;
            try
            {
                GetWindowThreadProcessId(intptr, out id);
                p = System.Diagnostics.Process.GetProcessById(id);
                if (p != null)
                {
                    p.Kill();
                    p.Dispose();
                }
            }
            catch (Exception ex)
            {
              //  System.Windows.Forms.MessageBox.Show("KillExcel:" + ex.Message);
            }
        }
        */

        private void ExportDatatableToExcel(DataTable dt, string sufixExportFile) //Заполнение таблицы в Excel  данными
        {
            _ProgressBar1Start();
            int rows = 1;
            int rowsInTable = dt.Rows.Count;

            int columnsInTable = dt.Columns.Count; // p.Length;

            int stepOfProgressCount = (rowsInTable * columnsInTable) / 100;

            string lastCell = GetColumnName(columnsInTable) + rowsInTable;
            _ProgressWork1Step();
            Excel.Application excel = new Excel.Application
            {
                Visible = false, //делаем объект не видимым
                SheetsInNewWorkbook = 1//количество листов в книге
            };

            Excel.Workbooks workbooks = excel.Workbooks;
            excel.Workbooks.Add(); //добавляем книгу
            Excel.Workbook workbook = workbooks[1];
            Excel.Sheets sheets = workbook.Worksheets;
            Excel.Worksheet sheet = sheets.get_Item(1);
            sheet.Name = Path.GetFileNameWithoutExtension(filepathLoadedData);
            _ProgressWork1Step();

            for (int k = 1; k < columnsInTable; k++)
            {
                sheet.Cells[k].WrapText = true;
                sheet.Cells[1, k].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                sheet.Cells[1, k].VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                sheet.Cells[1, k + 1].Value = dt.Columns[k].ColumnName;
                //string columnName = dt.Columns[0].Caption;

                sheet.Columns[k].Font.Size = 8;
                sheet.Columns[k].Font.Name = "Tahoma";

                //colourize of collumns
                sheet.Cells[1, k].Interior.Color = System.Drawing.Color.Silver;
                _ProgressWork1Step();
            }

            //input data and set type of cells - numbers /text
            int stepCount = stepOfProgressCount;
            foreach (DataRow row in dt.Rows)
            {
                rows++;
                foreach (DataColumn column in dt.Columns)
                {
                    if (rows > 1)
                    {
                        if (row[column.Ordinal].GetType().ToString().ToLower().Contains("string"))
                        { sheet.Columns[column.Ordinal + 1].NumberFormat = "@"; }
                        else
                        { sheet.Columns[column.Ordinal + 1].NumberFormat = "0" + System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator + "00"; }
                    }
                    sheet.Cells[rows, column.Ordinal + 1].Value = row[column.Ordinal];
                    stepCount--;
                    if (stepCount == 0)
                    {
                        _ProgressWork1Step("Обработано " + rows + " строк, из " + rowsInTable);
                        stepCount = stepOfProgressCount;
                    }
                    //  sheet.Columns[column.Ordinal + 1].AutoFit();
                }
            }

            //Autofilter                
            Excel.Range range = sheet.UsedRange;  //sheet.Cells.Range["A1", lastCell];

            //ширина колонок - авто
            range.Cells.EntireColumn.AutoFit();
            _ProgressWork1Step();

            range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            range.Cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            range.Select();
            _ProgressWork1Step();

            range.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

            workbook.SaveAs(
                Path.GetDirectoryName(filepathLoadedData) + @"\" + Path.GetFileNameWithoutExtension(filepathLoadedData) + sufixExportFile,
                Excel.XlFileFormat.xlOpenXMLWorkbook,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                Excel.XlSaveAsAccessMode.xlExclusive, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

            workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            workbooks.Close();
            _ProgressWork1Step(" ");

            lastCell = null;
            releaseObject(range);
            releaseObject(sheet);
            releaseObject(sheets);
            releaseObject(workbook);
            releaseObject(workbooks);
            excel.Quit();
            releaseObject(excel);
            _ProgressBar1Stop();
        }

        private void _ProgressWork1Step(string text = "") //add into progressBar Value 2 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    if (ProgressBar1.Value > 99)
                    { ProgressBar1.Value = 0; }
                    ProgressBar1.Maximum = 100;
                    ProgressBar1.Value += 1;
                    if (text.Length > 0)
                        _ToolStripStatusLabelSetItsText(StatusLabel1, text);
                }));
            else
            {
                if (ProgressBar1.Value > 99)
                { ProgressBar1.Value = 0; }
                ProgressBar1.Maximum = 100;
                ProgressBar1.Value += 1;
                if (text.Length > 0)
                    _ToolStripStatusLabelSetItsText(StatusLabel1, text);
            }
        }

        private void _ProgressBar1Start() //Set progressBar Value into 0 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    ProgressBar1.Value = 0;
                }));
            else
            {
                ProgressBar1.Value = 0;
            }
        }

        private void _ProgressBar1Stop() //Set progressBar Value into 100 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    ProgressBar1.Value = 100;
                }));
            else
            {
                ProgressBar1.Value = 100;
            }
        }


        private void ExportFullDataTableToExcel() //Заполнение таблицы в Excel всеми данными
        {
            int rows = 1;
            int rowsInTable = dtMobile.Rows.Count;
            int columnsInTable = pListParseStrings.Length; // p.Length;
            string lastCell = GetColumnName(columnsInTable) + rowsInTable;

            Excel.Application excel = new Excel.Application
            {
                Visible = false, //делаем объект не видимым
                SheetsInNewWorkbook = 1//количество листов в книге
            };

            Excel.Workbooks workbooks = excel.Workbooks;
            excel.Workbooks.Add(); //добавляем книгу
            Excel.Workbook workbook = workbooks[1];
            Excel.Sheets sheets = workbook.Worksheets;
            Excel.Worksheet sheet = sheets.get_Item(1);
            sheet.Name = Path.GetFileNameWithoutExtension(filePathTxt);
            // sheet.Names.Add("next", "=" + Path.GetFileNameWithoutExtension(filePathTxt) + "!$A$1", true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            HashSet<string> listCollumnsHide = new HashSet<string>(pTranslate);
            listCollumnsHide.ExceptWith(new HashSet<string>(pToAccount));

            for (int k = 0; k < columnsInTable; k++)
            {
                sheet.Cells[k + 1].WrapText = true;
                sheet.Cells[1, k + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                sheet.Cells[1, k + 1].VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                sheet.Cells[1, k + 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.Cells[1, k + 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                sheet.Cells[1, k + 1].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.Cells[1, k + 1].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;

                sheet.Cells[1, k + 1].Value = pTranslate[k];

                sheet.Columns[k + 1].Font.Size = 8;
                sheet.Columns[k + 1].Font.Name = "Tahoma";

                //colourize of collumns
                if (pTranslate[k].Equals("Итого по контракту, грн"))
                { sheet.Columns[k + 1].Interior.Color = System.Drawing.Color.DarkSeaGreen; }
                else if (pTranslate[k].Equals("К оплате владельцем номера, грн"))
                { sheet.Columns[k + 1].Interior.Color = System.Drawing.Color.SandyBrown; }
                else { sheet.Cells[1, k + 1].Interior.Color = System.Drawing.Color.Silver; }
            }

            //input data and set type of cells - numbers /text
            foreach (DataRow row in dtMobile.Rows)
            {
                rows++;
                foreach (DataColumn column in dtMobile.Columns)
                {
                    if (rows == 2)
                    {
                        if (row[column.Ordinal].GetType().ToString().ToLower().Contains("string"))
                        { sheet.Columns[column.Ordinal + 1].NumberFormat = "@"; }
                        else
                        { sheet.Columns[column.Ordinal + 1].NumberFormat = "0" + System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator + "00"; }
                    }
                    sheet.Cells[rows, column.Ordinal + 1].Value = row[column.Ordinal];
                    sheet.Cells[rows, column.Ordinal + 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    sheet.Cells[rows, column.Ordinal + 1].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                    sheet.Cells[rows, column.Ordinal + 1].Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    sheet.Cells[rows, column.Ordinal + 1].Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
                    sheet.Columns[column.Ordinal + 1].AutoFit();
                }
            }

            //Область сортировки   
            Excel.Range range = sheet.Range["A2", lastCell];

            //По какому столбцу сортировать
            string nameColumnSorted = GetColumnName(Array.IndexOf(pTranslate, "Номер телефона абонента") + 1);
            Excel.Range rangeKey = sheet.Range[nameColumnSorted + (rowsInTable - 1)];

            //Добавляем параметры сортировки
            sheet.Sort.SortFields.Add(rangeKey);
            sheet.Sort.SetRange(range);
            sheet.Sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
            sheet.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
            sheet.Sort.Apply();

            //Очищаем фильтр
            sheet.Sort.SortFields.Clear();

            for (int k = 0; k < pTranslate.Length; k++)
            {
                foreach (string str in listCollumnsHide)
                {
                    if (pTranslate[k].Equals(str))
                    {
                        sheet.Columns[k + 1].Hidden = true;
                    }
                }
            }

            //Autofilter                
            range = sheet.UsedRange;  //sheet.Cells.Range["A1", lastCell];
            range.Select();
            range.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

            workbook.SaveAs(
                Path.GetDirectoryName(filePathTxt) + @"\" + Path.GetFileNameWithoutExtension(filePathTxt) + @"_full.xlsx",
                Excel.XlFileFormat.xlOpenXMLWorkbook,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                Excel.XlSaveAsAccessMode.xlExclusive, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);

            workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            workbooks.Close();

            listCollumnsHide = null;
            nameColumnSorted = null;
            lastCell = null;
            releaseObject(range);
            releaseObject(rangeKey);
            releaseObject(sheet);
            releaseObject(sheets);
            releaseObject(workbook);
            releaseObject(workbooks);
            excel.Quit();
            releaseObject(excel);

            //  autofill. manualy set number in D1 and D2, then use function
            //rng = this.Application.get_Range("D1","D2");
            //Excel.Range rng.AutoFill(this.Application.get_Range("D1", "D5"), Excel.XlAutoFillType.xlFillSeries);
            //  add comment:
            //Excel.Range dateComment = this.Application.get_Range("A1");
            //dateComment.AddComment("Comment added " + DateTime.Now.ToString());
            //  delete comment:
            //if (dateComment.Comment != null) { dateComment.Comment.Delete(); }

            // sheet.Cells[1, k + 1].Font.Bold = true;
            // (sheet.Cells[1, column.Ordinal + 1] as Microsoft.Office.Interop.Excel.Range).Font.Size = 8;

            //объединение ячеек
            //sheet.get_Range(sheet.Cells[2, 2], sheet.Cells[4, 4]).Merge(missing);
            //(sheet.Columns).ColumnWidth = 15;
            // sheet.Columns.Font.Size = System.Drawing.Color.LightPink;
        }

        private void ExportDataTableToExcelForAccount() //Заполнение таблицы в Excel данными для бухгалтерии
        {
            int[] pIdxToAccount = new int[]
           {
                // для бухгалтерии
                dtMobile.Columns.IndexOf("Дата счета"),
                dtMobile.Columns.IndexOf("Номер телефона абонента"),
                dtMobile.Columns.IndexOf("ФИО сотрудника"),
                dtMobile.Columns.IndexOf("Затраты по номеру, грн"),
                dtMobile.Columns.IndexOf("НДС, грн"),
                dtMobile.Columns.IndexOf("ПФ, грн"),
                dtMobile.Columns.IndexOf("Итого по контракту, грн"),
                dtMobile.Columns.IndexOf("Общая сумма в роуминге, грн"),
                dtMobile.Columns.IndexOf("Подразделение"),
                dtMobile.Columns.IndexOf("Табельный номер"),
                dtMobile.Columns.IndexOf("ТАРИФНАЯ МОДЕЛЬ"),
                dtMobile.Columns.IndexOf("К оплате владельцем номера, грн")
           };

            int rows = 1;
            int rowsInTable = dtMobile.Rows.Count;
            int columnsInTable = pIdxToAccount.Length; // p.Length;

            Excel.Application excel = new Excel.Application
            {
                Visible = false, //делаем объект не видимым
                SheetsInNewWorkbook = 1//количество листов в книге
            };
            Excel.Workbooks workbooks = excel.Workbooks;
            excel.Workbooks.Add(); //добавляем книгу
            Excel.Workbook workbook = workbooks[1];
            Excel.Sheets sheets = workbook.Worksheets;
            Excel.Worksheet sheet = sheets.get_Item(1);
            sheet.Name = Path.GetFileNameWithoutExtension(filePathTxt);
            //sheet.Names.Add("next", "=" + Path.GetFileNameWithoutExtension(filePathTxt) + "!$A$1", true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            for (int k = 0; k < columnsInTable; k++)
            {
                sheet.Cells[k + 1].WrapText = true;
                sheet.Cells[k + 1].Interior.Color = System.Drawing.Color.Silver;
                sheet.Cells[k + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                sheet.Cells[k + 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                sheet.Cells[1, k + 1].Value = pToAccount[k];
                sheet.Columns[k + 1].Font.Size = 8;
                sheet.Columns[k + 1].Font.Name = "Tahoma";

                switch (k)
                {
                    case 0:
                    case 1:
                    case 2:
                    case 8:
                    case 9:
                    case 10:
                        {
                            sheet.Columns[k + 1].NumberFormat = "@";
                            break;
                        }
                    case 3:
                    case 4:
                    case 5:
                    case 6:
                    case 7:
                    case 11:
                        {
                            sheet.Columns[k + 1].NumberFormat = "0" + System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator + "00";
                            sheet.Columns[k + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            break;
                        }
                }
            }

            //colourize of collumns
            sheet.Columns[7].Interior.Color = System.Drawing.Color.DarkSeaGreen;  //"Итого по контракту, грн"
            sheet.Columns[columnsInTable].Interior.Color = System.Drawing.Color.SandyBrown;  //"К оплате владельцем номера, грн"

            //input data and set type of cells - numbers /text
            foreach (DataRow row in dtMobile.Rows)
            {
                rows++;
                for (int column = 0; column < columnsInTable; column++)
                {
                    sheet.Cells[rows, column + 1].Value = row[pIdxToAccount[column]];
                }
            }

            //Область сортировки          
            Excel.Range range = sheet.Range["A2", GetColumnName(columnsInTable) + (rows - 1)];

            //По какому столбцу сортировать
            string nameColumnSorted = GetColumnName(Array.IndexOf(pIdxToAccount, dtMobile.Columns.IndexOf("Номер телефона абонента")) + 1);
            Excel.Range rangeKey = sheet.Range[nameColumnSorted + (rowsInTable - 1)];

            //Добавляем параметры сортировки
            sheet.Sort.SortFields.Add(rangeKey);
            sheet.Sort.SetRange(range);
            sheet.Sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
            sheet.Sort.SortMethod = Excel.XlSortMethod.xlPinYin;
            sheet.Sort.Apply();
            //Очищаем фильтр
            sheet.Sort.SortFields.Clear();

            //Autofilter
            range = sheet.UsedRange; //sheet.Cells.Range["A1", GetColumnName(columnsInTable) + rowsInTable];
            range.Select();

            //Форматирование колонок (стиль линий обводки)
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
            range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = Excel.XlBorderWeight.xlThin;
            range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = Excel.XlBorderWeight.xlThin;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].Weight = Excel.XlBorderWeight.xlThin;
            range.AutoFilter(1, Type.Missing, Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);

            workbook.SaveAs(Path.GetDirectoryName(filePathTxt) + @"\" + Path.GetFileNameWithoutExtension(filePathTxt) + @".xlsx",
                Excel.XlFileFormat.xlOpenXMLWorkbook,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                Excel.XlSaveAsAccessMode.xlExclusive,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            workbooks.Close();

            releaseObject(range);
            releaseObject(rangeKey);
            releaseObject(sheet);
            releaseObject(sheets);
            releaseObject(workbook);
            releaseObject(workbooks);
            excel.Quit();
            releaseObject(excel);
            MessageBox.Show("Отчет готов и сохранен:\n" + Path.GetDirectoryName(filePathTxt) + @"\" + Path.GetFileNameWithoutExtension(filePathTxt) + @".xlsx");
        }

        private void releaseObject(object obj)
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            obj = null;
        }

        /*private static string GetColumnName1(int columnNumber)  // получить букву колонки в Екселе по ее номеру (нумерация идет от 1)
        {
            int dividend = columnNumber;
            string columnName = "A";
            try
            {
                int modulo;
                while (dividend > 0)
                {
                    modulo = (dividend - 1) % 26;
                    columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                    dividend = (int)((dividend - modulo) / 26);
                }
            }
            catch(Exception expt) { MessageBox.Show(expt.ToString()); }
            return columnName;
        }*/

        static string GetColumnName(int number)
        {
            string result;
            if (number > 0)
            {
                int alphabets = (number - 1) / 26;
                int remainder = (number - 1) % 26;
                result = ((char)('A' + remainder)).ToString();
                if (alphabets > 0)
                    result = GetColumnName(alphabets) + result;
            }
            else
                result = null;
            return result;
        }

        private void GetDataWithModel()  // получение данных из базы ТФактура
        {
            dataStart = _ControlReturnItsText(labelPeriod).Split('-')[0].Trim(); //'01.05.2018'
            dataEnd = _ControlReturnItsText(labelPeriod).Split('-')[1].Trim();  //'31.05.2018'
            string dataStartSearch = dataStart.Split('.')[2] + "-" + dataStart.Split('.')[1] + "-" + dataStart.Split('.')[0]; //'2018-05-01'
            string dataEndSearch = dataEnd.Split('.')[2] + "-" + dataEnd.Split('.')[1] + "-" + dataEnd.Split('.')[0]; //'2018-05-31'
            contractNumberFound = 0;
            listTarifData = new HashSet<string>();
            string sSqlQuery = "SELECT t1.*, t1.descr AS main," +
                                   " t2.emp_cd AS NAV, t2.emp_id AS t2emp_id," +
                                   " t3.contract_id as t3contract_id, t3.pay_model_id," +
                                   " t4.name AS model_name, " +
                                   " t5.tariff_package_name AS tariff, t5.begin_dt AS first_data , t5.end_dt AS last_data" +
                                   " FROM v_rs_contract_detail t1" +
                                   " INNER JOIN os_emp t2 ON t1.emp_id = t2.emp_id" +
                                   " LEFT JOIN (SELECT* FROM os_contract_link WHERE till_dt IS NULL)  t3 ON t1.contract_id = t3.contract_id" +
                                   " LEFT JOIN rs_pay_model t4 ON t3.pay_model_id = t4.pay_model_id" +
                                   " RIGHT JOIN (SELECT contract_id, tariff_package_name, begin_dt, end_dt, contract_bill_id FROM v_dp_contract_bill_detail_ex) t5 ON t1.contract_id = t5.contract_id" +
                                   " WHERE t1.emp_id IS NOT NULL" +
                                   " AND" +
                                   " t5.end_dt = '" + dataEndSearch + "'" +
                                   // t5.end_dt = '2018-05-31'
                                   // (DATEPART(yy, t5.end_dt) = 2018 AND DATEPART(mm, t5.end_dt) = 05 AND DATEPART(dd, t5.end_dt) = 31) 
                                   " AND " +
                                   " t5.begin_dt = '" + dataStartSearch + "'" +
                                   " AND (" +
                                   " t1.till_dt IS null" +
                                   " OR" +
                                   " t1.till_dt > '" + dataEndSearch + "'" +
                                   " ) AND" +
                                   " t1.from_dt < '" + dataEndSearch + "'" +
                                   " ORDER by t1.phone_no, t1.emp_name ;";
            try
            {
                using (System.Data.SqlClient.SqlConnection sqlConnection = new System.Data.SqlClient.SqlConnection(sConnection))
                {
                    sqlConnection.Open();
                    dtTarif.Rows.Clear();
                    using (System.Data.SqlClient.SqlCommand sqlCommand = new System.Data.SqlClient.SqlCommand(sSqlQuery, sqlConnection))
                    {
                        using (System.Data.SqlClient.SqlDataReader sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (System.Data.Common.DbDataRecord record in sqlReader)
                            {
                                if (record != null && record.ToString().Length > 0 && record["phone_no"].ToString().Length > 0)
                                {
                                    DataRow row = dtTarif.NewRow();
                                    row["Номер телефона"] = MakeCommonViewPhone(record["phone_no"].ToString());
                                    row["ФИО"] = record["emp_name"].ToString().Trim();
                                    row["NAV"] = record["NAV"].ToString().Trim();
                                    row["Подразделение"] = record["org_unit_name"].ToString().Trim();
                                    row["Основной"] = DefineMainPhone(record["main"].ToString());
                                    row["Действует c"] = record["from_dt"].ToString().Trim().Split(' ')[0];
                                    row["Модель компенсации"] = "T[" + record["pay_model_id"].ToString().Trim() + "] " + record["model_name"].ToString().Trim();

                                    //record contracts with error
                                    if (record["pay_model_id"].ToString().Trim().Length == 0) sbError.AppendLine(row["Номер телефона"].ToString().Trim() + ", " + row["ФИО"].ToString().Trim() + " - " + row["Модель компенсации"]);

                                    //if( record["model_name"].ToString().Trim().Length>0 ) listTarifData.Add(record["model_name"].ToString().Trim());
                                    listTarifData.Add(record["model_name"].ToString().Trim());
                                    dtTarif.Rows.Add(row);
                                    contractNumberFound++;
                                }
                            }
                        }
                    }
                }

            }
            catch (System.Data.SqlClient.SqlException expt) { MessageBox.Show(expt.ToString()); }
            catch (Exception expt) { MessageBox.Show(expt.ToString()); }

            sSqlQuery = null;
        }

        private string MakeCommonViewPhone(string sPrimaryPhone) //Normalize Phone to +380504197443
        {
            string sPhone = sPrimaryPhone.Trim();
            string sTemp1 = "", sTemp2 = "";
            sTemp1 = sPhone.Replace(" ", "");
            sTemp2 = sTemp1.Replace("-", "");
            sTemp1 = sTemp2.Replace(")", "");
            sTemp2 = sTemp1.Replace("(", "");
            sTemp1 = sTemp2.Replace("/", "");
            sTemp2 = sTemp1.Replace("_", "");

            if (sTemp2.StartsWith("+") && sTemp2.Length == 13) sPhone = sTemp2;
            else if (sTemp2.StartsWith("380") && sTemp2.Length == 12) sPhone = "+" + sTemp2;
            else if (sTemp2.StartsWith("80") && sTemp2.Length == 11) sPhone = "+3" + sTemp2;
            else if (sTemp2.StartsWith("0") && sTemp2.Length == 10) sPhone = "+38" + sTemp2;
            else if (sTemp2.Length == 9) sPhone = "+380" + sTemp2;
            else sPhone = sTemp2;

            sTemp1 = ""; sTemp2 = "";
            return sPhone;
        }

        private string DefineMainPhone(string sDescription)
        {
            if (sDescription.Trim() == "!") { return "Да"; }
            else { return ""; }
        }

        private void CheckNewTarif()
        {
            string pathToNewModels = Application.StartupPath + @"\VodafoneInvoiceModifierNewModels.txt";
            listTarifData.ExceptWith(new HashSet<string>(arrayTarif));
            if (listTarifData.Count > 0)
            {
                int i = 0;
                StringBuilder sb = new StringBuilder(String.Empty);
                DateTime localDate = DateTime.Now;

                strNewModels = "";
                try
                {
                    if (File.Exists(pathToNewModels))
                    { File.Delete(pathToNewModels); }
                    sb.AppendLine(@"; This VodafoneInvoiceModifier.ini for " + myFileVersionInfo.ProductName);
                    sb.AppendLine(@"; " + @"Author " + myFileVersionInfo.LegalCopyright);
                    sb.AppendLine(@"");
                    sb.AppendLine(@"; Дата обновления файла:  " + localDate.ToString());
                    sb.AppendLine(@";");
                    sb.AppendLine(@"; Найдены новые не учтенные модели компенсации затрат сотрудников:");
                    sb.AppendLine(@"");
                    sb.AppendLine(@"");

                    foreach (string str in listTarifData)
                    {
                        i++;
                        strNewModels += i + ". \"" + str + "\"\n";
                        sb.AppendLine(i + ". \"" + str + "\"");
                    }
                    sb.AppendLine(@"");

                    File.WriteAllText(pathToNewModels, sb.ToString(), Encoding.GetEncoding(1251));
                    File.AppendAllText(pathToNewModels, sbError.ToString(), Encoding.GetEncoding(1251));
                }
                catch (Exception Expt)
                { MessageBox.Show(Expt.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                finally
                { sb = null; }
                infoStatusBar = "В базе найдены новые, не добавленные ранее, модели компенсации затрат сотрудников!";

                DialogResult result = MessageBox.Show(
                    "В базе найдены новые не учтенные модели компенсации затрат сотрудников!\n\n" + strNewModels +
                    "\n\nДля их учета необходимо, внести изменения в модели рассчета в программе!\n\n" +
                    "Для прерывания дальнейших рассчетов нажмите кнопку\n\"Yes\"(Да)\nдля продолжения:\n\"No\"(Нет)",
                    "Внимание!",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button1);
                if (result == DialogResult.Yes)
                { newModels = true; }
            }
        }


        //Access to Control from other threads
        private string _OpenFileDialogReturnPath(OpenFileDialog ofd) //Return its name 
        {
            string filePath = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    ofd.FileName = @"";
                    ofd.Filter = "Текстовые файлы (*.txt)|*.txt|All files (*.*)|*.*";
                    ofd.ShowDialog();
                    filePath = ofd.FileName;
                }));
            else
            {
                ofd.FileName = @"";
                ofd.Filter = "Текстовые файлы (*.txt)|*.txt|All files (*.*)|*.*";
                ofd.ShowDialog();
                filePath = ofd.FileName;
            }
            return filePath;
        }

        private string _ControlReturnItsText(Control controlText) //Return its name 
        {
            string tBox = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tBox = controlText.Text.Trim(); }));
            else
                tBox = controlText.Text.Trim();
            return tBox;
        }

        private void _ControlSetItsText(Control control, string text) //Set its name 
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.Text = text; }));
            else
                control.Text = text;
        }

        private void _ControlVisibleEnabled(Control control, bool visible) //Set its name 
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.Visible = visible; }));
            else
                control.Visible = visible;
        }

        private void _ToolStripStatusLabelSetItsText(ToolStripStatusLabel control, string text) //Set its name 
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.Text = text; }));
            else
                control.Text = text;
        }

      /*  private void _ToolStripMenuItemVisibleEnabled(ToolStripMenuItem control, bool visible) //Set its name 
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.Visible = visible; }));
            else
                control.Visible = visible;
        }*/

        private void _ToolStripMenuItemEnabled(ToolStripMenuItem control, bool enabled) //Set its name 
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.Enabled = enabled; }));
            else
                control.Enabled = enabled;
        }

        private void _TextboxAppendText(TextBox control, string text) //Set its name 
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.AppendText(text); }));
            else
                control.AppendText(text);
        }

        private void _TextboxClear(TextBox control) //Set its name 
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.Clear(); }));
            else
                control.Clear();
        }




        //Save and Recover Data in Registry

        public void ListsRegistryDataCheck() //Read previously Saved Parameters from Registry
        {
            listSavedServices = new List<string>();
            listSavedNumbers = new List<string>();
            string[] getValue;

            try
            {
                using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                      myRegKey,
                      Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree,
                      System.Security.AccessControl.RegistryRights.ReadKey))
                {
                    getValue = (string[])EvUserKey.GetValue("ListServices");

                    try
                    {
                        foreach (string line in getValue)
                        {
                            listSavedServices.Add(line.Trim());
                        }
                        foundSavedData = true;

                    }
                    catch (Exception exct) { textBoxLog.AppendText("\n" + exct.ToString() + "\n"); }

                    getValue = (string[])EvUserKey.GetValue("ListNumbers");

                    try
                    {
                        foreach (string line in getValue)
                        {
                            listSavedNumbers.Add(line.Trim());
                        }
                        _ControlSetItsText(labelContracts, listSavedNumbers.Count.ToString() + " шт.");
                        _ControlVisibleEnabled(labelContracts, true);
                        foundSavedData = true;

                    }
                    catch (Exception exct) { textBoxLog.AppendText("\n" + exct.ToString() + "\n"); }

                    try
                    {
                        strSavedPathToInvoice = (string)EvUserKey.GetValue("PathToLastInvoice");

                        if (strSavedPathToInvoice.Trim().Length > 3)
                        { prepareBillItem.Enabled = true; }

                        string period = (string)EvUserKey.GetValue("PeriodLastInvoice");

                        if (period.Length > 6)
                        {
                            _ControlSetItsText(labelPeriod, period);
                            _ControlVisibleEnabled(labelPeriod, true);
                        }
                    }
                    catch (Exception exct) { textBoxLog.AppendText("\n" + exct.ToString() + "\n"); }

                    try
                    {
                        textBoxLog.AppendText("Настроенные данные для сбора данных для маркетинга:\n");
                        textBoxLog.AppendText("===================================================\n\n");

                        if (listSavedServices.Count > 0)
                        {
                            selectedServices = true;
                            textBoxLog.AppendText("Загруженный список сервисов:\n");
                            foreach (string service in listSavedServices)
                            { textBoxLog.AppendText(service + "\n"); }
                            textBoxLog.AppendText("===================================================\n\n");
                        }

                        if (listSavedNumbers.Count > 0)
                        {
                            selectedNumbers = true;
                            textBoxLog.AppendText("Загруженный список номеров:\n");
                            foreach (string number in listSavedNumbers)
                            { textBoxLog.AppendText(number + "\n"); }
                            textBoxLog.AppendText("===================================================\n");
                        }
                        textBoxLog.AppendText("===================================================\n\n");

                    }
                    catch { }
                }
            }
            catch (Exception exct) { textBoxLog.AppendText("\n" + exct.ToString() + "\n"); }
        }

        public void ListServicesRegistrySave() //Save Parameters into Registry and variables
        {
            if (listServices != null && listServices.Count > 0)
            {
                try
                {
                    using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                    {
                        EvUserKey.SetValue("ListServices", listServices.ToArray(),
                            Microsoft.Win32.RegistryValueKind.MultiString);
                    }
                    foundSavedData = true;
                }
                catch { MessageBox.Show("Ошибки с доступом для записи списка сервисов в реестр. Данные сохранены не корректно."); }
            }
        }

        public void ListNumbersRegistrySave() //Save inputed Credintials and Parameters into Registry and variables
        {
            try
            {
                using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                {
                    EvUserKey.SetValue("ListNumbers", listNumbers.ToArray(),
                        Microsoft.Win32.RegistryValueKind.MultiString);
                }
                foundSavedData = true;
            }
            catch { MessageBox.Show("Ошибки с доступом для записи списка номеров в реестр. Данные сохранены не корректно."); }
        }

        public void ParameterLastInvoiceRegistrySave() //Save Parameters into Registry and variables
        {
            try
            {
                using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                {
                    if (filepathLoadedData != null && filepathLoadedData.Length > 0)
                    { EvUserKey.SetValue("PathToLastInvoice", filepathLoadedData, Microsoft.Win32.RegistryValueKind.String); }

                    if (_ControlReturnItsText(labelPeriod).Length > 0)
                    { EvUserKey.SetValue("PeriodLastInvoice", periodInvoice, Microsoft.Win32.RegistryValueKind.String); }
                }
                foundSavedData = true;
            }
            catch { MessageBox.Show("Ошибки с доступом для записи пути к счету. Данные сохранены не корректно."); }
        }
    }

    class MobileContractPerson
    {
        public string ownerName = "";
        public string contractName = "";
        public string mobNumberName = "";
        public string tarifPackageName = "";
        public double monthCost = 0;
        public double roming = 0;
        public double discount = 0;
        public double totalCost = 0;
        public double tax = 0;
        public double pF = 0;
        public double totalCostWithTax = 0;
        public double romingData = 0;
        public double extraServiceOrdered = 0;
        public double extraInternetOrdered = 0;
        public double outToCity = 0;
        public double extraService = 0;
        public double content = 0;
        public string dateBillStart = "";
        public string dateBillEnd = "";

        public string NAV = "";
        public string orgUnit = "";
        public string startDate;
        public string modelCompensation = "";
        public double payOwner = 0;
        public bool isUsed = false;
        public bool isUnblocked = false;

        public string toString()
        {
            return dateBillStart + "|" +
                   ownerName + "|" +
                   NAV + "|" +
                   contractName + "|" +
                   mobNumberName + "|" +
                   tarifPackageName + "|" +
                   monthCost + "|" +
                   roming + "|" +
                   discount + "|" +
                   totalCost + "|" +
                   tax + "|" +
                   pF + "|" +
                   totalCostWithTax + "|" +
                   romingData + "|" +

                   outToCity + "|" +
                   extraService + "|" +
                   content + "|" +
                   orgUnit + "|" +
                   modelCompensation + "|" +
                   extraServiceOrdered + "|" +
                   extraInternetOrdered + "|" +
                   payOwner + "|" +
                   isUsed + "|" +
                   isUnblocked
                   ;
        }

        public string toStringName()
        {
            return "dateBillStart|" +
                   "ownerName|" +
                   "NAV|" +
                   "contractName|" +
                   "mobNumberName|" +
                   "tarifPackageName|" +
                   "monthCost|" +
                   "roming|" +
                   "discount|" +
                   "totalCost|" +
                   "tax|" +
                   "pF|" +
                   "totalCostWithTax|" +
                   "romingData|" +
                   "outToCity|" +
                   "extraService|" +
                   "content|" +
                   "orgUnit|" +
                   "modelCompensation|" +
                   "extraServiceOrdered|" +
                   "extraInternetOrdered|" +
                   "payOwner|" +
                   "isUsed|" +
                   "isUnblocked"
                   ;
        }
    }
}
