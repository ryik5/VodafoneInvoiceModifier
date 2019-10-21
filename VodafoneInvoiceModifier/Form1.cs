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
using System.Drawing;

namespace BillReportsGenerator
{
    public partial class Form1 : Form
    {
        System.Diagnostics.FileVersionInfo myFileVersionInfo;
        ContextMenu contextMenu1;
        string myRegKey;
        string pathToIni; //path to ini of tools

        string pStop = @"ЗАГАЛОМ ЗА ВСІМА КОНТРАКТАМИ";

        //Скидка на счет = pBillDeliveryCostDiscount/pBillDeliveryCost  (в процентах)
        string pBillDeliveryCost = @"Інші послуги на особовому рахунку"; //Стоимость услуги доставки электронного счета
        double BillDeliveryCost = 0; //Стоимость услуги доставки электронного счета
        string pBillDeliveryCostDiscount = @"Знижка на суму особового рахунку"; //Скидка на стоимость услуги по доставке электронного счета
        double BillDeliveryCostDiscount = 0; //Скидка на стоимость услуги по доставке электронного счета

        string about = "";
        string dataStart = ""; // дата начала периода счета
        string dataEnd = "";  // дата конца периода счета
        string periodInvoice = ""; //Период
        bool checkRahunok = false;
        bool checkNomerRahunku = false;
        bool checkPeriod = false;

        //  private string pConnection = ""; //string connection to MS SQL DB
        string pConnectionServer = ""; //string connection to MS SQL DB
        string pConnectionUserName = ""; //string connection to MS SQL DB
        string pConnectionUserPasswords = ""; //string connection to MS SQL DB
        const string NUMBER_OF_CONTRACT = @"Контракт №";
        const string MOBILE_NUMBER = @"Моб.номер";
        const string NAME_OF_TARIF = @"Ціновий Пакет";
        readonly string[] p = new string[] //Features of the mobile contract and db that have the values
       {
            // со счета
            @"Владелец",                                        //0     //owner
            NUMBER_OF_CONTRACT,                                    //1     //number of contract
            MOBILE_NUMBER,                                             //2     //number
            NAME_OF_TARIF,                                      //3     //name of tarif package
            @"ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА",                 //4     //cost of package
            @"ПОСЛУГИ МІЖНАРОДНОГО РОУМІНГУ",                   //5     //rouming
            @"ЗНИЖКИ",                                          //6     //discount
            @"ЗАГАЛОМ ЗА КОНТРАКТОМ (БЕЗ ПДВ ТА ПФ)",           //7     //total without tax and pf
            @"ПДВ",                                             //8     //Tax
            @"ПФ",                                              //9     //PF
            @"Загалом з податками",                             //10    //total with tax and pf
            @"GPRS/CDMA з'єд.  Роумінг",                        //11    //GPRS in rouming
            @"Передача даних - вартість пакету послуг",         //12    //transmission of data. cost of package
            @"Вихідні дзвінки  Міські номери",                  //13    //outgoing to city numbers
            @"ПОСЛУГИ, НАДАНІ ЗА МЕЖАМИ ПАКЕТА",                //14    //services outside the package
            @"НАДАНІ КОНТЕНТ-ПОСЛУГИ",                          //15    //content services
            @"Дата счета",                                      //16    //Invoice date
            @"Дата кінця періоду",                              //17    //Date of the end of period
            // из базы
            @"Таб. номер",                                      //18    //staff number
            @"Отдел",                                           //19    //department
            @"Действует c",                                     //20    //doing since
            @"Модель",                                          //21    //model
            @"Оплата владельцем",                               //22    //paid by owner
            // со счета
            @"ПОСЛУГИ ЗА МЕЖАМИ ПАКЕТА",                        //23    //services outside of the package
            // анализ
            @"Контракт использовался",                          //24    //contract was used
            @"Контракт не заблокирован",                        //25    //contract is not blocked         
            // доп.признаки строк
            @"Вх",                                           //26       //ingoing
            @"Вих",                                         //27        //outgoing
            @"Переадр",                                         //28    //redirected
            @"GPRS",                                        //29        //GPRS
            @"CDMA"                                        //30         //CDMA
       };
        readonly string[] pTranslate = new string[]
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
        readonly string[] pToAccount = new string[]
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
        StringBuilder sbError = new StringBuilder();
        DataTable dtMobile = new DataTable("MobileData");
      readonly  DataColumn[] dcMobile ={
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

        List<string> listTempContract = new List<string>();

        // из базы
        //  private List<string> lTarifData = new List<string>();
        DataTable dtOwnerOfMobileWithinSelectedPeriod = new DataTable("TarifListData");
      readonly  DataColumn[] dcTarif ={
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

        HashSet<string> listTarifData = new HashSet<string>(); //will write models in modelToPayment()

        readonly string[] arrayTarif = new string[] {
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

        string infoStatusBar = "";
        bool newModels = false; //stop calculating data
        string strNewModels = "";

        string filePathTxt; //path to the selected bill

        List<string> listNumbers = new List<string>(); //list of numbers for the marketing report
        List<string> listServices = new List<string>();//list of services for the marketing report

        string parametrStart = "Контракт";

        //скидка в текущем счете
        double resultOfCalculatingDiscount = 30;
        double amountBillAfterDiscount = 0.70; //  = 1 - (resultOfCalculatingDiscount / 100)


        bool loadedBill = false;
        bool selectedServices = false;
        bool selectedNumbers = false;

       readonly DataColumn[] dcFullBill ={
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
        DataTable dtMarket = new DataTable("MarketReport");

        List<string> listSavedServices = new List<string>();
        List<string> listSavedNumbers = new List<string>();
        string filepathLoadedData = "";  //current path to invoice
        string strSavedPathToInvoice = "";  //previous session path to invoice
        bool foundSavedData = false;


        public Form1()
        { InitializeComponent(); }

        private void Form1_Load(object sender, EventArgs e)
        {
            myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath);

            myRegKey = @"SOFTWARE\RYIK\" + myFileVersionInfo.ProductName;
            pathToIni = Application.StartupPath + @"\" + myFileVersionInfo.ProductName + ".ini"; //path to ini of tools

            about = myFileVersionInfo.Comments + " ver." + myFileVersionInfo.FileVersion + " " + myFileVersionInfo.LegalCopyright;
            StatusLabel1.Text = myFileVersionInfo.ProductName + " ver." + myFileVersionInfo.FileVersion + " " + myFileVersionInfo.LegalCopyright;
            StatusLabel1.Alignment = ToolStripItemAlignment.Right;

            notifyIcon1.Text = myFileVersionInfo.ProductName + " " + myFileVersionInfo.LegalCopyright;
            notifyIcon1.BalloonTipText = about;

            contextMenu1 = new ContextMenu();  //Context Menu on notify Icon
            notifyIcon1.ContextMenu = contextMenu1;
            contextMenu1.MenuItems.Add(Properties.Resources.About, AboutSoft);
            contextMenu1.MenuItems.Add(Properties.Resources.Exit, ApplicationExit);
            notifyIcon1.Text = myFileVersionInfo.ProductName + "\nv." + myFileVersionInfo.FileVersion + "\n" + myFileVersionInfo.CompanyName;
            this.Text = myFileVersionInfo.Comments;
            ProgressBar1.Value = 0;

            groupBox1.BackColor = System.Drawing.Color.Ivory;

            labelAccount.Visible = false;
            labelPeriod.Visible = false;
            labelBill.Visible = false;
            labelContracts.Visible = false;
            ReadStringsWithParametersFromIniFile();

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
            dtOwnerOfMobileWithinSelectedPeriod.Columns.AddRange(dcTarif);
          //  dtFullBill.Columns.AddRange(dcFullBill);
            dtMarket.Columns.AddRange(dcFullBill);
            ListsRegistryDataCheck();
            useSavedDataItem.Enabled = foundSavedData;

        }


        private void AboutSoft()
        {
            string strVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

            MessageBox.Show(
                myFileVersionInfo.Comments + "\n\nВерсия: " + myFileVersionInfo.FileVersion + "\nBuild: " +
                strVersion + "\n" + myFileVersionInfo.LegalCopyright,
             Properties.Resources.InfoApp,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1);
        }

        private void ApplicationExit()
        {
            WriteStringsWithParametersIntoIniFile();
            contextMenu1?.Dispose();
            dtMarket?.Dispose();
            dtMobile?.Dispose();
            dtOwnerOfMobileWithinSelectedPeriod?.Dispose();
            Application.Exit();
        }

        private void openBillItem_Click(object sender, EventArgs e)//Menu "Open"
        {
            OpenBill();
        }

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
            string strTemp;
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
                    strTemp = MakeCommonFormFromPhoneNumber(s);

                    if (strTemp.Length == 13)  //check Length of a number == 13
                    { listNumbers.Add(strTemp); }
                    else
                    { listWrongString.Add(strTemp); }
                }

                if (0 < listWrongString.Count)
                {
                    textBoxLog.AppendText("List of first 300 wrong rows in the selected list:\n");
                    textBoxLog.AppendText(Properties.Resources.RowDashedLinesBefore2NewLines);
                    int wrongRow = 0;
                    foreach (string s in listWrongString)
                    {
                        textBoxLog.AppendText(s + "\n");
                        wrongRow++;

                        if (wrongRow > limitWrongNumber)
                        { break; }
                    }
                    textBoxLog.AppendText("\n\n");
                }

                if (0 < listNumbers?.Count && listNumbers?.Count < 500)
                {
                    selectedNumbers = true;
                    SaveListStringsInRegistry("ListNumbers", listNumbers);
                    _ControlSetItsText(labelContracts, listNumbers.Count.ToString() + " шт.");
                    _ControlVisibleEnabled(labelContracts, true);

                    textBoxLog.AppendText("List of numbers:\n");
                    textBoxLog.AppendText(Properties.Resources.RowDashedLinesBefore2NewLines);

                    foreach (string s in listNumbers)
                    { textBoxLog.AppendText(s + "\n"); }
                }
                else
                { textBoxLog.AppendText("Check the list of numbers.\nIn the list was found: " + listNumbers.Count + " number(s)"); }
            }
            CheckConditionEnableMarketingReport();
        }

        private void selectListServicesItem_Click(object sender, EventArgs e)
        { PrepareListServicesToMakeReport(); }

        //limit of services <100
        private void PrepareListServicesToMakeReport() //Prepare list of services for the marketing report - listServices
        {
            selectedServices = false;
            makeReportMarketingItem.Enabled = false;
            textBoxLog.Clear();

            listServices.Clear();
            listServices = LoadDataIntoList();

            if (0 < listServices?.Count && listServices?.Count < 100)
            {
                textBoxLog.AppendText("\nList of services:\n");

                foreach (string s in listServices)
                { textBoxLog.AppendText(s + "\n"); }

                selectedServices = true;

                SaveListStringsInRegistry("ListServices", listServices);
            }
            else
            {
                textBoxLog.AppendText(Properties.Resources.RowDashedLinesBetween2NewLines);
                textBoxLog.AppendText("The selected list is wrong!\nWill check the file!\nIt has to contain from 1 to 100 services.");
                textBoxLog.AppendText(Properties.Resources.RowDashedLinesBetween2NewLines);
            }
            CheckConditionEnableMarketingReport();
        }
        
        private async void prepareBillItem_Click(object sender, EventArgs e)
        {
            dtMarket.Rows.Clear();
            await Task.Run(() => LoadBillIntoMemoryToFilter());
        }

        private void LoadBillIntoMemoryToFilter()
        {
            _ProgressBar1Start();
            _TextboxClear(textBoxLog);
            _ToolStripMenuItemEnabled(fileMenuItem, false);
            _ControlVisibleEnabled(labelPeriod, true);

            loadedBill = false;
            WrittingOfTextAtFile textWritting = new WrittingOfTextAtFile();

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

            string exceptedStringContains = @". . .";
            // NUMBER_OF_CONTRACT,       //1     //number of contract
            // MOBILE_NUMBER,           //2     //number
            // NAME_OF_TARIF,            //3     //name of tarif package

            p[1] = _ControlReturnText(textBoxP1);
            p[2] = _ControlReturnText(textBoxP2);

            List<string> filterBill = new List<string>();
            filterBill.Add(p[1]);
            filterBill.Add(p[2]);

            if (listServices.Count == 0)
            { listServices = listSavedServices; }
            if (listNumbers.Count == 0)
            { listNumbers = listSavedNumbers; }

            _ProgressWork1Step();

            foreach (string service in listServices)
            {
                filterBill.Add(service);
            }

            _ProgressWork1Step();

            List<string> loadedBillWithServicesFiltered = LoadDataUsingParameters(filterBill, parametrStart, pStop, exceptedStringContains);

            textWritting.Write(Path.GetDirectoryName(filepathLoadedData) + @"\selectedRows.csv", loadedBillWithServicesFiltered);
            
            int allRow = loadedBillWithServicesFiltered.Count * listServices.Count * (dtOwnerOfMobileWithinSelectedPeriod.Rows.Count + listNumbers.Count); //всего строк для борабоки

            int counterstep = (dtOwnerOfMobileWithinSelectedPeriod.Rows.Count + listNumbers.Count);
            int countStepProgressBar = counterstep;
            int countRowsInTable = 0;

            if (loadedBillWithServicesFiltered?.Count > 0)
            {
              //  dtFullBill.Rows.Clear();
                StringBuilder sb = new StringBuilder();

                //todo parsing strings of the filtered bill
                foreach (string sRowBill in loadedBillWithServicesFiltered)
                {
                    if (sRowBill.StartsWith(p[1]))
                    {
                        try
                        {
                            kontrakt = Regex.Split(sRowBill.Substring(sRowBill.IndexOf('№') + 1).Trim(), " ")[0].Trim();
                            tempRow = sRowBill.Substring(sRowBill.IndexOf(':') + 1).Trim();

                            if (tempRow.StartsWith("+"))
                            { numberMobile = tempRow; }
                            else
                            { numberMobile = "+" + tempRow; } //set format number like '+380...'
                        }
                        catch
                        {
                            MessageBox.Show("Проверьте правильность выбора файла с контрактами с детализацией разговоров!\n" +
                                "Возможно поменялся формат.\n" +
                                "Правильный формат начала каждого контракта:\n" +
                                NUMBER_OF_CONTRACT + " 000000000  _номер_: 380000000000\n\n" +
                                "Данная строка с началом разбираемого контракта имеет вид:\n" +
                                sRowBill
                                );
                        }
                    }
                    else
                    {
                        //parse a string of the contract 
                        //start position of a symbol and last one in the parse string 
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
                            serviceName = sRowBill?.Substring(0, 38)?.Trim();
                            numberB = sRowBill?.Substring(38, 13)?.Trim();
                            date = sRowBill?.Substring(52, 10)?.Trim();
                            time = sRowBill?.Substring(65, 8)?.Trim();
                            durationA = sRowBill?.Substring(74, 9)?.Trim();
                            durationB = sRowBill?.Substring(84, 9)?.Trim();
                            cost = sRowBill?.Substring(95)?.Trim();
                            
                            foreach (DataRow rowTarif in dtOwnerOfMobileWithinSelectedPeriod.Rows)
                            {
                                if (rowTarif["Номер телефона"].ToString().Contains(numberMobile))
                                {
                                    fio = rowTarif["ФИО"].ToString();
                                    nav = rowTarif["NAV"].ToString();
                                    department = rowTarif["Подразделение"].ToString();
                                    break;
                                }

                            }

                            tempRow = numberMobile + "\t" + fio + "\t" + nav + "\t" + department + "\t" + serviceName + "\t" + numberB + "\t" + date + "\t" + time + "\t" + durationA + "\t" + durationB + "\t" + cost;

                            foreach (string sNumber in listNumbers)
                            {
                                if (tempRow.StartsWith(sNumber))
                                {
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

                                    rowMarket["ФИО"] = fio;
                                    rowMarket["NAV"] = nav;
                                    rowMarket["Подразделение"] = department;
                                    
                                    dtMarket.Rows.Add(rowMarket);
                                    /*
                                    listParsedStrings.Add(
                                        new ParsedStringOfBillWithContractOwner
                                            {
                                            contract = kontrakt,
                                            numberOwner = numberMobile,
                                            serviceName = serviceName,
                                            numberTarget = numberB,
                                            date = date,
                                            time = time,
                                            durationA = durationA,
                                            durationB = durationB,
                                            cost = cost,
                                            fio = fio,
                                            nav = nav,
                                            department = department
                                            }
                                        ); 
                                    */
                                    countRowsInTable++;
                                    sb.AppendLine(tempRow);
                                    break;
                                }

                                countStepProgressBar--;
                                if (countStepProgressBar == 0)
                                {
                                    string s = String.Format("В отчет добавлено {0,20 }, строк из {1,10}", countRowsInTable, loadedBillWithServicesFiltered.Count);
                                    _ProgressWork1Step(s);
                                    countStepProgressBar = counterstep;
                                }
                            }

                        }
                        catch (Exception expt) { MessageBox.Show(sRowBill + "\n" + expt.ToString(), expt.Message); }

                    }
                }
                loadedBill = true;
                {
                    _TextboxAppendText(textBoxLog, "\n");
                    _TextboxAppendText(textBoxLog, "Сформировано для генерации отчета " + countRowsInTable + " строк номерами мобильных подпадающими под фильтр.");
                    _TextboxAppendText(textBoxLog, "\n");
                }

                textWritting.Write(Path.GetDirectoryName(filepathLoadedData) + @"\listMarketingCollectRows.csv", sb.ToString());
            }
            else
            { _TextboxAppendText(textBoxLog, "Нет в выборке ничего для указанных номеров!\n"); }

            CheckConditionEnableMarketingReport();
            _ToolStripStatusLabelSetText(StatusLabel1, "Файл сохранен в папку: " + Path.GetDirectoryName(filepathLoadedData));

            _ToolStripMenuItemEnabled(fileMenuItem, true);
            _ProgressBar1Stop();
        }

        private void makeReportMarketingItem_Click(object sender, EventArgs e)
        { MakeExcelReport(ExportMarketReport); }

        //Заполнение таблицы в Excel  данными
        private void ExportMarketReport()
        {
            ExportDatatableToExcel(dtMarket, "_Marketing.xlsx");
        }

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
                        _ToolStripStatusLabelSetText(StatusLabel1, "Обрабатываю файл:  " + filepathLoadedData);
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

        private List<string> LoadDataUsingParameters(List<string> listParameters, string startStringLoad, string endStringLoad, string excepted) //max List length = 500 000 rows
        {
            checkRahunok = false;
            checkNomerRahunku = false;
            checkPeriod = false;
            int? countParameters = listParameters?.Count;
            int countStepProgressBar = 500;
            int listMaxLength = 500000;
            List<string> listRows = new List<string>(listMaxLength);
            string loadedString = "";
            bool oldSavedInvoice = strSavedPathToInvoice?.Length > 2 ? true : false;
            bool currentInvoice = filepathLoadedData?.Length > 2 ? true : false;
            try
            {
                bool startLoadData = false;
                bool endLoadData = false;
                var Coder = Encoding.GetEncoding(1251);
                if (countParameters > 0)
                {
                    if (oldSavedInvoice)
                    {
                        DialogResult result = MessageBox.Show(
                              "Использовать предыдущий выбор файла?\n" + strSavedPathToInvoice,
                              "Внимание!",
                              MessageBoxButtons.YesNo,
                              MessageBoxIcon.Exclamation,
                              MessageBoxDefaultButton.Button1);
                        if (result == DialogResult.No)
                        {
                            filepathLoadedData = _OpenFileDialogReturnPath(openFileDialog1);
                        }
                        else
                        {
                            filepathLoadedData = strSavedPathToInvoice;
                        }
                    }
                    else if (!currentInvoice)
                    {
                        filepathLoadedData = _OpenFileDialogReturnPath(openFileDialog1);
                    }

                    if (filepathLoadedData?.Length > 2 && File.Exists(filepathLoadedData))
                    {
                        _ToolStripStatusLabelSetText(StatusLabel1, "Обрабатываю файл:  " + filepathLoadedData);
                        int counter = 0;
                        try
                        {
                            using (StreamReader Reader = new StreamReader(filepathLoadedData, Coder))
                            {
                                while ((loadedString = Reader.ReadLine()?.Trim()) != null && !endLoadData && listRows.Count < listMaxLength)
                                {
                                    //Set label Date
                                    if (loadedString.Contains("Особовий рахунок")) { checkRahunok = true; }
                                    if (loadedString.Contains("Номер рахунку")) { checkNomerRahunku = true; }
                                    if (loadedString.Contains("Розрахунковий період"))
                                    {
                                        string[] substrings = Regex.Split(loadedString, ": ");
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
                                            if (loadedString.StartsWith(parameterString)&&!loadedString.Contains(excepted))
                                            {
                                                listRows.Add(loadedString);
                                                counter++;
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
                        _TextboxAppendText(textBoxLog, "\n");
                        _TextboxAppendText(textBoxLog, "Из файла-счета: \n");
                        _TextboxAppendText(textBoxLog, filepathLoadedData);
                        _TextboxAppendText(textBoxLog, "\n");
                        _TextboxAppendText(textBoxLog, "отобрано для построения отчета " + counter + " строк с требуемыми сервисами");
                        _TextboxAppendText(textBoxLog, "\n");
                        if (listMaxLength - 2 < listRows.Count || listRows.Count == 0)
                        { MessageBox.Show("Error was happened on " + (listRows.Count) + " row\n You've been chosen the long file!"); }
                    }
                    else { MessageBox.Show("Did not select File!"); }
                }
            }
            catch (Exception expt) { MessageBox.Show(expt.ToString()); }
            return listRows;
        }

        private void UseSavedDataItem_Click(object sender, EventArgs e)
        {
            if (strSavedPathToInvoice?.Length > 1)
            { filepathLoadedData = strSavedPathToInvoice; }
            else { strSavedPathToInvoice = ""; }

            if (listSavedNumbers?.Count > 0)
            { listNumbers = listSavedNumbers; }

            if (listSavedServices?.Count > 0)
            { listServices = listSavedServices; }

            if (listSavedNumbers?.Count > 0 && listSavedServices?.Count > 0)
            { prepareBillItem.Enabled = true; }
        }

        private async void OpenBill()
        {
            dtMobile.Rows.Clear();
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
            p[1] = textBoxP1.Text;
            p[2] = textBoxP2.Text;
            p[3] = textBoxP3.Text;
            p[4] = textBoxP4.Text;
            p[5] = textBoxP5.Text;
            p[6] = textBoxP6.Text;
            p[7] = textBoxP7.Text;
            pStop = textBoxP8.Text;

            StatusLabel1.Text = "Обрабатываю исходные данные...";
            bool billCorrect = TryToReadBillToPrepareList();

            if (billCorrect)
            {
                StatusLabel1.Text = "Получаю данные с базы Tfactura...";

                await Task.Run(() => GetDataWithModel());
                if (dtOwnerOfMobileWithinSelectedPeriod.Rows.Count < 2)
                {
                    MessageBox.Show("Выбранный счет в базу данных Tfactura еще не импортирован!\nПеред обработкой счета, предварительно необходимо импортировать счет в базу!");
                    StatusLabel1.Text = "Обработка счета прекращена! Предварительно импортируйте счет в Tfactura!";
                    StatusLabel1.BackColor = System.Drawing.Color.SandyBrown;
                }
                else
                {
                    await Task.Run(() => CheckNewTarif());

                    //clear log if it was found a problem
                    if (listTarifData.Count > 0)
                    { textBoxLog.Clear(); }

                    if (!newModels)
                    {
                        ParseStringsOfPreparedListIntoTable();
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
                        textBoxLog.AppendText("-= Дата счета:  " + dtMobile.Rows[1][16].ToString() + " =-"); //Дата счета
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("====================================================\n");
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("\n");


                        //////////////////////////////
                        if (listTarifData.Count > 0)
                        {
                            textBoxLog.AppendText("-= Список тарифных схем, не существующих в программе =-\n");
                            textBoxLog.AppendText("'" + columnName5 + "' - " + columnName1 + " (" + columnName2 + ")\n");

                            foreach (string str in listTarifData)
                            {
                                textBoxLog.AppendText(str + "\n");
                            }
                            textBoxLog.AppendText("\n");
                            textBoxLog.AppendText(Properties.Resources.RowDashedLinesBetween2NewLines);
                        }

                        /////////////////
                        results = dtMobile.Select("NumberUsed='False' AND NumberNoBlock='True'", sortOrder, DataViewRowState.Added);
                        if (results.Length > 0)
                        {
                            textBoxLog.AppendText("-= Список контрактов, по которым не велась работа =-\n");
                            textBoxLog.AppendText(
                                 string.Format("{0,-40}", columnName1) +
                                 string.Format("{0,-15}", columnName2) +
                                 string.Format("{0,-30}", columnName3) +
                                 string.Format("{0,-10}", columnName4) +
                                 string.Format("{0,-30}", columnName5) +
                                 "\n");
                            for (int i = 0; i < results.Length; i++)
                            {

                                textBoxLog.AppendText(
                                 string.Format("{0,-40}", results[i][0].ToString()) +
                                 string.Format("{0,-15}", results[i][2].ToString()) +
                                 string.Format("{0,-30}", results[i][3].ToString()) +
                                 string.Format("{0,-10}", results[i][10].ToString()) +
                                 string.Format("{0,-30}", results[i][21].ToString()) +
                                 "\n"
                                  );
                            }
                            textBoxLog.AppendText("\n");
                            textBoxLog.AppendText(Properties.Resources.RowDashedLinesBetween2NewLines);
                        }

                        /////////////////
                        results = dtMobile.Select("NumberNoBlock='False'", sortOrder, DataViewRowState.Added);
                        if (results.Length > 0)
                        {
                            textBoxLog.AppendText("-= Список заблокированных контрактов =-\n");
                            textBoxLog.AppendText(
                                 string.Format("{0,-40}", columnName1) +
                                 string.Format("{0,-15}", columnName2) +
                                 string.Format("{0,-30}", columnName3) +
                                 string.Format("{0,-10}", columnName4) +
                                 string.Format("{0,-30}", columnName5) +
                                 "\n");
                            for (int i = 0; i < results.Length; i++)
                            {
                                textBoxLog.AppendText(
                                 string.Format("{0,-40}", results[i][0].ToString()) +
                                 string.Format("{0,-15}", results[i][2].ToString()) +
                                 string.Format("{0,-30}", results[i][3].ToString()) +
                                 string.Format("{0,-10}", results[i][10].ToString()) +
                                 string.Format("{0,-30}", results[i][21].ToString()) +
                                 "\n"
                                  );
                            }
                            textBoxLog.AppendText("\n");
                            textBoxLog.AppendText(Properties.Resources.RowDashedLinesBetween2NewLines);
                        }

                        /////////////////
                        textBoxLog.AppendText("---= Все =---\n");
                        results = dtMobile.Select(dtMobile.Columns[0].ColumnName.Length + " > 0", sortOrder, DataViewRowState.Added);
                        textBoxLog.AppendText(
                             string.Format("{0,-40}", columnName1) +
                             string.Format("{0,-15}", columnName2) +
                             string.Format("{0,-30}", columnName3) +
                             string.Format("{0,-10}", columnName4) +
                             string.Format("{0,-10}", columnName6) +
                             string.Format("{0,-30}", columnName5) +
                             string.Format("{0,-12}", columnName10) +
                             string.Format("{0,-12}", columnName11) +
                             "\n");
                        for (int i = 0; i < results.Length; i++)
                        {

                            textBoxLog.AppendText(
                             string.Format("{0,-40}", results[i][0].ToString().Trim()) +
                             string.Format("{0,-15}", results[i][2].ToString()) +
                             string.Format("{0,-30}", results[i][3].ToString()) +
                             string.Format("{0,-10}", results[i][10].ToString()) +
                             string.Format("{0,-10}", results[i][5].ToString()) +

                             string.Format("{0,-30}", results[i][21].ToString()) +
                             string.Format("{0,-12}", results[i][24].ToString()) +
                             string.Format("{0,-12}", results[i][25].ToString()) +
                             "\n"
                              );
                        }
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("\n");
                        textBoxLog.AppendText("====================================================\n");
                        /////////////////


                        makeReportAccountantItem.Enabled = true;
                        makeFullReportItem.Enabled = true;

                        StatusLabel1.Text = "Предварительная обработка счета из файла " + Path.GetFileName(filePathTxt) + " завершена!";
                        StatusLabel1.ToolTipText = "Данные для генерации отчета для бухгалтерии подготовлены";
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
            else { StatusLabel1.Text = "Файл с детализацией выбран не корректно!  "; }

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
        }

        private string ParseParameterNameAndValueFromReadString(string delimeter, string parameter, string defaultValue = null)
        {
            if (parameter == null || delimeter == null)
            {
                return null;
            }

            string tempString = Regex.Split(parameter, delimeter)?[1]?.Trim();

            if (tempString?.Length > 1)
                return tempString;
            else
            {
                if (defaultValue != null)
                    return (string)defaultValue.Clone();
                else return null;
            }
        }

        private async void ReadStringsWithParametersFromIniFile() //Чтение парсеров из ini файла
        {
            string s = "", info="";
            bool b1 = false, b2 = false;
            toolTip1.SetToolTip(this.groupBox1, "Использованы исходНые настройки программы");

            if (File.Exists(pathToIni))
            {
                var Coder = Encoding.GetEncoding(1251);
                using (StreamReader Reader = new StreamReader(pathToIni, Coder))
                {
                    while ((s = Reader.ReadLine()?.Trim()) != null)
                    {
                        if (s?.Length > 3)
                        {
                            //Проверка ini файла на наличие строк с авторством
                            if (s.Contains(myFileVersionInfo.ProductName))
                            { b1 = true; }
                            else if (s.Contains(@"Author " + myFileVersionInfo.LegalCopyright))
                            { b2 = true; }

                            //Далее - обработка ini файла только с наличием авторства
                            if (b1 && b2)
                            {
                                if (s.StartsWith(nameof(pConnectionServer) + "="))
                                {
                                    pConnectionServer = ParseParameterNameAndValueFromReadString("=", s, pConnectionServer);
                                }
                                else if (s.StartsWith(nameof(pConnectionUserName) + "="))
                                {
                                    pConnectionUserName = ParseParameterNameAndValueFromReadString("=", s, pConnectionUserName);
                                }
                                else if (s.StartsWith(nameof(pConnectionUserPasswords) + "="))
                                {
                                    pConnectionUserPasswords = ParseParameterNameAndValueFromReadString("=", s, pConnectionUserPasswords);
                                }
                                else if (s.StartsWith(nameof(parametrStart) + "="))
                                {
                                    parametrStart = ParseParameterNameAndValueFromReadString("=", s, parametrStart);
                                }
                                else if (s.StartsWith(nameof(pStop) + "="))
                                {
                                    pStop = ParseParameterNameAndValueFromReadString("=", s, pStop);
                                }
                                else if (s.StartsWith(nameof(pBillDeliveryCost) + "=")) //Строка с суммой стоимости доставки электронного счета до вычисления скидки и налогов
                                {
                                    pBillDeliveryCost = ParseParameterNameAndValueFromReadString("=", s, pBillDeliveryCost);
                                }
                                else if (s.StartsWith(nameof(pBillDeliveryCostDiscount) + "="))//Строка с суммой скидки на доставку электронного счет
                                {
                                    pBillDeliveryCostDiscount = ParseParameterNameAndValueFromReadString("=", s, pBillDeliveryCostDiscount);
                                }

                                for (int i = 0; i < p?.Length; i++)
                                {
                                    if (s.StartsWith("p" + i.ToString() + "="))
                                    {
                                        p[i] = ParseParameterNameAndValueFromReadString("=", s);
                                    }
                                }
                            }
                        }
                    }
                }

                if ((b1 && b2 == false) || (b2 && b1 == false))
                {
                    info += "Настройки из " + myFileVersionInfo.ProductName + ".ini проигнорированы. Изменен формат файла\n";
                }
                else
                {
                    info += "Парсеры модифицированы настройками из " + myFileVersionInfo.ProductName + ".ini\n";
                    groupBox1.BackColor = System.Drawing.Color.Tan;
                }

                toolTip1.SetToolTip(groupBox1, info);
            }

            textBoxP1.Text = p[1];
            textBoxP2.Text = p[2];
            textBoxP3.Text = p[3];
            textBoxP4.Text = p[4];
            textBoxP5.Text = p[5];
            textBoxP6.Text = p[6];
            textBoxP7.Text = p[7];
            textBoxP8.Text = pStop;
            if (!(pConnectionServer?.Length > 1 && pConnectionUserName?.Length > 1 && pConnectionUserPasswords?.Length > 1))
            {
                infoStatusBar = "Строка подключения к базе со счетами Tfactura неверно сконфигурирована";
                info +=  infoStatusBar + "\nПроверьте и добавьте в файл с настройками -\n\n" + 
                    pathToIni + "\n\nотсутствующие данные, необходимые для подключения к базе данных:\n\n" +
                      "pConnectionServer=" + pConnectionServer + 
                      "\npConnectionUserName=" + pConnectionUserName + 
                      "\npConnectionUserPasswords=" + pConnectionUserPasswords;
                MessageBox.Show(info,
                      "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                StatusLabel1.Text = infoStatusBar;
                StatusLabel1.ToolTipText= info;

                StatusLabel1.BackColor = System.Drawing.Color.SandyBrown;
            }
            else
            {
                fileMenuItem.Enabled = false;
                StatusLabel1.Text = "Проверяю доступность БД сервера";
                StatusLabel1.BackColor = System.Drawing.Color.PaleGoldenrod;

                _ProgressBar1Start();
                string infoStatus = null, infoStatusTooltip = null;
                System.Drawing.Color infoStatusBackColor= System.Drawing.SystemColors.Menu;
                using (Timer timer1 = new Timer { Interval = 200, Enabled = true })
                {
                    timer1.Tick += new System.EventHandler(this.timer1_Tick);
                    timer1.Start();

                    bool aliveServer = true;
                    await Task.Run(() => aliveServer = CheckAliveDbServer());
                   
                    if (!aliveServer)
                    {
                        infoStatusBar = "БД сервера со счетами Tfactura не доступна";
                        info += infoStatusBar + "\nПроверьте настройки в файле с настройками -\n\n" + pathToIni + "\nи исправьте не верные данные:\n\n" +
                            "pConnectionServer=" + pConnectionServer + "\npConnectionUserName=" + pConnectionUserName + "\npConnectionUserPasswords=" + pConnectionUserPasswords;
                        MessageBox.Show(info,
                            "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        infoStatusTooltip = info;
                        infoStatus=infoStatusBar;
                        infoStatusBackColor = System.Drawing.Color.SandyBrown;
                    }
                    else
                    {
                        fileMenuItem.Enabled = true;
                        infoStatusBackColor = System.Drawing.Color.PaleGreen;
                        infoStatus = "БД сервера со счетами Tfactura доступна для генерации отчетов";
                        infoStatusTooltip = "выберите счет мобильного оператора с которым планируете работать";
                    }
                    StatusLabel1.Text = infoStatus;
                    StatusLabel1.ToolTipText = infoStatusTooltip;
                    StatusLabel1.BackColor = infoStatusBackColor;

                    timer1.Enabled = false;
                    timer1.Stop();
                }
                _ProgressBar1Stop();
            }
            StatusLabel1.ForeColor = System.Drawing.Color.Black;
        }

        private bool CheckAliveDbServer()
        {
            bool state = false;
            string pConnection =
                "Data Source=" + pConnectionServer +
                "; Initial Catalog=EBP; Type System Version=SQL Server 2005; Persist Security Info =True" +
                "; User ID=" + pConnectionUserName +
                "; Password=" + pConnectionUserPasswords +
                "; Connect Timeout=5";

            string sqlQuery = @"SELECT database_id FROM sys.databases WHERE Name ='EBP'";
            using (var sqlConnection = new System.Data.SqlClient.SqlConnection(pConnection))
            {
                try
                {
                    sqlConnection.Open();

                    using (var sqlCommand = new System.Data.SqlClient.SqlCommand(sqlQuery, sqlConnection))
                    { sqlCommand.ExecuteScalar(); }

                    state = true;
                }
                catch { }
                finally { sqlConnection.Close(); }
            }

            return state;
        }

        private string ReturnPreparedStringWithParameterForIniFile(System.Linq.Expressions.Expression<Func<string>> parameter)
        {
            var me = (System.Linq.Expressions.MemberExpression)parameter.Body;
            var variableName = me.Member.Name;
            var variableValue = parameter.Compile()();

            if (variableValue?.Length > 0)
            { return (variableName + "=" + variableValue); }
            else { return variableName + "="; }
        }

        private void WriteStringsWithParametersIntoIniFile() //Запись всех рабочих парсеров в ini файл
        {
            StringBuilder sb = new StringBuilder(String.Empty);
            DateTime localDate = DateTime.Now;

            try
            {
                sb.AppendLine(@"; This " + myFileVersionInfo.ProductName + ".ini for " + myFileVersionInfo.ProductName);
                sb.AppendLine(@"; " + @"Author " + myFileVersionInfo.LegalCopyright);
                sb.AppendLine(@"");

                for (int i = 0; i < p.Length; i++)
                {
                    if (p[i]?.Length > 0)
                    { sb.AppendLine("p" + i + "=" + p[i]); }
                    else { sb.AppendLine("p" + i + "="); }
                }

                sb.AppendLine(ReturnPreparedStringWithParameterForIniFile(() => pConnectionServer));
                sb.AppendLine(ReturnPreparedStringWithParameterForIniFile(() => pConnectionUserName));
                sb.AppendLine(ReturnPreparedStringWithParameterForIniFile(() => pConnectionUserPasswords));
                sb.AppendLine(ReturnPreparedStringWithParameterForIniFile(() => pBillDeliveryCost));
                sb.AppendLine(ReturnPreparedStringWithParameterForIniFile(() => pBillDeliveryCostDiscount));
                sb.AppendLine(ReturnPreparedStringWithParameterForIniFile(() => parametrStart));
                sb.AppendLine(ReturnPreparedStringWithParameterForIniFile(() => pStop));

                sb.AppendLine(@"");
                sb.AppendLine(@"; Дата обновления файла:  " + localDate.ToString());

                File.WriteAllText(pathToIni, sb.ToString(), Encoding.GetEncoding(1251));
            }
            catch (Exception Expt)
            { MessageBox.Show(Expt.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            { sb = null; }
        }

        private bool TryToReadBillToPrepareList() //Чтение исходного файл, и первичный разбор счета (удаление ненужных данных)
        {
            bool ChosenFile = false;
            int i = 0; //amount contracts in the current bill
            listTempContract.Clear();
            filePathTxt = _OpenFileDialogReturnPath(openFileDialog1);

            if (filePathTxt == null || filePathTxt.Length < 1) { return false; }
            else
            {
                try
                {
                    _ControlSetItsText(labelFile, Path.GetFileName(filePathTxt));
                    toolTip1.SetToolTip(labelFile, "Выбранный счет для обработки");

                    var Coder = Encoding.GetEncoding(1251);

                    using (StreamReader Reader = new StreamReader(filePathTxt, Coder))
                    {
                        string s, tmp;
                        bool mystatusbegin = false;
                        bool startModuleWithDiscountWholeBill = false;
                        int lenghtData = 0;
                        _ToolStripStatusLabelSetText(StatusLabel1, "Обрабатываю файл:  " + filePathTxt);
                        while ((s = Reader.ReadLine()) != null)
                        {
                            if (s.Contains("Особовий рахунок"))
                            {
                                string[] substrings = Regex.Split(s, ":| ");
                                _ControlVisibleEnabled(labelAccount, true);
                                _ControlSetItsText(labelAccount, substrings[substrings.Length - 1].Trim());
                            }
                            else if (s.Contains("Номер рахунку"))
                            {
                                string[] substrings = Regex.Split(s, ":| ");
                                _ControlVisibleEnabled(labelBill, true);
                                _ControlSetItsText(labelBill, substrings[substrings.Length - 3].Trim());
                            }
                            else if (s.Contains(pStop)) //finished to look for contracts and start data for the bill's delivery cost
                            {
                                startModuleWithDiscountWholeBill = true;
                            }

                            else if (startModuleWithDiscountWholeBill && s.Contains(pBillDeliveryCost)) //discount calculating for the whole bill after all of contracts
                            {
                                lenghtData = s.Split(' ').Length;
                                tmp = s.Split(' ')[lenghtData - 1];

                                BillDeliveryCost = Convert.ToDouble(Regex.Replace(tmp, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
                            }
                            else if (startModuleWithDiscountWholeBill && s.Contains(pBillDeliveryCostDiscount)) //discount calculating for the whole bill after all of contracts
                            {
                                lenghtData = s.Split(' ').Length;
                                tmp = s.Split(' ')[lenghtData - 1];
                                BillDeliveryCostDiscount = Convert.ToDouble(Regex.Replace(tmp, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
                            }
                            else if (s.Contains("Розрахунковий період"))
                            {
                                string[] substrings = Regex.Split(s, ": ");
                                periodInvoice = substrings[substrings.Length - 1].Trim();
                                _ControlVisibleEnabled(labelPeriod, true);
                                _ControlSetItsText(labelPeriod, periodInvoice);
                            }

                            if (s.Contains(p[1]))
                            {
                                mystatusbegin = true;
                                i++;
                            }

                            foreach (string contractCollectedData in p)
                            {
                                if ((s.Contains(contractCollectedData) || s.Contains(pStop)) && mystatusbegin)
                                {
                                    listTempContract.Add(s.Trim());
                                    break;
                                }
                            }
                        }
                    }

                    _ControlVisibleEnabled(labelContracts, true);
                    _ControlSetItsText(labelContracts, " " + i + " шт.");

                    ChosenFile = true;

                    // вычисление скидки предоставленной Вудафон на данный счет(зависит от ИТОГОВОЙ суммы счета)
                    resultOfCalculatingDiscount = Math.Abs(BillDeliveryCostDiscount / BillDeliveryCost * 100);
                    amountBillAfterDiscount = 1 - Math.Abs(BillDeliveryCostDiscount / BillDeliveryCost);

                    _ControlVisibleEnabled(labelDiscount, true);
                    _ControlSetItsText(labelDiscount, resultOfCalculatingDiscount.ToString() + "%");

                    StatusLabel1.ToolTipText = "";

                    Dictionary<string, int> countParser = new Dictionary<string, int>();

                    foreach (string parser in p)
                    { countParser.Add(parser, 0); }

                    foreach (string str in listTempContract.ToArray())
                    {
                        foreach (string parser in p)
                        {
                            if (str.Contains(parser))
                            {
                                countParser[parser] += 1;
                            }
                        }
                    }

                    if (!(countParser[p[1]] != 0 &&                   //Количество контрактов должно быть больше нуля
                        countParser[p[1]] == countParser[p[2]] &&   //количество контрактов должно соответствовать 
                        countParser[p[2]] == countParser[p[3]]))     //количеству номеров и наименованию тарифных пакетов
                    {
                        ChosenFile = false;
                        string message = "Счет для анализа выбран с некорректными парсерами.\n" +
                                         "Количество этих параметров должны быть одинаковое и больше нуля:" +
                                         "\n'" + p[1] + @"' =  " + countParser[p[1]] +
                                         "\n'" + p[2] + @"' =  " + countParser[p[2]] +
                                         "\n'" + p[3] + @"' =  " + countParser[p[3]];
                        MessageBox.Show(message);
                        StatusLabel1.ToolTipText = message;
                    }
                }
                catch (Exception Expt)
                {
                    ChosenFile = false;
                    MessageBox.Show(Expt.ToString(), Expt.Message, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    StatusLabel1.ToolTipText = Expt.Message;
                }
            }

            return ChosenFile;
        }

        private double ClaculateAmountPaymentOfContractOwner(MobileContractPerson mobileContractPerson)
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
                    return result;
                }
            }
            return result;
        }

        private double CalculateTax(double valueBeforeTaxes)
        { return valueBeforeTaxes * 0.2; }

        private double CalculatePf(double valueBeforeTaxes)
        { return valueBeforeTaxes * 0.075; }

        private void ParseStringsOfPreparedListIntoTable() //Парсинг строк и передача результата текстовый редактор
        {
            _ToolStripStatusLabelSetText(StatusLabel1, "Обрабатываю полученные данные...");
            dataStart = labelPeriod.Text.Split('-')[0].Trim(); // дата начала периода счета
            dataEnd = labelPeriod.Text.Split('-')[1].Trim();  // дата конца периода счета

            DataRow row = dtMobile.NewRow();
            bool isUsedCurrent = false;
            bool isCheckFinishedTitles = false;

            string n = "", searchNumber;
            string[] substrings = new string[1];

            strNewModels = "";

            MobileContractPerson mcpCurrent = new MobileContractPerson();
            try
            {
                foreach (string s in listTempContract.ToArray())
                {
                    if (s.Contains(p[1]) || s.Contains(pStop))  //Начало учетов парсеров каждого кокретного контракта после упоминания ключевого слова в переменной 'p[1]'
                    {
                        //Начало учетов парсеров контракта начинаем после упоминания ключевого слова в переменной 'p[1]'
                        //перед началов учета парсеров этого контракта сначала записываем все собранные данные по предыдущему контракту
                        //для последнего в счете контракта маркером окночания данных является ключевое слово в переменной 'pStop'
                        isCheckFinishedTitles = false;
                        if (mcpCurrent.contractName.Length > 1)
                        {
                            mcpCurrent.dateBillStart = dataStart;
                            mcpCurrent.dateBillEnd = dataEnd;
                            mcpCurrent.tax = CalculateTax(mcpCurrent.totalCost);
                            mcpCurrent.pF = CalculatePf(mcpCurrent.totalCost);
                            mcpCurrent.totalCostWithTax = mcpCurrent.totalCost * 1.275;  //number spend+НДС+ПФ

                            searchNumber = mcpCurrent.mobNumberName;
                            foreach (DataRow dr in dtOwnerOfMobileWithinSelectedPeriod.Rows)
                            {
                                if (dr.ItemArray[0].ToString().Contains(searchNumber))
                                {
                                    mcpCurrent.ownerName = dr.ItemArray[1].ToString();
                                    mcpCurrent.NAV = dr.ItemArray[2].ToString();
                                    mcpCurrent.orgUnit = dr.ItemArray[3].ToString();
                                    mcpCurrent.startDate = dr.ItemArray[5].ToString();
                                    mcpCurrent.modelCompensation = dr.ItemArray[6].ToString();
                                    break;
                                }
                            }
                            mcpCurrent.payOwner = ClaculateAmountPaymentOfContractOwner(mcpCurrent);
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
                        }

                        mcpCurrent = new MobileContractPerson();
                        substrings = s.Split('№')[s.Split('№').Length - 1].Trim().Split(' ');
                        mcpCurrent.contractName = substrings[0].Trim();

                        if (s.Contains(p[2]))
                        {
                            substrings = s.Split(':')[s.Split(':').Length - 1].Trim().Split(' ');
                            mcpCurrent.mobNumberName = substrings[substrings.Length - 1].Trim();
                        }
                    }
                    else if (s.Contains(p[3]))
                    {
                        substrings = s.Split(':');
                        mcpCurrent.tarifPackageName = substrings[substrings.Length - 1].Trim();
                    }
                    else if (s.Contains(p[4]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.monthCost = Convert.ToDouble(Regex.Replace(n, "[,]",
                            System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * amountBillAfterDiscount * 1.275;
                    }
                    else if (s.Contains(p[5]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.roming = Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275;
                    }
                    else if (s.Contains(p[6]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.discount = Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
                    }
                    else if (s.Contains(p[7]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.totalCost = Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
                        isCheckFinishedTitles = true;
                        isUsedCurrent = false;
                    }
                    else if (s.Contains(p[11]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.romingData += Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275;
                    }
                    else if (s.Contains(p[12]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.extraInternetOrdered += Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275 * amountBillAfterDiscount;
                    }
                    else if (s.Contains(p[13]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.outToCity += Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275 * amountBillAfterDiscount;
                    }
                    else if (s.Contains(p[14]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.extraService += Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
                    }
                    else if (s.Contains(p[15]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.content += Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275;
                    }
                    else if (s.Contains(p[23]))
                    {
                        substrings = s.Split(' ');
                        n = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.extraServiceOrdered += Convert.ToDouble(Regex.Replace(n, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275 * amountBillAfterDiscount;
                    }
                    else if (isCheckFinishedTitles)
                    { isUsedCurrent = true; }
                }

                //additional payment for detalisation (at the end of the current bill)
                mcpCurrent = new MobileContractPerson();
                mcpCurrent.totalCost = Math.Abs(BillDeliveryCost * amountBillAfterDiscount);
                mcpCurrent.discount = Math.Abs(BillDeliveryCostDiscount);
                mcpCurrent.tax = CalculateTax(mcpCurrent.totalCost);
                mcpCurrent.pF = CalculatePf(mcpCurrent.totalCost);
                mcpCurrent.totalCostWithTax = mcpCurrent.totalCost * 1.275;  //number spend+НДС+ПФ

                row = dtMobile.NewRow();
                row[0] = "за детализацию счета, коррекция суммы";
                row[4] = Math.Round(BillDeliveryCost, 2);
                row[6] = Math.Round(mcpCurrent.discount, 2);
                row[7] = Math.Round(mcpCurrent.totalCost, 2);
                row[8] = Math.Round(mcpCurrent.tax, 2);
                row[9] = Math.Round(mcpCurrent.pF, 2);
                row[10] = Math.Round(mcpCurrent.totalCostWithTax, 2);
                row[16] = dataStart;
                row[17] = dataEnd;
                row[18] = "E22";
                row[19] = "IT-дирекция";
                row[21] = "T[6] L100% корпорация";
                dtMobile.Rows.Add(row);
            }
            catch (Exception Expt) { MessageBox.Show(Expt.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); }

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

        private void ExportFullDataTableToExcel() //Заполнение таблицы в Excel всеми данными
        {
            int rows = 1;
            int rowsInTable = dtMobile.Rows.Count;
            int columnsInTable = p.Length; // p.Length;
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
                            sheet.Columns[k + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
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
            string dataFromLabel = _ControlReturnText(labelPeriod);
            dataStart = dataFromLabel.Split('-')[0].Trim(); //'01.05.2018'
            dataEnd = dataFromLabel.Split('-')[1].Trim();  //'31.05.2018'
            string dataStartSearch =  dataStart.Split('.')[2] + "-" + dataStart.Split('.')[1] + "-" + dataStart.Split('.')[0]; //'2018-05-01'
           string dataEndSearch = dataEnd.Split('.')[2] + " - " + dataEnd.Split('.')[1] + "-" + dataEnd.Split('.')[0]; //'2018-05-31'

            listTarifData = new HashSet<string>();
            string sSqlQuery = "SELECT t1.*, t1.descr AS main," +
                                   " t2.emp_cd AS NAV, t2.emp_id AS t2emp_id," +
                                   " t3.contract_id as t3contract_id, t3.pay_model_id," +
                                   " t4.name AS model_name, " +
                                   " t5.tariff_package_name AS tariff, t5.begin_dt AS first_data , t5.end_dt AS last_data" +
                                   " FROM v_rs_contract_detail t1" +
                                   " INNER JOIN os_emp t2 ON t1.emp_id = t2.emp_id" +
                                   " LEFT JOIN (" +
                                   " SELECT * FROM os_contract_link WHERE till_dt IS NULL OR till_dt > '" + dataStartSearch + " 01:01:01 AM" + "'" +
                                   " ) t3 ON t1.contract_id = t3.contract_id AND t3.emp_id=t1.emp_id " +
                                   " LEFT JOIN rs_pay_model t4 ON t3.pay_model_id = t4.pay_model_id" +
                                   " RIGHT JOIN (" +
                                   " SELECT contract_id, tariff_package_name, begin_dt, end_dt, contract_bill_id FROM v_dp_contract_bill_detail_ex" +
                                   " ) t5" +
                                   " ON t1.contract_id = t5.contract_id" +
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
                                   " t1.till_dt > '" + dataStartSearch + "'" +
                                   " ) AND" +
                                   " t1.from_dt < '" + dataEndSearch + "'" +
                                   " ORDER by t1.phone_no, t1.emp_name ;";

            try
            {
                string pConnection = "Data Source=" + pConnectionServer +
                "; Initial Catalog=EBP;Type System Version=SQL Server 2005;Persist Security Info =True;User ID=" +
                pConnectionUserName + "; Password=" + pConnectionUserPasswords + "; Connect Timeout=180";

                using (System.Data.SqlClient.SqlConnection sqlConnection = new System.Data.SqlClient.SqlConnection(pConnection))
                {
                    sqlConnection.Open();
                    dtOwnerOfMobileWithinSelectedPeriod?.Rows?.Clear();
                    using (System.Data.SqlClient.SqlCommand sqlCommand = new System.Data.SqlClient.SqlCommand(sSqlQuery, sqlConnection))
                    {
                        using (System.Data.SqlClient.SqlDataReader sqlReader = sqlCommand.ExecuteReader())
                        {
                            foreach (System.Data.Common.DbDataRecord record in sqlReader)
                            {
                                if (record != null && record.ToString().Length > 0 && record["phone_no"].ToString().Length > 0)
                                {
                                    string mobileNumber = MakeCommonFormFromPhoneNumber(record["phone_no"].ToString());
                                    string fio = record["emp_name"].ToString().Trim();
                                    string model = record["model_name"].ToString().Trim();

                                    DataRow row = dtOwnerOfMobileWithinSelectedPeriod.NewRow();
                                    row["Номер телефона"] = mobileNumber;
                                    row["ФИО"] = fio;
                                    row["NAV"] = record["NAV"].ToString().Trim();
                                    row["Подразделение"] = record["org_unit_name"].ToString().Trim();
                                    row["Основной"] = DefineMainPhone(record["main"].ToString());
                                    row["Действует c"] = record["from_dt"].ToString().Trim().Split(' ')[0];
                                    row["Модель компенсации"] = "T[" + record["pay_model_id"].ToString().Trim() + "] " + model;

                                    //record contracts with error
                                    if (record["pay_model_id"].ToString().Trim().Length == 0) sbError.AppendLine(row["Номер телефона"].ToString().Trim() + ", " + row["ФИО"].ToString().Trim() + " - " + row["Модель компенсации"]);

                                    //if( record["model_name"].ToString().Trim().Length>0 ) listTarifData.Add(record["model_name"].ToString().Trim());
                                    listTarifData.Add("'" + model + "' - " + fio + " (" + mobileNumber + ")");
                                    dtOwnerOfMobileWithinSelectedPeriod.Rows.Add(row);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception expt) { MessageBox.Show(expt.ToString()); }
        }

        static string MakeCommonFormFromPhoneNumber(string sPrimaryPhone) //Normalize Phone to +380504197443
        {
            string sPhone = sPrimaryPhone.Trim();
            string sTemp1, sTemp2;
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

            return sPhone;
        }

        static string DefineMainPhone(string sDescription)
        {
            if (sDescription.Trim() == "!") { return "Да"; }
            else { return ""; }
        }

        private void CheckNewTarif()
        {
            string pathToNewModels = Application.StartupPath + @"\BillReportsGeneratorIsNotExistedPaymentModels.txt";
            string[] arrayData = listTarifData.ToArray();
            List<string> removeData = new List<string>();
            foreach (var tarif in arrayTarif)
            {
                for (int index = 0; index < arrayData.Length; index++)
                {
                    if (arrayData[index].Contains(tarif))
                    {
                        removeData.Add(arrayData[index]);
                    }
                }
            }

            listTarifData.ExceptWith(removeData);
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
                    sb.AppendLine(@"; This " + myFileVersionInfo.ProductName + ".ini for " + myFileVersionInfo.ProductName);
                    sb.AppendLine(@"; " + @"Author " + myFileVersionInfo.LegalCopyright);
                    sb.AppendLine(@"");
                    sb.AppendLine(@"; Дата обновления файла:  " + localDate.ToString());
                    sb.AppendLine(@";");
                    sb.AppendLine(@"; Найдены новые не учтенные модели компенсации затрат сотрудников привязанные к сотруднику в текущем счете:");
                    sb.AppendLine(@"");
                    sb.AppendLine(@"");

                    foreach (string str in listTarifData)
                    {
                        if (str?.Length > 0)
                        {
                            i++;
                            strNewModels += i + ". \"" + str + "\"\n";
                            sb.AppendLine(i + ". \"" + str + "\"");
                        }
                    }
                    sb.AppendLine(@"");

                    File.WriteAllText(pathToNewModels, sb.ToString(), Encoding.GetEncoding(1251));
                    File.AppendAllText(pathToNewModels, sbError.ToString(), Encoding.GetEncoding(1251));
                }
                catch (Exception e)
                { MessageBox.Show(e.ToString(), e.Message, MessageBoxButtons.OK, MessageBoxIcon.Error); }

                infoStatusBar = "В базе найдены новые, не добавленные ранее, модели компенсации затрат сотрудников!";

                DialogResult result = MessageBox.Show(
                    "В базе со счетами мобильного оператора на сервере " + pConnectionServer + " найдены не существующие в программе модели компенсации затрат сотрудников!\n\n" + strNewModels +
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
                        _ToolStripStatusLabelSetText(StatusLabel1, text);
                }));
            else
            {
                if (ProgressBar1.Value > 99)
                { ProgressBar1.Value = 0; }
                ProgressBar1.Maximum = 100;
                ProgressBar1.Value += 1;
                if (text.Length > 0)
                    _ToolStripStatusLabelSetText(StatusLabel1, text);
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


        private void timer1_Tick(object sender, EventArgs e) //Change a Color of the Font on Status by the Timer
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    if (StatusLabel1.ForeColor == Color.DarkBlue)
                    {
                        StatusLabel1.ForeColor = Color.DarkRed;
                    }
                    else
                    {
                        StatusLabel1.ForeColor = Color.DarkBlue;
                    }

                    if (ProgressBar1.Value > 99)
                    { ProgressBar1.Value = 0; }
                    ProgressBar1.Maximum = 100;
                    ProgressBar1.Value += 1;
                }));
            else
            {
                if (StatusLabel1.ForeColor == Color.DarkBlue)
                {
                    StatusLabel1.ForeColor = Color.DarkRed;
                }
                else
                {
                    StatusLabel1.ForeColor = Color.DarkBlue;
                }

                if (ProgressBar1.Value > 99)
                { ProgressBar1.Value = 0; }
                ProgressBar1.Maximum = 100;
                ProgressBar1.Value += 1;
            }
        }


        private string _ControlReturnText(Control controlText) //Return its name 
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

        private void _ToolStripStatusLabelSetText(ToolStripStatusLabel control, string text) //Set its name 
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.Text = text; }));
            else
                control.Text = text;
        }

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
            StringBuilder sb = new StringBuilder();
            string[] getValue;

            using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                  myRegKey,
                  Microsoft.Win32.RegistryKeyPermissionCheck.ReadSubTree,
                  System.Security.AccessControl.RegistryRights.ReadKey))
            {
                getValue = (string[])EvUserKey?.GetValue("ListServices");
                if (getValue?.Length > 0)
                {
                    foreach (string line in getValue)
                    {
                        if (!string.IsNullOrWhiteSpace(line))
                            listSavedServices.Add(line.Trim());
                    }
                    foundSavedData = true;
                }

                getValue = (string[])EvUserKey?.GetValue("ListNumbers");
                if (getValue?.Length > 0)
                {
                    foreach (string line in getValue)
                    {
                        if (!string.IsNullOrWhiteSpace(line))
                        { listSavedNumbers.Add(line.Trim()); }
                    }
                    _ControlSetItsText(labelContracts, listSavedNumbers.Count.ToString() + " шт.");
                    _ControlVisibleEnabled(labelContracts, true);
                    foundSavedData = true;
                }

                strSavedPathToInvoice = (string)EvUserKey?.GetValue("PathToLastInvoice");
                if (strSavedPathToInvoice?.Trim()?.Length > 3)
                { prepareBillItem.Enabled = true; }

                string period = (string)EvUserKey?.GetValue("PeriodLastInvoice");
                if (period?.Length > 6)
                {
                    _ControlSetItsText(labelPeriod, period);
                    _ControlVisibleEnabled(labelPeriod, true);
                }

                if (listSavedServices?.Count > 0 || listSavedNumbers?.Count > 0)
                {
                    sb.AppendLine("-= Данные для генерации маркетингового отчета =-");
                    sb.AppendLine("===================================================");
                }

                if (listSavedServices?.Count > 0)
                {
                    selectedServices = true;
                    sb.AppendLine("Учитываемые имена сервисов:");
                    foreach (string service in listSavedServices)
                    { sb.AppendLine(service ); }
                    sb.AppendLine("===================================================");
                }

                if (listSavedNumbers?.Count > 0)
                {
                    selectedNumbers = true;
                    sb.AppendLine("Список номеров по которым генерируется отчет:");
                    foreach (string number in listSavedNumbers)
                    {
                        sb.AppendLine(number);
                    }
                    sb.AppendLine("===================================================");
                }
            }

            textBoxLog.AppendText(sb.ToString());
        }

        public void SaveListStringsInRegistry(string parameterName, List<string> list) //Save List <string> into Registry as 'parameterName'
        {
            if (list?.Count > 0)
            {
                try
                {
                    using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                    {
                        EvUserKey.SetValue(parameterName, list.ToArray(),
                            Microsoft.Win32.RegistryValueKind.MultiString);
                    }
                    foundSavedData = true;
                }
                catch { MessageBox.Show("Ошибки с доступом для записи списка "+ parameterName + " в реестр. Данные не сохранены."); }
            }
        }

        public void ParameterLastInvoiceRegistrySave() //Save Parameters into Registry and variables
        {
            try
            {
                using (Microsoft.Win32.RegistryKey EvUserKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(myRegKey))
                {
                    if (filepathLoadedData?.Length > 0)
                    { EvUserKey.SetValue("PathToLastInvoice", filepathLoadedData, Microsoft.Win32.RegistryValueKind.String); }

                    if (_ControlReturnText(labelPeriod).Length > 0)
                    { EvUserKey.SetValue("PeriodLastInvoice", periodInvoice, Microsoft.Win32.RegistryValueKind.String); }
                }
                foundSavedData = true;
            }
            catch { _ = MessageBox.Show("Ошибки с доступом для записи пути к счету. Данные сохранены не корректно."); }
        }
    }

    public class MobileContractPerson
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
        public double totalCostWithoutTaxBeforDiscount = 0;
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
        /*
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
        */
    }
    public class WrittingOfTextAtFile
    {
        public WrittingOfTextAtFile() { }

        public void Write(string filePath, string text)
        {
            File.WriteAllText(
                filePath,
                text,
                Encoding.GetEncoding(1251));
        }

        public void Write(string filePath, List<string> listStrings)
        {
            File.WriteAllLines(
                filePath,
                listStrings,
                Encoding.GetEncoding(1251));
        }
    }

    public class ParsedStringOfBillWithContractOwner
    {
        internal string contract = "";
        internal string numberOwner = "";
        internal string serviceName = "";
        internal string numberTarget = "";
        internal string date = "";
        internal string time = "";
        internal string durationA = "";
        internal string durationB = "";
        internal string cost = "";

        internal string fio = "";
        internal string nav = "";
        internal string department = "";
    }
}
