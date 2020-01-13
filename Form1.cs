using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Threading.Tasks;
using System.Linq;
using System.Drawing;

namespace MobileNumbersDetailizationReportGenerator
{
    public partial class Form1 : Form
    {
        System.Diagnostics.FileVersionInfo myFileVersionInfo;
        private ContextMenu contextMenu1;
        string myRegKey;
        string pathToIni; //path to ini of tools

        string pStop = @"ЗАГАЛОМ ЗА ВСІМА КОНТРАКТАМИ";

        //Скидка на счет = pBillDeliveryCostDiscount/pBillDeliveryCost  (в процентах)
        string pBillDeliveryCost = @"Інші послуги на особовому рахунку"; //Стоимость услуги доставки электронного счета
        double BillDeliveryCost = 0; //Стоимость услуги доставки электронного счета
        string pBillDeliveryCostDiscount = @"Знижка на суму особового рахунку"; //Скидка на стоимость услуги по доставке электронного счета
        double BillDeliveryCostDiscount = 0; //Скидка на стоимость услуги по доставке электронного счета

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
        readonly DataColumn[] dcMobile ={
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

        DataTable dtOwnerOfMobileWithinSelectedPeriod = new DataTable("TarifListData");
        readonly DataColumn[] dcTarif ={
                                  new DataColumn("Номер телефона",typeof(string)),
                                  new DataColumn("ФИО",typeof(string)),
                                  new DataColumn("NAV",typeof(string)),
                                  new DataColumn("Подразделение",typeof(string)),
                                  new DataColumn("Основной",typeof(string)),
                                  new DataColumn("Действует c",typeof(string)),
                                  new DataColumn("Модель компенсации",typeof(string)),
                                  new DataColumn("Тарифный пакет",typeof(string))
                              };

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

        string filePathSourceTxt; //path to the selected bill

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
                                  new DataColumn("Стоимость",typeof(string)),
                              //    new DataColumn("Результат",typeof(decimal)),
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
            this.Text = myFileVersionInfo.Comments;

            myRegKey = $@"SOFTWARE\RYIK\{myFileVersionInfo.ProductName}";
            pathToIni = Path.Combine(Application.StartupPath, myFileVersionInfo.ProductName + ".ini"); //path to ini of tools

            string about = $"{myFileVersionInfo.Comments} ver.{myFileVersionInfo.FileVersion} {myFileVersionInfo.LegalCopyright}";

            StatusLabel1.Text = $"{myFileVersionInfo.ProductName} ver.{myFileVersionInfo.FileVersion} {myFileVersionInfo.LegalCopyright}";
            StatusLabel1.Alignment = ToolStripItemAlignment.Right;

            contextMenu1 = new ContextMenu();  //Context Menu on notify Icon
            contextMenu1.MenuItems.Add(Properties.Resources.About, AboutSoft);
            contextMenu1.MenuItems.Add(Properties.Resources.Exit, ApplicationExit);

            notifyIcon1.ContextMenu = contextMenu1;
            notifyIcon1.BalloonTipText = about;
            notifyIcon1.Text = $"{myFileVersionInfo.ProductName}{Environment.NewLine}v.{myFileVersionInfo.FileVersion}";

            ProgressBar1.Value = 0;

            groupBox1.BackColor = Color.Ivory;

            labelAccount.Visible = false;
            labelPeriod.Visible = false;
            labelBill.Visible = false;
            labelContracts.Visible = false;
            ReadStringsWithParametersFromIniFile();

            prepareBillItem.Enabled = false;

            openBillItem.ToolTipText = "Открыть счет Voodafon в текстовом формате." + Environment.NewLine + "Max количество строк - 500 000";
            makeFullReportItem.Enabled = false;
            makeFullReportItem.ToolTipText = "Подготовить полный отчет в Excel-файле." + Environment.NewLine + "Файл будет сохранен в папке с программой";
            makeReportAccountantItem.Enabled = false;
            makeReportAccountantItem.ToolTipText = "Подготовить отчет для бух. в Excel-файле." + Environment.NewLine + "Файл будет сохранен в папке с программой";
            labelDiscount.Text = "";
            clearTextboxItem.ToolTipText = "Убрать весь текст из окна просмотра";
            aboutItem.ToolTipText = "О программе";
            exitItem.ToolTipText = "Выйти из программы и сохранить настройки и парсеры счета";

            /*buttonReport2.FlatAppearance.MouseOverBackColor = Color.PaleGreen;
            buttonExit.FlatAppearance.MouseOverBackColor = Color.SandyBrown;
            */
            dtMobile.Columns.AddRange(dcMobile);
            dtOwnerOfMobileWithinSelectedPeriod.Columns.AddRange(dcTarif);
            dtMarket.Columns.AddRange(dcFullBill);
            ListsRegistryDataCheck();
            useSavedDataItem.Enabled = foundSavedData;
            useSavedDataItem.ToolTipText = "Использовать сохраненный список файлов и сервисов из предыдущей сессии";
        }


        private void AboutSoft()
        {
            string strVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

            MessageBox.Show(
                myFileVersionInfo.Comments + Environment.NewLine + "Версия: " + myFileVersionInfo.FileVersion + Environment.NewLine + "Build: " +
                strVersion + Environment.NewLine + myFileVersionInfo.LegalCopyright,
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
            contextMenu1?.Dispose();
            notifyIcon1?.Dispose();

            Application.Exit();
        }

        private void openBillItem_Click(object sender, EventArgs e)//Menu "Open"
        {
            textBoxLog.Clear();
            OpenBill();
        }

        private void makeFullReportItem_Click(object sender, EventArgs e)
        { ExportDataTableToExcelForAccount(true); }

        private void makeReportAccountantToolItem_Click(object sender, EventArgs e)
        { ExportDataTableToExcelForAccount(); }

        private void ExportDataTableToExcelForAccount(bool makePivot = false)
        {
            string[] columnsCollection = new string[]      // для бухгалтерии
                       {
                        "Дата счета",
                        "Номер телефона абонента",
                        "ФИО сотрудника",
                        "Затраты по номеру, грн",
                        "НДС, грн",
                        "ПФ, грн",
                        "Итого по контракту, грн",
                        "Общая сумма в роуминге, грн",
                        "Подразделение",
                        "Табельный номер",
                        "ТАРИФНАЯ МОДЕЛЬ",
                        "К оплате владельцем номера, грн"
                   };

            string pathToFile = Path.Combine(Path.GetDirectoryName(filePathSourceTxt), $"{Path.GetFileNameWithoutExtension(filePathSourceTxt)}.xlsx");
            string nameSheet = Path.GetFileNameWithoutExtension(filePathSourceTxt);
            string[] redColumns = { "К оплате владельцем номера, грн" };
            string[] greenColumns = { "Затраты по номеру, грн", "Итого по контракту, грн" };

            DataTable dt = dtMobile.Copy();
            if (makePivot)
            {
                pathToFile = Path.Combine(Path.GetDirectoryName(filePathSourceTxt), $"{Path.GetFileNameWithoutExtension(filePathSourceTxt)} pivot.xlsx");
                dt.AllowToEditTable()
                         .SetColumnsOrder(columnsCollection)
                         .SeteColumnsCollectionInDataTable(columnsCollection)
                         .ExportToExcelPivotTable(pathToFile, nameSheet, redColumns, greenColumns);
            }
            else
            {
                dt.AllowToEditTable()
                    .SetColumnsOrder(columnsCollection)
                    .SeteColumnsCollectionInDataTable(columnsCollection)
                    .ExportToExcel(pathToFile, nameSheet, redColumns, greenColumns);
            }

            MessageShow($"Отчет готов и сохранен:{Environment.NewLine}{pathToFile}");
        }

        private void clearTextBoxItem_Click(object sender, EventArgs e)
        { textBoxLog.Clear(); }

        private void AboutSoft(object sender, EventArgs e)
        { AboutSoft(); }

        private void ApplicationExit(object sender, EventArgs e)
        { ApplicationExit(); }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        { ApplicationExit(); }

        private void selectListNumbersItem_Click(object sender, EventArgs e)
        { PrepareListPhoneNumbers(); }

        /// <summary>
        /// limit of numbers <500. Prepare list of numbers for the marketing report
        /// </summary>
        private void PrepareListPhoneNumbers()
        {
            selectedNumbers = false;
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

                    if (strTemp.Length == 13)  //Correct Length of a formated mobile number == 13 //+380123456789
                    { listNumbers.Add(strTemp); }
                    else
                    { listWrongString.Add(strTemp); }
                }

                if (0 < listWrongString?.Count)
                {
                    textBoxLog.AppendLine("List of first 300 wrong rows in the selected list:");
                    textBoxLog.AppendLine(Properties.Resources.RowDashedLines);
                    int wrongRow = 0;
                    foreach (string s in listWrongString)
                    {
                        textBoxLog.AppendLine(s);
                        wrongRow++;

                        if (wrongRow > limitWrongNumber)
                        { break; }
                    }
                    textBoxLog.AppendLine();
                }

                if (0 < listNumbers?.Count && listNumbers?.Count < 500)
                {
                    selectedNumbers = true;
                    SaveListStringsInRegistry(Properties.Resources.ListOfNumbers, listNumbers);

                    textBoxLog.AppendLine(Properties.Resources.ListOfNumbers);
                    textBoxLog.AppendLine(Properties.Resources.RowDashedLines);

                    foreach (string s in listNumbers)
                    { textBoxLog.AppendLine(s); }
                }
                else
                { textBoxLog.AppendLine("Check the list of numbers." + Environment.NewLine + "In the list was found: " + listNumbers.Count + " number(s)"); }
            }
            CheckConditionEnableMarketingReport();
        }

        private void selectListServicesItem_Click(object sender, EventArgs e)
        { PrepareListServicesToMakeReport(); }

        /// <summary>
        /// limit of services <100. Prepare list of services for the marketing report
        /// </summary>        
        private void PrepareListServicesToMakeReport()
        {
            selectedServices = false;
            textBoxLog.Clear();

            listServices.Clear();
            listServices = LoadDataIntoList();

            if (0 < listServices?.Count && listServices?.Count < 100)
            {
                textBoxLog.AppendLine(Properties.Resources.RowDashedLines);
                textBoxLog.AppendLine(Properties.Resources.ListOfServices);
                textBoxLog.AppendLine(Properties.Resources.RowDashedLines);

                foreach (string s in listServices)
                { textBoxLog.AppendLine(s); }

                selectedServices = true;

                SaveListStringsInRegistry(Properties.Resources.ListOfServices, listServices);
            }
            else
            {
                textBoxLog.AppendLine(Properties.Resources.RowDashedLines);
                textBoxLog.AppendLine("The selected list is wrong!" + Environment.NewLine + "Will check the file!" + Environment.NewLine + "It has to contain from 1 to 100 services.");
                textBoxLog.AppendLine(Properties.Resources.RowDashedLines);
            }
            CheckConditionEnableMarketingReport();
        }

        private async void prepareBillItem_Click(object sender, EventArgs e)
        {
            dtMarket.Rows.Clear();
            string pathToFileMarketPivotTable;
            string pathToFileMarketTable;
            await Task.Run(() => LoadBillIntoMemoryToFilter());

            //test
            //  var typeResult = TypeData.DataStringMb;
            ConditionForMakingPivotTable condition = new ConditionForMakingPivotTable
            {                                                           // columns 'dcFullBill' in the table 'dtMarket'
                KeyColumnName = "Номер телефона",                       // column "ФИО" //groupby
                FilteringService = "internet",                          // it is used by column - "Номер В", //Передача даних  
                NameColumnWithFilteringService = "Номер В",             // column "Номер В",
                NameColumnWithFilteringServiceValue = "Длительность А", // column "Длительность А", it is used by column 'Summary'
                NameNewColumnWithSummary = "Суммарно, МБ",              // column 'Summary' - result data format for column Summary
                NameNewColumnWithCount = "Количество",
                //  TypeResultCalcultedData = typeResult,                   
                ColumnsCollectionAtRightOrder =
                new string[] { "Подразделение", "ФИО", "NAV", "Номер телефона", "Номер В" }
            };

            MakerPivotTable makingPivotData = new MakerPivotTable(dtMarket, condition);

            pathToFileMarketPivotTable = Path.Combine(Path.GetDirectoryName(filepathLoadedData), $"{Path.GetFileNameWithoutExtension(filepathLoadedData)} MarketPivotTable.xlsx");
            pathToFileMarketTable = Path.Combine(Path.GetDirectoryName(filepathLoadedData), $"{Path.GetFileNameWithoutExtension(filepathLoadedData)} MarketTable.xlsx");

            try { await Task.Run(() => dtMarket.ExportToExcel(pathToFileMarketTable, "Selected data")); }
            catch (Exception err) { MessageShow(err.ToString()); }

            try { await Task.Run(() => makingPivotData.MakePivot().ExportToExcel(pathToFileMarketPivotTable, "PivotTable")); }
            catch (Exception err) { MessageShow(err.ToString()); }

            textBoxLog.AppendLine("Генерация завершена");
            MessageShow("Готово!");
        }

        //  private void MessageShow(object sender, TextEventArgs e)
        // { Task.Run(() => MessageBox.Show(e.Message)); }

        private void MessageShow(string text)
        { Task.Run(() => MessageBox.Show(text)); }

        private void LoadBillIntoMemoryToFilter()
        {
            ProgressBar1Start();
            textBoxLog.Clear();
            ToolStripMenuItemEnabled(fileMenuItem, false);
            ControlVisibleEnabled(labelPeriod, true);

            loadedBill = false;

            dtOwnerOfMobileWithinSelectedPeriod = GetDataWithModel();

            string contract = "";
            string numberMobile = "";
            string tempRow;
            string exceptedStringContains = @". . .";
            // NUMBER_OF_CONTRACT,       //1     //number of contract
            // MOBILE_NUMBER,           //2     //number
            // NAME_OF_TARIF,            //3     //name of tarif package

            p[1] = ControlReturnText(textBoxP1);
            p[2] = ControlReturnText(textBoxP2);

            List<string> filterBill = new List<string>
            {
                p[1],
                p[2]
            };

            if (listServices?.Count == 0) { listServices = listSavedServices; }
            if (listNumbers?.Count == 0) { listNumbers = listSavedNumbers; }

            ProgressWork1Step();

            foreach (string service in listServices) { filterBill.Add(service); }

            ProgressWork1Step();

            List<string> loadedBillWithServicesFiltered = LoadDataUsingParameters(filterBill, parametrStart, pStop, exceptedStringContains);

            int counterstep = (dtOwnerOfMobileWithinSelectedPeriod.Rows.Count + listNumbers.Count);
            int countStepProgressBar = counterstep;
            int countRowsInTable = 0;

            if (loadedBillWithServicesFiltered?.Count > 0)
            {
                StringBuilder sb = new StringBuilder();

                //todo parsing strings of the filtered bill
                foreach (string sRowBill in loadedBillWithServicesFiltered)
                {
                    if (sRowBill.StartsWith(p[1]))
                    {
                        try
                        {
                            contract = Regex.Split(sRowBill.Substring(sRowBill.IndexOf('№') + 1).Trim(), " ")[0].Trim();
                            tempRow = sRowBill.Substring(sRowBill.IndexOf(':') + 1).Trim();

                            if (tempRow.StartsWith("+"))
                            { numberMobile = tempRow; }
                            else
                            { numberMobile = "+" + tempRow; } //set format number like '+380...'
                        }
                        catch (Exception err)
                        {
                            MessageShow("Проверьте правильность выбора файла с контрактами с детализацией разговоров!" + Environment.NewLine +
                                "Возможно поменялся формат." + Environment.NewLine +
                                "Правильный формат первых строк с новым контрактом:" + Environment.NewLine +
                                NUMBER_OF_CONTRACT + " 000000000  Моб.номер: 380000000000" + Environment.NewLine +
                                "Ціновий Пакет: название_пакета" + Environment.NewLine + "далее - детализацией разговоров контракта" + Environment.NewLine +
                                "В данном случае строка с началом разбираемого контракта имеет форму:" + Environment.NewLine +
                                sRowBill + Environment.NewLine +
                                "Ошибка: " + err.ToString()
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
                        ParsingStringDetalizationOfBill parsing = new ParsingStringDetalizationOfBill();
                        ParsedStringOfBill parsed;

                        try
                        {
                            foreach (string sNumber in listNumbers)
                            {
                                if (numberMobile.StartsWith(sNumber))
                                {
                                    //serviceName = sRowBill?.Substring(0, 38)?.Trim();
                                    //numberB = sRowBill?.Substring(38, 13)?.Trim();
                                    //date = sRowBill?.Substring(52, 10)?.Trim();
                                    //time = sRowBill?.Substring(65, 8)?.Trim();
                                    //durationA = sRowBill?.Substring(74, 9)?.Trim();
                                    //durationB = sRowBill?.Substring(84, 9)?.Trim();
                                    //cost = sRowBill?.Substring(95)?.Trim();

                                    parsing.SetString(sRowBill);
                                    parsed = parsing.ParseString();

                                    parsed.contract = contract;
                                    parsed.numberOwner = sNumber;

                                    foreach (DataRow rowTarif in dtOwnerOfMobileWithinSelectedPeriod.Rows)
                                    {
                                        if (rowTarif["Номер телефона"].ToString().Contains(sNumber))
                                        {
                                            parsed.fio = rowTarif["ФИО"].ToString();
                                            parsed.nav = rowTarif["NAV"].ToString();
                                            parsed.department = rowTarif["Подразделение"].ToString();
                                            break;
                                        }
                                    }

                                    //for dump
                                    tempRow = $"{parsed.numberOwner}\t{parsed.fio}\t{parsed.nav}\t{parsed.department}\t{parsed.serviceName}\t" +
                                        $"{parsed.numberTarget}\t{parsed.date}\t{parsed.time}\t{parsed.durationA}\t{parsed.durationB}\t{parsed.cost}";

                                    DataRow rowMarket = dtMarket.NewRow(); //for Market
                                    rowMarket["Контракт"] = parsed.contract;
                                    rowMarket["Номер телефона"] = parsed.numberOwner;
                                    rowMarket["Имя сервиса"] = parsed.serviceName;
                                    rowMarket["Номер В"] = parsed.numberTarget;
                                    rowMarket["Дата"] = parsed.date;
                                    rowMarket["Время"] = parsed.time;
                                    rowMarket["Длительность А"] = parsed.durationA;
                                    rowMarket["Длительность В"] = parsed.durationB;
                                    rowMarket["Стоимость"] = parsed.cost;
                                    rowMarket["ФИО"] = parsed.fio;
                                    rowMarket["NAV"] = parsed.nav;
                                    rowMarket["Подразделение"] = parsed.department;

                                    dtMarket.Rows.Add(rowMarket);
                                    countRowsInTable++;
                                    sb.AppendLine(tempRow);
                                    break;
                                }

                                countStepProgressBar--;
                                if (countStepProgressBar <= 0)
                                {
                                    string s = $"В отчет добавлено {countRowsInTable,20 }, строк из {loadedBillWithServicesFiltered.Count,15}";
                                    ProgressWork1Step(s);
                                    countStepProgressBar = counterstep;
                                }
                            }
                        }
                        catch (Exception err)
                        { MessageBox.Show($"Во время парсинга счета возникла ошибка в строке:{Environment.NewLine}{sRowBill}{Environment.NewLine}{err.ToString()}", err.Message); }
                    }
                }
                loadedBill = true;
                { textBoxLog.AppendLine($"Сформировано для генерации отчета {countRowsInTable} строк c номерами мобильных подпадающими под фильтр."); }

                sb.ToString()
                    .WriteAtFile(Path.Combine(Path.GetDirectoryName(filepathLoadedData), "listMarketingCollectRows.csv"));
            }
            else
            { textBoxLog.AppendLine("В выборке нет ничего для указанных номеров!"); }

            CheckConditionEnableMarketingReport();
            ToolStripStatusLabelSetText(StatusLabel1, "Файл сохранен в папку: " + Path.GetDirectoryName(filepathLoadedData));

            ToolStripMenuItemEnabled(fileMenuItem, true);
            ProgressBar1Stop();
        }

        private void CheckConditionEnableMarketingReport() //enableing Marketing report if load data is correct
        {
            if (selectedServices && selectedNumbers && loadedBill)
            {
                ToolStripMenuItemEnabled(prepareBillItem, true);
            }
            else if (selectedServices && selectedNumbers)
            {
                ToolStripMenuItemEnabled(prepareBillItem, true);
            }
        }

        private List<string> LoadDataIntoList() //max received List's length = 500 000 rows
        {
            int listMaxLength = 500000;
            List<string> listValue = new List<string>(listMaxLength);
            string s = "";
            int i = 0; // it is not empty's rows in the selected file

            string filepathLoadedData = OpenFileDialogReturnPath(openFileDialog1);
            if (filepathLoadedData?.Length > 0)
            {
                try
                {
                    var Coder = Encoding.GetEncoding(1251);
                    using (StreamReader Reader = new StreamReader(filepathLoadedData, Coder))
                    {
                        ToolStripStatusLabelSetText(StatusLabel1, "Обрабатываю файл:  " + filepathLoadedData);
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
                catch (Exception expt)
                { MessageShow("Ошибка произошла на " + i + " строке:" + Environment.NewLine + expt.ToString()); }

                if (i > listMaxLength - 10 || i == 0)
                { MessageShow("Error was happened on " + i + " row" + Environment.NewLine + " You've been chosen the long file!"); }
            }
            else
            { MessageShow("Не выбран файл со счетом."); }

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
                              "Использовать предыдущий выбор файла?" + Environment.NewLine + strSavedPathToInvoice,
                              Properties.Resources.Attention,
                              MessageBoxButtons.YesNo,
                              MessageBoxIcon.Exclamation,
                              MessageBoxDefaultButton.Button1);
                        if (result == DialogResult.No)
                        {
                            filepathLoadedData = OpenFileDialogReturnPath(openFileDialog1);
                        }
                        else
                        {
                            filepathLoadedData = strSavedPathToInvoice;
                        }
                    }
                    else if (!currentInvoice)
                    {
                        filepathLoadedData = OpenFileDialogReturnPath(openFileDialog1);
                    }

                    if (filepathLoadedData?.Length > 2 && File.Exists(filepathLoadedData))
                    {
                        ToolStripStatusLabelSetText(StatusLabel1, "Обрабатываю файл:  " + filepathLoadedData);
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
                                            if (loadedString.StartsWith(parameterString) && !loadedString.Contains(excepted))
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
                                        ProgressWork1Step();
                                        countStepProgressBar = 500;
                                    }
                                }
                            }

                            if (checkPeriod && checkRahunok && checkNomerRahunku)
                            {
                                ControlSetItsText(labelPeriod, periodInvoice);
                            }

                            ParameterLastInvoiceRegistrySave();
                        }
                        catch (Exception expt)
                        { MessageBox.Show("Error was happened on " + listRows.Count + " row" + Environment.NewLine + expt.ToString()); }
                        textBoxLog.AppendLine("Из файла-счета: " + Environment.NewLine);
                        textBoxLog.AppendLine(filepathLoadedData);
                        textBoxLog.AppendLine("отобрано для построения отчета " + counter + " строк с требуемыми сервисами");
                        if (listMaxLength - 2 < listRows.Count || listRows.Count == 0)
                        { MessageBox.Show("Error was happened on " + (listRows.Count) + " row" + Environment.NewLine + " You've been chosen the long file!"); }
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
            dtMobile?.Rows?.Clear();
            filePathSourceTxt = null;
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

                await Task.Run(() => dtOwnerOfMobileWithinSelectedPeriod = GetDataWithModel());
                if (dtOwnerOfMobileWithinSelectedPeriod.Rows.Count < 2)
                {
                    MessageBox.Show("Выбранный счет в базу данных Tfactura еще не импортирован!" + Environment.NewLine + "Перед обработкой счета, предварительно необходимо импортировать счет в базу!");
                    StatusLabel1.Text = "Обработка счета прекращена! Предварительно импортируйте счет в Tfactura!";
                    StatusLabel1.BackColor = Color.SandyBrown;
                }
                else
                {
                    await Task.Run(() => CheckNewTarif());

                    //clear log if it was found a problem
                    if (listTarifData?.Count > 0)
                    { textBoxLog.Clear(); }

                    if (!newModels)
                    {
                        ParseStringsOfPreparedListIntoTable();
                        DataRow[] results;

                        string columnName1 = dtMobile.Columns["ФИО сотрудника"].ColumnName.Remove(3);
                        string columnName2 = dtMobile.Columns["Номер телефона абонента"].ColumnName.Remove(14);
                        string columnName3 = dtMobile.Columns["Ціновий Пакет"].ColumnName;
                        string columnName4 = dtMobile.Columns["Итого по контракту, грн"].ColumnName.Remove(6);
                        string columnName5 = dtMobile.Columns["ТАРИФНАЯ МОДЕЛЬ"].ColumnName;
                        string columnName6 = "Роуминг";                     //dtMobile.Columns[5].ColumnName;
                        string columnName10 = dtMobile.Columns["NumberUsed"].ColumnName;
                        string columnName11 = dtMobile.Columns["NumberNoBlock"].ColumnName;

                        string sortOrder = dtMobile.Columns["ФИО сотрудника"].ColumnName + " ASC";

                        textBoxLog.AppendLine("-= Дата счета:  " + dtMobile.Rows[1]["Дата счета"].ToString() + " =-"); //Дата счета
                        textBoxLog.AppendLine(Properties.Resources.RowDozenOfEqualSymbols);

                        //////////////////////////////
                        if (listTarifData.Count > 0)
                        {
                            textBoxLog.AppendLine("-= Список тарифных схем, не существующих в программе =-");
                            textBoxLog.AppendLine("'" + columnName5 + "' - " + columnName1 + " (" + columnName2 + ")");

                            foreach (string str in listTarifData)
                            {
                                textBoxLog.AppendLine(str);
                            }
                            textBoxLog.AppendLine(Properties.Resources.RowDashedLines);
                        }

                        /////////////////
                        results = dtMobile.Select("NumberUsed='False' AND NumberNoBlock='True'", sortOrder, DataViewRowState.Added);
                        if (results.Length > 0)
                        {
                            textBoxLog.AppendLine("-= Список контрактов, по которым не велась работа =-");
                            textBoxLog.AppendLine(
                                 string.Format("{0,-40}", columnName1) +
                                 string.Format("{0,-15}", columnName2) +
                                 string.Format("{0,-30}", columnName3) +
                                 string.Format("{0,-10}", columnName4) +
                                 string.Format("{0,-30}", columnName5));
                            for (int i = 0; i < results.Length; i++)
                            {

                                textBoxLog.AppendLine(
                                 string.Format("{0,-40}", results[i][0].ToString()) +
                                 string.Format("{0,-15}", results[i][2].ToString()) +
                                 string.Format("{0,-30}", results[i][3].ToString()) +
                                 string.Format("{0,-10}", results[i][10].ToString()) +
                                 string.Format("{0,-30}", results[i][21].ToString()));
                            }
                            textBoxLog.AppendLine(Properties.Resources.RowDashedLines);
                        }

                        /////////////////
                        results = dtMobile.Select("NumberNoBlock='False'", sortOrder, DataViewRowState.Added);
                        if (results.Length > 0)
                        {
                            textBoxLog.AppendLine("-= Список заблокированных контрактов =-");
                            textBoxLog.AppendLine(
                                 string.Format("{0,-40}", columnName1) +
                                 string.Format("{0,-15}", columnName2) +
                                 string.Format("{0,-30}", columnName3) +
                                 string.Format("{0,-10}", columnName4) +
                                 string.Format("{0,-30}", columnName5));
                            for (int i = 0; i < results.Length; i++)
                            {
                                textBoxLog.AppendLine(
                                 string.Format("{0,-40}", results[i][0].ToString()) +
                                 string.Format("{0,-15}", results[i][2].ToString()) +
                                 string.Format("{0,-30}", results[i][3].ToString()) +
                                 string.Format("{0,-10}", results[i][10].ToString()) +
                                 string.Format("{0,-30}", results[i][21].ToString()));
                            }
                            textBoxLog.AppendLine(Properties.Resources.RowDashedLines);
                        }

                        /////////////////
                        textBoxLog.AppendLine("---= Все =---");
                        results = dtMobile.Select(dtMobile.Columns[0].ColumnName.Length + " > 0", sortOrder, DataViewRowState.Added);
                        textBoxLog.AppendLine(
                             string.Format("{0,-40}", columnName1) +
                             string.Format("{0,-15}", columnName2) +
                             string.Format("{0,-30}", columnName3) +
                             string.Format("{0,-10}", columnName4) +
                             string.Format("{0,-10}", columnName6) +
                             string.Format("{0,-30}", columnName5) +
                             string.Format("{0,-12}", columnName10) +
                             string.Format("{0,-12}", columnName11));
                        for (int i = 0; i < results.Length; i++)
                        {

                            textBoxLog.AppendLine(
                             string.Format("{0,-40}", results[i][0].ToString().Trim()) +
                             string.Format("{0,-15}", results[i][2].ToString()) +
                             string.Format("{0,-30}", results[i][3].ToString()) +
                             string.Format("{0,-10}", results[i][10].ToString()) +
                             string.Format("{0,-10}", results[i][5].ToString()) +

                             string.Format("{0,-30}", results[i][21].ToString()) +
                             string.Format("{0,-12}", results[i][24].ToString()) +
                             string.Format("{0,-12}", results[i][25].ToString()));
                        }
                        textBoxLog.AppendLine(Properties.Resources.RowDozenOfEqualSymbols);
                        /////////////////

                        makeReportAccountantItem.Enabled = true;
                        makeFullReportItem.Enabled = true;

                        StatusLabel1.Text = "Предварительная обработка счета из файла " + Path.GetFileName(filePathSourceTxt) + " завершена!";
                        StatusLabel1.ToolTipText = "Данные для генерации отчета для бухгалтерии подготовлены";
                    }
                    else
                    {
                        textBoxLog.AppendLine("В базе найдены новые, не настроенные в данной программе на обработку,");
                        textBoxLog.AppendLine("модели тарификации компенсации затрат сотрудников:");

                        int i = 0;
                        foreach (string str in listTarifData)
                        {
                            textBoxLog.AppendLine(++i + ". \"" + str);
                        }
                        textBoxLog.AppendLine(Properties.Resources.RowDozenOfEqualSymbols);
                        textBoxLog.AppendLine(sbError.ToString());
                    }

                    if (infoStatusBar.Length > 1)
                    {
                        StatusLabel1.Text = infoStatusBar;
                        StatusLabel1.BackColor = Color.SandyBrown;
                    }
                    makeReportAccountantItem.Enabled = true;
                    makeFullReportItem.Enabled = true;
                }

                filepathLoadedData = filePathSourceTxt;

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
            string s = "", info = "";
            bool b1 = false, b2 = false;
            toolTip1.SetToolTip(this.groupBox1, "Использованы исходные настройки программы");

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
                                if (s.StartsWith($"{nameof(pConnectionServer)}="))
                                {
                                    pConnectionServer = ParseParameterNameAndValueFromReadString("=", s, pConnectionServer);
                                }
                                else if (s.StartsWith($"{nameof(pConnectionUserName)}="))
                                {
                                    pConnectionUserName = ParseParameterNameAndValueFromReadString("=", s, pConnectionUserName);
                                }
                                else if (s.StartsWith($"{nameof(pConnectionUserPasswords)}="))
                                {
                                    pConnectionUserPasswords = ParseParameterNameAndValueFromReadString("=", s, pConnectionUserPasswords);
                                }
                                else if (s.StartsWith($"{nameof(parametrStart)}="))
                                {
                                    parametrStart = ParseParameterNameAndValueFromReadString("=", s, parametrStart);
                                }
                                else if (s.StartsWith($"{nameof(pStop)}="))
                                {
                                    pStop = ParseParameterNameAndValueFromReadString("=", s, pStop);
                                }
                                else if (s.StartsWith($"{nameof(pBillDeliveryCost)}=")) //Строка с суммой стоимости доставки электронного счета до вычисления скидки и налогов
                                {
                                    pBillDeliveryCost = ParseParameterNameAndValueFromReadString("=", s, pBillDeliveryCost);
                                }
                                else if (s.StartsWith($"{nameof(pBillDeliveryCostDiscount)}="))//Строка с суммой скидки на доставку электронного счет
                                {
                                    pBillDeliveryCostDiscount = ParseParameterNameAndValueFromReadString("=", s, pBillDeliveryCostDiscount);
                                }

                                for (int i = 0; i < p?.Length; i++)
                                {
                                    if (s.StartsWith($"p{i.ToString()}="))
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
                    info += $"Настройки из {myFileVersionInfo.ProductName}.ini проигнорированы. Изменен формат файла{Environment.NewLine}";
                }
                else
                {
                    info += $"Парсеры модифицированы настройками из {myFileVersionInfo.ProductName}.ini{Environment.NewLine}";
                    groupBox1.BackColor = Color.Tan;
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
                info += infoStatusBar + Environment.NewLine + "Проверьте и добавьте в файл с настройками - " + Environment.NewLine +
                    pathToIni + Environment.NewLine + "отсутствующие данные, необходимые для подключения к базе данных:" + Environment.NewLine + Environment.NewLine +
                      "pConnectionServer=" + pConnectionServer + Environment.NewLine +
                      "pConnectionUserName=" + pConnectionUserName + Environment.NewLine +
                      "pConnectionUserPasswords=" + pConnectionUserPasswords;
                MessageBox.Show(info, Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);

                StatusLabel1.Text = infoStatusBar;
                StatusLabel1.ToolTipText = info;

                StatusLabel1.BackColor = Color.SandyBrown;
            }
            else
            {
                fileMenuItem.Enabled = false;
                StatusLabel1.Text = "Проверяю доступность БД сервера";
                StatusLabel1.BackColor = Color.PaleGoldenrod;

                ProgressBar1Start();
                string infoStatus = null, infoStatusTooltip = null;
                System.Drawing.Color infoStatusBackColor = System.Drawing.SystemColors.Menu;
                using (Timer timer1 = new Timer { Interval = 200, Enabled = true })
                {
                    timer1.Tick += new System.EventHandler(this.timer1_Tick);
                    timer1.Start();

                    bool aliveServer = true;
                    await Task.Run(() => aliveServer = CheckAliveDbServer());

                    if (!aliveServer)
                    {
                        infoStatusBar = "БД сервера со счетами Tfactura не доступна";
                        info += infoStatusBar + Environment.NewLine + "Проверьте настройки в файле с настройками -" + Environment.NewLine +
                            pathToIni + "и исправьте не верные данные:" + Environment.NewLine +
                            "pConnectionServer=" + pConnectionServer + Environment.NewLine +
                            "pConnectionUserName=" + pConnectionUserName + Environment.NewLine +
                            "pConnectionUserPasswords=" + pConnectionUserPasswords;
                        MessageBox.Show(info, Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);

                        infoStatusTooltip = info;
                        infoStatus = infoStatusBar;
                        infoStatusBackColor = Color.SandyBrown;
                    }
                    else
                    {
                        fileMenuItem.Enabled = true;
                        infoStatusBackColor = Color.PaleGreen;
                        infoStatus = "БД сервера со счетами Tfactura доступна для генерации отчетов";
                        infoStatusTooltip = "выберите счет мобильного оператора с которым планируете работать";
                    }
                    StatusLabel1.Text = infoStatus;
                    StatusLabel1.ToolTipText = infoStatusTooltip;
                    StatusLabel1.BackColor = infoStatusBackColor;

                    timer1.Enabled = false;
                    timer1.Stop();
                }
                ProgressBar1Stop();
            }
            StatusLabel1.ForeColor = Color.Black;
        }

        private bool CheckAliveDbServer()
        {
            bool state = false;
            string pConnection =
                $"Data Source={pConnectionServer}; Initial Catalog=EBP; Type System Version=SQL Server 2005; Persist Security Info =True" +
                $"; User ID={pConnectionUserName}; Password={pConnectionUserPasswords}; Connect Timeout=5";

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

                sb.ToString().WriteAtFile(pathToIni);
            }
            catch (Exception Expt)
            { MessageBox.Show(Expt.ToString(), Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally
            { sb = null; }
        }

        private bool TryToReadBillToPrepareList() //Чтение исходного файл, и первичный разбор счета (удаление ненужных данных)
        {
            bool ChosenFile;
            int i = 0; //amount contracts in the current bill
            listTempContract.Clear();
            filePathSourceTxt = OpenFileDialogReturnPath(openFileDialog1);

            if (filePathSourceTxt?.Length > 3)
            {
                try
                {
                    Invoice invoice = new Invoice
                    {
                        invoicePathToFile = filePathSourceTxt,
                        invoiceFileName = Path.GetFileName(filePathSourceTxt)
                    };

                    ControlSetItsText(labelFile, invoice.invoiceFileName);
                    toolTip1.SetToolTip(labelFile, Properties.Resources.SelectedInvoice);

                    using (StreamReader Reader = new StreamReader(invoice.invoicePathToFile, Encoding.GetEncoding(1251)))
                    {
                        string s, tmp;
                        bool mystatusbegin = false;
                        bool startModuleWithDiscountWholeBill = false;
                        int lenghtData = 0;

                        ToolStripStatusLabelSetText(StatusLabel1, "Обрабатываю файл:  " + invoice.invoicePathToFile);
                        while ((s = Reader.ReadLine()) != null)
                        {
                            if (s.Contains("Особовий рахунок"))
                            {
                                string[] substrings = Regex.Split(s, ":| ");
                                invoice.invoiceInternalHoldingNumber = substrings[substrings.Length - 1].Trim();

                                ControlVisibleEnabled(labelAccount, true);
                                ControlSetItsText(labelAccount, invoice.invoiceInternalHoldingNumber);
                            }
                            else if (s.Contains("Номер рахунку"))
                            {
                                string[] substrings = Regex.Split(s, ":| ");
                                invoice.invoiceNumber = substrings[substrings.Length - 3].Trim();

                                ControlVisibleEnabled(labelBill, true);
                                ControlSetItsText(labelBill, invoice.invoiceNumber);
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
                                invoice.invoiceDeliveryCost = BillDeliveryCost;
                            }
                            else if (startModuleWithDiscountWholeBill && s.Contains(pBillDeliveryCostDiscount)) //discount calculating for the whole bill after all of contracts
                            {
                                lenghtData = s.Split(' ').Length;
                                tmp = s.Split(' ')[lenghtData - 1];
                                BillDeliveryCostDiscount = Convert.ToDouble(Regex.Replace(tmp, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
                                invoice.invoiceDeliveryCostDiscount = BillDeliveryCostDiscount;
                            }
                            else if (s.Contains("Розрахунковий період"))
                            {
                                string[] substrings = Regex.Split(s, ": ");
                                periodInvoice = substrings[substrings.Length - 1].Trim();
                                invoice.invoicePeriod = periodInvoice;

                                ControlVisibleEnabled(labelPeriod, true);
                                ControlSetItsText(labelPeriod, periodInvoice);
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

                    ControlVisibleEnabled(labelContracts, true);
                    ControlSetItsText(labelContracts, " " + i + " шт.");

                    ChosenFile = true;

                    // вычисление скидки предоставленной Вудафон на данный счет(зависит от ИТОГОВОЙ суммы счета)
                    resultOfCalculatingDiscount = Math.Abs(BillDeliveryCostDiscount / BillDeliveryCost * 100);
                    amountBillAfterDiscount = 1 - Math.Abs(BillDeliveryCostDiscount / BillDeliveryCost);

                    ControlVisibleEnabled(labelDiscount, true);
                    ControlSetItsText(labelDiscount, resultOfCalculatingDiscount.ToString() + "%");

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
                        string message = "Счет для анализа выбран с некорректными парсерами." + Environment.NewLine +
                                         "Количество этих параметров должны быть одинаковое и больше нуля:" + Environment.NewLine +
                                         "'" + p[1] + @"' =  " + countParser[p[1]] + Environment.NewLine +
                                         "'" + p[2] + @"' =  " + countParser[p[2]] + Environment.NewLine +
                                         "'" + p[3] + @"' =  " + countParser[p[3]];
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
            else { return false; }

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

        private static double CalculateTax(double valueBeforeTaxes)
        { return valueBeforeTaxes * 0.2; }

        private static double CalculatePf(double valueBeforeTaxes)
        { return valueBeforeTaxes * 0.075; }

        private void ParseStringsOfPreparedListIntoTable() //Парсинг строк и передача результата текстовый редактор
        {
            ToolStripStatusLabelSetText(StatusLabel1, Properties.Resources.WorkingWithData);
            dataStart = labelPeriod.Text.Split('-')[0].Trim(); // дата начала периода счета
            dataEnd = labelPeriod.Text.Split('-')[1].Trim();  // дата конца периода счета

            DataRow row;
            bool isUsedCurrent = false;
            bool isCheckFinishedTitles = false;

            string temp, searchNumber;
            string[] substrings;

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
                        if (mcpCurrent.contractName?.Length > 1)
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
                        temp = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.monthCost = Convert.ToDouble(Regex.Replace(temp, "[,]",
                            System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * amountBillAfterDiscount * 1.275;
                    }
                    else if (s.Contains(p[5]))
                    {
                        substrings = s.Split(' ');
                        temp = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.roming = Convert.ToDouble(Regex.Replace(temp, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275;
                    }
                    else if (s.Contains(p[6]))
                    {
                        substrings = s.Split(' ');
                        temp = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.discount = Convert.ToDouble(Regex.Replace(temp, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
                    }
                    else if (s.Contains(p[7]))
                    {
                        substrings = s.Split(' ');
                        temp = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.totalCost = Convert.ToDouble(Regex.Replace(temp, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
                        isCheckFinishedTitles = true;
                        isUsedCurrent = false;
                    }
                    else if (s.Contains(p[11]))
                    {
                        substrings = s.Split(' ');
                        temp = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.romingData += Convert.ToDouble(Regex.Replace(temp, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275;
                    }
                    else if (s.Contains(p[12]))
                    {
                        substrings = s.Split(' ');
                        temp = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.extraInternetOrdered += Convert.ToDouble(Regex.Replace(temp, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275 * amountBillAfterDiscount;
                    }
                    else if (s.Contains(p[13]))
                    {
                        substrings = s.Split(' ');
                        temp = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.outToCity += Convert.ToDouble(Regex.Replace(temp, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275 * amountBillAfterDiscount;
                    }
                    else if (s.Contains(p[14]))
                    {
                        substrings = s.Split(' ');
                        temp = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.extraService += Convert.ToDouble(Regex.Replace(temp, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
                    }
                    else if (s.Contains(p[15]))
                    {
                        substrings = s.Split(' ');
                        temp = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.content += Convert.ToDouble(Regex.Replace(temp, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275;
                    }
                    else if (s.Contains(p[23]))
                    {
                        substrings = s.Split(' ');
                        temp = substrings[substrings.Length - 1].Trim();
                        mcpCurrent.extraServiceOrdered += Convert.ToDouble(Regex.Replace(temp, "[,]", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator)) * 1.275 * amountBillAfterDiscount;
                    }
                    else if (isCheckFinishedTitles)
                    { isUsedCurrent = true; }
                }

                //additional payment for detalisation (at the end of the current bill)
                mcpCurrent = new MobileContractPerson
                {
                    totalCost = Math.Abs(BillDeliveryCost * amountBillAfterDiscount),
                    discount = Math.Abs(BillDeliveryCostDiscount)
                };
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
            catch (Exception Expt) { MessageBox.Show(Expt.ToString(), Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error); }

            listTempContract.Clear();
        }


        private DataTable GetDataWithModel()  // получение данных из базы ТФактура
        {
            DataTable dt = dtOwnerOfMobileWithinSelectedPeriod.Clone();

            string dataFromLabel = ControlReturnText(labelPeriod);
            dataStart = dataFromLabel.Split('-')[0].Trim(); //'01.05.2018'
            dataEnd = dataFromLabel.Split('-')[1].Trim();  //'31.05.2018'
            string dataStartSearch = dataStart.Split('.')[2] + "-" + dataStart.Split('.')[1] + "-" + dataStart.Split('.')[0]; //'2018-05-01'
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
                                   " ) t3 ON t1.contract_id = t3.contract_id AND t1.emp_id = t3.emp_id" +
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

#pragma warning disable CA2100 // Review SQL queries for security vulnerabilities
                    using (System.Data.SqlClient.SqlCommand sqlCommand = new System.Data.SqlClient.SqlCommand(sSqlQuery, sqlConnection))
#pragma warning restore CA2100 // Review SQL queries for security vulnerabilities
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

                                    DataRow row = dt.NewRow();
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
                                    dt.Rows.Add(row);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception expt) { MessageBox.Show(expt.ToString()); }
            return dt;
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

#pragma warning disable CA1307 // Specify StringComparison
            if (sTemp2.StartsWith("+") && sTemp2.Length == 13) sPhone = sTemp2;
            else if (sTemp2.StartsWith("380") && sTemp2.Length == 12) sPhone = "+" + sTemp2;
            else if (sTemp2.StartsWith("80") && sTemp2.Length == 11) sPhone = "+3" + sTemp2;
            else if (sTemp2.StartsWith("0") && sTemp2.Length == 10) sPhone = "+38" + sTemp2;
#pragma warning restore CA1307 // Specify StringComparison
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
                            strNewModels += i + ". \"" + str + Environment.NewLine;
                            sb.AppendLine(i + ". \"" + str + "\"");
                        }
                    }
                    sb.AppendLine(@"");

                    sb.ToString().WriteAtFile(pathToNewModels);
                    sbError.ToString().AppendAtFile(pathToNewModels);
                }
                catch (Exception e)
                { MessageBox.Show(e.ToString(), e.Message, MessageBoxButtons.OK, MessageBoxIcon.Error); }

                infoStatusBar = "В базе найдены новые, не добавленные ранее, модели компенсации затрат сотрудников!";

                DialogResult result = MessageBox.Show(
                    "В базе со счетами мобильного оператора на сервере " + pConnectionServer + " найдены не существующие в программе модели компенсации затрат сотрудников!" +
                    Environment.NewLine + strNewModels + Environment.NewLine +
                    "Для их учета необходимо, внести изменения в модели рассчета в программе!" + Environment.NewLine +
                    "Для прерывания дальнейших рассчетов нажмите кнопку" + Environment.NewLine + "\"Yes\"(Да)" + Environment.NewLine +
                    "для продолжения:" + Environment.NewLine + "\"No\"(Нет)",
                    Properties.Resources.Attention,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button1);
                if (result == DialogResult.Yes)
                { newModels = true; }
            }
        }


        //Access to Control from other threads
        private string OpenFileDialogReturnPath(OpenFileDialog ofd) //Return its name 
        {
            string filePath = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    ofd.FileName = @"";
                    ofd.Filter = Properties.Resources.OpenDialogTextFiles;
                    ofd.ShowDialog();
                    filePath = ofd.FileName;
                }));
            else
            {
                ofd.FileName = @"";
                ofd.Filter = Properties.Resources.OpenDialogTextFiles;
                ofd.ShowDialog();
                filePath = ofd.FileName;
            }
            return filePath;
        }

        private void ProgressWork1Step(string text = "") //add into progressBar Value 2 from other threads
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate
                {
                    if (ProgressBar1.Value > 99)
                    { ProgressBar1.Value = 0; }
                    ProgressBar1.Maximum = 100;
                    ProgressBar1.Value += 1;
                    if (text.Length > 0)
                        ToolStripStatusLabelSetText(StatusLabel1, text);
                }));
            else
            {
                if (ProgressBar1.Value > 99)
                { ProgressBar1.Value = 0; }
                ProgressBar1.Maximum = 100;
                ProgressBar1.Value += 1;
                if (text.Length > 0)
                    ToolStripStatusLabelSetText(StatusLabel1, text);
            }
        }

        private void ProgressBar1Start() //Set progressBar Value into 0 from other threads
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

        private void ProgressBar1Stop() //Set progressBar Value into 100 from other threads
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


        private string ControlReturnText(Control controlText) //Return its name 
        {
            string tBox = "";
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { tBox = controlText.Text.Trim(); }));
            else
                tBox = controlText.Text.Trim();
            return tBox;
        }

        private void ControlSetItsText(Control control, string text) //Set its name 
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.Text = text; }));
            else
                control.Text = text;
        }

        private void ControlVisibleEnabled(Control control, bool visible) //Set its name 
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.Visible = visible; }));
            else
                control.Visible = visible;
        }

        private void ToolStripStatusLabelSetText(ToolStripStatusLabel control, string text) //Set its name 
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.Text = text; }));
            else
                control.Text = text;
        }

        private void ToolStripMenuItemEnabled(ToolStripMenuItem control, bool enabled) //Set its name 
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { control.Enabled = enabled; }));
            else
                control.Enabled = enabled;
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
                getValue = (string[])EvUserKey?.GetValue(Properties.Resources.ListOfServices);
                if (getValue?.Length > 0)
                {
                    foreach (string line in getValue)
                    {
                        if (!string.IsNullOrWhiteSpace(line))
                            listSavedServices.Add(line.Trim());
                    }
                    foundSavedData = true;
                }

                getValue = (string[])EvUserKey?.GetValue(Properties.Resources.ListOfNumbers);
                if (getValue?.Length > 0)
                {
                    foreach (string line in getValue)
                    {
                        if (!string.IsNullOrWhiteSpace(line))
                        { listSavedNumbers.Add(line.Trim()); }
                    }
                    ControlSetItsText(labelContracts, listSavedNumbers.Count.ToString() + " шт.");
                    ControlVisibleEnabled(labelContracts, true);
                    foundSavedData = true;
                }

                strSavedPathToInvoice = (string)EvUserKey?.GetValue("PathToLastInvoice");
                if (strSavedPathToInvoice?.Trim()?.Length > 3)
                { prepareBillItem.Enabled = true; }

                string period = (string)EvUserKey?.GetValue("PeriodLastInvoice");
                if (period?.Length > 6)
                {
                    ControlSetItsText(labelPeriod, period);
                    ControlVisibleEnabled(labelPeriod, true);
                }

                if (listSavedServices?.Count > 0 || listSavedNumbers?.Count > 0)
                {
                    sb.AppendLine("-= Данные для генерации маркетингового отчета =-");
                    sb.AppendLine(Properties.Resources.RowDozenOfEqualSymbols);
                }

                if (listSavedServices?.Count > 0)
                {
                    selectedServices = true;
                    sb.AppendLine(Properties.Resources.ListOfServices);
                    foreach (string service in listSavedServices)
                    { sb.AppendLine(service); }
                    sb.AppendLine(Properties.Resources.RowDozenOfEqualSymbols);
                }

                if (listSavedNumbers?.Count > 0)
                {
                    selectedNumbers = true;
                    sb.AppendLine(Properties.Resources.ListOfNumbers);
                    foreach (string number in listSavedNumbers)
                    {
                        sb.AppendLine(number);
                    }
                    sb.AppendLine(Properties.Resources.RowDozenOfEqualSymbols);
                }
            }

            textBoxLog.AppendLine(sb.ToString());
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
                catch (Exception expt) { MessageShow("Ошибки с доступом для записи списка " + parameterName + " в реестр. Данные не сохранены.\n"+ expt.Message); }
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

                    if (ControlReturnText(labelPeriod).Length > 0)
                    { EvUserKey.SetValue("PeriodLastInvoice", periodInvoice, Microsoft.Win32.RegistryValueKind.String); }
                }
                foundSavedData = true;
            }
            catch (Exception expt) { MessageShow("Ошибки с доступом для записи пути к счету. Данные сохранены не корректно.\n"+expt.Message); }
        }
    }
}
