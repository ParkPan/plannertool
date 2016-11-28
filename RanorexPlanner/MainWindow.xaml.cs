using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
//using System.Windows.Shapes;
using System.IO;
using System.Data.SqlClient;
using System.Collections.ObjectModel;
using TaskScheduler;
using System.Runtime.InteropServices;
using System.Xml;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;
using System.Threading;
using System.ComponentModel;
using System.Xml.Linq;
using System.Threading.Tasks;
using System.Collections.Specialized;

namespace RanorexPlanner
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string siteName = null;
        private string siteIP = null;
        private int siteID;
        private string currentUser = null;
        private string currentUserEmail = null;
        private ObservableCollection<TestCaseInfo> allCaseListInSite = new ObservableCollection<TestCaseInfo>();
        private List<TestCaseInfo> tmpAllCases = new List<TestCaseInfo>();
        private ObservableCollection<TestCaseInfo> filterByConfigCaseList = new ObservableCollection<TestCaseInfo>();
        private List<TestCaseInfo> testCaseList = new List<TestCaseInfo>();
        private List<string> caseConfigList = new List<string>();
        private const string CASEINFOXMLFILE = "TfsInformation.xml";
        private const string METADATAFILE = "Document.MetaData";
        private const string TARANOREXSCHEDULENAME = "RanorexTAExecutor";
        private const string RANOREXTALOCALROOTDIR = @"D:\TA_Ranorex";
        private const string EMAILHOST = "140.231.210.53";  //"SHAI001A.WW005.SIEMENS.NET";  //"CNSHA01075.CN001.SIEMENS.NET";
        private const string TAPRESETTINGDIR = @"C:\HELIUM\PlanEnv";
        private const string TAPRESETTINGSCHEDULENAME = "RreTASetting";
        private const string TAKILLEXECUTORDIR = @"C:\HELIUM\KillExecutor";
        private const string TAKILLEXECUTORSCHEDULENAME = "TAKillExecutorProcess";
        private bool metaDataInfoChange = false;
        private bool caseOrderChange = false;
        private Dictionary<string, List<TestCaseInfo>> caseAuthors = new Dictionary<string, List<TestCaseInfo>>();
        private Dictionary<string, string> authorEmail = new Dictionary<string, string>();

        private string filterConfig, filterPart;
        private List<string> casePartList = new List<string>();

        private string toolSiteIP = null;
        private string toolSiteName = null;
        private ObservableCollection<Node> ToolNodes;
        private readonly List<BackgroundWorker> myActiveWorkers = new List<BackgroundWorker>();
        private const string TOOLSCHEDULENAME = "RanorexToolExecutor";
        private const string TATOOLEXECUTORDIR = @"C:\HELIUM\ToolExecutor";

        private const int TASKCOUNT = 12;
        private string lockObj = "thisIsALock";

        private AutoResetEvent autoEvent;

        public MainWindow()
        {
            InitializeComponent();
            allTestCaseLV.ItemsSource = allCaseListInSite;
            ToolNodes = new ObservableCollection<Node>();
            autoEvent = new AutoResetEvent(false);
            casePartList.Add("all");
            casePartList.Add("hz_taf_scanprotocols");
            casePartList.Add("hz_taf_generalhazard");
            casePartList.Add("hz_taf_imageorientation");
            casePartList.Add("taf_autorangeperformance");
            casePartList.Add("taf_autoregionscan");
            casePartList.Add("tap_performance");
            casePartList.Add("taf_scanguide");
            casePartList.Add("taf_tubepower");
            casePartList.Add("taf_generalui");
            casePartList.Add("taf_scanprotocols");
            casePartList.Add("taf_modeparameter");
            cmbPartFilter.ItemsSource = casePartList;
            cmbPartFilter.SelectedIndex = 0;
            //immediacyBtn.IsEnabled = false;
            //StartBtn.IsEnabled = false;
        }

        private void combxSiteName_SelectChange(object sender, SelectionChangedEventArgs e)
        {
            List<string> tmpAuthorNotInDB = new List<string>();
            labelVersion.Text = "";
            filterConfig = "";
            filterPart = "";
            txtProductNum.Text = "";
            allCaseListInSite.Clear();
            tmpAllCases.Clear();
            caseConfigList.Clear();
            testCaseList.Clear();
            bePlanedCaseLB.Items.Clear();
            cmbConfigFilter.Items.Clear();
            statusBar.Items.Clear();
            string tmpDirStr = null;
            allTestCaseLV.ItemsSource = allCaseListInSite;
            cmbPartFilter.SelectedIndex = 0;
            if (combxSiteName.SelectedItem.ToString().Equals("For-Refresh"))
                return;
            siteIP = ((TestSiteInfo)combxSiteName.SelectedItem).Value;
            siteName = ((TestSiteInfo)combxSiteName.SelectedItem).Text;
            siteID = ((TestSiteInfo)combxSiteName.SelectedItem).ID;
            ExecuteCommand(@"net use \\" + siteIP + @" ASDzxc123!@#456 /user:SYADMIN");
            using (IdentityScope identityScope = new IdentityScope("SYADMIN", "ASDzxc123!@#456", siteIP))
            {
                string taRootDirPath = @"\\" + siteIP + @"\D$\TA_Ranorex";
                string taCaseDirPath = null;
                // Initialize case list info to combobox and MetaData info to textfields
                #region
                try
                {
                    //XmlDocument xmlDoc = new XmlDocument();
                    // Initialize case list
                    string[] tempSubDirs = Directory.GetDirectories(taRootDirPath);
                    foreach (string tmpSubDir in tempSubDirs)
                    {
                        tmpDirStr = System.IO.Path.GetFileName(tmpSubDir);
                        if (!tmpDirStr.Equals("Resources") && !tmpDirStr.Equals("TestApplication"))
                        {
                            taCaseDirPath = tmpDirStr;
                            break;
                        }
                    }
                    if (taCaseDirPath == null || taCaseDirPath.Equals(""))
                    {
                        MessageBox.Show("TA case director does not exist, please check...", "Ranorex Planner");
                        return;
                    }
                    string[] tempCaseDirs = Directory.GetDirectories(taRootDirPath + System.IO.Path.DirectorySeparatorChar + taCaseDirPath);
                    //foreach (string tmpCaseDir in tempCaseDirs)
                    Parallel.ForEach<string>(tempCaseDirs, tmpCaseDir =>
                    {
                        lock (lockObj)
                        {
                            XmlDocument xmlDoc1 = new XmlDocument();
                            tmpDirStr = System.IO.Path.GetFileName(tmpCaseDir);
                            if (!tmpDirStr.StartsWith("Plan_"))
                                return;//continue;
                            TestCaseInfo tmpCaseInfo = new TestCaseInfo();
                            Match match = Regex.Match(tmpDirStr, @"^Plan_\d+-\w+_\d+-(?<casename>\w+)_\d+.*");
                            if (match.Success)
                                tmpCaseInfo.CaseName = match.Groups["casename"].Value;
                            else
                            {
                                MessageBox.Show("case directory: [" + tmpDirStr + "] name is not correct format, ignore this case and connect case author please...", "Ranorex Planner");
                                tmpCaseInfo = null;
                                return;//continue;
                            }
                            tmpCaseInfo.CasePath = tmpCaseDir;
                            //lock (lockObj)
                            //{
                            xmlDoc1.Load(tmpCaseDir + System.IO.Path.DirectorySeparatorChar + CASEINFOXMLFILE);
                            //}
                            XmlNodeList xmlNodeList = xmlDoc1.GetElementsByTagName("TestConfigurationName");
                            tmpCaseInfo.CaseConfiguration = xmlNodeList[0].InnerText;
                            xmlNodeList = xmlDoc1.GetElementsByTagName("TestCaseAutor");
                            tmpCaseInfo.CaseAuthor = xmlNodeList[0].InnerText;
                            xmlNodeList = xmlDoc1.GetElementsByTagName("OrderId");
                            try
                            {
                                tmpCaseInfo.CaseOrderID = Convert.ToInt32(xmlNodeList[0].InnerText);
                            }
                            catch
                            {
                                tmpCaseInfo.CaseOrderID = -1;
                                MessageBox.Show("Convert case [" + tmpCaseInfo.CaseName + "] orderID to number type failed...", "Ranorex Planner");
                            }
                            xmlNodeList = xmlDoc1.GetElementsByTagName("Results");
                            if (!String.IsNullOrEmpty(xmlNodeList[0].InnerText))
                            {
                                xmlNodeList = xmlDoc1.GetElementsByTagName("State");
                                tmpCaseInfo.CaseState = xmlNodeList[xmlNodeList.Count - 1].InnerText;
                            }
                            else
                            {
                                tmpCaseInfo.CaseState = "";
                            }
                            if (!caseAuthors.Keys.Contains(tmpCaseInfo.CaseAuthor))
                            {
                                if (!tmpAuthorNotInDB.Contains(tmpCaseInfo.CaseAuthor))
                                {
                                    MessageBox.Show("The author [" + tmpCaseInfo.CaseAuthor + "] does not include in person list of DB, the email cannot be sent to " + tmpCaseInfo.CaseAuthor + "! Please contact admin to let author join to list of DB...", "Ranorex Planner");
                                    tmpAuthorNotInDB.Add(tmpCaseInfo.CaseAuthor);
                                }
                            }
                            tmpAllCases.Add(tmpCaseInfo);
                            //allCaseListInSite.Add(tmpCaseInfo);
                            //if (!caseConfigList.Contains(tmpCaseInfo.CaseConfiguration))
                            //    caseConfigList.Add(tmpCaseInfo.CaseConfiguration);
                            xmlDoc1 = null;
                        }
                    });
                    tmpAllCases.OrderBy(item => item.CaseName).ToList<TestCaseInfo>().ForEach(tmpItem => { allCaseListInSite.Add(tmpItem); });
                    allCaseListInSite.ToList<TestCaseInfo>().ForEach(tmpItem => { if (!caseConfigList.Contains(tmpItem.CaseConfiguration)) caseConfigList.Add(tmpItem.CaseConfiguration); });
                    caseConfigList.Sort();
                    // Initialize MetaData
                    XmlDocument xmlDoc = new XmlDocument();
                    if (File.Exists(taRootDirPath + System.IO.Path.DirectorySeparatorChar + "TestApplication" + System.IO.Path.DirectorySeparatorChar + METADATAFILE))
                    {
                        xmlDoc.Load(taRootDirPath + System.IO.Path.DirectorySeparatorChar + "TestApplication" + System.IO.Path.DirectorySeparatorChar + METADATAFILE);
                        XmlNodeReader reader = new XmlNodeReader(xmlDoc);
                        while (reader.Read())
                        {
                            switch (reader.NodeType)
                            {
                                case XmlNodeType.Element:
                                    if (reader.Name.Equals("Department"))
                                        txtDepartment.Text = reader.ReadElementContentAsString();
                                    if (reader.Name.Equals("Organization"))
                                        txtOrganization.Text = reader.ReadElementContentAsString();
                                    if (reader.Name.Equals("ProductName"))
                                        txtProductName.Text = reader.ReadElementContentAsString();
                                    if (reader.Name.Equals("ProjectNumber"))
                                        txtProductNum.Text = reader.ReadElementContentAsString();
                                    if (reader.Name.Equals("SoftwareVersion"))
                                    {
                                        string tmpSWVersion = reader.ReadElementContentAsString();
                                        if (tmpSWVersion.Equals("#AUTODETECT#"))
                                        {
                                            txtSoftVersion.Text = "";
                                            cbxSoftVersion.IsChecked = true;
                                        }
                                        else
                                        {
                                            txtSoftVersion.Text = tmpSWVersion;
                                            cbxSoftVersion.IsChecked = false;
                                        }
                                    }
                                    if (reader.Name.Equals("Language"))
                                        txtLanguage.Text = reader.ReadElementContentAsString();
                                    if (reader.Name.Equals("TracesDisabled"))
                                        txtTrace.Text = reader.ReadElementContentAsString();
                                    break;
                            }
                        }
                    }
                    xmlDoc = null;
                    metaDataInfoChange = false;
                    // Get SW version of system and display on UI
                    labelVersion.Text = getSWVersion();
                    // Get trace info from remote test site and confirm local info on PC
                    #region Get trace info and Confirm
                    string tmpTraceInfo = getTestSiteTraceInfo();
                    if (tmpTraceInfo != null && !tmpTraceInfo.Equals(""))
                    {
                        tmpTraceInfo = tmpTraceInfo.Trim().Split(new string[] { "    " }, StringSplitOptions.None)[2];
                        if (!"Y".Equals(tmpTraceInfo.Trim().ToUpper()) && !"YES".Equals(tmpTraceInfo.Trim().ToUpper()))
                        {
                            if (!txtTrace.Text.ToUpper().Equals("YES"))
                            {
                                txtTrace.Text = "Yes";
                                metaDataInfoChange = true;
                            }
                            MessageBox.Show("Get remote test site CTTrace is OFF, it should be ON. Please check trace info by manual with controller...", "Ranorex Planner");
                        }
                        else
                        {
                            if (!txtTrace.Text.ToUpper().Equals("NO"))
                            {
                                txtTrace.Text = "No";
                                metaDataInfoChange = true;
                                MessageBox.Show("Get remote test site CTTrace info is different with old value. Please check...", "Ranorex Planner");
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Get remote test site CTTrace info failed, default value for it now. It's better to check trace info by manual...", "Ranorex Planner");
                    }
                    #endregion
                    // Get system version from remote test site and confirm local info on PC
                    #region Get system version and Confirm
                    string tmpSystemVersion = getTestSiteSystemVersion();
                    if (tmpSystemVersion != null && !tmpSystemVersion.Equals(""))
                    {
                        try
                        {
                            tmpSystemVersion = tmpSystemVersion.Trim().Split(new string[] { "    " }, StringSplitOptions.None)[2];
                            if (tmpSystemVersion.Contains("18"))
                                tmpSystemVersion = tmpSystemVersion.Replace("18", "15") + "(2010)";
                        }
                        catch
                        {
                            tmpSystemVersion = "";
                        }
                        if (String.IsNullOrEmpty(txtProductNum.Text))
                        {
                            txtProductNum.Text = tmpSystemVersion + " ICS";
                            metaDataInfoChange = true;
                        }
                        else
                        {
                            if (!String.IsNullOrEmpty(tmpSystemVersion))
                            {
                                if (!tmpSystemVersion.Equals(txtProductNum.Text))
                                {
                                    MessageBox.Show("Get remote test site System Version [" + tmpSystemVersion + "] is different with old value [" + txtProductNum.Text + "], the remote value will be used. Please check...", "Ranorex Planner");
                                    txtProductNum.Text = tmpSystemVersion + " ICS";
                                    metaDataInfoChange = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Get remote test site System Version (Product Number) failed, default value for it now. It's better to check trace info by manual...", "Ranorex Planner");
                    }
                    #endregion
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Get TA case director failed, please check director path exist..., more exception info: " + ex.Message, "Ranorex Planner");
                    return;
                }
                #endregion
                // Initialize config info to combobox
                #region
                if (caseConfigList.Count > 1)
                    cmbConfigFilter.Items.Add("all");
                foreach (string tmpCaseConfig in caseConfigList)
                    cmbConfigFilter.Items.Add(tmpCaseConfig);
                #endregion

                tmpAuthorNotInDB = null;
            }
        }

        private void mainWin_init(object sender, EventArgs e)
        {
            // init combox of site name control
            #region combxSiteName
            combxSiteName.Items.Clear();
            combxSiteName.Items.Add("For-Refresh");
            List<TestSiteInfo> itemList = new List<TestSiteInfo>();
            string sqlStr = "Data Source=SHAI203A;Initial Catalog=SSME_Ranorex;User ID=zope_usr;pwd=zope_usr";
            string sqlComStr = "select TestSiteName, SiteIP, SiteType, TestSite_ID from TESTSITE where USED='Y'";
            using (SqlConnection sqlCon = new SqlConnection(sqlStr))
            using (SqlCommand sqlCom = new SqlCommand(sqlComStr, sqlCon))
            {
                sqlCon.Open();
                using (SqlDataReader sqlReader = sqlCom.ExecuteReader())
                {
                    while (sqlReader.Read())
                    {
                        TestSiteInfo tmpItem = new TestSiteInfo();
                        tmpItem.Text = sqlReader.GetString(0);
                        tmpItem.Value = sqlReader.GetString(1);
                        tmpItem.Type = sqlReader.IsDBNull(2) ? "" : sqlReader.GetString(2);
                        tmpItem.ID = sqlReader.GetInt32(3);
                        itemList.Add(tmpItem);
                    }
                    sqlReader.Close();
                    sqlReader.Dispose();
                }
                sqlCom.Dispose();
                sqlCon.Close();
                sqlCon.Dispose();
            }
            if (itemList.Count > 0)
            {
                itemList.Sort();
                foreach (TestSiteInfo item in itemList)
                {
                    combxSiteName.Items.Add(item);
                    combxToolTestSite.Items.Add(item);   // init test site combox control on tool tab page
                }
            }
            itemList.Clear();
            itemList = null;
            #endregion
            // init radDateTimePicker control
            string date = DateTime.Now.ToShortDateString();
            string time = DateTime.Now.ToShortTimeString();
            radDateTimePicker.DateTimeText = date + " " + time;
            radDateTimePickerTool.DateTimeText = date + " " + time;
            // init case author email dictionary list.
            #region authorEmail List
            sqlStr = "Data Source=SHAI203A;Initial Catalog=SSME_Ranorex;User ID=zope_usr;pwd=zope_usr";
            sqlComStr = "select AccountName, Email from TESTPERSON";
            using (SqlConnection sqlCon = new SqlConnection(sqlStr))
            using (SqlCommand sqlCom = new SqlCommand(sqlComStr, sqlCon))
            {
                sqlCon.Open();
                using (SqlDataReader sqlReader = sqlCom.ExecuteReader())
                {
                    while (sqlReader.Read())
                    {
                        if (!authorEmail.Keys.Contains(sqlReader.GetString(0)))
                        {
                            authorEmail.Add(sqlReader.GetString(0), sqlReader.GetString(1));
                            caseAuthors.Add(sqlReader.GetString(0), new List<TestCaseInfo>());
                        }
                    }
                    sqlReader.Close();
                    sqlReader.Dispose();
                }
                sqlCom.Dispose();
                sqlCon.Close();
                sqlCon.Dispose();
            }
            #endregion
            // init current user on operation PC
            #region current user on PC
            currentUser = Environment.UserName;
            sqlStr = "Data Source=SHAI203A;Initial Catalog=SSME_Ranorex;User ID=zope_usr;pwd=zope_usr";
            sqlComStr = "select Email from TESTPERSON where Account='" + Environment.UserName + "'";
            using (SqlConnection sqlCon = new SqlConnection(sqlStr))
            using (SqlCommand sqlCom = new SqlCommand(sqlComStr, sqlCon))
            {
                sqlCon.Open();
                using (SqlDataReader sqlReader = sqlCom.ExecuteReader())
                {
                    while (sqlReader.Read())
                    {
                        currentUserEmail = sqlReader.GetString(0);
                    }
                    sqlReader.Close();
                    sqlReader.Dispose();
                }
                sqlCom.Dispose();
                sqlCon.Close();
                sqlCon.Dispose();
            }
            #endregion
        }

        private void textChange_MetadataInfo(object sender, TextChangedEventArgs e)
        {
            metaDataInfoChange = true;
        }

        private void cbxSoftVersionClick_MetadataInfo(object sender, RoutedEventArgs e)
        {
            metaDataInfoChange = true;
        }

        private void cmbConfigFilter_SelectedChange(object sender, SelectionChangedEventArgs e)
        {
            if (cmbConfigFilter.Items.Count == 0)
                return;
            filterConfig = cmbConfigFilter.SelectedItem.ToString();
            filterByConfigCaseList.Clear();
            if (allCaseListInSite.Count > 0)
            {
                if ((filterConfig.Equals("all") || String.IsNullOrEmpty(filterConfig)) && filterPart.Equals("all"))
                    allTestCaseLV.ItemsSource = allCaseListInSite;
                else
                {
                    if (filterPart.Equals("all"))
                        filterByConfigCaseList = new ObservableCollection<TestCaseInfo>(allCaseListInSite.Where<TestCaseInfo>(item => item.CaseConfiguration.Equals(filterConfig)).ToList<TestCaseInfo>());
                    else if (filterConfig.Equals("all") || String.IsNullOrEmpty(filterConfig))
                        filterByConfigCaseList = new ObservableCollection<TestCaseInfo>(allCaseListInSite.Where<TestCaseInfo>(item => item.CaseName.StartsWith(filterPart)).ToList<TestCaseInfo>());
                    else
                        filterByConfigCaseList = new ObservableCollection<TestCaseInfo>(allCaseListInSite.Where<TestCaseInfo>(item => item.CaseName.StartsWith(filterPart) && item.CaseConfiguration.Equals(filterConfig)).ToList<TestCaseInfo>());
                    allTestCaseLV.ItemsSource = filterByConfigCaseList;
                }
            }
            else
                MessageBox.Show("No cases on current test site. IP:  " + siteIP, "Ranorex Planner");
        }

        private void cmbPartFilter_SelectedChange(object sender, SelectionChangedEventArgs e)
        {
            if (cmbPartFilter.Items.Count == 0)
                return;
            filterPart = cmbPartFilter.SelectedItem.ToString();
            filterByConfigCaseList.Clear();
            if (allCaseListInSite.Count > 0)
            {
                if ((filterConfig.Equals("all") || String.IsNullOrEmpty(filterConfig)) && filterPart.Equals("all"))
                    allTestCaseLV.ItemsSource = allCaseListInSite;
                else
                {
                    if (filterPart.Equals("all"))
                        filterByConfigCaseList = new ObservableCollection<TestCaseInfo>(allCaseListInSite.Where<TestCaseInfo>(item => item.CaseConfiguration.Equals(filterConfig)).ToList<TestCaseInfo>());
                    else if (filterConfig.Equals("all") || String.IsNullOrEmpty(filterConfig))
                        filterByConfigCaseList = new ObservableCollection<TestCaseInfo>(allCaseListInSite.Where<TestCaseInfo>(item => item.CaseName.StartsWith(filterPart)).ToList<TestCaseInfo>());
                    else
                        filterByConfigCaseList = new ObservableCollection<TestCaseInfo>(allCaseListInSite.Where<TestCaseInfo>(item => item.CaseName.StartsWith(filterPart) && item.CaseConfiguration.Equals(filterConfig)).ToList<TestCaseInfo>());
                    allTestCaseLV.ItemsSource = filterByConfigCaseList;
                }
            }
        }

        private void planCaseBT_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(siteIP))
            {
                MessageBox.Show("Please select one test site first... ", "Ranorex Planner");
                return;
            }

            if (String.IsNullOrEmpty(txtDepartment.Text.Trim()) || String.IsNullOrEmpty(txtOrganization.Text.Trim()) ||
                String.IsNullOrEmpty(txtProductName.Text.Trim()) || String.IsNullOrEmpty(txtProductNum.Text.Trim()) ||
                String.IsNullOrEmpty(txtLanguage.Text.Trim()) || String.IsNullOrEmpty(txtTrace.Text.Trim()) ||
                (cbxSoftVersion.IsChecked != true && (String.IsNullOrEmpty(txtSoftVersion.Text.Trim()))))
            {
                MessageBox.Show("Metadata Info message cannot be empty, please fill it... ", "Ranorex Planner");
                return;
            }

            ExecuteCommand(@"net use \\" + siteIP + @" ASDzxc123!@#456 /user:SYADMIN");

            var process = new Process();
            process.StartInfo.FileName = "PsExec.exe";
            //process.StartInfo.Arguments = string.Format(@"\\{0} -u meduser -p ASDzxc123!@#456 -d -i -w {1} {2}", siteIP, TAKILLEXECUTORDIR,
            //    TAKILLEXECUTORDIR + System.IO.Path.DirectorySeparatorChar + "KillExecutor.bat");
            //process.StartInfo.Arguments = string.Format(@"\\{0} -u meduser -p ASDzxc123!@#456 -d -i -w {1} {2}", siteIP, @"\\" + siteIP +@"\CtElog",
            //    @"D:\DATABASE\CtElog" + System.IO.Path.DirectorySeparatorChar + "CreateTAScheduler.exe createnew tesTest 9/10/2015-16:50:50");
            process.StartInfo.Arguments = string.Format(@"\\{0} -u syadmin -p ASDzxc123!@#456 -d -i -w {1} {2}", siteIP, @"D:\DATABASE\CtElog",
                @"D:\DATABASE\CtElog" + System.IO.Path.DirectorySeparatorChar + "CreateTAScheduler.exe createnew tesTest 9/10/2015-16:50:50");
            //process.StartInfo.Arguments = string.Format(@"\\{0} -u meduser -p ASDzxc123!@#456 -d -i -w {1} {2}", siteIP, @"C:\HELIUM",
            //    @"C:\HELIUM" + System.IO.Path.DirectorySeparatorChar + "CreateTAScheduler.exe createnew tesTest 9/10/2015-16:50:50");
            process.StartInfo.UseShellExecute = false;
            process.Start();
            //Thread.Sleep(2000);
            //process.StartInfo.FileName = "PsExec.exe";
            //process.StartInfo.Arguments = string.Format(@"\\{0} -u meduser -p meduser1 -d -i -w {1} {2}", siteIP, RANOREXTALOCALROOTDIR,
            //    RANOREXTALOCALROOTDIR + System.IO.Path.DirectorySeparatorChar + "TAOfflineTestExecutor.exe -arguments /immediately");
            //process.StartInfo.UseShellExecute = false;
            //process.Start();

            //if (bePlanedCaseLB.Items.Count > 0 && bePlanedCaseLB.Items.Count == testCaseList.Count)
            //{
            //    modifyCaseInfoByPlanList();

            //    //create the schedule task command on the test site
            //    #region create schedule task
            //    deleteTAScheduler();
            //    try
            //    {
            //        // Get remote test site local time and caculate time according current PC time to plan
            //        #region caculate remote PC time
            //        string remoteDate, remoteTime, scheduleTime;
            //        scheduleTime = radDateTimePicker.DateTimeText;
            //        if (getTestSiteLocalTime(out remoteDate, out remoteTime, siteIP))
            //        {
            //            caculateRemoteTimeByCurrent(ref scheduleTime, ref remoteDate, ref remoteTime);
            //        }
            //        else
            //        {
            //            remoteDate = "Get NO";
            //            remoteTime = "Remote Time";
            //            MessageBox.Show("Get remote time of test site [" + siteName + "] failed, take notice of using current PC time now... ", "Ranorex Planner");
            //        }
            //        #endregion
            //        DateTime tmpDT = covertSchedulerDateTime(scheduleTime);

            //        if (preSettingCbx.IsChecked == true)
            //        {
            //            if (!createTAEnvironmentTask(tmpDT))
            //            {
            //                MessageBox.Show("Create pre-Setting Task command failed on [" + siteName + "] test site... ", "Ranorex Planner");
            //            }
            //            tmpDT = tmpDT.AddMinutes(15.0);
            //        }

            //        if (!createTAScheduler(tmpDT))
            //        {
            //            MessageBox.Show("Create Schedule Task command failed on [" + siteName + "] test site... ", "Ranorex Planner");
            //            return;
            //        }
            //        else
            //        {
            //            MessageBox.Show("Schedule Task has been created on [" + siteName + "] test site... ", "Ranorex Planner");
            //            ParameterizedThreadStart paraThread = new ParameterizedThreadStart(logInfoThisPlan);
            //            Thread tmpThread = new Thread(paraThread);
            //            if (!remoteDate.Equals("Get NO") && !remoteDate.Equals("Caculate Fail For"))
            //                tmpThread.Start(scheduleTime);
            //            else
            //                tmpThread.Start(scheduleTime + " (PC time)");
            //        }
            //    }
            //    catch
            //    {
            //        deleteTAScheduler();
            //        MessageBox.Show("Exception... when Create Schedule Task on [" + siteName + "] test site... ", "Ranorex Planner");
            //        return;
            //    }
            //    #endregion

            //    //send email to authors
            //    #region send email
            //    if (MessageBox.Show("Do you want to email case authors? :)... ", "Ranorex Planner", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            //    {
            //        //Thread thread = new Thread(sendEmailToAuthors);
            //        //thread.Start();
            //        ParameterizedThreadStart paramThread = new ParameterizedThreadStart(sendEmailToAuthors);
            //        Thread thread = new Thread(paramThread);
            //        object param = (object)(siteName + "=>" + siteIP);
            //        thread.Start(param);
            //    }
            //    #endregion
            //}
            //else
            //{
            //    if (bePlanedCaseLB.Items.Count == 0)
            //        MessageBox.Show("No cases have been selected on [" + siteName + "] test site... ", "Ranorex Planner");
            //    else
            //        MessageBox.Show("Case number of list in UI is not equal number list in array, please contact tool admin to solve this issue... ", "Ranorex Planner");
            //}
        }

        private void immediacyBtn_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(siteIP))
            {
                MessageBox.Show("Please select one test site first... ", "Ranorex Planner");
                return;
            }

            string startupPath = System.Windows.Forms.Application.StartupPath;

            if (bePlanedCaseLB.Items.Count > 0 && bePlanedCaseLB.Items.Count == testCaseList.Count)
            {
                modifyCaseInfoByPlanList();
                try
                {
                    //ExecuteCommand(@"net use \\" + siteIP + @" ASDzxc123!@#456 /user:SYADMIN");

                    //var process = new Process();
                    //process.StartInfo.FileName = startupPath + System.IO.Path.DirectorySeparatorChar + "PsExec.exe";
                    //process.StartInfo.Arguments = string.Format(@"\\{0} -u meduser -p meduser1 -d -i -w {1} {2}", siteIP, TAKILLEXECUTORDIR,
                    //    TAKILLEXECUTORDIR + System.IO.Path.DirectorySeparatorChar + "KillExecutor.bat");
                    //process.StartInfo.UseShellExecute = false;
                    //process.Start();
                    //Thread.Sleep(2000);
                    //process.StartInfo.FileName = startupPath + System.IO.Path.DirectorySeparatorChar + "PsExec.exe";
                    //process.StartInfo.Arguments = string.Format(@"\\{0} -u meduser -p meduser1 -d -i -w {1} {2}", siteIP, RANOREXTALOCALROOTDIR,
                    //    RANOREXTALOCALROOTDIR + System.IO.Path.DirectorySeparatorChar + "TAOfflineTestExecutor.exe -arguments /immediately");
                    //process.StartInfo.UseShellExecute = false;
                    //process.Start();

                    deleteTAScheduler();
                    try
                    {
                        if (!createTAScheduler(DateTime.Now, true, true))
                        {
                            MessageBox.Show("Schedule Task is running failed on [" + siteName + "] test site... ", "Ranorex Planner");
                            return;
                        }
                        else
                        {
                            MessageBox.Show("Schedule Task is running on [" + siteName + "] test site... ", "Ranorex Planner");
                            ParameterizedThreadStart paraThread = new ParameterizedThreadStart(logInfoThisPlan);
                            Thread tmpThread = new Thread(paraThread);
                            tmpThread.Start("Start Now");
                        }
                    }
                    catch
                    {
                        deleteTAScheduler();
                        MessageBox.Show("Exception... when running on [" + siteName + "] test site... ", "Ranorex Planner");
                        return;
                    }

                    //send email to authors
                    #region send email
                    if (MessageBox.Show("Do you want to email case authors? :)... ", "Ranorex Planner", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        ParameterizedThreadStart paramThread = new ParameterizedThreadStart(sendEmailToAuthors);
                        Thread thread = new Thread(paramThread);
                        object param = (object)(siteName + "=>" + siteIP);
                        thread.Start(param);
                    }
                    #endregion
                }
                catch
                {
                    MessageBox.Show("Exception... when Execute Task immediately on [" + siteName + "] test site... ", "Ranorex Planner");
                }
            }
            else
            {
                if (bePlanedCaseLB.Items.Count == 0)
                    MessageBox.Show("No cases have been selected on [" + siteName + "] test site... ", "Ranorex Planner");
                else
                    MessageBox.Show("Case number of list in UI is not equal number list in array, please contact tool admin to solve this issue... ", "Ranorex Planner");
            }
        }

        private void deleteTaskBT_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(siteIP))
            {
                MessageBox.Show("Please select one test site first... ", "Ranorex Planner");
                return;
            }
            if (deleteTAScheduler())
                MessageBox.Show("Delete schedule task successfully on [" + siteName + "] test site... ", "Ranorex Planner");
            else
                MessageBox.Show("Delete schedule task unsuccessfully  on [" + siteName + "] test site... ", "Ranorex Planner");
        }

        private void addCaseBT_Click(object sender, RoutedEventArgs e)
        {
            List<TestCaseInfo> tmpList = new List<TestCaseInfo>(testCaseList);
            if (allTestCaseLV.SelectedItems.Count > 0)
            {
                foreach (TestCaseInfo tmpCase in allTestCaseLV.SelectedItems)
                    if (!tmpList.Contains(tmpCase))
                    {
                        tmpCase.OriginalState = tmpCase.CaseState;
                        tmpCase.CaseState = "Planned";
                        tmpList.Add(tmpCase);
                    }
            }
            //System.ComponentModel.ICollectionView view = CollectionViewSource.GetDefaultView(allTestCaseLV.ItemsSource);
            //view.Refresh();
            allTestCaseLV.Items.Refresh();
            bePlanedCaseLB.Items.Clear();
            testCaseList.Clear();
            tmpList.OrderBy(item => item.CaseOrderID).ToList<TestCaseInfo>().ForEach(item => testCaseList.Add(item));
            foreach (TestCaseInfo tempCase in testCaseList)
                bePlanedCaseLB.Items.Add(tempCase);
            statusBar.Items.Clear();
            TextBlock txt = new TextBlock();
            txt.Text = testCaseList.Count + " cases are selected to plan";
            statusBar.Items.Add(txt);
        }

        private void deleteCaseBT_Click(object sender, RoutedEventArgs e)
        {
            List<TestCaseInfo> tmpList = new List<TestCaseInfo>(testCaseList);
            if (bePlanedCaseLB.SelectedItems.Count > 0)
            {
                foreach (TestCaseInfo tmpCase in bePlanedCaseLB.SelectedItems)
                {
                    if (tmpList.Contains(tmpCase))
                    {
                        tmpCase.CaseState = tmpCase.OriginalState;
                        tmpList.Remove(tmpCase);
                    }
                }
            }
            allTestCaseLV.Items.Refresh();
            bePlanedCaseLB.Items.Clear();
            testCaseList.Clear();
            tmpList.OrderBy(item => item.CaseOrderID).ToList<TestCaseInfo>().ForEach(item => testCaseList.Add(item));
            foreach (TestCaseInfo tempCase in testCaseList)
                bePlanedCaseLB.Items.Add(tempCase);
            statusBar.Items.Clear();
            TextBlock txt = new TextBlock();
            txt.Text = testCaseList.Count + " cases are selected to plan";
            statusBar.Items.Add(txt);
        }

        private void MoveUpBT_Click(object sender, RoutedEventArgs e)
        {
            if (bePlanedCaseLB.Items.Count > 0 && bePlanedCaseLB.SelectedItems.Count > 0)
            {
                TestCaseInfo tempCase = (TestCaseInfo)bePlanedCaseLB.SelectedItem;
                int itemIndex = bePlanedCaseLB.SelectedIndex;
                if (itemIndex == 0)
                    return;
                bePlanedCaseLB.Items.RemoveAt(itemIndex);
                bePlanedCaseLB.Items.Insert(itemIndex - 1, tempCase);
                bePlanedCaseLB.SelectedIndex = itemIndex - 1;
                caseOrderChange = true;
            }
        }

        private void MoveDownBT_Click(object sender, RoutedEventArgs e)
        {
            if (bePlanedCaseLB.Items.Count > 0 && bePlanedCaseLB.SelectedItems.Count > 0)
            {
                TestCaseInfo tempCase = (TestCaseInfo)bePlanedCaseLB.SelectedItem;
                int itemIndex = bePlanedCaseLB.SelectedIndex;
                if (itemIndex == bePlanedCaseLB.Items.Count - 1)
                    return;
                bePlanedCaseLB.Items.RemoveAt(itemIndex);
                bePlanedCaseLB.Items.Insert(itemIndex + 1, tempCase);
                bePlanedCaseLB.SelectedIndex = itemIndex + 1;
                caseOrderChange = true;
            }
        }

        private void doCreateRunIDFiles(object args)
        {
            TaskParamters paramters = args as TaskParamters;
            if (paramters == null)
                return;
            List<TestCaseInfo> tmpPlannedCases = paramters.TestCasesQueue;
            using (IdentityScope identityScope = new IdentityScope("SYADMIN", "ASDzxc123!@#456", siteIP))
            {
                foreach (TestCaseInfo tmpPlannedCase in tmpPlannedCases)
                {
                    lock (lockObj)
                    {
                        createTestRunIDFile(tmpPlannedCase.CasePath, tmpPlannedCase.AdditionRunID);
                    }
                }
            }
        }

        private void doTaskPlan(object args)
        {
            TaskParamters paramters = args as TaskParamters;
            if (paramters == null)
                return;
            List<TestCaseInfo> tmpAllCases = paramters.TestCasesQueue;
            List<string> phaseNameErrorList = paramters.ErrorInfoList[0];
            List<string> caseNameErrorList = paramters.ErrorInfoList[1];
            List<string> caseConfigErrorList = paramters.ErrorInfoList[2];
            string xmlFilePath = null;
            string caseConfig = null;
            XmlNodeList xmlNodes = null;
            //if (caseOrderChange)
            //{
            //    refreshCastOrderToExecute();
            //    //caseOrderChange = false;
            //}
            string phaseName = null;
            int caseID, phaseID = -1, configID;//, runID;
            //List<string> caseNameErrorList = new List<string>(), phaseNameErrorList = new List<string>(), caseConfigErrorList = new List<string>();
            using (IdentityScope identityScope = new IdentityScope("SYADMIN", "ASDzxc123!@#456", siteIP))
            {
                XmlDocument xmlDoc = new XmlDocument();
                foreach (TestCaseInfo plannedCase in tmpAllCases)
                //Parallel.ForEach<TestCaseInfo>(tmpAllCases, plannedCase =>
                {
                    //XmlDocument xmlDoc = new XmlDocument();
                    xmlFilePath = plannedCase.CasePath;
                    xmlDoc.Load(xmlFilePath + System.IO.Path.DirectorySeparatorChar + CASEINFOXMLFILE);
                    if (String.IsNullOrEmpty(phaseName))
                    {
                        //xmlNodes = xmlDoc.GetElementsByTagName("TeamProjectName");
                        //phaseName = xmlNodes[0].InnerText;
                        xmlNodes = xmlDoc.GetElementsByTagName("TestSuiteName");
                        //phaseName = phaseName + " " + (xmlNodes[0].InnerText.Contains("TA_") ? xmlNodes[0].InnerText.Substring(3) : xmlNodes[0].InnerText);
                        phaseName = xmlNodes[0].InnerText.Contains("TA_") ? xmlNodes[0].InnerText.Substring(3) : xmlNodes[0].InnerText;
                        phaseID = getPhaseID(phaseName);
                        if (phaseID == -1)
                        {
                            //MessageBox.Show("This phase info does not exist in DB... Phase Name: [" + phaseName + "]", "Ranorex Planner");
                            if (!phaseNameErrorList.Contains(phaseName))
                            {
                                phaseNameErrorList.Add(phaseName);
                            }
                        }
                    }

                    xmlNodes = xmlDoc.GetElementsByTagName("TestSiteName");
                    //xmlNodes[0].InnerText = "";
                    xmlNodes[0].InnerText = siteName.Substring(0, 2).ToUpper().Equals("CT") ? siteName.Substring(2) : siteName;
                    xmlNodes = xmlDoc.GetElementsByTagName("TestConfigurationName");
                    caseConfig = xmlNodes[0].InnerText;
                    if (testCaseList.Contains(plannedCase) && caseConfig.Equals(plannedCase.CaseConfiguration))
                    {
                        caseID = getTestCaseID(plannedCase.CaseName);
                        if (caseID == -1)
                        {
                            //MessageBox.Show("This case does not exist in DB... Case Name: [" + plannedCase.CaseName + "]", "Ranorex Planner");
                            if (!caseNameErrorList.Contains(plannedCase.CaseName))
                            {
                                caseNameErrorList.Add(plannedCase.CaseName);
                            }
                        }
                        configID = getConfigID(caseConfig);
                        if (configID == -1)
                        {
                            //MessageBox.Show("This Config info does not exist in DB... Config Name: [" + caseConfig + "]", "Ranorex Planner");
                            if (!caseConfigErrorList.Contains(caseConfig))
                            {
                                caseConfigErrorList.Add(caseConfig);
                            }
                        }
                        //if (configID == -1 || phaseID == -1 || caseID == -1)
                        //{
                        //    runID = 0;
                        //    //MessageBox.Show("Don't insert the case info to DB because information is not full about this case... Case Name: [" + plannedCase.CaseName + "]", "Ranorex Planner");
                        //}
                        //else
                        //{
                        //    lock (lockObj)
                        //    {
                        //        setTestInfoToDB(phaseID, caseID, configID);
                        //        runID = getLastNewTestRunID();
                        //    }
                        //}
                        plannedCase.AdditionCaseID = caseID;
                        plannedCase.AdditionConfigID = configID;
                        plannedCase.AdditionPhaseID = phaseID;
                        xmlNodes = xmlDoc.GetElementsByTagName("PlannedForExecution");
                        //xmlNodes[0].InnerText = "";
                        xmlNodes[0].InnerText = "true";
                        if (caseOrderChange && plannedCase.CaseOrderID > 0)
                        {
                            xmlNodes = xmlDoc.GetElementsByTagName("OrderId");
                            xmlNodes[0].InnerText = "";
                            xmlNodes[0].InnerText = Convert.ToString(plannedCase.CaseOrderID);
                        }
                        //createTestRunIDFile(xmlFilePath, runID);
                    }
                    else
                    {
                        xmlNodes = xmlDoc.GetElementsByTagName("PlannedForExecution");
                        //xmlNodes[0].InnerText = "";
                        xmlNodes[0].InnerText = "false";
                        //createTestRunIDFile(xmlFilePath, -1);
                    }
                    xmlDoc.Save(xmlFilePath + System.IO.Path.DirectorySeparatorChar + CASEINFOXMLFILE);
                    //xmlDoc = null;
                    createTestRunIDFile(xmlFilePath, -1);
                }//);
                xmlDoc = null;
            }
        }

        private List<TestCaseInfo>[] splitAllCaseListInSite(IList<TestCaseInfo> splitList)
        {
            List<TestCaseInfo>[] casesQueue = new List<TestCaseInfo>[TASKCOUNT];
            //int itemsNum = allCaseListInSite.Count / TASKCOUNT;
            int itemsNum = splitList.Count / TASKCOUNT;
            for (int i = 0; i < casesQueue.Length; i++)
            {
                if (i == casesQueue.Length - 1)
                    //casesQueue[i] = allCaseListInSite.ToList<TestCaseInfo>().GetRange(i * itemsNum, allCaseListInSite.Count - itemsNum * (TASKCOUNT - 1));
                    casesQueue[i] = splitList.ToList<TestCaseInfo>().GetRange(i * itemsNum, splitList.Count - itemsNum * (TASKCOUNT - 1));
                else
                    //casesQueue[i] = allCaseListInSite.ToList<TestCaseInfo>().GetRange(i * itemsNum, itemsNum);
                    casesQueue[i] = splitList.ToList<TestCaseInfo>().GetRange(i * itemsNum, itemsNum);
            }
            return casesQueue;
        }

        private void modifyCaseInfoByPlanList()
        {
            using (IdentityScope identityScope = new IdentityScope("SYADMIN", "ASDzxc123!@#456", siteIP))
            {
                List<string> caseNameErrorList = new List<string>(), phaseNameErrorList = new List<string>(), caseConfigErrorList = new List<string>();
                //XmlDocument xmlDoc = new XmlDocument();
                string taRootDirPath = @"\\" + siteIP + @"\D$\TA_Ranorex";
                string taCaseDirPath = null;
                string tmpDirStr = null;
                //change metadata file on the test site
                #region change metadata file
                try
                {
                    string[] tempSubDirs = Directory.GetDirectories(taRootDirPath);
                    foreach (string tmpSubDir in tempSubDirs)
                    {
                        tmpDirStr = System.IO.Path.GetFileName(tmpSubDir);
                        if (!tmpDirStr.Equals("Resources") && !tmpDirStr.Equals("TestApplication"))
                        {
                            taCaseDirPath = tmpDirStr;
                            break;
                        }
                    }
                    if (metaDataInfoChange)
                    {
                        if (File.Exists(taRootDirPath + System.IO.Path.DirectorySeparatorChar + "TestApplication" + System.IO.Path.DirectorySeparatorChar + METADATAFILE))
                        {
                            File.Delete(taRootDirPath + System.IO.Path.DirectorySeparatorChar + "TestApplication" + System.IO.Path.DirectorySeparatorChar + METADATAFILE);
                            System.Threading.Thread.Sleep(1000);
                        }
                        createMetaDataXMLFile(taRootDirPath + System.IO.Path.DirectorySeparatorChar + "TestApplication" + System.IO.Path.DirectorySeparatorChar + METADATAFILE);
                        metaDataInfoChange = false;
                    }
                    else
                    {
                        if (!File.Exists(taRootDirPath + System.IO.Path.DirectorySeparatorChar + "TestApplication" + System.IO.Path.DirectorySeparatorChar + METADATAFILE))
                            createMetaDataXMLFile(taRootDirPath + System.IO.Path.DirectorySeparatorChar + "TestApplication" + System.IO.Path.DirectorySeparatorChar + METADATAFILE);
                    }
                }
                catch
                {
                    MessageBox.Show("Create MetaData file exception on [" + siteName + "] test site... ", "Ranorex Planner");
                    return;
                }
                #endregion
                //change planned case flag and site name on the test site
                #region change planned cases
                #region temp1
                //string xmlFilePath = null;
                //string caseConfig = null;
                //XmlNodeList xmlNodes = null;
                //if (caseOrderChange)
                //{
                //    refreshCastOrderToExecute();
                //    //caseOrderChange = false;
                //}
                //string phaseName = null;
                //int caseID, phaseID = -1, configID, runID;
                //List<string> caseNameErrorList = new List<string>(), phaseNameErrorList = new List<string>(), caseConfigErrorList = new List<string>();
                ////foreach (TestCaseInfo plannedCase in allCaseListInSite)
                //Parallel.ForEach<TestCaseInfo>(allCaseListInSite, plannedCase =>
                //{
                //    xmlFilePath = plannedCase.CasePath;
                //    xmlDoc.Load(xmlFilePath + System.IO.Path.DirectorySeparatorChar + CASEINFOXMLFILE);
                //    if (String.IsNullOrEmpty(phaseName))
                //    {
                //        //xmlNodes = xmlDoc.GetElementsByTagName("TeamProjectName");
                //        //phaseName = xmlNodes[0].InnerText;
                //        xmlNodes = xmlDoc.GetElementsByTagName("TestSuiteName");
                //        //phaseName = phaseName + " " + (xmlNodes[0].InnerText.Contains("TA_") ? xmlNodes[0].InnerText.Substring(3) : xmlNodes[0].InnerText);
                //        phaseName = xmlNodes[0].InnerText.Contains("TA_") ? xmlNodes[0].InnerText.Substring(3) : xmlNodes[0].InnerText;
                //        phaseID = getPhaseID(phaseName);
                //        if (phaseID == -1)
                //        {
                //            //MessageBox.Show("This phase info does not exist in DB... Phase Name: [" + phaseName + "]", "Ranorex Planner");
                //            if (!phaseNameErrorList.Contains(phaseName))
                //            {
                //                phaseNameErrorList.Add(phaseName);
                //            }
                //        }
                //    }

                //    xmlNodes = xmlDoc.GetElementsByTagName("TestSiteName");
                //    xmlNodes[0].InnerText = "";
                //    xmlNodes[0].InnerText = siteName;
                //    xmlNodes = xmlDoc.GetElementsByTagName("TestConfigurationName");
                //    caseConfig = xmlNodes[0].InnerText;
                //    if (testCaseList.Contains(plannedCase) && caseConfig.Equals(plannedCase.CaseConfiguration))
                //    {
                //        caseID = getTestCaseID(plannedCase.CaseName);
                //        if (caseID == -1)
                //        {
                //            //MessageBox.Show("This case does not exist in DB... Case Name: [" + plannedCase.CaseName + "]", "Ranorex Planner");
                //            if (!caseNameErrorList.Contains(plannedCase.CaseName))
                //            {
                //                caseNameErrorList.Add(plannedCase.CaseName);
                //            }
                //        }
                //        configID = getConfigID(caseConfig);
                //        if (configID == -1)
                //        {
                //            //MessageBox.Show("This Config info does not exist in DB... Config Name: [" + caseConfig + "]", "Ranorex Planner");
                //            if (!caseConfigErrorList.Contains(caseConfig))
                //            {
                //                caseConfigErrorList.Add(caseConfig);
                //            }
                //        }
                //        if (configID == -1 || phaseID == -1 || caseID == -1)
                //        {
                //            //MessageBox.Show("Don't insert the case info to DB because information is not full about this case... Case Name: [" + plannedCase.CaseName + "]", "Ranorex Planner");
                //        }
                //        else
                //        {
                //            setTestInfoToDB(phaseID, caseID, configID);
                //        }
                //        xmlNodes = xmlDoc.GetElementsByTagName("PlannedForExecution");
                //        xmlNodes[0].InnerText = "";
                //        xmlNodes[0].InnerText = "true";
                //        if (caseOrderChange && plannedCase.CaseOrderID > 0)
                //        {
                //            xmlNodes = xmlDoc.GetElementsByTagName("OrderId");
                //            xmlNodes[0].InnerText = "";
                //            xmlNodes[0].InnerText = Convert.ToString(plannedCase.CaseOrderID);
                //        }
                //        runID = getLastNewTestRunID();
                //        createTestRunIDFile(xmlFilePath, runID);
                //    }
                //    else
                //    {
                //        xmlNodes = xmlDoc.GetElementsByTagName("PlannedForExecution");
                //        xmlNodes[0].InnerText = "";
                //        xmlNodes[0].InnerText = "false";
                //        createTestRunIDFile(xmlFilePath, -1);
                //    }
                //    xmlDoc.Save(xmlFilePath + System.IO.Path.DirectorySeparatorChar + CASEINFOXMLFILE);
                //});
                #endregion

                if (caseOrderChange)
                {
                    refreshCastOrderToExecute();
                    //caseOrderChange = false;
                }
                List<TestCaseInfo>[] testCasesQueue = splitAllCaseListInSite(allCaseListInSite);
                Task[] planTasks = new Task[TASKCOUNT];
                for (int i = 0; i < planTasks.Length; i++)
                {
                    TaskParamters temParam = new TaskParamters
                    {
                        TestCasesQueue = testCasesQueue[i],
                        ErrorInfoList = new List<string>[] { phaseNameErrorList, caseNameErrorList, caseConfigErrorList }
                    };
                    planTasks[i] = new Task(doTaskPlan, temParam);
                    planTasks[i].Start();
                }
                Task.WaitAll(planTasks);
                // insert cases to DB with plan order
                foreach (TestCaseInfo tmpPlannedCase in testCaseList)
                {
                    if (tmpPlannedCase.AdditionPhaseID == -1 || tmpPlannedCase.AdditionConfigID == -1 || tmpPlannedCase.AdditionCaseID == -1)
                    {
                        tmpPlannedCase.AdditionRunID = 0;
                    }
                    else
                    {
                        setTestInfoToDB(tmpPlannedCase.AdditionPhaseID, tmpPlannedCase.AdditionCaseID, tmpPlannedCase.AdditionConfigID);
                        tmpPlannedCase.AdditionRunID = getLastNewTestRunID();
                    }
                }
                // create runid file in case folder
                List<TestCaseInfo>[] planCasesQueue = splitAllCaseListInSite(testCaseList);
                Task[] createRunIDFileTasks = new Task[TASKCOUNT];
                for (int i = 0; i < createRunIDFileTasks.Length; i++)
                {
                    TaskParamters temParam = new TaskParamters
                    {
                        TestCasesQueue = planCasesQueue[i],
                        //ErrorInfoList = new List<string>[] { phaseNameErrorList, caseNameErrorList, caseConfigErrorList }
                    };
                    createRunIDFileTasks[i] = new Task(doCreateRunIDFiles, temParam);
                    createRunIDFileTasks[i].Start();
                }
                Task.WaitAll(createRunIDFileTasks);

                caseOrderChange = false;
                //xmlDoc = null;

                if (caseNameErrorList.Count > 0 || phaseNameErrorList.Count > 0 || caseConfigErrorList.Count > 0)
                {
                    string errorString = "";
                    if (caseNameErrorList.Count > 0)
                    {
                        errorString += "The following cases do not exist in DB, those cases cannot insert into DB ... \r\n";
                        foreach (string tmpError in caseNameErrorList)
                        {
                            errorString += tmpError + "\r\n";
                        }
                    }
                    if (phaseNameErrorList.Count > 0)
                    {
                        errorString += "The following phase info do not exist in DB, cases of those phase cannot insert into DB ... \r\n";
                        foreach (string tmpError in phaseNameErrorList)
                        {
                            errorString += tmpError + "\r\n";
                        }
                    }
                    if (caseConfigErrorList.Count > 0)
                    {
                        errorString += "The following case configs do not exist in DB, cases of those config cannot insert into DB ... \r\n";
                        foreach (string tmpError in caseConfigErrorList)
                        {
                            errorString += tmpError + "\r\n";
                        }
                    }
                    MessageBox.Show(errorString, "Ranorex Planner");
                }
                #endregion
            }
        }

        private bool deleteTAScheduler(bool bPlanCase = true)
        {
            bool returnValue = true;
            try
            {
                TaskScheduler.TaskScheduler Scheduler = new TaskScheduler.TaskScheduler();
                if (bPlanCase)
                    Scheduler.Connect(siteIP, "SYADMIN", "localhost", "ASDzxc123!@#456"); //run as current user.
                else
                    Scheduler.Connect(toolSiteIP, "SYADMIN", "localhost", "ASDzxc123!@#456"); //run as current user.
                ITaskFolder root = Scheduler.GetFolder("\\");
                if (bPlanCase)
                {
                    if (!String.IsNullOrEmpty(TARANOREXSCHEDULENAME))
                        try
                        {
                            root.DeleteTask(TARANOREXSCHEDULENAME, 0);
                            returnValue = returnValue ? true : returnValue;
                        }
                        catch
                        {
                            // TODO logger this exception as info
                            returnValue = false;
                        }
                    if (preSettingCbx.IsChecked == true)
                    {
                        if (!String.IsNullOrEmpty(TAPRESETTINGSCHEDULENAME))
                            try
                            {
                                root.DeleteTask(TAPRESETTINGSCHEDULENAME, 0);
                                returnValue = returnValue ? true : returnValue;
                            }
                            catch
                            {
                                // TODO logger this exception as info 
                                returnValue = false;
                            }
                    }
                    if (!String.IsNullOrEmpty(TAKILLEXECUTORSCHEDULENAME))
                        try
                        {
                            root.DeleteTask(TAKILLEXECUTORSCHEDULENAME, 0);
                            returnValue = returnValue ? true : returnValue;
                        }
                        catch
                        {
                            // TODO logger this exception as info 
                            returnValue = false;
                        }
                }
                else
                {
                    if (!String.IsNullOrEmpty(TOOLSCHEDULENAME))
                        try
                        {
                            root.DeleteTask(TOOLSCHEDULENAME, 0);
                            returnValue = returnValue ? true : returnValue;
                        }
                        catch
                        {
                            // TODO logger this exception as info
                            returnValue = false;
                        }
                }
            }
            catch
            {
                returnValue = false;
            }
            return returnValue;
        }

        private bool createTAScheduler(DateTime dt, bool bPlanCase = true, bool bImmediately = false)
        {
            bool returnValue = false;
            try
            {
                TaskScheduler.TaskScheduler Scheduler = new TaskScheduler.TaskScheduler();
                if (bPlanCase)
                    Scheduler.Connect(siteIP, "SYADMIN", "localhost", "ASDzxc123!@#456");
                else
                    Scheduler.Connect(toolSiteIP, "SYADMIN", "localhost", "ASDzxc123!@#456");
                ITaskDefinition taskDef = Scheduler.NewTask(0);
                ITrigger trigger;
                if (bImmediately)
                    trigger = taskDef.Triggers.Create(_TASK_TRIGGER_TYPE2.TASK_TRIGGER_REGISTRATION);
                else
                    trigger = (ITimeTrigger)taskDef.Triggers.Create(_TASK_TRIGGER_TYPE2.TASK_TRIGGER_TIME);
                var action = (IExecAction)taskDef.Actions.Create(_TASK_ACTION_TYPE.TASK_ACTION_EXEC);
                ITaskFolder root = Scheduler.GetFolder("\\");
                if (bPlanCase)
                {
                    try
                    {
                        // add those for create kill executor schedule
                        if (!bImmediately)
                        {
                            DateTime preDT = dt.AddSeconds(-10.0);
                            trigger.StartBoundary = preDT.ToString("s", System.Globalization.CultureInfo.InvariantCulture);
                        }
                        action.Path = TAKILLEXECUTORDIR + Path.DirectorySeparatorChar + "killExecutor.bat";
                        action.WorkingDirectory = TAKILLEXECUTORDIR;

                        root.RegisterTaskDefinition(
                            TAKILLEXECUTORSCHEDULENAME,
                            taskDef,
                            (int)_TASK_CREATION.TASK_CREATE_OR_UPDATE,
                            "meduser", // user
                            "ASDzxc123!@#456", // password
                            _TASK_LOGON_TYPE.TASK_LOGON_INTERACTIVE_TOKEN,
                            //User must already be logged on. The task will be run only in an existing interactive session.
                            "" //SDDL
                            );

                        Thread.Sleep(1500);

                        if(!bImmediately)
                            trigger.StartBoundary = dt.ToString("s", System.Globalization.CultureInfo.InvariantCulture);
                        // TODO build command
                        action.Path = RANOREXTALOCALROOTDIR + Path.DirectorySeparatorChar + "TAOfflineTestExecutor.exe";
                        action.WorkingDirectory = RANOREXTALOCALROOTDIR;
                        action.Arguments = "/immediately";

                        root.RegisterTaskDefinition(
                            TARANOREXSCHEDULENAME,
                            taskDef,
                            (int)_TASK_CREATION.TASK_CREATE_OR_UPDATE,
                            "meduser", // user
                            "ASDzxc123!@#456", // password
                            _TASK_LOGON_TYPE.TASK_LOGON_INTERACTIVE_TOKEN,
                            //User must already be logged on. The task will be run only in an existing interactive session.
                            "" //SDDL
                            );

                        returnValue = true;
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                        if (!bImmediately)
                            trigger.StartBoundary = dt.ToString("s", System.Globalization.CultureInfo.InvariantCulture);

                        // TODO build command
                        action.Path = TATOOLEXECUTORDIR + Path.DirectorySeparatorChar + "MultiTestsuiteRunnerExecutor.exe";
                        action.WorkingDirectory = TATOOLEXECUTORDIR;

                        root.RegisterTaskDefinition(
                            TOOLSCHEDULENAME,
                            taskDef,
                            (int)_TASK_CREATION.TASK_CREATE_OR_UPDATE,
                            "meduser", // user
                            "ASDzxc123!@#456", // password
                            _TASK_LOGON_TYPE.TASK_LOGON_INTERACTIVE_TOKEN,
                            //User must already be logged on. The task will be run only in an existing interactive session.
                            "" //SDDL
                            );
                        returnValue = true;
                    }
                    catch
                    { }
                }
            }
            catch
            { }
            return returnValue;
        }

        private bool createTAEnvironmentTask(DateTime dt)
        {
            bool retValue = false;
            try
            {
                TaskScheduler.TaskScheduler Scheduler = new TaskScheduler.TaskScheduler();
                Scheduler.Connect(siteIP, "SYADMIN", "localhost", "ASDzxc123!@#456");
                ITaskDefinition taskDef = Scheduler.NewTask(0);
                var trigger = (ITimeTrigger)taskDef.Triggers.Create(_TASK_TRIGGER_TYPE2.TASK_TRIGGER_TIME);

                trigger.StartBoundary = dt.ToString("s", System.Globalization.CultureInfo.InvariantCulture);

                var action = (IExecAction)taskDef.Actions.Create(_TASK_ACTION_TYPE.TASK_ACTION_EXEC);

                // TODO build command
                action.Path = TAPRESETTINGDIR + System.IO.Path.DirectorySeparatorChar + "PrepareEnvironment.exe";
                action.WorkingDirectory = TAPRESETTINGDIR;

                ITaskFolder root = Scheduler.GetFolder("\\");

                root.RegisterTaskDefinition(
                    TAPRESETTINGSCHEDULENAME,
                    taskDef,
                    (int)_TASK_CREATION.TASK_CREATE_OR_UPDATE,
                    "meduser", // user
                    "ASDzxc123!@#456", // password
                    _TASK_LOGON_TYPE.TASK_LOGON_INTERACTIVE_TOKEN,
                    //User must already be logged on. The task will be run only in an existing interactive session.
                    "" //SDDL
                    );
                retValue = true;
            }
            catch
            {
            }
            return retValue;
        }

        private DateTime covertSchedulerDateTime(string dateTime)
        {
            int year, month, day, hour, minute, second;
            string[] tmpDateTime = dateTime.Split(new string[] { " " }, StringSplitOptions.None);
            string[] tmpDates = tmpDateTime[0].Split(new string[] { "/" }, StringSplitOptions.None);
            year = Convert.ToInt32(tmpDates[2]);
            month = Convert.ToInt32(tmpDates[0]);
            day = Convert.ToInt32(tmpDates[1]);
            string[] tmpTimes = tmpDateTime[1].Split(new string[] { ":" }, StringSplitOptions.None);
            if (tmpDateTime[2].Equals("PM"))
                hour = Convert.ToInt32(tmpTimes[0]) + 12;
            else
                hour = Convert.ToInt32(tmpTimes[0]);
            minute = Convert.ToInt32(tmpTimes[1]);
            second = 20;
            return new DateTime(year, month, day, hour, minute, second);
        }

        private void createMetaDataXMLFile(string filePath)
        {
            XmlTextWriter writer = new XmlTextWriter(filePath, null);
            writer.Formatting = Formatting.Indented;
            writer.WriteStartDocument();
            writer.WriteStartElement("TestResultMetaData");
            writer.WriteStartAttribute("xmlns:xsi");
            writer.WriteString("http://www.w3.org/2001/XMLSchema-instance");
            writer.WriteStartAttribute("xmlns:xsd");
            writer.WriteString("http://www.w3.org/2001/XMLSchema");
            writer.WriteEndAttribute();
            writer.WriteElementString("Department", txtDepartment.Text);
            writer.WriteElementString("Organization", txtOrganization.Text);
            writer.WriteElementString("ProductName", txtProductName.Text);
            writer.WriteElementString("ProjectNumber", txtProductNum.Text);
            if (cbxSoftVersion.IsChecked == true)
                writer.WriteElementString("SoftwareVersion", "#AUTODETECT#");
            else
                writer.WriteElementString("SoftwareVersion", txtSoftVersion.Text);
            writer.WriteElementString("Language", txtLanguage.Text);
            writer.WriteElementString("TracesDisabled", txtTrace.Text);
            writer.WriteEndElement();
            writer.Close();
        }

        private void caculateRemoteTimeByCurrent(ref string wanTime, ref string remDate, ref string remTime)
        {
            string oldTime = wanTime;
            try
            {
                DateTime currentDT = DateTime.Now;
                currentDT = Convert.ToDateTime(currentDT.ToShortDateString() + " " + currentDT.ToShortTimeString());
                DateTime wantDT = Convert.ToDateTime(oldTime);
                TimeSpan difTS = wantDT.Subtract(currentDT);
                DateTime scheduleDT = Convert.ToDateTime(remDate + " " + remTime);
                scheduleDT = scheduleDT.Add(difTS);
                wanTime = scheduleDT.ToShortDateString() + " " + scheduleDT.ToShortTimeString();
            }
            catch
            {
                MessageBox.Show("Caculate remote PC time failed, take notice of using current PC time now... ", "Ranorex Planner");
                wanTime = oldTime;
                remDate = "Caculate Fail For";
                remTime = "Remote Time";
            }
        }

        private bool getTestSiteLocalTime(out string date, out string time, string remoteSiteIP)
        {
            bool retValue = false;
            string strLocalTime = "", strTemp;
            string[] tmpStr;
            try
            {
                Process process = new Process();
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "cmd.exe";
                info.UseShellExecute = false;
                info.RedirectStandardInput = true;
                info.RedirectStandardOutput = true;
                info.RedirectStandardError = true;
                info.CreateNoWindow = true;
                process.StartInfo = info;
                process.Start();
                StreamReader sr = process.StandardOutput;
                process.StandardInput.WriteLine(@"net use \\" + remoteSiteIP + @" ASDzxc123!@#456 /user:SYADMIN");
                process.StandardInput.WriteLine(@"net time \\" + remoteSiteIP);
                process.StandardInput.WriteLine(@"exit");
                while (!sr.EndOfStream)
                {
                    strTemp = sr.ReadLine();
                    if (strTemp.Contains("Current time"))
                    {
                        strLocalTime = strTemp;
                        retValue = true;
                    }
                    if (strTemp.Contains("Local time"))
                    {
                        strLocalTime = strTemp;
                        retValue = true;
                        break;
                    }
                }
                if (retValue)
                {
                    strLocalTime = strLocalTime.Remove(0, strLocalTime.IndexOf("is") + 2).Trim();
                    tmpStr = strLocalTime.Split(new string[] { " " }, StringSplitOptions.None);
                    date = tmpStr[0];
                    time = (tmpStr[1]).Substring(0, tmpStr[1].LastIndexOf(":")) + " " + tmpStr[2];
                }
                else
                {
                    date = "";
                    time = "";
                }
            }
            catch
            {
                retValue = false;
                date = "";
                time = "";
            }
            return retValue;
        }

        private void refreshAuthorCaseList()
        {
            foreach (string key in caseAuthors.Keys)
            {
                caseAuthors[key].Clear();
            }
            foreach (TestCaseInfo tmpCase in testCaseList)
            {
                if (caseAuthors.Keys.Contains(tmpCase.CaseAuthor))
                    caseAuthors[tmpCase.CaseAuthor].Add(tmpCase);
            }
        }

        private void sendEmailToAuthors(object siteInfo)
        {
            //bool tmpVal = true;
            string siteName = ((string)siteInfo).Split(new string[] { "=>" }, StringSplitOptions.None)[0];
            string siteIP = ((string)siteInfo).Split(new string[] { "=>" }, StringSplitOptions.None)[1];
            if (currentUserEmail == null || currentUserEmail.Trim().Equals(""))
                currentUserEmail = "TATeam@siemens.com";
            refreshAuthorCaseList();

            SmtpClient client = new SmtpClient();
            client.Host = EMAILHOST;
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential(currentUserEmail, "");
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            MailMessage message = new MailMessage(currentUserEmail, "SSMEALLCTRDTCAutomationSom5.healthcare@siemens.com");
            message.CC.Add("wu.bw.bin@siemens.com");
            message.Subject = "TA planner has planned cases on " + siteName + " (" + siteIP + ")";
            message.Body = "hello all,\t\n case list: \t\n";
            foreach (string key in caseAuthors.Keys)
            {
                if (caseAuthors[key] != null && caseAuthors[key].Count > 0)
                {
                    //SmtpClient client = new SmtpClient();
                    //client.Host = EMAILHOST;
                    //client.UseDefaultCredentials = false;
                    //client.Credentials = new NetworkCredential(currentUserEmail, "");
                    //client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    try
                    {
                        //MailMessage message = new MailMessage(currentUserEmail, authorEmail[key]);
                        //message.CC.Add("wu.bw.bin@siemens.com");
                        //message.Subject = "TA planner has planned " + caseAuthors[key].Count + " cases on " + siteName + " (" + siteIP + ")";
                        //message.Body = "hello " + key + "\t\n case list: \t\n";
                        foreach (TestCaseInfo tmpCase in caseAuthors[key])
                            message.Body += "   " + tmpCase.CaseName + "                " + tmpCase.CaseConfiguration + "   " + key + "\t\n";
                        //message.BodyEncoding = Encoding.UTF8;
                        //message.IsBodyHtml = false;
                        //client.Send(message);
                        //tmpVal = tmpVal && true;
                    }
                    catch
                    {                        
                        MessageBox.Show("Send email to author [" + key + "] failed, please contact case author by self... ", "Ranorex Planner");
                        //tmpVal = false;
                    }
                }
            }
            message.BodyEncoding = Encoding.UTF8;
            message.IsBodyHtml = false;
            client.Send(message);
            //return tmpVal;
        }

        private void logInfoThisPlan(object sitePlanTime)
        {
            string filePath = Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + "logs";
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
                Thread.Sleep(1000);
            }
            DateTime tmpDate = DateTime.Now;
            string fileName = "log_" + siteName + "_" + tmpDate.Year.ToString() + String.Format("{0:D2}", tmpDate.Month) + String.Format("{0:D2}", tmpDate.Day) + "_" +
                String.Format("{0:D2}", tmpDate.Hour) + String.Format("{0:D2}", tmpDate.Minute) + String.Format("{0:D2}", tmpDate.Second) + ".txt";
            string logInfo = "Planed " + testCaseList.Count.ToString() + " cases on test site [" + siteName + "] at test site time [" + sitePlanTime.ToString() + "], following is the case list: ";
            using (StreamWriter sw = new StreamWriter(filePath + Path.DirectorySeparatorChar + fileName))
            {
                sw.WriteLine(logInfo);
                foreach (TestCaseInfo tmpCase in testCaseList)
                    sw.WriteLine(tmpCase.CaseName + "             " + tmpCase.CaseAuthor);
                sw.Flush();
                sw.Close();
            }
        }

        private string getSWVersion()
        {
            string retValue = "undefined";
            string versionFilePath = @"\\" + siteIP + @"\c$\Somaris\bin\version.txt";
            try
            {
                Process process = new Process();
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "cmd.exe";
                info.UseShellExecute = false;
                info.RedirectStandardInput = true;
                info.RedirectStandardOutput = true;
                info.RedirectStandardError = true;
                info.CreateNoWindow = true;
                process.StartInfo = info;
                process.Start();
                StreamReader sr = process.StandardOutput;
                process.StandardInput.WriteLine(@"net use \\" + siteIP + @" ASDzxc123!@#456 /user:SYADMIN");
                for (int i = 1; i <= 3; i++)
                {
                    process.StandardInput.WriteLine(@"reg query \\" + siteIP + @"\HKLM\SOFTWARE\Wow6432Node\Siemens\SOMARIS\Config\Site\General\SERVICE_PACKS\L1 /v V" + i.ToString());
                }
                process.StandardInput.WriteLine(@"exit");
                System.Collections.Generic.List<string> PatchStrs = new System.Collections.Generic.List<string>();
                string PatchStr;
                PatchStrs.Clear();
                while (!sr.EndOfStream)
                {
                    PatchStr = sr.ReadLine();
                    if (PatchStr.Contains("V1    REG_SZ") || PatchStr.Contains("V2    REG_SZ") || PatchStr.Contains("V3    REG_SZ"))
                    {
                        PatchStrs.Add(PatchStr.Trim().Split(new string[] { "    " }, StringSplitOptions.None)[2]);
                    }
                }
                string[] tmpVersion = File.ReadAllLines(versionFilePath);
                if (PatchStrs.Count == 0)
                {
                    retValue = tmpVersion[0];
                }
                else
                {
                    PatchStr = "";
                    for (int j = 0; j < PatchStrs.Count; j++)
                        PatchStr += "+" + PatchStrs[j];
                    retValue = tmpVersion[0] + PatchStr;
                }
            }
            catch
            {
                retValue = "undefined";
            }
            return retValue;
        }

        private string getTestSiteTraceInfo()
        {
            string regValue = null;
            bool retValue = false;
            try
            {
                Process process = new Process();
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "cmd.exe";
                info.UseShellExecute = false;
                info.RedirectStandardInput = true;
                info.RedirectStandardOutput = true;
                info.RedirectStandardError = true;
                info.CreateNoWindow = true;
                process.StartInfo = info;
                process.Start();
                StreamReader sr = process.StandardOutput;
                process.StandardInput.WriteLine(@"net use \\" + siteIP + @" ASDzxc123!@#456 /user:SYADMIN");
                process.StandardInput.WriteLine(@"reg query \\" + siteIP + @"\HKLM\SOFTWARE\Wow6432Node\Siemens\MedCom\TraceConfig /v CTTraceON");
                process.StandardInput.WriteLine(@"exit");
                while (!sr.EndOfStream)
                {
                    regValue = sr.ReadLine();
                    if (regValue.Contains("CTTraceON    REG_SZ"))
                    {
                        retValue = true;
                        break;
                    }
                }
                if (!retValue)
                    regValue = "";
            }
            catch
            {
                regValue = null;
            }
            return regValue;
        }

        private string getTestSiteSystemVersion()
        {
            string regValue = null;
            bool retValue = false;
            try
            {
                Process process = new Process();
                ProcessStartInfo info = new ProcessStartInfo();
                info.FileName = "cmd.exe";
                info.UseShellExecute = false;
                info.RedirectStandardInput = true;
                info.RedirectStandardOutput = true;
                info.RedirectStandardError = true;
                info.CreateNoWindow = true;
                process.StartInfo = info;
                process.Start();
                StreamReader sr = process.StandardOutput;
                process.StandardInput.WriteLine(@"net use \\" + siteIP + @" ASDzxc123!@#456 /user:SYADMIN");
                process.StandardInput.WriteLine(@"reg query \\" + siteIP + @"\HKLM\SOFTWARE\Wow6432Node\Siemens\SOMARIS\Config\Site\General\MODEL_TYPE_INHOUSE /v V1");
                process.StandardInput.WriteLine(@"exit");
                while (!sr.EndOfStream)
                {
                    regValue = sr.ReadLine();
                    if (regValue.Contains("V1    REG_SZ"))
                    {
                        retValue = true;
                        break;
                    }
                }
                if (!retValue)
                    regValue = "";
            }
            catch
            {
                regValue = null;
            }
            return regValue;
        }

        private void refreshTestCaseListOrder()
        {
            if (bePlanedCaseLB.Items.Count > 0 && testCaseList.Count == bePlanedCaseLB.Items.Count)
            {
                testCaseList.Clear();
                foreach (TestCaseInfo tmpCase in bePlanedCaseLB.Items)
                {
                    testCaseList.Add(tmpCase);
                }
            }
        }

        private void refreshCastOrderToExecute()
        {
            int skipCount = 0;
            refreshTestCaseListOrder();
            List<int> orderNumberList = new List<int>();
            foreach (TestCaseInfo tempCase in testCaseList)
            {
                if (tempCase.CaseOrderID > 0)
                    orderNumberList.Add(tempCase.CaseOrderID);
            }
            orderNumberList.Sort();
            for (int i = 0; i < testCaseList.Count; i++)
            {
                if (testCaseList[i].CaseOrderID > 0)
                    testCaseList[i].CaseOrderID = orderNumberList[i - skipCount];
                else
                    skipCount++;
            }
        }

        private void setTestInfoToDB(int phaseID, int caseID, int configID)
        {
            string tempVersion, tempTime;
            if (String.IsNullOrEmpty(txtSoftVersion.Text))
                tempVersion = getSWVersion();
            else
                tempVersion = txtSoftVersion.Text;
            DateTime dateTime = System.DateTime.Now;
            tempTime = dateTime.Year.ToString() + "-" + String.Format("{0:D2}", dateTime.Month) + "-" + String.Format("{0:D2}", dateTime.Day) + " " +
                String.Format("{0:D2}", dateTime.Hour) + ":" + String.Format("{0:D2}", dateTime.Minute) + ":" + String.Format("{0:D2}", dateTime.Second);
            string sqlStr = "Data Source=SHAI203A;Initial Catalog=SSME_Ranorex;User ID=zope_usr;pwd=zope_usr";
            string sqlComStr = "insert into TESTRUN(STATUS_ID, PHASE_ID, TESTCASE_ID, CONFIG_ID, TESTSITE_ID, SOFTWAREVERSION, LOGPATH, PLANNER, PLANTIME) " +
                                        "values(111, " + phaseID + ", " + caseID + ", " + configID + ", " + siteID + ", '" + tempVersion + "', '', '" + currentUser + "', '" + tempTime + "')";
            using (SqlConnection sqlCon = new SqlConnection(sqlStr))
            using (SqlCommand sqlCom = new SqlCommand(sqlComStr, sqlCon))
            {
                sqlCon.Open();
                sqlCom.ExecuteNonQuery();

                sqlCom.Dispose();
                sqlCon.Close();
                sqlCon.Dispose();
            }
        }

        private int getLastNewTestRunID()
        {
            string sqlComStr = "Select TESTRUN_ID from TESTRUN order by TESTRUN_ID DESC";
            return getIDFromDBySql(sqlComStr);
        }

        private int getPhaseID(string phaseName)
        {
            string sqlComStr = "Select PHASE_ID from PHASE where PHASENAME='" + phaseName + "'";
            return getIDFromDBySql(sqlComStr);
        }

        private int getTestCaseID(string caseName)
        {
            string sqlComStr = "Select TESTCASE_ID from TESTCASE where TESTCASENAME='" + caseName + "'";
            return getIDFromDBySql(sqlComStr);
        }

        private int getConfigID(string configName)
        {
            string sqlComStr = "Select CONFIG_ID from CONFIG where CONFIG='" + configName + "'";
            return getIDFromDBySql(sqlComStr);
        }

        private int getIDFromDBySql(string sqlComStr)
        {
            int runID = -1;
            string sqlStr = "Data Source=SHAI203A;Initial Catalog=SSME_Ranorex;User ID=zope_usr;pwd=zope_usr";
            using (SqlConnection sqlCon = new SqlConnection(sqlStr))
            using (SqlCommand sqlCom = new SqlCommand(sqlComStr, sqlCon))
            {
                sqlCon.Open();
                using (SqlDataReader sqlReader = sqlCom.ExecuteReader())
                {
                    if (sqlReader.Read())
                        runID = sqlReader.GetInt32(0);
                    sqlReader.Close();
                    sqlReader.Dispose();
                }
                sqlCom.Dispose();
                sqlCon.Close();
                sqlCon.Dispose();
            }
            return runID;
        }

        private void createTestRunIDFile(string idFilePath, int testRunID)
        {
            string fileName = "PlanRunID.xml";
            if (File.Exists(idFilePath + System.IO.Path.DirectorySeparatorChar + fileName))
            {
                File.Delete(idFilePath + System.IO.Path.DirectorySeparatorChar + fileName);
            }
            if (testRunID != -1)
            {
                System.Threading.Thread.Sleep(1000);
                XmlTextWriter writer = new XmlTextWriter(idFilePath + System.IO.Path.DirectorySeparatorChar + fileName, null);
                writer.Formatting = Formatting.Indented;
                writer.WriteStartDocument();
                writer.WriteStartElement("PlanTestRunData");
                writer.WriteStartAttribute("xmlns:xsi");
                writer.WriteString("http://www.w3.org/2001/XMLSchema-instance");
                writer.WriteStartAttribute("xmlns:xsd");
                writer.WriteString("http://www.w3.org/2001/XMLSchema");
                writer.WriteEndAttribute();
                writer.WriteElementString("RunID", testRunID.ToString());
                writer.WriteElementString("Status", "Not_Executed");
                writer.WriteElementString("TestPointPath", idFilePath);
                writer.WriteEndElement();
                writer.Close();
            }
        }

        private string ExecuteCommand(string command)
        {
            string retVal = null;
            try
            {
                Process proc = new Process();
                proc.StartInfo.FileName = "cmd.exe";
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardInput = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.CreateNoWindow = true;
                proc.Start();
                proc.StandardInput.WriteLine(command);
                proc.StandardInput.WriteLine("exit");
                while (proc.HasExited == false)
                {
                    proc.WaitForExit(1000);
                }
                retVal = proc.StandardOutput.ReadToEnd();
                proc.StandardOutput.Close();
            }
            catch (Exception)
            {
                retVal = "Execute command: " + command + " failed.";
            }
            return retVal;
        }

        private void toolTestSite_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ToolNodes.Clear();
            toolSiteIP = ((TestSiteInfo)combxToolTestSite.SelectedItem).Value;
            toolSiteName = ((TestSiteInfo)combxToolTestSite.SelectedItem).Text;
            ExecuteCommand(@"net use \\" + toolSiteIP + @" ASDzxc123!@#456 /user:SYADMIN");
            initToolNodes();
        }

        private void initToolNodes()
        {
            try
            {
                string toolFilePath = @"\\" + toolSiteIP + @"\D$\TA_Ranorex\TestApplication";
                var testsuitsDirectory = new DirectoryInfo(toolFilePath);
                FileInfo[] fileInfos = testsuitsDirectory.GetFiles();
                fileInfos = fileInfos.Where(item => item.Extension == Constants.TESTSUITE_FILE_EXTENSION).ToArray();
                List<string> filenames = fileInfos.Select(item => item.FullName).ToList();
                foreach (string filePath in filenames)
                {
                    AddTestsuiteTree(filePath);
                }
                toolView.ItemsSource = ToolNodes;
                CheckTreeAndChildItems(ToolNodes, false);
            }
            catch
            {
                MessageBox.Show("Get and Init TA tools info exception on [" + toolSiteName + "] test site... ", "Ranorex Planner");
            }
        }

        private void AddTestsuiteTree(string filePath)
        {
            XElement testSuiteElement = XElement.Load(@"" + filePath + "");
            XElement contentElement = testSuiteElement.Element(Constants.XML_TAG_CONTENT);

            XAttribute attribute = testSuiteElement.Attribute(Constants.XML_ATTRIBUTE_NAME);
            if (attribute != null)
            {
                var testsuiteNode = new Node { Text = attribute.Value };

                AddChildNodes(contentElement, testsuiteNode, Path.GetFileNameWithoutExtension(filePath) + Constants.EXECUTABLE_FILE_EXTENSION);

                ToolNodes.Add(testsuiteNode);
            }
            else
            {
                throw new Exception("Couldn't find the attribute name");
            }
        }

        private void AddChildNodes(XElement parentElement, Node parentNode, string fileName)
        {
            if (parentElement.Elements().Any())
            {
                foreach (XElement xElement in parentElement.Elements())
                {
                    if (xElement.Name == Constants.XML_TAG_FOLDER || xElement.Name == Constants.XML_TAG_TESTCASE)
                    {
                        XAttribute xAttribute = xElement.Attribute(Constants.XML_ATTRIBUTE_NAME);
                        if (xAttribute != null)
                        {
                            var node = new Node { Text = xAttribute.Value, FileName = fileName };
                            parentNode.Children.Add(node);
                            node.Parent.Add(parentNode);
                            AddChildNodes(xElement, node, fileName);
                            if (xElement.Name == Constants.XML_TAG_TESTCASE)
                            {
                                node.IsTestcase = true;
                            }
                        }
                        else
                        {
                            throw new Exception("Couldn't find the attribute name");
                        }
                    }
                }
            }
        }

        private void CheckTreeAndChildItems(IEnumerable<Node> items, bool isChecked)
        {
            foreach (Node item in items)
            {
                item.IsChecked = isChecked;
                if (item.Children.Count != 0)
                {
                    CheckTreeAndChildItems(item.Children, isChecked);
                }
            }
        }

        private void OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var currentCheckBox = (CheckBox)sender;
            CheckBoxId.Id = currentCheckBox.Uid;
        }

        private void ButtonCheckAllClick(object sender, RoutedEventArgs e)
        {
            CheckTreeAndChildItems(ToolNodes, true);
        }

        private void ButtonUncheckAllClick(object sender, RoutedEventArgs e)
        {
            CheckTreeAndChildItems(ToolNodes, false);
        }

        private void ButtonExpandClick(object sender, RoutedEventArgs e)
        {
            ExpandTree(ToolNodes, true);
        }

        private void ButtonCollapseClick(object sender, RoutedEventArgs e)
        {
            ExpandTree(ToolNodes, false);
        }

        private void ButtonInvertClick(object sender, RoutedEventArgs e)
        {
            CheckBoxId.Id = "";
            InvertTree(ToolNodes);
        }

        private void ButtonRefreshClick(object sender, RoutedEventArgs e)
        {
            ToolNodes.Clear();
            initToolNodes();
        }

        private void ExpandTree(IEnumerable<Node> items, bool isExpanded)
        {
            foreach (Node item in items)
            {
                item.IsExpanded = isExpanded;
                if (item.Children.Count != 0)
                {
                    ExpandTree(item.Children, isExpanded);
                }
            }
        }

        private void InvertTree(IEnumerable<Node> items)
        {
            foreach (Node item in items)
            {
                if (item.IsChecked != null)
                {
                    item.IsChecked = item.IsChecked != true;
                    if (item.Parent.Count != 0 && item.Parent[0].IsChecked == true)
                    {
                        item.IsChecked = true;
                    }
                    if (item.Parent.Count != 0 && item.Parent[0].IsChecked == false)
                    {
                        item.IsChecked = false;
                    }
                }
                if (item.Children.Count != 0)
                {
                    InvertTree(item.Children);
                }
            }
        }

        private void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(toolSiteIP))
            {
                MessageBox.Show("Please select one test site first... ", "Ranorex Planner");
                return;
            }

            List<Node> checkedTestcases = GetCheckedTestcases();
            int runTimes;
            try
            {
                runTimes = int.Parse(txtRunTime.Text);
                if (runTimes <= 0)
                    runTimes = 1;
                else if (runTimes >= 1000)
                    runTimes = 999;
            }
            catch
            {
                runTimes = 1;
            }
            object argument = new object[] { checkedTestcases, toolSiteIP, autoEvent, runTimes };
            StartWorker(Utils.StoreAtContainer, argument);
            if (autoEvent.WaitOne(5000))
            {
                //string startupPath = System.Windows.Forms.Application.StartupPath;
                //try
                //{
                //    //ExecuteCommand(@"net use \\" + toolSiteIP + @" ASDzxc123!@#456 /user:SYADMIN");

                //    //var process = new Process();
                //    //process.StartInfo.FileName = startupPath + System.IO.Path.DirectorySeparatorChar + "PsExec.exe";
                //    //process.StartInfo.Arguments = string.Format(@"\\{0} -u meduser -p meduser1 -d -i -w {1} {2}", toolSiteIP, @"C:\HELIUM\ToolExecutor",
                //    //    @"C:\HELIUM\ToolExecutor" + System.IO.Path.DirectorySeparatorChar + "MultiTestsuiteRunnerExecutor.exe");
                //    //process.StartInfo.UseShellExecute = false;
                //    //process.Start();
                //}
                //catch
                //{
                //    MessageBox.Show("Exception... when Execute TA Tool immediately on [" + toolSiteName + "] test site... ", "Ranorex Planner");
                //}
                deleteTAScheduler(false);
                try
                {
                    if (!createTAScheduler(DateTime.Now, false, true))
                    {
                        MessageBox.Show("Schedule Task is running failed on [" + toolSiteName + "] test site... ", "Ranorex Planner");
                        return;
                    }
                    MessageBox.Show("Schedule Task is running on [" + toolSiteName + "] test site... ", "Ranorex Planner");
                }
                catch
                {
                    deleteTAScheduler(false);
                    MessageBox.Show("Exception... when running on [" + toolSiteName + "] test site... ", "Ranorex Planner");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Waiting to create tool-cases list time out on [" + toolSiteName + "] test site... ", "Ranorex Planner");
            }
            autoEvent.Reset();
        }

        private void PlanToolBtn_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(toolSiteIP))
            {
                MessageBox.Show("Please select one test site first... ", "Ranorex Planner");
                return;
            }
            int runTimes;
            try
            {
                runTimes = int.Parse(txtRunTime.Text);
                if (runTimes <= 0)
                    runTimes = 1;
                else if (runTimes >= 1000)
                    runTimes = 999;
            }
            catch
            {
                runTimes = 1;
            }
            List<Node> checkedTestcases = GetCheckedTestcases();
            object argument = new object[] { checkedTestcases, toolSiteIP, autoEvent, runTimes };
            StartWorker(Utils.StoreAtContainer, argument);
            if (autoEvent.WaitOne(5000))
            {
                deleteTAScheduler(false);
                try
                {
                    // Get remote test site local time and caculate time according current PC time to plan
                    #region caculate remote PC time
                    string remoteDate, remoteTime, scheduleTime;
                    scheduleTime = radDateTimePickerTool.DateTimeText;
                    if (getTestSiteLocalTime(out remoteDate, out remoteTime, toolSiteIP))
                    {
                        caculateRemoteTimeByCurrent(ref scheduleTime, ref remoteDate, ref remoteTime);
                    }
                    else
                    {
                        MessageBox.Show("Get remote time of test site [" + toolSiteName + "] failed, take notice of using current PC time now... ", "Ranorex Planner");
                    }
                    #endregion
                    DateTime tmpDT = covertSchedulerDateTime(scheduleTime);

                    if (!createTAScheduler(tmpDT, false))
                    {
                        MessageBox.Show("Create Schedule Task command failed on [" + toolSiteName + "] test site... ", "Ranorex Planner");
                        return;
                    }
                    MessageBox.Show("Create Tool Schedule Task successfully on [" + toolSiteName + "] test site... ", "Ranorex Planner");
                }
                catch
                {
                    deleteTAScheduler(false);
                    MessageBox.Show("Exception... when Create Schedule Task on [" + toolSiteName + "] test site... ", "Ranorex Planner");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Waiting to create tool-cases list time out on [" + toolSiteName + "] test site... ", "Ranorex Planner");
            }
            autoEvent.Reset();
        }

        private void PlanToolDeleteBtn_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(toolSiteIP))
            {
                MessageBox.Show("Please select one test site first... ", "Ranorex Planner");
                return;
            }
            if (deleteTAScheduler(false))
                MessageBox.Show("Delete schedule task successfully on [" + toolSiteName + "] test site... ", "Ranorex Planner");
            else
                MessageBox.Show("Delete schedule task unsuccessfully  on [" + toolSiteName + "] test site... ", "Ranorex Planner");
        }

        private void OpenFolderBtn_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(toolSiteIP))
            {
                MessageBox.Show("Please select one test site first... ", "Ranorex Planner");
                return;
            }

            StartWorker(Utils.ShowReportFolder, toolSiteIP);
        }

        private IEnumerable<Node> GetCheckedItems(IEnumerable<Node> items)
        {
            var checkedItems = new List<Node>();

            foreach (Node item in items)
            {
                if (item.IsChecked != false)
                {
                    checkedItems.Add(item);

                    if (item.Children.Count != 0)
                    {
                        checkedItems.AddRange(GetCheckedItems(item.Children));
                    }
                }
            }
            return checkedItems;
        }

        private List<Node> GetCheckedTestcases()
        {
            IEnumerable<Node> checkedItems = GetCheckedItems(ToolNodes);
            return checkedItems.Where(item => item.IsTestcase).ToList();
        }

        private void StartWorker(DoWorkEventHandler eventHandler, object argument)
        {
            var worker = new BackgroundWorker();
            worker.DoWork += eventHandler;
            worker.RunWorkerCompleted += WorkerCompleted;
            myActiveWorkers.Add(worker);

            worker.RunWorkerAsync(argument);
        }

        private void WorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            var worker = (BackgroundWorker)sender;
            myActiveWorkers.Remove(worker);
        }
    }

    public class TestSiteInfo : IComparable<TestSiteInfo>
    {
        public string Text { get; set; }
        public string Value { get; set; }
        public string Type { get; set; }
        public int ID { get; set; }

        public int CompareTo(TestSiteInfo otherObj)
        {
            return this.Text.CompareTo(otherObj.Text);
        }

        public override string ToString()
        {
            return Text;
        }
    }

    public class TestCaseInfo
    {
        public string CaseName { get; set; }
        public string CaseConfiguration { get; set; }
        public string CasePath { get; set; }
        public string CaseAuthor { get; set; }
        public string AuthorEMailAdd { get; set; }
        public int CaseOrderID { get; set; }
        public string CaseState { get; set; }
        public string OriginalState { get; set; }
        public override string ToString()
        {
            return CaseName + "  |  " + CaseConfiguration;
        }
        public int AdditionCaseID { get; set; }
        public int AdditionPhaseID { get; set; }
        public int AdditionConfigID { get; set; }
        public int AdditionRunID { get; set; }
    }

    public class IdentityScope : IDisposable
    {
        [DllImport("advapi32.dll", SetLastError = true)]
        static extern bool LogonUser(string pszUsername, string pszDomain, string pszPassword,
            int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        static extern bool CloseHandle(IntPtr handle);

        [DllImport("advapi32.dll")]
        static extern bool ImpersonateLoggedOnUser(IntPtr hToken);

        [DllImport("advapi32.dll")]
        static extern bool RevertToSelf();

        const int LOGON32_PROVIDER_DEFAULT = 0;
        const int LOGON32_LOGON_NEWCREDENTIALS = 9;
        private bool disposed;

        public IdentityScope(string sUsername, string sPassword, string sDomain)
        {
            IntPtr pExistingTokenHandle = new IntPtr(0);
            try
            {
                bool bImpersonated = LogonUser(sUsername, sDomain, sPassword, LOGON32_LOGON_NEWCREDENTIALS,
                     LOGON32_PROVIDER_DEFAULT, ref pExistingTokenHandle);
                if (true == bImpersonated)
                {
                    if (!ImpersonateLoggedOnUser(pExistingTokenHandle))
                    {
                        int nErrorCode = Marshal.GetLastWin32Error();
                        throw new Exception("ImpersonateLoggedOnUser error, code=" + nErrorCode);
                    }
                }
                else
                {
                    int nErrorCode = Marshal.GetLastWin32Error();
                    throw new Exception("LogonUser error, code=" + nErrorCode);
                }
            }
            finally
            {
                if (pExistingTokenHandle != IntPtr.Zero)
                    CloseHandle(pExistingTokenHandle);
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                RevertToSelf();
                disposed = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }
    }

    public static class Utils
    {
        private static string TOOLSLISTFILEFOLDER = @"\\{0}\C$\HELIUM\ToolExecutor";
        private static string TOOLSLISTFILEPATH = TOOLSLISTFILEFOLDER + Path.DirectorySeparatorChar + @"MultiTestSuiteRunnerExecutorFileList.txt";
        private static string TOOLREPORTFOLDER = @"\\{0}\D$\RanorexReports";

        private static IEnumerable<string> CreateTestcaseList(List<Node> checkedTestcases, int runTimes)
        {
            var filelist = new List<string>();

            if (checkedTestcases.Count != 0)
            {
                foreach (Node node in checkedTestcases)
                {
                    string testcaseName = node.Text;
                    string fileName = node.FileName;
                    for (int i = 0; i < runTimes; i++)
                        filelist.Add(string.Format("{0};{1}", fileName, testcaseName));
                }
                return filelist;
            }
            throw new InvalidOperationException("No testcase was chosen!");
        }

        public static void StoreAtContainer(object sender, DoWorkEventArgs eventArgs)
        {
            var parameters = (object[])eventArgs.Argument;
            var checkedTestcases = (List<Node>)parameters[0];
            var hostname = (string)parameters[1];
            var autoEventObj = (AutoResetEvent)parameters[2];
            var runTimes = (int)parameters[3];

            try
            {
                string folderPath = string.Format(TOOLSLISTFILEFOLDER, hostname);
                string filePath = string.Format(TOOLSLISTFILEPATH, hostname);
                if (!Directory.Exists(folderPath))
                {
                    MessageBox.Show("The Directory to store tool case list file does not exists, please check your system of execution...");
                    autoEventObj.Set();
                    return;
                }
                IEnumerable<string> batchScript = CreateTestcaseList(checkedTestcases, runTimes);
                File.WriteAllLines(filePath, batchScript);
                autoEventObj.Set();
            }
            catch (Exception e)
            {

            }
        }

        public static void ShowReportFolder(object sender, DoWorkEventArgs eventArgs)
        {
            var hostname = (string)eventArgs.Argument;

            var process = new Process();
            process.StartInfo.FileName = "explorer.exe";
            process.StartInfo.Arguments = string.Format(TOOLREPORTFOLDER, hostname);
            process.Start();
        }

    }

    public class Node : INotifyPropertyChanged
    {
        public Node()
        {
            myID = Guid.NewGuid().ToString();
        }

        private readonly ObservableCollection<Node> myChildren = new ObservableCollection<Node>();
        private readonly ObservableCollection<Node> myParent = new ObservableCollection<Node>();
        private string myText;
        private string myID;
        private bool? myIsChecked = true;
        private bool myIsExpanded;
        public bool IsTestcase = false;
        public string FileName;

        public ObservableCollection<Node> Children
        {
            get { return myChildren; }
        }

        public ObservableCollection<Node> Parent
        {
            get { return myParent; }
        }

        public bool? IsChecked
        {
            get { return myIsChecked; }
            set
            {
                myIsChecked = value;
                RaisePropertyChanged("IsChecked");
            }
        }

        public string Text
        {
            get { return myText; }
            set
            {
                myText = value;
                RaisePropertyChanged("Text");
            }
        }

        public bool IsExpanded
        {
            get { return myIsExpanded; }
            set
            {
                myIsExpanded = value;
                RaisePropertyChanged("IsExpanded");
            }
        }

        public string Id
        {
            get { return myID; }
            set { myID = value; }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void RaisePropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
            const int COUNT_CHECK = 0;
            if (propertyName == "IsChecked")
            {
                if (Id == CheckBoxId.Id && Parent.Count == 0 && Children.Count != 0)
                {
                    CheckedTreeParent(Children, IsChecked);
                }
                if (Id == CheckBoxId.Id && Parent.Count > 0 && Children.Count > 0)
                {
                    CheckedTreeChildMiddle(Parent, Children, IsChecked);
                }
                if (Id == CheckBoxId.Id && Parent.Count > 0 && Children.Count == 0)
                {
                    CheckedTreeChild(Parent, COUNT_CHECK);
                }
            }
        }

        private void CheckedTreeChildMiddle(IEnumerable<Node> itemsParent, IEnumerable<Node> itemsChild, bool? isChecked)
        {
            const int COUNT_CHECK = 0;
            CheckedTreeParent(itemsChild, isChecked);
            CheckedTreeChild(itemsParent, COUNT_CHECK);
        }

        private void CheckedTreeParent(IEnumerable<Node> items, bool? isChecked)
        {
            foreach (Node item in items)
            {
                item.IsChecked = isChecked;
                if (item.Children.Count != 0)
                {
                    CheckedTreeParent(item.Children, isChecked);
                }
            }
        }

        private void CheckedTreeChild(IEnumerable<Node> items, int countCheck)
        {
            bool isNull = false;
            foreach (Node paren in items)
            {
                foreach (Node child in paren.Children)
                {
                    if (child.IsChecked == true || child.IsChecked == null)
                    {
                        countCheck++;
                        if (child.IsChecked == null)
                        {
                            isNull = true;
                        }
                    }
                }
                if (countCheck != paren.Children.Count && countCheck != 0)
                {
                    paren.IsChecked = null;
                }
                else if (countCheck == 0)
                {
                    paren.IsChecked = false;
                }
                else if (countCheck == paren.Children.Count && isNull)
                {
                    paren.IsChecked = null;
                }
                else if (countCheck == paren.Children.Count && !isNull)
                {
                    paren.IsChecked = true;
                }
                if (paren.Parent.Count != 0)
                {
                    CheckedTreeChild(paren.Parent, 0);
                }
            }
        }
    }

    public struct CheckBoxId
    {
        public static string Id;
    }

    public class TaskParamters
    {
        public List<TestCaseInfo> TestCasesQueue
        {
            get;
            set;
        }

        public List<string>[] ErrorInfoList
        {
            get;
            set;
        }
    }

    //public class AsyncObservableCollection<T> : ObservableCollection<T>
    //{
    //    //private SynchronizationContext _synchronizationContext = SynchronizationContext.Current;
    //    //public AsyncObservableCollection() { }
    //    //public AsyncObservableCollection(IEnumerable<T> list) : base(list) { }
    //    //protected override void OnCollectionChanged(NotifyCollectionChangedEventArgs e)
    //    //{
    //    //    if (SynchronizationContext.Current == _synchronizationContext)
    //    //    {
    //    //        RaiseCollectionChanged(e);
    //    //    }
    //    //    else
    //    //    {   
    //    //        _synchronizationContext.Post(RaiseCollectionChanged, e);
    //    //    }
    //    //}
    //    //private void RaiseCollectionChanged(object param)
    //    //{
    //    //    base.OnCollectionChanged((NotifyCollectionChangedEventArgs)param);
    //    //}
    //    //protected override void OnPropertyChanged(PropertyChangedEventArgs e)
    //    //{
    //    //    if (SynchronizationContext.Current == _synchronizationContext)
    //    //    {
    //    //        RaisePropertyChanged(e);
    //    //    }
    //    //    else
    //    //    {
    //    //        _synchronizationContext.Post(RaisePropertyChanged, e);
    //    //    }
    //    //}
    //    //private void RaisePropertyChanged(object param)
    //    //{
    //    //    base.OnPropertyChanged((PropertyChangedEventArgs)param);
    //    //}
    //    protected override void OnCollectionChanged(NotifyCollectionChangedEventArgs args)
    //    {
    //        var notifyCollectionChangedEventHandler = CollectionChanged;

    //        if (notifyCollectionChangedEventHandler == null)
    //            return;

    //        foreach (NotifyCollectionChangedEventHandler handler in notifyCollectionChangedEventHandler.GetInvocationList())
    //        {
    //            var dispatcherObject = handler.Target as DispatcherObject;

    //            if (dispatcherObject != null && !dispatcherObject.CheckAccess())
    //            {
    //                dispatcherObject.Dispatcher.Invoke(DispatcherPriority.DataBind, handler, this, args);
    //            }
    //            else
    //                handler(this, args); // note : this does not execute handler in target thread's context
    //        }
    //    }
    //}
}
