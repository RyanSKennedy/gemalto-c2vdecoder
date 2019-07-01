using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.IO;
using Sentinel.Ldk.LicGen;
using System.Collections.ObjectModel;
using Microsoft.Office.Interop.Excel;

// Nick Name 
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace C2V_Decoder
{
    /// <summary>
    /// Logics for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        #region Variable declaration
        public DispatcherTimer dispatcherTimer;
        public int repeatNumber = 0;
        public int curentPosition = 0;
        public int addCoef = 0;
        public string baseDir;
        public static string defaultPathToVC;
        public List<string> vendorCodes;
        public FingerprintData fingerprintData = new FingerprintData();
        private ObservableCollection<Criteria> _sfpData = new ObservableCollection<Criteria>();
        private ObservableCollection<Criteria> _rfpData = new ObservableCollection<Criteria>();
        #endregion

        #region Declare user classes
        public class FingerprintData
        {
            public string keyId;
            public string keyType;
            public string keyUpdateCounter;
            public string vendorId;
            public Dictionary<KeyValuePair<int, string>, Dictionary<string, string>> fingerprintDataCompare;
            public string haspVlibVersion;
            public string fridgeVersion;
            public bool cloneDetected;
            public List<ProductInfo> productInfo;

            public FingerprintData()
            { }
        }

        public class ProductInfo
        {
            public string productName;
            public int productId;
            public bool productLocked;
            public bool productFingerprintChange;
            public string productCloneProtectionExForPhysicalMachine;
            public string productCloneProtectionExForVirtualMachine;


            public ProductInfo (int i, bool pl, bool pfc, string pcpefpm, string pcpefvm, string n = "")
            {
                this.productName = n;
                this.productId = i;
                this.productLocked = pl;
                this.productFingerprintChange = pfc;
                this.productCloneProtectionExForPhysicalMachine = pcpefpm;
                this.productCloneProtectionExForVirtualMachine = pcpefvm;
            }
        }

        public class Criteria
        {
            public string CriteriaName { get; set; }
            public string Value { get; set; }
        }
        #endregion

        public MainWindow()
        {
            int majorVersion = 0;
            int minorVersion = 0;
            int buildServer = 0;
            int buildNumber = 0;

            InitializeComponent();
            // Set Base dir for app
            //============================================= 
            System.Reflection.Assembly a = System.Reflection.Assembly.GetEntryAssembly();
            baseDir = System.IO.Path.GetDirectoryName(a.Location);
            //=============================================

            // get path for Vendor Code
            //============================================= 
            defaultPathToVC = baseDir + System.IO.Path.DirectorySeparatorChar + "VendorCode" + System.IO.Path.DirectorySeparatorChar;
            //=============================================

            Load_VendorCodes();

            sntl_lg_status_t status = LicGenAPIHelper.sntl_lg_get_version(ref majorVersion, ref minorVersion, ref buildServer, ref buildNumber);
            if (sntl_lg_status_t.SNTL_LG_STATUS_OK == status)
            {
                /*handle success*/
                this.Title += " (LicGen version: " + majorVersion + "." + minorVersion + "." + buildServer + "." + buildNumber + ")";
            }
            else
            {
                /*handle error*/
                this.Title += " (LicGen version: Unknown)";
            }

            DataGrid_SystemFingerPrint.IsReadOnly = true;
            DataGrid_SystemFingerPrint.CanUserResizeColumns = false;
            DataGrid_SystemFingerPrint.CanUserResizeRows = false;
            
            DataGrid_ReferenceFingerPrint.IsReadOnly = true;
            DataGrid_ReferenceFingerPrint.CanUserResizeColumns = false;
            DataGrid_ReferenceFingerPrint.CanUserResizeRows = false;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //  DispatcherTimer setup
            // Timer using for background checking Vendor code dir
            dispatcherTimer = new System.Windows.Threading.DispatcherTimer();
            dispatcherTimer.Tick += new EventHandler(DispatcherTimer_Tick);
            dispatcherTimer.Interval = new TimeSpan(0, 0, 5);
            dispatcherTimer.Start();
        }

        private void DispatcherTimer_Tick(object sender, EventArgs e)
        {
            Load_VendorCodes();

            // Forcing the CommandManager to raise the RequerySuggested event
            CommandManager.InvalidateRequerySuggested();
        }

        private void Button_Browse_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Filter = "C2V file (.c2v)|*.c2v"; // Filter files by extension

            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                TextBox_PathToC2V.Text = dlg.FileName;
                Button_Start.IsEnabled = true;
            }
        }

        private void Button_Start_Click(object sender, RoutedEventArgs e)
        {
            // Reset data in variables
            TextBox_Results.Text = "";
            DataGrid_ReferenceFingerPrint.ItemsSource = null;
            DataGrid_SystemFingerPrint.ItemsSource = null;

            string initParam = null;
            string vendorCode = System.IO.File.ReadAllText(defaultPathToVC + System.IO.Path.DirectorySeparatorChar + ComboBox_VendorCode.SelectedValue);
            string currentState = System.IO.File.ReadAllText(TextBox_PathToC2V.Text);
            string readableState = null;

            LicGenAPIHelper licGenHelper = new LicGenAPIHelper();

            sntl_lg_status_t status = licGenHelper.sntl_lg_initialize(initParam);

            if (sntl_lg_status_t.SNTL_LG_STATUS_OK != status)
            {
                /*handle error*/
                TextBox_Results.Text += "Init Error: " + status + Environment.NewLine;
                return;
            }

            status = licGenHelper.sntl_lg_decode_current_state(vendorCode, currentState, ref readableState);

            if (sntl_lg_status_t.SNTL_LG_STATUS_OK != status)
            {
                /*handle error*/
                TextBox_Results.Text += "Decode Error: " + status + Environment.NewLine;
                return;
            }
            
            CheckData(readableState);
        }

        private void Button_SaveAs_Click(object sender, RoutedEventArgs e)
        {
            ExportToExcel();
        }

        private void TextBox_Results_TextChanged(object sender, TextChangedEventArgs e)
        {
            // No auto scrolling to down! If you need auto scrolling - uncomment this block
            //TextBox_Results.SelectionStart = TextBox_Results.Text.Length;
            //TextBox_Results.ScrollToEnd();
        }

        private void ComboBox_VendorCode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Enable/Disable Browse button
            if (ComboBox_VendorCode.SelectedIndex >= 0)
            {
                Button_Browse.IsEnabled = true;
            }
            else
            {
                Button_Browse.IsEnabled = false;
            }
        }

        private void DataGrid_ReferenceFingerPrint_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            // Auto scrolling mechanism for both linked DataGrid
            if (e.VerticalChange != 0)
            {
                int coef = 8;

                if ((curentPosition < (int)e.VerticalOffset) && ((int)e.VerticalOffset) < coef)
                {
                    addCoef = coef;
                }
                else if ((curentPosition > (int)e.VerticalOffset) && ((int)e.VerticalOffset) > coef)
                {
                    addCoef = 0;
                }
                else if ((curentPosition > (int)e.VerticalOffset) && ((int)e.VerticalOffset) == coef)
                {
                    addCoef = 0;
                }
                else if ((curentPosition > (int)e.VerticalOffset) && ((int)e.VerticalOffset) < coef)
                {
                    addCoef = 0;
                }
                else
                {
                    addCoef = coef;
                }

                object currentPos = DataGrid_SystemFingerPrint.Items[(int)e.VerticalOffset + addCoef];
                DataGrid_SystemFingerPrint.ScrollIntoView(currentPos);
                DataGrid_SystemFingerPrint.UpdateLayout();
                curentPosition += (curentPosition < (int)e.VerticalOffset) ? 1 : -1;
            }
        }
        
        public void Load_VendorCodes()
        {
            // Check and load Vendor codes from dir to ListBox
            if (Directory.Exists(defaultPathToVC))
            {
                List<string> files = Directory.GetFiles(defaultPathToVC, "*.hvc").ToList<string>();

                if (files.Count() > 0)
                {
                    repeatNumber = 0;

                    vendorCodes = files;
                    Update_CheckBox(vendorCodes);
                }
                else if (repeatNumber == 0)
                {
                    ComboBox_VendorCode.SelectedIndex = -1;
                    ComboBox_VendorCode.Items.Clear();
                    TextBox_PathToC2V.Text = "...";
                    Button_Browse.IsEnabled = false;
                    Button_Start.IsEnabled = false;
                    MessageBox.Show("Not found any *.hvc file in directory VendorCode!" + Environment.NewLine +
                    "Please do this steps:" + Environment.NewLine +
                    "1. Close c2vdecoder.exe;" + Environment.NewLine +
                    "2. Put your Vendor code files (*.hvc) in directory VendorCode;" + Environment.NewLine +
                    "3. Run c2vdecoder.exe again." + Environment.NewLine, "Error");
                    repeatNumber++;
                }
            }
            else if (repeatNumber == 1)
            {
                ComboBox_VendorCode.SelectedIndex = -1;
                ComboBox_VendorCode.Items.Clear();
                TextBox_PathToC2V.Text = "...";
                Button_Browse.IsEnabled = false;
                Button_Start.IsEnabled = false;
                MessageBox.Show("Not found VendorCode directory!" + Environment.NewLine +
                    "Please do this steps:" + Environment.NewLine +
                    "1. Close c2vdecoder.exe;" + Environment.NewLine +
                    "2. Create new folder \"VendorCode\" in directory with c2vdecoder.exe;" + Environment.NewLine +
                    "3. Put your Vendor code files (*.hvc) in this directory;" + Environment.NewLine +
                    "4. Run c2vdecoder.exe again." + Environment.NewLine, "Error");
                repeatNumber++;
            }
        }

        public void Update_CheckBox(List<string> val)
        {
            // Mechanism for updating ListBox with recheck exist values and add/remove instances
            foreach (string el in val)
            {
                string[] tmpMass = el.Split(System.IO.Path.DirectorySeparatorChar);
                if (!ComboBox_VendorCode.Items.Contains(tmpMass[tmpMass.Length - 1]))
                {
                    ComboBox_VendorCode.Items.Add(tmpMass[tmpMass.Length - 1]);
                }
            }

            List<string> valForRemove = new List<string>();
            foreach (string el2 in ComboBox_VendorCode.Items)
            {
                bool candidat = true;
                foreach (string el3 in val)
                {
                    if (el3.Contains(el2))
                    {
                        candidat = false;
                        break;
                    }
                }

                if (candidat)
                {
                    valForRemove.Add(el2);
                }
            }

            if (valForRemove.Count() > 0)
            {
                foreach (string el4 in valForRemove)
                {
                    ComboBox_VendorCode.Items.Remove(el4);
                }

                if (ComboBox_VendorCode.Items.Count == 0)
                {
                    ComboBox_VendorCode.SelectedIndex = -1;
                }
            }
        }
        
        public void CheckData(string data)
        {
            // Analysing decoded C2V and convert data for report and load it to DataGrid viewer and to result's TextBox
            int i = 0;

            #region Test cases
            //// test with dif val cpu in system_fingerprint
            //System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex("<value>905426015</value>");
            //data = reg.Replace(data, "<value>905426017</value>", 1);
            //// ---

            //// test with +1 new criteria in system_fingerprint
            //System.Text.RegularExpressions.Regex reg2 = new System.Text.RegularExpressions.Regex("<criteria>\n            <name>ip_address</name>\n            <value>431891170</value>\n          </criteria>\n");
            //data = reg2.Replace(data, "<criteria>\n            <name>ip_address</name>\n            <value>431891170</value>\n          </criteria>\n          <criteria>\n            <name>ip_TEST</name>\n            <value>Hahaha</value>\n          </criteria>\n", 1);
            //// ---

            //// test with dif val cpu in reference_fingerprint
            //System.Text.RegularExpressions.Regex reg3 = new System.Text.RegularExpressions.Regex("<reference_fingerprint>\n        <fingerprint_control_type>ISV Managed</fingerprint_control_type>\n        <raw_data>MXhJSW8JQTZwFMCrJtw6RR844No9yaHMAgjSdKDEaY0D6pSfTqQEgIpe18KmPiRIzvn6NagAmINmRH4BXwKOmjhwowhBt2hDbpEaKU9emJlT3QAIEH3RjoiAD0wKMPBQCcBk6gWkHohkg6yKjAE0AAYvaEECq0EAV16dqzcM8FWLAfJCODxV</raw_data>\n        <fingerprint_info>\n          <criteria>\n            <name>cpu</name>\n            <value>905426015</value>\n          </criteria>\n          <criteria>\n            <name>cpu</name>\n            <value>905426015</value>\n");
            //data = reg3.Replace(data, "<reference_fingerprint>\n        <fingerprint_control_type>ISV Managed</fingerprint_control_type>\n        <raw_data>MXhJSW8JQTZwFMCrJtw6RR844No9yaHMAgjSdKDEaY0D6pSfTqQEgIpe18KmPiRIzvn6NagAmINmRH4BXwKOmjhwowhBt2hDbpEaKU9emJlT3QAIEH3RjoiAD0wKMPBQCcBk6gWkHohkg6yKjAE0AAYvaEECq0EAV16dqzcM8FWLAfJCODxV</raw_data>\n        <fingerprint_info>\n          <criteria>\n            <name>cpu</name>\n            <value>905426015</value>\n          </criteria>\n          <criteria>\n            <name>cpu</name>\n            <value>905426021</value>\n", 1);
            //// ---

            //// test with +1 new criteria in reference_fingerprint
            //System.Text.RegularExpressions.Regex reg4 = new System.Text.RegularExpressions.Regex("<reference_fingerprint>\n        <fingerprint_control_type>ISV Managed</fingerprint_control_type>\n        <raw_data>MXhJSW8JQTZwFMCrJtw6RR844No9yaHMAgjSdKDEaY0D6pSfTqQEgIpe18KmPiRIzvn6NagAmINmRH4BXwKOmjhwowhBt2hDbpEaKU9emJlT3QAIEH3RjoiAD0wKMPBQCcBk6gWkHohkg6yKjAE0AAYvaEECq0EAV16dqzcM8FWLAfJCODxV</raw_data>\n        <fingerprint_info>\n          <criteria>\n            <name>cpu</name>\n            <value>905426015</value>\n          </criteria>\n          <criteria>\n            <name>cpu</name>\n            <value>905426015</value>\n          </criteria>\n");
            //data = reg4.Replace(data, "<reference_fingerprint>\n        <fingerprint_control_type>ISV Managed</fingerprint_control_type>\n        <raw_data>MXhJSW8JQTZwFMCrJtw6RR844No9yaHMAgjSdKDEaY0D6pSfTqQEgIpe18KmPiRIzvn6NagAmINmRH4BXwKOmjhwowhBt2hDbpEaKU9emJlT3QAIEH3RjoiAD0wKMPBQCcBk6gWkHohkg6yKjAE0AAYvaEECq0EAV16dqzcM8FWLAfJCODxV</raw_data>\n        <fingerprint_info>\n          <criteria>\n            <name>cpu</name>\n            <value>905426015</value>\n          </criteria>\n          <criteria>\n            <name>cpu</name>\n            <value>905426021</value>\n          </criteria>\n          <criteria>\n            <name>cpu</name>\n            <value>905426015</value>\n          </criteria>\n", 1);
            //// ---
            #endregion

            XmlDocument xmlFullData = new XmlDocument();
            xmlFullData.LoadXml(data);

            if (!(data.Contains("<reference_fingerprint>") && data.Contains("<system_fingerprint>")))
            {
                TextBox_Results.Text += "Error: <system_fingerprint> or <reference_fingerprint> doesn't exist in C2V!" + Environment.NewLine;
                TextBox_Results.Text += "Decoded C2V is:" + Environment.NewLine;
                TextBox_Results.Text += "//---" + Environment.NewLine;
                TextBox_Results.Text += XDocument.Parse(xmlFullData.InnerXml) + Environment.NewLine;
                TextBox_Results.Text += "//---" + Environment.NewLine;

                Button_SaveAs.IsEnabled = false;
                return;
            }
            else
            {
                Button_SaveAs.IsEnabled = true;
            }

            XmlDocument xmlReferenceFingerprint = new XmlDocument();
            xmlReferenceFingerprint.LoadXml("<reference_fingerprint>" + xmlFullData.SelectSingleNode(@"//sentinel_ldk_info/key/configuration_info/reference_fingerprint").InnerXml + "</reference_fingerprint>");

            XmlDocument xmlSystemFingerprint = new XmlDocument();
            xmlSystemFingerprint.LoadXml("<system_fingerprint>" + xmlFullData.SelectSingleNode(@"//sentinel_ldk_info/key/configuration_info/system_fingerprint").InnerXml + "</system_fingerprint>");

            fingerprintData.fingerprintDataCompare = new Dictionary<KeyValuePair<int, string>, Dictionary<string, string>>();

            if (data.Contains("<id>"))
            {
                fingerprintData.keyId = xmlFullData.SelectSingleNode(@"//sentinel_ldk_info/key/id").InnerText;
            }

            if (data.Contains("<type>"))
            {
                fingerprintData.keyType = xmlFullData.SelectSingleNode(@"//sentinel_ldk_info/key/type").InnerText;
            }

            if (data.Contains("<update_counter>"))
            {
                fingerprintData.keyUpdateCounter = xmlFullData.SelectSingleNode(@"//sentinel_ldk_info/key/update_counter").InnerText;
            }

            if (data.Contains("<vendor>"))
            {
                fingerprintData.vendorId = xmlFullData.SelectSingleNode(@"//sentinel_ldk_info/key/vendor/id").InnerText;
            }

            if (data.Contains("<vlib_version>"))
            {
                fingerprintData.haspVlibVersion = xmlFullData.SelectSingleNode(@"//sentinel_ldk_info/key/configuration_info/vlib_version").InnerText;
            }

            if (data.Contains("<fridge_version>"))
            {
                fingerprintData.fridgeVersion = xmlFullData.SelectSingleNode(@"//sentinel_ldk_info/key/configuration_info/fridge_version").InnerText;
            }

            if (data.Contains("<clone_detected>"))
            {
                fingerprintData.cloneDetected = (xmlFullData.SelectSingleNode(@"//sentinel_ldk_info/key/clone_detected").InnerText == "Yes") ? true : false;
            }
            
            fingerprintData.productInfo = new List<ProductInfo>();
            foreach (XmlElement elp in xmlFullData.SelectNodes(@"//sentinel_ldk_info/key/product"))
            {
                fingerprintData.productInfo.Add(new ProductInfo(
                    Convert.ToInt32(elp.SelectSingleNode(@"id").InnerText),
                    (elp.SelectSingleNode(@"//locked").InnerText == "Yes") ? true : false,
                    (elp.SelectSingleNode(@"//fingerprint_change").InnerText == "Yes") ? true : false,
                    elp.SelectSingleNode(@"//clone_protection_ex/physical_machine").InnerText,
                    elp.SelectSingleNode(@"//clone_protection_ex/virtual_machine").InnerText,
                    elp.SelectSingleNode(@"name").InnerText));
            }
            
            i = 0;
            string currentCriteria = "";
            Dictionary<string, Dictionary<int, KeyValuePair<string, bool>>> tmpDataSF = new Dictionary<string, Dictionary<int,KeyValuePair<string, bool>>>();
            foreach (XmlElement elsf in xmlSystemFingerprint.SelectNodes(@"//system_fingerprint/fingerprint_info/criteria"))
            {
                if (elsf.FirstChild.InnerText == currentCriteria)
                {
                    tmpDataSF[currentCriteria].Add(i, new KeyValuePair<string, bool>(elsf.LastChild.InnerText, false));
                }
                else
                {
                    currentCriteria = elsf.FirstChild.InnerText;
                    tmpDataSF.Add(currentCriteria, new Dictionary<int, KeyValuePair<string, bool>>() { { i, new KeyValuePair<string, bool>(elsf.LastChild.InnerText, false) } });
                }
                i++;
            }

            i = 0;
            Dictionary<string, Dictionary<int, KeyValuePair<string, bool>>> tmpDataRF = new Dictionary<string, Dictionary<int, KeyValuePair<string, bool>>>();
            foreach (XmlElement elrf in xmlReferenceFingerprint.SelectNodes(@"//reference_fingerprint/fingerprint_info/criteria"))
            {
                if (elrf.FirstChild.InnerText == currentCriteria)
                {
                    tmpDataRF[currentCriteria].Add(i, new KeyValuePair<string, bool>(elrf.LastChild.InnerText, false));
                }
                else
                {
                    currentCriteria = elrf.FirstChild.InnerText;
                    tmpDataRF.Add(currentCriteria, new Dictionary<int, KeyValuePair<string, bool>>() { { i, new KeyValuePair<string, bool>(elrf.LastChild.InnerText, false) } });
                }
                i++;
            }

            var tmpDataSFsorted = tmpDataSF.OrderBy(x => x.Key);
            tmpDataSF = tmpDataSFsorted.ToDictionary(t => t.Key, t => t.Value);

            var tmpDataRFsorted = tmpDataRF.OrderBy(x => x.Key);
            tmpDataRF = tmpDataRFsorted.ToDictionary(t => t.Key, t => t.Value);

            var tmpDataFBeforeSorting = ConvertData(tmpDataSF, tmpDataRF, new Dictionary<KeyValuePair<int, string>, Dictionary<string, string>>());
            var tmpDataFSorted = tmpDataFBeforeSorting.OrderBy(x => x.Key.Value);
            fingerprintData.fingerprintDataCompare = tmpDataFSorted.ToDictionary(t => t.Key, t => t.Value);

            DataGrid_SystemFingerPrint.AlternatingRowBackground = Brushes.LightGray;
            DataGrid_ReferenceFingerPrint.AlternatingRowBackground = Brushes.LightGray;
            
            _sfpData = new ObservableCollection<Criteria>();
            _rfpData = new ObservableCollection<Criteria>();
            foreach (var el in fingerprintData.fingerprintDataCompare)
            {
                _sfpData.Add(new Criteria()
                {
                    CriteriaName = el.Key.Value,
                    Value = (el.Value["system_fingerprint"] != "" ? el.Value["system_fingerprint"] : "-")
                });

                _rfpData.Add(new Criteria()
                {
                    CriteriaName = el.Key.Value,
                    Value = (el.Value["reference_fingerprint"] != "" ? el.Value["reference_fingerprint"] : "-")
                });
                
                DataGrid_SystemFingerPrint.ItemsSource = _sfpData;
                DataGrid_ReferenceFingerPrint.ItemsSource = _rfpData;

                DataGrid_SystemFingerPrint.UpdateLayout();
                var sfRow = (DataGridRow)DataGrid_SystemFingerPrint.ItemContainerGenerator.ContainerFromItem((Criteria)DataGrid_SystemFingerPrint.Items[DataGrid_SystemFingerPrint.Items.Count - 1]);

                DataGrid_ReferenceFingerPrint.UpdateLayout();
                var rfRow = (DataGridRow)DataGrid_ReferenceFingerPrint.ItemContainerGenerator.ContainerFromItem((Criteria)DataGrid_ReferenceFingerPrint.Items[DataGrid_ReferenceFingerPrint.Items.Count - 1]);
                
                if (el.Value["system_fingerprint"] != el.Value["reference_fingerprint"] && (sfRow != null && rfRow != null))
                {
                    sfRow.Background = Brushes.LightPink;
                    rfRow.Background = Brushes.LightPink;
                }
            }

            DataGrid_SystemFingerPrint.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
            DataGrid_SystemFingerPrint.IsReadOnly = true;
            DataGrid_ReferenceFingerPrint.IsReadOnly = true;

            DataGrid_SystemFingerPrint.Columns[0].MinWidth = 125;
            DataGrid_SystemFingerPrint.Columns[0].MaxWidth = 125;
            DataGrid_SystemFingerPrint.Columns[1].MinWidth = 127;

            DataGrid_ReferenceFingerPrint.Columns[0].MinWidth = 125;
            DataGrid_ReferenceFingerPrint.Columns[0].MaxWidth = 125;
            DataGrid_ReferenceFingerPrint.Columns[1].MinWidth = 110;

            TextBox_Results.Text = "";
            TextBox_Results.Text += "Key " + fingerprintData.keyType + Environment.NewLine;
            TextBox_Results.Text += "ID: " + fingerprintData.keyId + Environment.NewLine;
            TextBox_Results.Text += "Key update counter: " + fingerprintData.keyUpdateCounter + Environment.NewLine;
            TextBox_Results.Text += "Vendor ID: " + fingerprintData.vendorId + Environment.NewLine;
            TextBox_Results.Text += "HaspVlib version: " + fingerprintData.haspVlibVersion + Environment.NewLine;
            TextBox_Results.Text += "Fridge version: " + fingerprintData.fridgeVersion + Environment.NewLine;
            TextBox_Results.Text += "Clone detected: " + (fingerprintData.cloneDetected ? "YES" : "NO") + Environment.NewLine;
            TextBox_Results.Text += "//---" + Environment.NewLine;

            bool haveBlocked = false;
            foreach (var el in fingerprintData.productInfo)
            {
                if (el.productFingerprintChange)
                {
                    TextBox_Results.Text += "Product: " + el.productName + "(ID=" + el.productId + ") - is blocked!" + Environment.NewLine;
                    TextBox_Results.Text += "PM Scheme is: " + el.productCloneProtectionExForPhysicalMachine + Environment.NewLine;
                    TextBox_Results.Text += "VM Scheme is: " + el.productCloneProtectionExForVirtualMachine + Environment.NewLine;
                    TextBox_Results.Text += "//---" + Environment.NewLine;
                    haveBlocked = true;
                }
            }

            if (!haveBlocked)
            {
                TextBox_Results.Text += "All Product's is unblocked!" + Environment.NewLine;
                TextBox_Results.Text += "//---" + Environment.NewLine;
            }

            foreach (var elCriteria in fingerprintData.fingerprintDataCompare)
            {
                if (elCriteria.Value["system_fingerprint"] != elCriteria.Value["reference_fingerprint"])
                {
                    TextBox_Results.Text += "Criteria " + elCriteria.Key.Value + " is different: " +
                        (elCriteria.Value["system_fingerprint"] != "" ? elCriteria.Value["system_fingerprint"] : "none") + "(SystemFP) != " +
                        (elCriteria.Value["reference_fingerprint"] != "" ? elCriteria.Value["reference_fingerprint"] : "none") + "(ReferenceFP)" + Environment.NewLine;
                }
            }
        }

        private Dictionary<KeyValuePair<int, string>, Dictionary<string, string>> ConvertData(Dictionary<string, Dictionary<int, KeyValuePair<string, bool>>> tmpDataSF, Dictionary<string, Dictionary<int, KeyValuePair<string, bool>>> tmpDataRF, Dictionary<KeyValuePair<int, string>, Dictionary<string, string>> tmpRes)
        {
            // Convert-compare data from System and Reference Fingerprint
            foreach (var el1 in tmpDataSF.ToList())
            {
                foreach (var el2 in tmpDataRF.ToList())
                {
                    if (el1.Key == el2.Key)
                    {
                        foreach (var subEl1 in el1.Value.ToList())
                        {
                            foreach (var subEl2 in el2.Value.ToList())
                            {
                                if (subEl2.Value.Key == subEl1.Value.Key && !subEl2.Value.Value && !subEl1.Value.Value)
                                {
                                    tmpRes.Add(new KeyValuePair<int, string>(tmpRes.Count, el1.Key), new Dictionary<string, string> { { "system_fingerprint", subEl1.Value.Key }, { "reference_fingerprint", subEl2.Value.Key } });
                                    tmpDataSF[el1.Key][subEl1.Key] = new KeyValuePair<string, bool>(tmpDataSF[el1.Key][subEl1.Key].Key, true);
                                    tmpDataRF[el2.Key][subEl2.Key] = new KeyValuePair<string, bool>(tmpDataRF[el2.Key][subEl2.Key].Key, true);
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            foreach (var reCheckSFCriteria in tmpDataSF.ToList())
            {
                foreach (var reCheckSFCriteriaVal in reCheckSFCriteria.Value.ToList())
                {
                    if (reCheckSFCriteriaVal.Value.Value == false)
                    {
                        tmpRes.Add(new KeyValuePair<int, string>(tmpRes.Count, reCheckSFCriteria.Key), new Dictionary<string, string> { { "system_fingerprint", reCheckSFCriteriaVal.Value.Key }, { "reference_fingerprint", "" } });
                        
                        tmpDataSF[reCheckSFCriteria.Key][reCheckSFCriteriaVal.Key] = new KeyValuePair<string, bool>(reCheckSFCriteriaVal.Value.Key, true);
                    }
                }
            }

            foreach (var reCheckRFCriteria in tmpDataRF.ToList())
            {
                foreach (var reCheckRFCriteriaVal in reCheckRFCriteria.Value.ToList())
                {
                    if (reCheckRFCriteriaVal.Value.Value == false)
                    {
                        if (tmpRes.Keys.Where(x => x.Value == reCheckRFCriteria.Key).Count() <= 0)
                        {
                            tmpRes.Add(new KeyValuePair<int, string>(tmpRes.Count, reCheckRFCriteria.Key), new Dictionary<string, string> { { "system_fingerprint", "" }, { "reference_fingerprint", reCheckRFCriteriaVal.Value.Key } });
                        }
                        else
                        {
                            bool isDone = false;
                            var tmpKey = tmpRes.Keys.Where(x => x.Value == reCheckRFCriteria.Key);
                            foreach (var key in tmpKey)
                            {
                                if (tmpRes[key].ContainsKey("reference_fingerprint") && tmpRes[key]["reference_fingerprint"] == "")
                                {
                                    tmpRes[key]["reference_fingerprint"] = reCheckRFCriteriaVal.Value.Key;
                                    isDone = true;
                                }
                                else if (!tmpRes[key].ContainsKey("reference_fingerprint"))
                                {
                                    tmpRes[key].Add("reference_fingerprint", reCheckRFCriteriaVal.Value.Key);
                                    isDone = true;
                                }
                            }

                            if (!isDone)
                            {
                                tmpRes.Add(new KeyValuePair<int, string>(tmpRes.Count, reCheckRFCriteria.Key), new Dictionary<string, string> { { "system_fingerprint", "" }, { "reference_fingerprint", reCheckRFCriteriaVal.Value.Key } });
                            }
                        }

                        tmpDataRF[reCheckRFCriteria.Key][reCheckRFCriteriaVal.Key] = new KeyValuePair<string, bool>(reCheckRFCriteriaVal.Value.Key, true);
                    }
                }
            }

            return tmpRes;
        }
        
        private void ExportToExcel()
        {
            // Method for export data to excel report
            System.Globalization.CultureInfo Oldci = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");

            try
            {
                Excel.Application exApp = new Excel.Application();
                exApp.Visible = false;
                Object missing = Type.Missing;
                exApp.Workbooks.Add(missing);

                Worksheet workSheet = (Worksheet)exApp.ActiveSheet;

                #region Set sheet name
                workSheet.Name = "Report";
                ((Excel.Range)workSheet.Columns[1]).ColumnWidth = 20;
                ((Excel.Range)workSheet.Columns[2]).ColumnWidth = 24;
                ((Excel.Range)workSheet.Columns[3]).ColumnWidth = 15;
                ((Excel.Range)workSheet.Columns[4]).ColumnWidth = 15;
                ((Excel.Range)workSheet.Columns[5]).ColumnWidth = 15;
                ((Excel.Range)workSheet.Columns[6]).ColumnWidth = 20;
                #endregion

                //-------------------

                #region Set key info headers
                (workSheet.Cells[1, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[1, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[1, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[1, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[1, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[1, 1] = "Vendor ID:";

                (workSheet.Cells[2, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[2, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[2, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[2, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[2, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[2, 1] = "Key ID:";

                (workSheet.Cells[3, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[3, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[3, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[3, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[3, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[3, 1] = "Key type:";

                (workSheet.Cells[4, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[4, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[4, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[4, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[4, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[4, 1] = "Update counter:";

                (workSheet.Cells[5, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[5, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[5, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[5, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[5, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[5, 1] = "Clone detected:";

                (workSheet.Cells[6, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[6, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[6, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[6, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[6, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[6, 1] = "Haspvlib version:";

                (workSheet.Cells[7, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[7, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[7, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[7, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[7, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[7, 1] = "Fridge version:";
                #endregion

                #region Set key info data 
                (workSheet.Cells[1, 2] as Excel.Range).NumberFormat = "@";
                (workSheet.Cells[1, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                (workSheet.Cells[1, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[1, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[1, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[1, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[1, 2] = fingerprintData.vendorId;

                (workSheet.Cells[2, 2] as Excel.Range).NumberFormat = "@";
                (workSheet.Cells[2, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                (workSheet.Cells[2, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[2, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[2, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[2, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[2, 2] = fingerprintData.keyId;

                (workSheet.Cells[3, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                (workSheet.Cells[3, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[3, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[3, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[3, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[3, 2] = fingerprintData.keyType;

                (workSheet.Cells[4, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                (workSheet.Cells[4, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[4, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[4, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[4, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[4, 2] = fingerprintData.keyUpdateCounter;

                (workSheet.Cells[5, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                (workSheet.Cells[5, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[5, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[5, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[5, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[5, 2] = (fingerprintData.cloneDetected ? "YES" : "NO");

                (workSheet.Cells[6, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                (workSheet.Cells[6, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[6, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[6, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[6, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[6, 2] = fingerprintData.haspVlibVersion;

                (workSheet.Cells[7, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                (workSheet.Cells[7, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[7, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[7, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[7, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[7, 2] = fingerprintData.fridgeVersion;
                #endregion

                //-----------------

                #region Set Fingerprint info headers
                (workSheet.Cells[9, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[9, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                (workSheet.Cells[9, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[9, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[9, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[9, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[9, 1] = "Criteria";

                (workSheet.Cells[9, 2] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[9, 2] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                (workSheet.Cells[9, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[9, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[9, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[9, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[9, 2] = "Reference FP";

                (workSheet.Cells[9, 3] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[9, 3] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[9, 3] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                (workSheet.Cells[9, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[9, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[9, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[9, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[9, 3] = "System FP";
                #endregion

                #region Set Fingerprint info data
                int rowExcel = 10;
                foreach (var elCriteria in fingerprintData.fingerprintDataCompare)
                {
                    if (elCriteria.Value["system_fingerprint"] != elCriteria.Value["reference_fingerprint"])
                    {
                        (workSheet.Cells[rowExcel, 1] as Excel.Range).Interior.Color = Excel.XlRgbColor.rgbRed;
                        (workSheet.Cells[rowExcel, 2] as Excel.Range).Interior.Color = Excel.XlRgbColor.rgbRed;
                        (workSheet.Cells[rowExcel, 3] as Excel.Range).Interior.Color = Excel.XlRgbColor.rgbRed;
                    }

                    (workSheet.Cells[rowExcel, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    (workSheet.Cells[rowExcel, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    workSheet.Cells[rowExcel, 1] = elCriteria.Key.Value;

                    (workSheet.Cells[rowExcel, 2] as Excel.Range).NumberFormat = "@";
                    (workSheet.Cells[rowExcel, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    (workSheet.Cells[rowExcel, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    workSheet.Cells[rowExcel, 2] = elCriteria.Value["reference_fingerprint"];

                    (workSheet.Cells[rowExcel, 3] as Excel.Range).NumberFormat = "@";
                    (workSheet.Cells[rowExcel, 3] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    (workSheet.Cells[rowExcel, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    workSheet.Cells[rowExcel, 3] = elCriteria.Value["system_fingerprint"];

                    rowExcel++;
                }
                #endregion

                //-----------------

                #region Set Product info headers
                rowExcel += 2;
                (workSheet.Cells[rowExcel, 1] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[rowExcel, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[rowExcel, 1] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                (workSheet.Cells[rowExcel, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[rowExcel, 1] = "Product ID";

                (workSheet.Cells[rowExcel, 2] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[rowExcel, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[rowExcel, 2] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                (workSheet.Cells[rowExcel, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[rowExcel, 2] = "Product Name";

                (workSheet.Cells[rowExcel, 3] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[rowExcel, 3] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[rowExcel, 3] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                (workSheet.Cells[rowExcel, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[rowExcel, 3] = "PM Scheme";

                (workSheet.Cells[rowExcel, 4] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[rowExcel, 4] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[rowExcel, 4] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                (workSheet.Cells[rowExcel, 4] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 4] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 4] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 4] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[rowExcel, 4] = "VM Scheme";

                (workSheet.Cells[rowExcel, 5] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[rowExcel, 5] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[rowExcel, 5] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                (workSheet.Cells[rowExcel, 5] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 5] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 5] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 5] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[rowExcel, 5] = "FP Changed";

                (workSheet.Cells[rowExcel, 6] as Excel.Range).Font.Bold = true;
                (workSheet.Cells[rowExcel, 6] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                (workSheet.Cells[rowExcel, 6] as Excel.Range).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                (workSheet.Cells[rowExcel, 6] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 6] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 6] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                (workSheet.Cells[rowExcel, 6] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                workSheet.Cells[rowExcel, 6] = "Locked";
                #endregion

                #region Set Product info data
                rowExcel++;
                foreach (var elProduct in fingerprintData.productInfo)
                {
                    (workSheet.Cells[rowExcel, 1] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    (workSheet.Cells[rowExcel, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 1] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    workSheet.Cells[rowExcel, 1] = elProduct.productId;

                    (workSheet.Cells[rowExcel, 2] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    (workSheet.Cells[rowExcel, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 2] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    workSheet.Cells[rowExcel, 2] = (elProduct.productName != "" ? elProduct.productName : "-");

                    (workSheet.Cells[rowExcel, 3] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    (workSheet.Cells[rowExcel, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 3] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    workSheet.Cells[rowExcel, 3] = elProduct.productCloneProtectionExForPhysicalMachine;

                    (workSheet.Cells[rowExcel, 4] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    (workSheet.Cells[rowExcel, 4] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 4] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 4] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 4] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    workSheet.Cells[rowExcel, 4] = elProduct.productCloneProtectionExForVirtualMachine;

                    (workSheet.Cells[rowExcel, 5] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    (workSheet.Cells[rowExcel, 5] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 5] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 5] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 5] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    workSheet.Cells[rowExcel, 5] = elProduct.productFingerprintChange ? "YES" : "NO";

                    (workSheet.Cells[rowExcel, 6] as Excel.Range).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    (workSheet.Cells[rowExcel, 6] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 6] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 6] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                    (workSheet.Cells[rowExcel, 6] as Excel.Range).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                    workSheet.Cells[rowExcel, 6] = elProduct.productLocked ? "YES" : "NO(Unlocked/Trialware)";

                    rowExcel++;
                }
                #endregion

                //-----------------

                #region Save to file
                SaveFileDialog saveFileDialogForExport = new SaveFileDialog();
                saveFileDialogForExport.Filter = "xlsx files (*.xlsx)|*.xlsx";
                if (saveFileDialogForExport.ShowDialog() == true)
                {
                    try
                    {
                        workSheet.SaveAs(saveFileDialogForExport.FileName);
                        exApp.Workbooks.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Could not create file! Error: " + ex);
                    }
                }
                #endregion
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: " + e.Message + 
                    (e.Message.Contains("REGDB_E_CLASSNOTREG") ?
                    Environment.NewLine + Environment.NewLine + "Maybe on your PC not installed Microsoft Excel." + 
                    Environment.NewLine + "Please install it for saving report, or using reporting view in app." 
                    : ""),
                    "Error in process of saving report", 
                    MessageBoxButton.OK);
            }
        }
    }
}
