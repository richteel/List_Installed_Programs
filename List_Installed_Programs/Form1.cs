using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace List_Installed_Programs
{
    public partial class Form1 : Form
    {
        private readonly DateTime NullDateTime = new DateTime(1970, 1, 1, 0, 0, 0);

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //List<SoftwareInfo> programList = new List<SoftwareInfo>();
            Dictionary<string, SoftwareInfo> programList = new Dictionary<string, SoftwareInfo>();

            txtProgramsList.AppendText("DisplayName\tDisplayVersion\tInstallDateRaw\tInstallDate\tPublisher\tURLInfoAbout\tCount\r\n");
            GetInstalledPrograms(programList, SoftwareTypes.bit64);
            GetInstalledPrograms(programList, SoftwareTypes.bit32);

            //programList

            foreach (KeyValuePair<string, SoftwareInfo> softwareInfo in programList.OrderBy(key => key.Key))
            {
                txtProgramsList.AppendText(string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\r\n",
                    softwareInfo.Value.DisplayName,
                    softwareInfo.Value.DisplayVersion,
                    softwareInfo.Value.InstallDateRaw,
                    softwareInfo.Value.InstallDate,
                    softwareInfo.Value.Publisher,
                    softwareInfo.Value.URLInfoAbout,
                    softwareInfo.Value.KeyCount));
            }
        }

        private void GetInstalledPrograms(Dictionary<string, SoftwareInfo> ProgramList, SoftwareTypes SoftwareType = SoftwareTypes.bit64)
        {
            if (ProgramList == null)
                return;

            string uninstallKey = "";

            switch (SoftwareType)
            {
                case SoftwareTypes.bit32:
                    uninstallKey = @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
                    break;
                case SoftwareTypes.bit64:
                default:
                    uninstallKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
                    break;
            }

            using (RegistryKey rk = Registry.LocalMachine.OpenSubKey(uninstallKey))
            {
                foreach (string skName in rk.GetSubKeyNames())
                {
                    using (RegistryKey sk = rk.OpenSubKey(skName))
                    {
                        SoftwareInfo softwareInfo = GetSoftwareInfoForKey(sk);

                        if (!string.IsNullOrEmpty(softwareInfo.DisplayName))
                        {
                            if(ProgramList.Keys.Contains(softwareInfo.DisplayName))
                            {
                                SoftwareInfo si = ProgramList[softwareInfo.DisplayName];

                                si.KeyCount++;

                                if (string.IsNullOrEmpty(si.DisplayVersion) && !string.IsNullOrEmpty(softwareInfo.DisplayVersion))
                                    si.DisplayVersion = softwareInfo.DisplayVersion;

                                if (string.IsNullOrEmpty(si.InstallDateRaw) && !string.IsNullOrEmpty(softwareInfo.InstallDateRaw))
                                    si.InstallDateRaw = softwareInfo.InstallDateRaw;

                                if (si.InstallDate < softwareInfo.InstallDate)
                                    si.InstallDate = softwareInfo.InstallDate;

                                if (string.IsNullOrEmpty(si.Publisher) && !string.IsNullOrEmpty(softwareInfo.Publisher))
                                    si.Publisher = softwareInfo.Publisher;

                                if (string.IsNullOrEmpty(si.URLInfoAbout) && !string.IsNullOrEmpty(softwareInfo.URLInfoAbout))
                                    si.URLInfoAbout = softwareInfo.URLInfoAbout;

                                ProgramList[softwareInfo.DisplayName] = si;
                            }
                            else
                            {
                                ProgramList.Add(softwareInfo.DisplayName, softwareInfo);
                            }
                        }
                    }
                }
            }
        }

        private SoftwareInfo GetSoftwareInfoForKey(RegistryKey sk)
        {
            string DisplayName = "";
            string DisplayVersion = "";
            string InstallDateRaw = "";
            DateTime InstallDate = NullDateTime;
            string Publisher = "";
            string URLInfoAbout = "";
            Int64 InstallDateAsInt = 0;

            if (sk.GetValue("DisplayName") != null)
                DisplayName = sk.GetValue("DisplayName").ToString().Replace("\0", string.Empty).Trim();

            if (sk.GetValue("DisplayVersion") != null)
                DisplayVersion = sk.GetValue("DisplayVersion").ToString().Replace("\0", string.Empty).Trim();

            if (sk.GetValue("InstallDate") != null)
                InstallDateRaw = sk.GetValue("InstallDate").ToString().Replace("\0", string.Empty).Trim();

            if (sk.GetValue("InstallDate") != null && !string.IsNullOrEmpty(sk.GetValue("InstallDate").ToString()))
            {
                string s = sk.GetValue("InstallDate").ToString().Replace("\0", string.Empty).Trim();

                if (s.Contains("/"))
                    DateTime.TryParse(s, out InstallDate);
                else if(s.Length == 8)
                {
                    if (int.TryParse(s.Substring(0, 4), out int year) && int.TryParse(s.Substring(4, 2), out int month) && int.TryParse(s.Substring(6, 2), out int day))
                        InstallDate = new DateTime(year, month, day);
                }
                else if (Int64.TryParse(s, out InstallDateAsInt))
                    InstallDate = InstallDate.AddSeconds(InstallDateAsInt);
            }

            if (sk.GetValue("Publisher") != null)
                Publisher = sk.GetValue("Publisher").ToString().Replace("\0", string.Empty).Trim();

            if (sk.GetValue("URLInfoAbout") != null)
                URLInfoAbout = sk.GetValue("URLInfoAbout").ToString().Replace("\0", string.Empty).Trim();


            return new SoftwareInfo(DisplayName, DisplayVersion, InstallDateRaw, InstallDate, Publisher, URLInfoAbout);
        }
    }

    public enum SoftwareTypes
    {
        bit32 = 0,
        bit64 = 1
    };

    public struct SoftwareInfo
    {
        public string DisplayName { get; }

        public string DisplayVersion { get; set; }

        public string InstallDateRaw { get; set; }

        public DateTime InstallDate { get; set; }

        public string Publisher { get; set; }

        public string URLInfoAbout { get; set; }

        public int KeyCount { get; set; }

        public SoftwareInfo(string DisplayName, string DisplayVersion, string InstallDateRaw, DateTime InstallDate, string Publisher, string URLInfoAbout)
        {
            this.DisplayName = DisplayName;
            this.DisplayVersion = DisplayVersion;
            this.InstallDateRaw = InstallDateRaw;
            this.InstallDate = InstallDate;
            this.Publisher = Publisher;
            this.URLInfoAbout = URLInfoAbout;
            this.KeyCount = 1;
        }
    }
}
