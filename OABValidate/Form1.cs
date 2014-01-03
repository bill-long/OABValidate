using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.DirectoryServices;
using System.Threading;
using System.IO;
using System.DirectoryServices.Protocols;

namespace OABValidate
{
	public partial class Form1 : Form
	{
        struct Attribute
        {
            public string Name;
            public bool IsBacklink;
        }

        Attribute[] defaultAttList = new Attribute[] {
            new Attribute() { Name = "altRecipient", IsBacklink = false },
            new Attribute() { Name = "altRecipientBL", IsBacklink = true },
            new Attribute() { Name = "assistant", IsBacklink = false },
            new Attribute() { Name = "authOrig", IsBacklink = false },
            new Attribute() { Name = "authOrigBL", IsBacklink = true },
            new Attribute() { Name = "bridgeheadServerListBL", IsBacklink = true },
            new Attribute() { Name = "defaultClassStore", IsBacklink = false },
            new Attribute() { Name = "directReports", IsBacklink = false },
            new Attribute() { Name = "distinguishedName", IsBacklink = false },
            new Attribute() { Name = "dLMemRejectPerms", IsBacklink = false },
            new Attribute() { Name = "dLMemRejectPermsBL", IsBacklink = true },
            new Attribute() { Name = "dLMemSubmitPerms", IsBacklink = false },
            new Attribute() { Name = "dLMemSubmitPermsBL", IsBacklink = true },
            new Attribute() { Name = "dynamicLDAPServer", IsBacklink = false },
            new Attribute() { Name = "homeMDB", IsBacklink = false },
            new Attribute() { Name = "homeMTA", IsBacklink = false },
            new Attribute() { Name = "isPrivilegeHolder", IsBacklink = false },
            new Attribute() { Name = "kMServer", IsBacklink = false },
            new Attribute() { Name = "lastKnownParent", IsBacklink = false },
            new Attribute() { Name = "managedObjects", IsBacklink = true },
            new Attribute() { Name = "managedBy", IsBacklink = false },
            new Attribute() { Name = "manager", IsBacklink = false },
            new Attribute() { Name = "member", IsBacklink = false },
            new Attribute() { Name = "memberOf", IsBacklink = true },
            new Attribute() { Name = "msExchConferenceMailboxBL", IsBacklink = true },
            new Attribute() { Name = "msExchControllingZone", IsBacklink = false },
            new Attribute() { Name = "msExchIMVirtualServer", IsBacklink = false },
            new Attribute() { Name = "msExchQueryBaseDN", IsBacklink = false },
            new Attribute() { Name = "msExchUseOAB", IsBacklink = false },
            new Attribute() { Name = "netbootSCPBL", IsBacklink = false },
            new Attribute() { Name = "nonSecurityMemberBL", IsBacklink = false },
            new Attribute() { Name = "ownerBL", IsBacklink = true },
            new Attribute() { Name = "preferredOU", IsBacklink = false },
            new Attribute() { Name = "publicDelegates", IsBacklink = false },
            new Attribute() { Name = "publicDelegatesBL", IsBacklink = true },
            new Attribute() { Name = "queryPolicyBL", IsBacklink = true },
            new Attribute() { Name = "secretary", IsBacklink = false },
            new Attribute() { Name = "seeAlso", IsBacklink = false },
            new Attribute() { Name = "serverReferenceBL", IsBacklink = true },
            new Attribute() { Name = "showInAddressBook", IsBacklink = false },
            new Attribute() { Name = "siteObjectBL", IsBacklink = true },
            new Attribute() { Name = "unAuthOrig", IsBacklink = false },
            new Attribute() { Name = "unAuthOrigBL", IsBacklink = true }
        };

        string[] defaultProplist = new string[] { 
        "altRecipient", 
        "altRecipientBL", 
        "assistant", 
        "authOrig",
        "authOrigBL", 
        "bridgeheadServerListBL",
        "defaultClassStore", 
        "directReports",
        "distinguishedName", 
        "dLMemRejectPerms",
        "dLMemRejectPermsBL",
        "dLMemSubmitPerms",
        "dLMemSubmitPermsBL", 
        "dynamicLDAPServer", 
        "homeMDB", 
        "homeMTA",
        "isPrivilegeHolder",
        "kMServer", 
        "lastKnownParent", 
        "managedObjects",
        "managedBy",
        "manager",
        "masteredBy",
        "member", 
        "memberOf", 
        "msExchConferenceMailboxBL",
        "msExchControllingZone", 
        "msExchIMVirtualServer", 
        "msExchQueryBaseDN", 
        "msExchUseOAB", 
        "netbootSCPBL", 
        "nonSecurityMemberBL", 
        "ownerBL",
        "preferredOU",
        "publicDelegates",
        "publicDelegatesBL", 
        "queryPolicyBL",
        "secretary",
        "seeAlso", 
        "serverReferenceBL", 
        "showInAddressBook",
        "siteObjectBL",
        "unAuthOrig",
        "unAuthOrigBL"
            };

        int objectsFound = 0;
        int objectsProcessed = 0;
        int problemObjects = 0;
        string logFile = "";
        bool startImmediately = false;
        string customFilter = null;
        string[] customPropList = null;
        Attribute[] customAttributeList = null;
        delegate void SetIntCallback(int integer);
        delegate void LogTextCallback(string text);
        delegate void SetBoolCallback(bool boolean);
        Queue<string> goodGuids = new Queue<string>();
        Queue<string> goodDNs = new Queue<string>();
        LdapConnectionCache domainConnectionCache = null;
        string linkCheckCachedObject = "";
        string linkCheckCachedAttribute = "";
        List<string> linkCheckCachedAttributeValues = new List<string>();
        bool debugLogging = false;
        bool generateRandomFailures = false;
        string dcList = null;
	    private string logFolder = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\OABValidate\\";
	    private bool configFileError = false;

		public Form1() : this(null, null, null, null)
		{
		}

        public Form1(string gcName, string customFilter, string customPropList, string dcList)
        {
            this.InitializeComponent();
#if DEBUG
            this.Text = "OABValidate DEBUG BUILD";
#endif
            this.dcList = dcList;

            try
            {
                string debugLoggingSetting = System.Configuration.ConfigurationManager.AppSettings["DebugLogging"];
                if (debugLoggingSetting != null)
                {
                    debugLogging = bool.Parse(debugLoggingSetting);
                    if (debugLogging)
                    {
                        richTextBox1.AppendText("DebugLogging is " + debugLogging.ToString() + "\n");
                    }
                }

                string generateRandomFailuresSetting = System.Configuration.ConfigurationManager.AppSettings["GenerateRandomFailures"];
                if (generateRandomFailuresSetting != null)
                {
                    generateRandomFailures = bool.Parse(generateRandomFailuresSetting);
                    if (generateRandomFailures)
                    {
                        richTextBox1.AppendText("GenerateRandomFailures is " + generateRandomFailures.ToString() + "\n");
                    }
                }

                string logFolderFromConfigFile = System.Configuration.ConfigurationManager.AppSettings["LogFolder"];
                if (logFolderFromConfigFile != null)
                {
                    // logFolder must end in a \
                    if (!logFolderFromConfigFile.EndsWith("\\"))
                    {
                        logFolderFromConfigFile += "\\";
                    }

                    logFolder = logFolderFromConfigFile;
                }
            }
            catch (Exception exc)
            {
                this.richTextBox1.AppendText("Exception encountered reading configuration file.\n");
                this.richTextBox1.AppendText("Any settings in the file will be ignored.\n");
                this.richTextBox1.AppendText(exc.ToString() + "\n");
                configFileError = true;
            }

            if (gcName == null)
            {
                try
                {
                    DirectoryEntry entry = new DirectoryEntry("GC://RootDSE");
                    this.textBoxGC.Text = entry.Properties["dnsHostName"][0].ToString();
                }
                catch (Exception)
                {
                }
            }
            else
            {
                startImmediately = true;
                this.textBoxGC.Text = gcName;
            }

            if (customFilter != null)
            {
                this.customFilter = customFilter;
            }

            if (customPropList != null)
            {
                this.customPropList = customPropList.Split(new char[] { ',' });
                this.customAttributeList = new Attribute[this.customPropList.Length];
                for (int x = 0; x < this.customPropList.Length; x++)
                {
                    this.customAttributeList[x] = new Attribute { Name = this.customPropList[x], IsBacklink = false };
                }
            }

            if (this.textBoxGC.Text.Length > 0)
            {
                this.ReadOABs();
            }

            if (startImmediately)
            {
                this.HandleCreated += new EventHandler(Form1_HandleCreated);
            }
        }

        void Form1_HandleCreated(object sender, EventArgs e)
        {
            StartValidationThread();
        }

        private int ObjectsFound
        {
            get { return this.objectsFound; }
            set
            {
                this.objectsFound = value;
                this.textBoxObjectsFound.Text = value.ToString();
            }
        }

        private void SetObjectsFound(int newValue)
        {
            this.ObjectsFound = newValue;
        }

        private int ObjectsProcessed
        {
            get { return this.objectsProcessed; }
            set
            {
                this.objectsProcessed = value;
                this.textBoxObjectsProcessed.Text = value.ToString();
            }
        }

        private void SetObjectsProcessed(int newValue)
        {
            this.ObjectsProcessed = newValue;
        }

        private int ProblemObjects
        {
            get { return this.problemObjects; }
            set
            {
                this.problemObjects = value;
                this.textBoxProblemObjects.Text = value.ToString();
            }
        }

        private void SetProblemObjects(int newValue)
        {
            this.ProblemObjects = newValue;
        }

        private void SetButtonGoEnabled(bool enabled)
        {
            this.buttonGo.Enabled = enabled;
        }

        private void SetButtonGetOABsEnabled(bool enabled)
        {
            this.buttonGetOABs.Enabled = enabled;

            // This means we're done running. I'm going to cheat and
            // just add this here.
            if (startImmediately)
            {
                Application.Exit();
            }
        }

		private void buttonGetOABs_Click(object sender, EventArgs e)
		{
			this.ReadOABs();
		}

		private void buttonGo_Click(object sender, EventArgs e)
		{
            this.StartValidationThread();
		}

        private void StartValidationThread()
        {
            this.buttonGo.Enabled = false;
            this.buttonGetOABs.Enabled = false;
            new Thread(new ThreadStart(this.Go)).Start();
        }

		private void Go()
		{
            SetBoolCallback setButtonGoEnabledCallback = new SetBoolCallback(this.SetButtonGoEnabled);
            SetBoolCallback setButtonGetOABsEnabledCallback = new SetBoolCallback(this.SetButtonGetOABsEnabled);
            try
            {
                DateTime now = DateTime.Now;
                string dateTimeString = now.Year.ToString() + now.Month.ToString("D2") + now.Day.ToString("D2") + now.Hour.ToString("D2") + now.Minute.ToString("D2") + now.Second.ToString("D2");
                string logPath = logFolder + dateTimeString + "-" + this.textBoxGC.Text;
                if (System.IO.Directory.Exists(logPath))
                {
                    throw new Exception("Folder already exists: " + logPath + "\nPlease delete any files out of the OABValidate folder and try again.");
                }

                System.IO.Directory.CreateDirectory(logPath);
                this.logFile = logPath + "\\log.txt";
                if (debugLogging)
                {
                    DebugLog("Starting at " + dateTimeString + " using GC " + this.textBoxGC.Text);
                }

                if (configFileError)
                {
                    log("Exception encountered reading configuration file at startup.\n");
                    log("Any settings in the file will be ignored.\n");
                }

                DirectoryEntry searchRoot = new DirectoryEntry("GC://" + this.textBoxGC.Text);

                SetIntCallback setObjectsFoundCallback = new SetIntCallback(this.SetObjectsFound);
                SetIntCallback setObjectsProcessedCallback = new SetIntCallback(this.SetObjectsProcessed);
                SetIntCallback setProblemObjectsCallback = new SetIntCallback(this.SetProblemObjects);
                this.Invoke(setObjectsFoundCallback, 0);
                this.Invoke(setObjectsProcessedCallback, 0);
                this.Invoke(setProblemObjectsCallback, 0);
                int objectsFound = 0;
                int objectsProcessed = 0;
                int problemObjects = 0;
                Dictionary<string, int> problemAttributes = new Dictionary<string,int>();

                this.log("Files will be written to: " + logPath);

                domainConnectionCache = new LdapConnectionCache();
                if (dcList != null)
                {
                    foreach (string dcName in dcList.Split(new char[] {','}))
                    {
                        var dcRoot = new DirectoryEntry("LDAP://" + dcName + "/RootDSE");
                        var namingContext = dcRoot.Properties["defaultNamingContext"][0].ToString();
                        var dnsHostName = dcRoot.Properties["dnsHostName"][0].ToString();
                        this.log("Adding specified DC " + dnsHostName + " to connection cache for naming context " + namingContext + ".");
                        domainConnectionCache[namingContext] = new LdapConnection(dnsHostName);
                    }
                }

                string filter;
                if (customFilter == null)
                {
                    this.log("Validating objects in OABs: ");
                    foreach (ListViewItem oabItem in listViewChooseOAB.CheckedItems)
                    {
                        this.log("     " + oabItem.Text);
                    }

                    this.log("These OABs include the following address lists:");
                    // We're going to read the address list from each OAB and then combine them
                    // This way we won't add the filter twice if the address list appears twice
                    List<DirectoryEntry> addressListEntries = new List<DirectoryEntry>();
                    List<string> addressListDns = new List<string>();
                    foreach (ListViewItem oabItem in listViewChooseOAB.CheckedItems)
                    {
                        // Can't read offlineABContainers from the GC port, have to use LDAP
                        var oabItemEntry = new DirectoryEntry("LDAP://" + this.textBoxGC.Text + "/" + oabItem.Tag.ToString());
                        foreach (object addressListDnEntry in oabItemEntry.Properties["offlineABContainers"])
                        {
                            string addressListDn = addressListDnEntry.ToString();
                            if (!(addressListDns.Contains(addressListDn)))
                            {
                                addressListDns.Add(addressListDn);
                            }
                        }
                    }

                    foreach (string addressListDn in addressListDns)
                    {
                        DirectoryEntry addressListEntry = new DirectoryEntry("LDAP://" + this.textBoxGC.Text + "/" + addressListDn);
                        addressListEntries.Add(addressListEntry);
                    }

                    foreach (DirectoryEntry addressListEntry in addressListEntries)
                    {
                        this.log("     " + addressListEntry.Properties["cn"][0].ToString());
                    }

                    // Build the filter to query for everyone in these address lists
                    List<string> addressListFilters = new List<string>();
                    foreach (DirectoryEntry addressListEntry in addressListEntries)
                    {
                        addressListFilters.Add(addressListEntry.Properties["purportedSearch"][0].ToString());
                    }

                    StringBuilder filterBuilder = new StringBuilder();
                    filterBuilder.Append("(|");
                    foreach (DirectoryEntry addressListEntry in addressListEntries)
                    {
                        filterBuilder.Append(addressListEntry.Properties["purportedSearch"][0].ToString());
                    }

                    filterBuilder.Append(")");
                    filter = filterBuilder.ToString();
                }
                else
                {
                    filter = this.customFilter;
                }

                Attribute[] attributeListToUse;
                string[] propListToUse;
                if (this.customPropList == null)
                {
                    attributeListToUse = this.defaultAttList;
                    propListToUse = this.defaultProplist;
                }
                else
                {
                    attributeListToUse = this.customAttributeList;
                    propListToUse = this.customPropList;
                }

                this.log("");
                this.log("Filter: " + filter);
                this.log("Validating objects...");

                bool foundLingeringLinksOrObjects = false;
                bool encounteredLdapExceptions = false;

                int pageSize = 1000;
                LdapConnection connection = new LdapConnection(this.textBoxGC.Text + ":3268");
                connection.Timeout = TimeSpan.FromMinutes(5);
                connection.Bind();
                SearchRequest request = new SearchRequest("", filter, System.DirectoryServices.Protocols.SearchScope.Subtree, propListToUse);
                ExtendedDNControl edc = new System.DirectoryServices.Protocols.ExtendedDNControl(ExtendedDNFlag.StandardString);
                request.Controls.Add(edc);
                PageResultRequestControl prc = new PageResultRequestControl(pageSize);
                request.Controls.Add(prc);

                do
                {
                    SearchResponse response = connection.SendRequest(request) as SearchResponse;
                    foreach (DirectoryControl control in response.Controls)
                    {
                        if (control is PageResultResponseControl)
                        {
                            prc.Cookie = ((PageResultResponseControl)control).Cookie;
                            break;
                        }
                    }

                    objectsFound += response.Entries.Count;
                    this.Invoke(setObjectsFoundCallback, objectsFound);

                    foreach (SearchResultEntry entry in response.Entries)
                    {
                        if (debugLogging)
                        {
                            DebugLog("Evaluating object: " + entry.DistinguishedName);
                        }

                        bool thisObjectIsWrittenToImportFile = false;
                        bool foundBadAttribute = false;

                        foreach (Attribute attribute in attributeListToUse)
                        {
                            // We don't check backlinks anymore, because it shouldn't be possible to have a bad
                            // DN in a backlink.
                            if (attribute.Name != "distinguishedName" && !attribute.IsBacklink && entry.Attributes.Contains(attribute.Name))
                            {
                                string[] values = (string[])entry.Attributes[attribute.Name].GetValues(typeof(string));
                                if (values.Length == 0)
                                {
                                    foreach (string attName in entry.Attributes.AttributeNames)
                                    {
                                        if (attName.StartsWith(attribute.Name + ";", StringComparison.OrdinalIgnoreCase))
                                        {
                                            if (debugLogging)
                                            {
                                                DebugLog("Ranged attribute found: " + attName);
                                            }

                                            values = (string[])entry.Attributes[attName].GetValues(typeof(string));

                                            if (!(attName.EndsWith("*")))
                                            {
                                                List<string> allAttributeValues = new List<string>(values);
                                                allAttributeValues.AddRange(RetrieveAllAttributeValues(connection, entry.DistinguishedName.Substring(entry.DistinguishedName.LastIndexOf(';') + 1), attribute.Name, allAttributeValues.Count));
                                                values = allAttributeValues.ToArray();
                                            }

                                            break;
                                        }
                                    }
                                }

                                if (debugLogging)
                                {
                                    DebugLog(values.Length.ToString() + " values in attribute: " + attribute.Name);
                                }

                                List<string> badValues = new List<string>();
                                foreach (string guidDnString in values)
                                {
                                    ValidationResult validationResult = ValidateLink(entry.DistinguishedName.Substring(entry.DistinguishedName.LastIndexOf(';') + 1), attribute, guidDnString);
                                    if (validationResult != ValidationResult.Good)
                                    {
                                        if (!foundBadAttribute)
                                        {
                                            // This means this is the first attribute for this object that has failed
                                            // So we need to output the DN
                                            
                                            this.log(entry.DistinguishedName.Substring(entry.DistinguishedName.LastIndexOf(';') + 1));
                                            foundBadAttribute = true;
                                            problemObjects++;
                                            this.Invoke(setProblemObjectsCallback, problemObjects);
                                        }

                                        // Log it
                                        this.log("     " + attribute.Name + " " + validationResult);

                                        if (validationResult == ValidationResult.LingeringLink)
                                        {
                                            this.log("          Warning! This appears to be a lingering link and cannot be fixed by the import file.");
                                            foundLingeringLinksOrObjects = true;
                                        }
                                        else if (validationResult == ValidationResult.LingeringObject)
                                        {
                                            this.log("          Warning! This appears to be a lingering object and cannot be fixed by the import file.");
                                            foundLingeringLinksOrObjects = true;
                                        }
                                        else if (validationResult == ValidationResult.ObjectGuidMismatch)
                                        {
                                            this.log("          Warning! The DN resolves to an object with a different objectGUID! This is probably a lingering link and will not be added to the import file.");
                                            foundLingeringLinksOrObjects = true;
                                        }
                                        else if (validationResult == ValidationResult.LdapException)
                                        {
                                            this.log("          Warning! This link could not be validated due to an LdapException. A DC may be unreachable or offline.");
                                            encounteredLdapExceptions = true;
                                        }
                                        else
                                        {
                                            // Add it to the bad values for this attribute in case this is multivalued
                                            // We use this when we write the ldf file
                                            badValues.Add(guidDnString);
                                        }

                                        // Write it to the tab-delimited file
                                        StreamWriter tsvWriter = new StreamWriter(logPath + "\\ProblemAttributes.txt", true);
                                        tsvWriter.WriteLine(entry.DistinguishedName + "\t" + attribute.Name + "\t" + validationResult + "\t" + guidDnString);
                                        tsvWriter.Close();

                                        // Add it to the collection we use for statistics
                                        if (problemAttributes.ContainsKey(attribute.Name))
                                        {
                                            problemAttributes[attribute.Name]++;
                                        }
                                        else
                                        {
                                            problemAttributes.Add(attribute.Name, 1);
                                        }
                                    }
                                }

                                if (badValues.Count > 0)
                                {
                                    // We have to create a separate import file for each domain
                                    string domainNC = entry.DistinguishedName.Substring(entry.DistinguishedName.IndexOf("DC="));
                                    StreamWriter importWriter = new StreamWriter(logPath + "\\Fix-" + domainNC + ".ldf", true);

                                    if (!thisObjectIsWrittenToImportFile)
                                    {
                                        importWriter.WriteLine("dn: " + entry.DistinguishedName.Substring(entry.DistinguishedName.LastIndexOf(';') + 1));
                                        importWriter.WriteLine("changetype: modify");
                                        thisObjectIsWrittenToImportFile = true;
                                    }

                                    if (values.Length > 1)
                                    {
                                        if (badValues.Count < values.Length)
                                        {
                                            importWriter.WriteLine("replace: " + attribute.Name);
                                            foreach (string val in values)
                                            {
                                                if (!badValues.Contains(val))
                                                {
                                                    string[] valSplit = val.Split(new char[] { ';' });
                                                    string valDn = valSplit[valSplit.Length - 1];
                                                    importWriter.WriteLine(attribute.Name + ": " + valDn);
                                                }
                                            }
                                            importWriter.WriteLine("-");
                                            importWriter.Close();
                                        }
                                        else
                                        {
                                            importWriter.WriteLine("delete: " + attribute.Name);
                                            importWriter.WriteLine("-");
                                            importWriter.Close();
                                        }
                                    }
                                    else
                                    {
                                        importWriter.WriteLine("delete: " + attribute.Name);
                                        importWriter.WriteLine("-");
                                        importWriter.Close();
                                    }
                                }
                            }
                        }

                        if (thisObjectIsWrittenToImportFile)
                        {
                            string domainNC = entry.DistinguishedName.Substring(entry.DistinguishedName.IndexOf("DC="));
                            StreamWriter importWriter = new StreamWriter(logPath + "\\Fix-" + domainNC + ".ldf", true);
                            importWriter.WriteLine("");
                            importWriter.Close();
                        }

                        objectsProcessed++;
                        this.Invoke(setObjectsProcessedCallback, objectsProcessed);
                    }

                } while (prc.Cookie.Length > 0);

                StreamWriter statsWriter = new StreamWriter(logPath + "\\Statistics.txt");
                foreach (string attributeName in problemAttributes.Keys)
                {
                    statsWriter.WriteLine(attributeName + "," + problemAttributes[attributeName].ToString());
                }
                statsWriter.Close();

                this.log("");
                this.log(objectsFound.ToString() + " objects were processed.");
                this.log(problemObjects.ToString() + " objects had problems.");
                
                if (foundLingeringLinksOrObjects)
                {
                    this.log("WARNING! Lingering links were found. Please use the log or the tab-delimited file");
                    this.log("     to identify and correct them. They cannot be corrected with the import file.");
                }

                if (encounteredLdapExceptions)
                {
                    this.log("WARNING! LDAP Exceptions were encountered. This could be due to a network problem");
                    this.log("     or a domain controller being offline. As a result, some links could not be");
                    this.log("     validated.");
                }
            }
            catch (Exception exception)
            {
                this.log(exception.ToString());
                this.log("Operation aborted.");
            }
			this.log("");
			this.log("Finished.");

            domainConnectionCache.Dispose();
            goodDNs.Clear();
            goodGuids.Clear();
            linkCheckCachedObject = "";
            linkCheckCachedAttribute = "";
            linkCheckCachedAttributeValues.Clear();
            this.Invoke(setButtonGoEnabledCallback, true);
            this.Invoke(setButtonGetOABsEnabledCallback, true);
        }

        enum LinkCheckResult
        {
            NormalLink,
            LingeringLink,
            LingeringObject,
            LdapException
        }

        private LinkCheckResult CheckForLingeringLink(string objectDN, Attribute attribute, string link)
        {
            if (debugLogging)
            {
                DebugLog("CheckForLingeringLink called with:");
                DebugLog("     objectDN: " + objectDN);
                DebugLog("    attribute: " + attribute.Name);
                DebugLog("         link: " + link);
            }

            if (generateRandomFailures)
            {
                Random rand = new Random();
                int randomNumber = rand.Next(10);
                if (randomNumber < 1)
                {
                    if (debugLogging)
                    {
                        DebugLog("Generating random LingeringLink failure in CheckForLingeringLink.");
                    }

                    return LinkCheckResult.LingeringLink;
                }
                if (randomNumber < 2)
                {
                    if (debugLogging)
                    {
                        DebugLog("Generating random LingeringObject failure in CheckForLingeringLink.");
                    }

                    return LinkCheckResult.LingeringObject;
                }
            }

            if (linkCheckCachedObject == objectDN && linkCheckCachedAttribute == attribute.Name)
            {
                if (linkCheckCachedAttributeValues.Contains(link))
                {
                    if (debugLogging)
                    {
                        DebugLog("Link check was satisfied from attribute cache as NormalLink.");
                    }

                    return LinkCheckResult.NormalLink;
                }
                else
                {
                    if (debugLogging)
                    {
                        DebugLog("Link check was satisfied from attribute cache as LingeringLink.");
                    }

                    return LinkCheckResult.LingeringLink;
                }
            }

            LinkCheckResult returnValue;
            LdapConnection domainConnection = domainConnectionCache[objectDN];

            linkCheckCachedAttributeValues.Clear();
            linkCheckCachedObject = "";
            linkCheckCachedAttribute = "";

            SearchRequest request = new SearchRequest(objectDN, "(objectClass=*)", System.DirectoryServices.Protocols.SearchScope.Base, attribute.Name);
            ExtendedDNControl edc = new ExtendedDNControl(ExtendedDNFlag.StandardString);
            request.Controls.Add(edc);
            SearchResponse response = null;

            try
            {
                response = domainConnection.SendRequest(request) as SearchResponse;
                if (response == null || response.Entries == null)
                {
                    if (debugLogging)
                    {
                        if (response == null)
                        {
                            DebugLog("response was null.");
                        }
                        else
                        {
                            DebugLog("Entries was null. Response ResultCode is: " + response.ResultCode);
                        }
                    }

                    returnValue = LinkCheckResult.LingeringObject;
                }
                else if (response.Entries.Count < 1)
                {
                    if (debugLogging)
                    {
                        DebugLog("There were 0 Entries. Response ResultCode is: " + response.ResultCode);
                    }

                    returnValue = LinkCheckResult.LingeringObject;
                }
                else
                {
                    SearchResultEntry entry = response.Entries[0];

                    if (debugLogging)
                    {
                        DebugLog("Response contained " + entry.Attributes.Count.ToString() + " attributes.");
                    }

                    string[] values = (string[])entry.Attributes[attribute.Name].GetValues(typeof(string));
                    if (values.Length == 0)
                    {
                        foreach (string attName in entry.Attributes.AttributeNames)
                        {
                            if (attName.StartsWith(attribute.Name + ";", StringComparison.OrdinalIgnoreCase))
                            {
                                if (debugLogging)
                                {
                                    DebugLog("Ranged attribute found: " + attName);
                                }

                                values = (string[])entry.Attributes[attName].GetValues(typeof(string));

                                if (!(attName.EndsWith("*")))
                                {
                                    List<string> allAttributeValues = new List<string>(values);
                                    allAttributeValues.AddRange(RetrieveAllAttributeValues(domainConnection, objectDN, attribute.Name, allAttributeValues.Count));
                                    values = allAttributeValues.ToArray();
                                }

                                break;
                            }
                        }
                    }

                    linkCheckCachedAttributeValues.AddRange(values);
                }

                bool linkIsNormal = false;
                if (linkCheckCachedAttributeValues.Count > 0)
                {
                    linkCheckCachedObject = objectDN;
                    linkCheckCachedAttribute = attribute.Name;

                    if (debugLogging)
                    {
                        DebugLog("Populated attribute cache with " + linkCheckCachedAttributeValues.Count.ToString() + " values for " + attribute.Name + " attribute.");
                    }

                    if (linkCheckCachedAttributeValues.Contains(link))
                    {
                        linkIsNormal = true;
                    }
                }

                if (linkIsNormal)
                    returnValue = LinkCheckResult.NormalLink;
                else
                    returnValue = LinkCheckResult.LingeringLink;
            }
            catch (LdapException)
            {
                // We couldn't contact a writable DC for the domain for some reason.
                returnValue = LinkCheckResult.LdapException;
            }
            catch (Exception exc)
            {
                if (debugLogging)
                {
                    DebugLog("CheckForLingeringLink encountered exception at domainConnection.SendRequest():");
                    DebugLog(exc.ToString());
                }

                // We interpret any error here as a lingering object. However, this could happen
                // for other reasons.
                // This shouldn't be a problem since the user is required to manually correct
                // lingering objects anyway - he or she will undoubtedly identify the real
                // problem when the lingering object error is manually investigated.
                returnValue = LinkCheckResult.LingeringObject;
            }

            if (debugLogging)
            {
                DebugLog("Returning from CheckForLingeringLink with result: " + returnValue);
            }

            return returnValue;
        }

        private List<string> RetrieveAllAttributeValues(LdapConnection connection, string objectDN, string attribute, int startingOffset)
        {
            if (debugLogging)
            {
                LdapDirectoryIdentifier ldapDirId = connection.Directory as LdapDirectoryIdentifier;
                DebugLog("Performing ranged retrieval of all values for: ");
                DebugLog("     object: " + objectDN);
                DebugLog("     attribute: " + attribute);
                DebugLog("     server: " + ldapDirId.Servers[0]);
                DebugLog("     this operation is starting at: " + DateTime.Now.ToLongTimeString());
            }

            int lowRange = startingOffset;
            List<string> rangedValues = new List<string>();
            SearchRequest rangedAttributeRequest = new SearchRequest(objectDN, "(objectClass=*)", System.DirectoryServices.Protocols.SearchScope.Base);
            ExtendedDNControl edc = new System.DirectoryServices.Protocols.ExtendedDNControl(ExtendedDNFlag.StandardString);
            rangedAttributeRequest.Controls.Add(edc);
            string desiredAttribute = "";
            do
            {
                string rangeString = attribute + ";Range=" + lowRange.ToString() + "-*";
                rangedAttributeRequest.Attributes.Clear();
                rangedAttributeRequest.Attributes.Add(rangeString);
                SearchResponse response = connection.SendRequest(rangedAttributeRequest) as SearchResponse;
                if (response.Entries.Count != 1)
                {
                    if (debugLogging)
                    {
                        DebugLog("Error in ranged retrieval. Response ResultCode is: " + response.ResultCode);
                    }

                    break;
                }
                else
                {
                    desiredAttribute = "";
                    SearchResultEntry entry = response.Entries[0];
                    foreach (string attName in entry.Attributes.AttributeNames)
                    {
                        if (attName.StartsWith(attribute + ";", StringComparison.OrdinalIgnoreCase))
                        {
                            desiredAttribute = attName;
                        }
                    }

                    if (desiredAttribute != "")
                    {
                        if (debugLogging)
                        {
                            DebugLog("Received ranged attribute response: " + desiredAttribute);
                        }

                        string[] values = (string[])entry.Attributes[desiredAttribute].GetValues(typeof(string));
                        if (values.Length > 0)
                        {
                            rangedValues.AddRange(values);
                        }
                        else
                        {
                            break;
                        }
                    }

                    lowRange = startingOffset + rangedValues.Count;
                }
            }
            while (!desiredAttribute.EndsWith("*"));

            if (debugLogging)
            {
                DebugLog("Ranged retrieval done.");
                DebugLog("     this operation finished at: " + DateTime.Now.ToLongTimeString());
            }

            return rangedValues;
        }

        private enum ValidationResult
        {
            Good,
            NullValue,
            UnresolvableDN,
            ObjectGuidMismatch,
            LingeringLink,
            LingeringObject,
            LdapException,
            SimulatedDebugFailure
        }

        private ValidationResult ValidateLink(string objectDN, Attribute attribute, string link)
        {
            // Because of the ExtendedDNControl, all DN values will come back in the format:
            // <GUID=nnnn>;<SID=nnnn>;DN=whatever
            // SID is only present for security principals.
            // See http://msdn.microsoft.com/en-us/library/aa366980(VS.85).aspx
            // This is necessary so we can catch lingering links where an object with the
            // same DN has been created, but has a different GUID than what is present in
            // the DN-valued attribute.
            string[] guidDnStringSplit = link.Split(new char[] { ';' });
            string guidString;
            string dnString;
            if (guidDnStringSplit.Length > 1)
            {
                guidString = guidDnStringSplit[0].Substring(6).TrimEnd(new char[] { '>' });
                dnString = guidDnStringSplit[guidDnStringSplit.Length - 1];
            }
            else
            {
                guidString = "";
                dnString = link;
            }

            if (debugLogging)
            {
                DebugLog("ValidateLink called for link: " + link);
                DebugLog("     Guid: " + guidString);
                DebugLog("       Dn: " + dnString);

                if (generateRandomFailures)
                {
                    Random rand = new Random();
                    int randomNumber = rand.Next(100);
                    if ((randomNumber % 20) == 0)
                    {
                        if (debugLogging)
                        {
                            DebugLog("Generating random failure in ValidateLink.");
                        }

                        return ValidationResult.SimulatedDebugFailure;
                    }
                }
            }

            if (guidString != "" && goodGuids.Contains(guidString))
            {
                return ValidationResult.Good;
            }
            else if (guidString == "" && goodDNs.Contains(dnString))
            {
                return ValidationResult.Good;
            }

            bool isValidDn = false;
            bool guidMatches = false;
            try
            {
                if (dnString != "")
                {
                    string escapedDn = dnString.Replace("/", @"\/");
                    DirectoryEntry testEntry = new DirectoryEntry("GC://" + this.textBoxGC.Text + "/" + escapedDn);
                    string testDn = testEntry.Properties["distinguishedName"][0].ToString();
                    isValidDn = true;
                    // The DN resolved, but is this really the same object? Need to check the GUID
                    if (guidString != "")
                    {
                        Guid resolvedObjectGuid = new Guid((byte[])testEntry.Properties["objectGUID"][0]);

                        if (debugLogging)
                        {
                            DebugLog("Comparing guid: " + guidString + " to resolvedObjectGuid: " + resolvedObjectGuid.ToString());
                        }

                        if (guidString == resolvedObjectGuid.ToString())
                        {
                            guidMatches = true;
                        }
                    }
                    else
                    {
                        guidMatches = true;
                    }
                }
            }
            catch
            {
            }

            ValidationResult validationResult;
            if (isValidDn)
            {
                if (guidMatches)
                {
                    // Now we need to determine if this is a lingering link or not. A lingering link
                    // is when we have a GUID+DN pair in a linked attribute when looking at this object on
                    // the GC, but that same value is not present when looking at it on a DC
                    // (i.e. a writable copy of the domain). A lingering link cannot be fixed by
                    // ldifde import and needs to be flagged in the log and the tab-delimited file.
                    LinkCheckResult linkType = CheckForLingeringLink(objectDN, attribute, link);
                    if (linkType == LinkCheckResult.NormalLink)
                    {
                        validationResult = ValidationResult.Good;
                    }
                    else if (linkType == LinkCheckResult.LingeringLink)
                    {
                        validationResult = ValidationResult.LingeringLink;
                    }
                    else if (linkType == LinkCheckResult.LingeringObject)
                    {
                        validationResult = ValidationResult.LingeringObject;
                    }
                    else if (linkType == LinkCheckResult.LdapException)
                    {
                        validationResult = ValidationResult.LdapException;
                    }
                    else
                        throw new Exception("Unhandled link type");
                }
                else
                {
                    validationResult = ValidationResult.ObjectGuidMismatch;
                }
            }
            else
            {
                if (dnString == "")
                {
                    validationResult = ValidationResult.NullValue;
                }
                else
                {
                    validationResult = ValidationResult.UnresolvableDN;
                }
            }

            if (validationResult == ValidationResult.Good)
            {
                if (guidString != "")
                {
                    goodGuids.Enqueue(guidString);
                    if (goodGuids.Count > 10000)
                    {
                        goodGuids.Dequeue();
                    }
                }
                else
                {
                    goodDNs.Enqueue(dnString);
                    if (goodDNs.Count > 10000)
                    {
                        goodDNs.Dequeue();
                    }
                }
            }

            return validationResult;
        }

		private void log(string logstring)
		{
            LogTextCallback d = new LogTextCallback(this.richTextBox1.AppendText);
            this.Invoke(d, (logstring + "\n"));

            StreamWriter writer = new StreamWriter(this.logFile, true);
            writer.WriteLine(logstring);
            writer.Close();
		}

		private void ReadOABs()
		{
			this.listViewChooseOAB.Items.Clear();

            try
            {
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + this.textBoxGC.Text + "/RootDSE");
                DirectoryEntry searchRoot = new DirectoryEntry("LDAP://CN=Microsoft Exchange,CN=Services," + entry.Properties["configurationNamingContext"][0].ToString());
                DirectorySearcher searcher = new DirectorySearcher(searchRoot, "(objectClass=msExchOAB)", new string[] { "cn", "distinguishedName", "offlineABContainers" }, System.DirectoryServices.SearchScope.Subtree);
                foreach (SearchResult result in searcher.FindAll())
                {
                    ListViewItem item = new ListViewItem(result.Properties["cn"][0].ToString());
                    item.Tag = result.Properties["distinguishedName"][0].ToString();
                    item.Checked = true;
                    this.listViewChooseOAB.Items.Add(item);
                }

                if (this.listViewChooseOAB.Items.Count > 0)
                {
                    this.listViewChooseOAB.Enabled = true;
                }
            }
            catch
            {
            }
		}

        private void DebugLog(string logstring)
        {
            StreamWriter writer = new StreamWriter(this.logFolder + "OabValidateDebug.log", true);
            writer.WriteLine(logstring);
            writer.Close();
        }
	}
}
