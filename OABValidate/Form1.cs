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

		public Form1() : this(null, null, null)
		{
		}

        public Form1(string gcName, string customFilter, string customPropList)
        {
            this.InitializeComponent();
#if DEBUG
            this.Text = "OABValidate DEBUG BUILD";
#endif
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
                string logPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\OABValidate\\" + dateTimeString + "-" + this.textBoxGC.Text;
                if (System.IO.Directory.Exists(logPath))
                {
                    throw new Exception("Folder already exists: " + logPath + "\nPlease delete any files out of the OABValidate folder and try again.");
                }

                System.IO.Directory.CreateDirectory(logPath);
                this.logFile = logPath + "\\log.txt";
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
                Queue<string> goodDNs = new Queue<string>();
#if DEBUG
                this.log("WARNING! This is a DEBUG build and will generate a lot of FALSE failures for testing purposes. Importing the LDIF file will mess up perfectly good objects!");
#endif
                this.log("Files will be written to: " + logPath);

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

                PageResultRequestControl prc = null;

                bool foundLingeringLinksOrObjects = false;

                int pageSize = 1000;
                LdapConnection connection = new LdapConnection(this.textBoxGC.Text + ":3268");
                connection.Timeout = TimeSpan.FromMinutes(5);
                connection.Bind();
                SearchRequest request = new SearchRequest("", filter, System.DirectoryServices.Protocols.SearchScope.Subtree, propListToUse);
                prc = new PageResultRequestControl(pageSize);
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
                        bool thisObjectIsWrittenToImportFile = false;
                        bool foundBadAttribute = false;

                        foreach (Attribute attribute in attributeListToUse)
                        {
                            // We don't check backlinks anymore, because it shouldn't be possible to have a bad
                            // DN in a backlink.
                            if (attribute.Name != "distinguishedName" && !attribute.IsBacklink && entry.Attributes.Contains(attribute.Name))
                            {
                                string[] values = (string[])entry.Attributes[attribute.Name].GetValues(typeof(string));
                                List<string> badValues = new List<string>();
                                foreach (string dn in values)
                                {
                                    string validationResult;
                                    if (goodDNs.Contains(dn))
                                    {
                                        validationResult = "";
                                    }
                                    else
                                    {
                                        validationResult = ValidateDn(dn);
                                        if (validationResult == "")
                                        {
                                            goodDNs.Enqueue(dn);
                                            if (goodDNs.Count > 10000)
                                            {
                                                goodDNs.Dequeue();
                                            }
                                        }
                                    }

                                    if (validationResult != "")
                                    {
                                        // The DN did not resolve
                                        if (!foundBadAttribute)
                                        {
                                            // This means this is the first attribute for this object that has failed
                                            // So we need to output the DN
                                            
                                            this.log(entry.DistinguishedName);
                                            foundBadAttribute = true;
                                            problemObjects++;
                                            this.Invoke(setProblemObjectsCallback, problemObjects);
                                        }

                                        // Log it
                                        this.log("     " + attribute.Name + " " + validationResult);

                                        // Now we need to determine if this is a lingering link or not. A lingering link
                                        // is when we have a bad DN in a linked attribute when looking at this object on
                                        // the GC, but that same bad value is not present when looking at it on a DC
                                        // (i.e. a writable copy of the domain). A lingering link cannot be fixed by
                                        // ldifde import and needs to be flagged in the log and the tab-delimited file.
                                        LinkCheckResult linkCheckResult = CheckForLingeringLink(entry.DistinguishedName, attribute, dn);
                                        if (linkCheckResult == LinkCheckResult.LingeringLink)
                                        {
                                            this.log("          Warning! This appears to be a lingering link and cannot be fixed by the import file.");
                                        }
                                        else if (linkCheckResult == LinkCheckResult.LingeringObject)
                                        {
                                            this.log("          Warning! This appears to be a lingering object and cannot be fixed by the import file.");
                                        }

                                        if (linkCheckResult == LinkCheckResult.NormalLink)
                                        {
                                            // Add it to the bad values for this attribute in case this is multivalued
                                            // We use this when we write the ldf file
                                            badValues.Add(dn);
                                        }
                                        else
                                        {
                                            foundLingeringLinksOrObjects = true;
                                        }

                                        // Write it to the tab-delimited file
                                        StreamWriter tsvWriter = new StreamWriter(logPath + "\\ProblemAttributes.txt", true);
                                        tsvWriter.WriteLine(entry.DistinguishedName + "\t" + attribute.Name + "\t" + linkCheckResult.ToString() + "\t" + dn);
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
                                        importWriter.WriteLine("dn: " + entry.DistinguishedName);
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
                                                    importWriter.WriteLine(attribute.Name + ": " + val);
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
                this.log(problemObjects.ToString() + " objects had unresolvable DNs.");
                
                if (foundLingeringLinksOrObjects)
                {
                    this.log("WARNING! Lingering links were found. Please use the log or the CSV to");
                    this.log("     identify and correct them. They cannot be corrected with the import file.");
                }
            }
            catch (Exception exception)
            {
                this.log("Exception: " + exception.Message);
                this.log("Stack: " + exception.StackTrace);
                this.log(exception.ToString());
                this.log("Operation aborted.");
            }
			this.log("");
			this.log("Finished.");
            foreach (string domain in domainConnections.Keys)
            {
                domainConnections[domain].Dispose();
            }
            domainConnections.Clear();
            this.Invoke(setButtonGoEnabledCallback, true);
            this.Invoke(setButtonGetOABsEnabledCallback, true);
        }

        enum LinkCheckResult
        {
            NormalLink,
            LingeringLink,
            LingeringObject
        }

        Dictionary<string, LdapConnection> domainConnections = new Dictionary<string, LdapConnection>();

        private LinkCheckResult CheckForLingeringLink(string objectDN, Attribute attribute, string link)
        {
#if DEBUG
            Random rand = new Random();
            int randomNumber = rand.Next(10);
            if (randomNumber < 1)
            {
                return LinkCheckResult.LingeringLink;
            }
            if (randomNumber < 2)
            {
                return LinkCheckResult.LingeringObject;
            }
#endif
            LinkCheckResult returnValue;
            string domainNC = objectDN.Substring(objectDN.IndexOf("DC="));
            string domain = domainNC.Replace(",DC=", ".").Substring(3);
            LdapConnection domainConnection;
            if (domainConnections.ContainsKey(domain))
            {
                domainConnection = domainConnections[domain];
            }
            else
            {
                domainConnection = new LdapConnection(domain + ":389");
                domainConnection.Timeout = TimeSpan.FromMinutes(5);
                domainConnection.SessionOptions.AutoReconnect = true;
                domainConnection.AutoBind = true;
                domainConnections.Add(domain, domainConnection);
            }

            SearchRequest request = new SearchRequest(objectDN, "(objectClass=*)", System.DirectoryServices.Protocols.SearchScope.Base, attribute.Name);
            SearchResponse response = null;
            try
            {
                response = domainConnection.SendRequest(request) as SearchResponse;
            }
            catch (Exception)
            {
                // We interpret any error here as a lingering object. However, this could happen
                // for other reasons, such as not being able to contact a DC for that domain.
                // This shouldn't be a problem since the user is required to manually correct
                // lingering objects anyway - he or she will undoubtedly identify the real
                // problem when the lingering object error is manually investigated.
                returnValue = LinkCheckResult.LingeringObject;
            }

            if (response == null || response.Entries == null)
            {
                returnValue = LinkCheckResult.LingeringObject;
            }
            else if (response.Entries.Count < 1)
            {
                returnValue = LinkCheckResult.LingeringObject;
            }
            else
            {
                bool linkIsNormal = false;
                SearchResultEntry entry = response.Entries[0];
                if (entry.Attributes.Contains(attribute.Name))
                {
                    string[] values = (string[])entry.Attributes[attribute.Name].GetValues(typeof(string));
                    foreach (string val in values)
                    {
                        if (val == link)
                        {
                            linkIsNormal = true;
                            break;
                        }
                    }
                }

                if (linkIsNormal)
                    returnValue = LinkCheckResult.NormalLink;
                else
                    returnValue = LinkCheckResult.LingeringLink;
            }

            return returnValue;
        }

        private string ValidateDn(string dn)
        {
#if DEBUG
            Random rand = new Random();
            int randomNumber = rand.Next(100);
            if ((randomNumber % 2) == 0)
            {
                return "Simulated failure for debug.";
            }
#endif
            bool isValidDn = false;
            try
            {
                if (dn != "")
                {
                    string escapedDn = dn.Replace("/", @"\/");
                    DirectoryEntry testEntry = new DirectoryEntry("GC://" + this.textBoxGC.Text + "/" + escapedDn);
                    string testDn = testEntry.Properties["distinguishedName"][0].ToString();
                    isValidDn = true;
                }
            }
            catch
            {
            }

            if (!isValidDn)
            {
                if (dn == "")
                {
                    return ("contains null value");
                }
                else
                {
                    return ("contains unresolvable DN: " + dn);
                }
            }
            else
            {
                return "";
            }
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
	}
}
