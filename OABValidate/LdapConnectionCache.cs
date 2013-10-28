using System;
using System.Collections.Generic;
using System.DirectoryServices.Protocols;
using System.Text;

namespace OABValidate
{
    public class LdapConnectionCache : IDisposable
    {
        private Dictionary<string, LdapConnection> domainConnections = new Dictionary<string, LdapConnection>();

        public LdapConnection this[string domainOrObjectDN]
        {
            get
            {
                string domainNC = domainOrObjectDN.Substring(domainOrObjectDN.IndexOf("DC="));
                string domain = domainNC.Replace(",DC=", ".").Substring(3);

                LdapConnection domainConnection;
                if (domainConnections.ContainsKey(domain))
                {
                    domainConnection = domainConnections[domain];
                }
                else
                {
                    domainConnection = new LdapConnection(domain + ":3268");
                    domainConnection.Timeout = TimeSpan.FromMinutes(5);
                    domainConnection.SessionOptions.AutoReconnect = true;
                    domainConnection.AutoBind = true;
                    domainConnections.Add(domain, domainConnection);
                }

                return domainConnection;
            }

            set
            {
                string domainNC = domainOrObjectDN.Substring(domainOrObjectDN.IndexOf("DC="));
                string domain = domainNC.Replace(",DC=", ".").Substring(3);

                if (domainConnections.ContainsKey(domain))
                {
                    domainConnections[domain] = value;
                }
                else
                {
                    domainConnections.Add(domain, value);
                }
            }
        }

        public void Dispose()
        {
            foreach (string domain in domainConnections.Keys)
            {
                domainConnections[domain].Dispose();
            }

            domainConnections.Clear();
        }
    }
}
