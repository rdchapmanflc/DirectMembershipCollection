using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.ConfigurationManagement.ManagementProvider;
using Microsoft.ConfigurationManagement.ManagementProvider.WqlQueryEngine;
using System.Security;
using System.Runtime.InteropServices;

namespace DirectMembershipCollection
{
    class Program
    {
        static void Main(string[] args)
        {
            //command line arguments
            //server FQDN args[0]
            //username args[1]
            //input file args[2]
            //interger for input file hostname field args[3]
            //string for limit to collection ID args[4]
            //string for new collection name and comment--will be the same args[5]

            //read in list of hostnames
            string fileName = args[2];
            int hostnameFieldNumber = int.Parse(args[3]);

            string record = String.Empty;
            List<string> computers = computers = readInputFile(fileName,
                                                               hostnameFieldNumber);

            //get the user password and setup connection to SMS provider
            SCCMConnection myConnection = new SCCMConnection();
            string serverName = args[0];
            string userName = args[1];

            string limitCollectionId = args[4];
            string newCollectionName = args[5];

            SecureString password = readPassword();
            String plainPassword = secureStringToPlainString(password);
            WqlConnectionManager aConnection = null;
            
            //the connection to our SMS Provider using the SCCMConnection class below
            try
            {
                aConnection = myConnection.Connect(serverName, userName, plainPassword);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            try
            {
                CreateSCCMCollection createSCCMCollection = new CreateSCCMCollection(aConnection,
                                                                 newCollectionName,
                                                                 newCollectionName,
                                                                 true,
                                                                 limitCollectionId,
                                                                 computers);
                createSCCMCollection.createStaticCollection();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception at createSCCMCollection " + ex.ToString());
            }

            //clean up passwords
            password.Dispose();
            plainPassword = null;

            Console.WriteLine("Press enter to exit...");
            Console.ReadLine();
        }

        //helper to read input file and return List<string> of computers
        private static List<string> readInputFile(string filePath,
                                                  int fieldNumber)
        {
            string record = String.Empty;
            List<string> computers = new List<string>();
            using (StreamReader streamReader = new StreamReader(filePath))
            {
                while (!streamReader.EndOfStream)
                {
                    record = streamReader.ReadLine();
                    string[] fields = record.Split(',');
                    //for data copied out of deployment monitoring in SCCM Console
                    //computers.Add(fields[7]);

                    //data from Script output in SCCM
                    computers.Add(fields[fieldNumber]);
                }
            }

            //foreach (string c in computers)
            //{
            //    Console.WriteLine(c);
            //}

            //Console.ReadLine();
            return computers;
        }

        //helpers for password handling, from my SCCM-SDK-Demo-Project
        private static SecureString readPassword()
        {
            //new secure string
            SecureString securePassword = new SecureString();

            ConsoleKeyInfo key;

            Console.Write("Please enter your password: ");
            do
            {
                key = Console.ReadKey(true);

                //Ignore out of range characters
                if ((int)key.Key >= 32 && (int)key.Key <= 122)
                {
                    //append the character returned by key.Key to the password
                    securePassword.AppendChar(key.KeyChar);
                    Console.Write("*");

                }

            } while (key.Key != ConsoleKey.Enter);
            Console.WriteLine();

            return securePassword;
        }

        //Marshal secure string to plain string
        //https://stackoverflow.com/questions/818704/how-to-convert-securestring-to-system-string
        private static String secureStringToPlainString(SecureString value)
        {
            IntPtr valuePtr = IntPtr.Zero;
            try
            {
                valuePtr = Marshal.SecureStringToGlobalAllocUnicode(value);
                return Marshal.PtrToStringUni(valuePtr);
            }
            finally
            {
                Marshal.ZeroFreeGlobalAllocUnicode(valuePtr);
            }
        }

    }
    
    //internal class to create the collection
    //keeping it quick and dirty as a prototype
    internal class CreateSCCMCollection
    {
        //from
        //https://docs.microsoft.com/en-us/previous-versions/system-center/developer/jj885570(v=cmsdk.12)

        private WqlConnectionManager _connection { get; set; }
        private string collectionName;
        private string collectionComment;
        private bool ownedByThisSite;
        private string limitToCollectionId;
        private List<String> computers;

        //private WqlConnectionManager wqlConnection { get { return _wqlConnection; } set { _wqlConnection = value; } }

        public CreateSCCMCollection(WqlConnectionManager pWQLConnection,
                            string pCollectionName,
                            string pCollectionComment,
                            bool pOwnedByThisSite,
                            string pLimitToCollectionId,
                            List<string>pComputers)
        {
            _connection = pWQLConnection;
            this.collectionName = pCollectionName;
            this.collectionComment = pCollectionComment;
            this.ownedByThisSite = pOwnedByThisSite;
            this.limitToCollectionId = pLimitToCollectionId;
            this.computers = pComputers;

        }

        //static collection (Direct membership)
        public void createStaticCollection()
        {

            IResultObject newCollection = null;

            //try and create the collection and setup base required properties
            try
            {
                newCollection = _connection.CreateInstance("SMS_Collection");
                newCollection["Name"].StringValue = this.collectionName;
                newCollection["Comment"].StringValue = this.collectionComment;
                newCollection["OwnedByThisSite"].BooleanValue = ownedByThisSite;
                newCollection["LimitToCollectionID"].StringValue = limitToCollectionId;

                //push to SMS Provided using WMI Put and then retrive with Get
                newCollection.Put();
                newCollection.Get();
            }
            catch (SmsException ex)
            {
                //Handle error
                Console.WriteLine(ex.ToString());
            }

            //setup membership rules
            //loop over computer list adding membership rule for each item
            foreach (string computer in computers)
            {
                //direct membership WMI class
                //connection to SMS Provider
                IResultObject directMembershipRule = _connection.CreateInstance("SMS_CollectionRuleDirect");
                int resourceId = 0;

                //need to locate the client ResourceId based on Resource name
                string resourceQuery = String.Format("select ResourceId from SMS_R_System where Name = '{0}'", computer.ToString());
                IResultObject theResource = _connection.QueryProcessor.ExecuteQuery(resourceQuery);
                //WqlResultObject queryResults = null;
                if (theResource != null)
                {
                    foreach (WqlResultObject r in theResource)
                    {
                        resourceId = int.Parse(r.PropertyList["ResourceId"]);
                    }
                }

                try
                {
                    //directMembershipRule["ResourceClassName"].StringValue = computer.ToString();
                    directMembershipRule["ResourceClassName"].StringValue = "SMS_R_System";

                    directMembershipRule["ResourceID"].IntegerValue = resourceId;

                    //add the rule
                    Dictionary<string, object> addMembershipRuleParameters = new Dictionary<string, object>();
                    addMembershipRuleParameters.Add("collectionRule", directMembershipRule);
                    IResultObject staticID = newCollection.ExecuteMethod("AddMembershipRule", addMembershipRuleParameters);

                    Console.WriteLine($"Resource {computer} with Id {resourceId.ToString()} added to collection");
                }
                catch (SmsException ex)
                {
                    Console.WriteLine(ex.ToString());
                    throw;
                }

            }

            try
            {
                ////start the collection evaluator
                Dictionary<string, object> requestRefreshParameters = new Dictionary<string, object>();
                requestRefreshParameters.Add("IncludeSubCollections", false);
                newCollection.ExecuteMethod("RequestRefresh", requestRefreshParameters);
            }
            catch (SmsException ex)
            {
                Console.WriteLine("Unable to start collection refresh " + ex.ToString());
                throw;
            }


        }
    }


    //To keep this prototype quick and dirty using internal class
    //code from my SCCM-SDK-Demo Project
    internal class SCCMConnection
    {
        public WqlConnectionManager Connect(string pServerName,
                                            string pUserName,
                                            string pPassword)
        {
            try
            {
                SmsNamedValuesDictionary namedValues = new SmsNamedValuesDictionary();
                WqlConnectionManager connection = new WqlConnectionManager(namedValues);
                if (System.Net.Dns.GetHostName().ToUpper() == pServerName.ToUpper())
                {
                    connection.Connect(pServerName);
                }
                else
                {
                    connection.Connect(pServerName, pUserName, pPassword);
                }
                return connection;

            }
            catch (SmsException ex)
            {
                Console.WriteLine("Failed to connect. Error: " + ex.Message);
                return null;

            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine("Failed to authenticate. Error: " + ex.Message);
                throw;
            }

        }
    }

}
