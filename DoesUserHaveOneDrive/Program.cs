using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace DoesUserHaveOneDrive
{
    class Program
    {
        private static String AdminUrl { get; set; }
        private static String TennantAdmin { get; set; }
        private static SecureString Password { get; set; }
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine(@"Usage: OneDrive Checker <admin url> <tennant admin>");
                Environment.Exit(-1);
            }

            AdminUrl = args[0];
            TennantAdmin = args[1];
            Password = GetPassword();

            Stopwatch sw = new Stopwatch();
            sw.Start();
            try
            {
                Console.WriteLine("User to check:");
                var userTocheck = Console.ReadLine();
                string startMsg = "Checking OneDrive - please wait...";
                Console.WriteLine(startMsg);
                Console.SetCursorPosition(startMsg.Length + 2, Console.CursorTop - 1);
                ProcessOneDriveUsers(userTocheck);
                sw.Stop();
            }
            finally
            {

                Console.WriteLine();
                Console.WriteLine("OneDrive Check completed: {0:g}", sw.Elapsed);
                Console.ReadKey();
            }

        }

        private static void ProcessOneDriveUsers(string userTocheck)
        {
            string personalSpace = DoesUserAlreadyHaveAOneDrive(getContext(), userTocheck);

            if (personalSpace.Contains("Person.aspx?accountname"))
            {
                Console.WriteLine("No OneDrive found for this user");
            }
            else
            {
                Console.WriteLine("Users OneDrive Url: " + personalSpace);
            }
        }

        private static string DoesUserAlreadyHaveAOneDrive(ClientContext cc, string email)
        {
            var user = cc.Web.EnsureUser(email.Replace("\"", ""));
            cc.Load(user);
            cc.ExecuteQuery();
            PeopleManager pmanger = new PeopleManager(cc);
            var properties = pmanger.GetPropertiesFor(user.LoginName);
            cc.Load(properties, ondrive => ondrive.PersonalUrl);
            cc.ExecuteQuery();

            return properties.PersonalUrl;
        }

        private static ClientContext getContext()
        {
            ClientContext cc = new ClientContext(AdminUrl);
            SecureString secPass = Password;
            cc.Credentials = new SharePointOnlineCredentials(TennantAdmin, secPass);
            return cc;

        }

        private static SecureString GetPassword()
        {
            SecureString sStrPwd = new SecureString();
            try
            {
                Console.Write("Password : ");

                for (ConsoleKeyInfo keyInfo = Console.ReadKey(true); keyInfo.Key != ConsoleKey.Enter; keyInfo = Console.ReadKey(true))
                {
                    if (keyInfo.Key == ConsoleKey.Backspace)
                    {
                        if (sStrPwd.Length > 0)
                        {
                            sStrPwd.RemoveAt(sStrPwd.Length - 1);
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                            Console.Write(" ");
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        }
                    }
                    else if (keyInfo.Key != ConsoleKey.Enter)
                    {
                        Console.Write("*");
                        sStrPwd.AppendChar(keyInfo.KeyChar);
                    }

                }
                Console.WriteLine("");
            }
            catch (Exception e)
            {
                sStrPwd = null;
                Console.WriteLine(e.Message);
            }

            return sStrPwd;
        }
    }
}
