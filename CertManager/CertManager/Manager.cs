
using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace CertManager
{
    class Manager
    {
        private const string FileExtension = ".pdf";

        private Dictionary<string, string> appSettings;
        private readonly string uniqueKey;
        private string emailBodyTemplate;
        public Manager()
        {
            appSettings = GetConfigurationUsingSection("appSettings");
            uniqueKey = appSettings["KeyColumn"];
        }

        private HashSet<string> getExcludeUsers()
        {
            string filePath = appSettings["UserListExclude"];

            HashSet<string> excludeUsers = new HashSet<string>();

            foreach (String line in File.ReadAllLines(filePath))
            {
                string[] cells = line.Split(',');
                if (cells.Length != 0 && !string.IsNullOrWhiteSpace(cells[0]) && !excludeUsers.Contains(cells[0].Trim()))
                { excludeUsers.Add(cells[0].Trim()); }
            }

            return excludeUsers;
        }

        public void Execute()
        {
            CSVTable allUsers = new CSVTable(File.ReadAllText(appSettings["UserList"]), appSettings["Delimeter"].ToCharArray()[0], Convert.ToBoolean(appSettings["AllValuesRequired"]));

            CSVTable tempUsers = allUsers;

            var excludeUsers = getExcludeUsers();

            List<CSVRecord> removeUsers = new List<CSVRecord>();


            //foreach (var user in allUsers.Records)
            //{
            //    if (excludeUsers.Contains(user[uniqueKey])) {
            //        removeUsers.Add(user);
            //    }
            //}

            int removedCount= allUsers.Records.RemoveAll(u => excludeUsers.Contains(u[uniqueKey]));

            Console.WriteLine("Removed users:"+removedCount);
            if (Convert.ToBoolean(appSettings["GenerateCertificate"]))
            {
                Console.WriteLine("Generate Certificates");
                GenerateCertificates(allUsers);
            }

            if (Convert.ToBoolean(appSettings["SendEmail"]))
            {
                Console.WriteLine("Sending Emails");
                emailBodyTemplate = File.ReadAllText(appSettings["EmailBodyFileName"]);
                SendEmail(allUsers);
            }
        }

        private void GenerateCertificates(CSVTable users)
        {

            //create a Presentation instance and load the template PowerPoint file
            Presentation presentation = new Presentation();
            string filePath = appSettings["Template"];

            string outputFolder = appSettings["OutputFolder"];

            foreach (CSVRecord user in users.Records)
            {
                Console.WriteLine($"Loading file:{filePath}");

                presentation.LoadFromFile(filePath);
                Console.WriteLine("File loaded");

                FindAndReplaceTags(presentation.Slides[0], user);

                //save and launch the result file
                string outputFile = Path.Combine(outputFolder, user[uniqueKey] + FileExtension);
                presentation.SaveToFile(outputFile, FileFormat.PDF);
                Console.WriteLine($"File saved: {outputFile}");
            }
        }

        /// [summary]
        /// Find and replace existing strings in slide with new strings.
        /// [/summary]
        /// [param name="slide"]specify the specific slide[/param]
        /// [param name="dictionary"]where keys are strings to place, values are strings for replacement[/param]
        public static void FindAndReplaceTags(Spire.Presentation.ISlide slide, Dictionary<string, string> dictionary)
        {
            foreach (IShape shape in slide.Shapes)
            {
                if (shape is IAutoShape)
                {
                    foreach (TextParagraph tp in (shape as IAutoShape).TextFrame.Paragraphs)
                    {
                        foreach (string key in dictionary.Keys)
                        {
                            if (tp.Text.Equals($"${key}$"))
                            {
                                tp.Text = tp.Text.Replace($"${key}$", dictionary[key]);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Get config section from the app config file
        /// </summary>
        /// <param name="sectionName">Section to be fetched</param>
        /// <returns>Dictionary containing section</returns>
        public static Dictionary<string, string> GetConfigurationUsingSection(string sectionName)
        {
            NameValueCollection applicationSettings = ConfigurationManager.GetSection(sectionName) as NameValueCollection;
            if (applicationSettings.Count == 0)
            {
                Console.WriteLine("Application Settings are not defined");
            }
            else
            {
                Dictionary<string, string> dictionary = new Dictionary<string, string>();

                foreach (string key in applicationSettings.AllKeys)
                {
                    dictionary.Add(key, applicationSettings[key]);
                    //Console.WriteLine(key + " = " + applicationSettings[key]);
                }

                return dictionary;
            }

            return null;
        }

        /// <summary>
        /// Prepare the Email Body for each user
        /// </summary>
        /// <param name="users">User list</param>
        /// <returns>Dictionary containing messages</returns>
        Dictionary<string, string> GetMessages(CSVTable users)
        {
            Dictionary<string, string> messages = new Dictionary<string, string>(users.Records.Count);

            foreach (CSVRecord user in users.Records)
            {
                messages[user[uniqueKey]] = PrepareEmailBody(user);
            }

            return messages;
        }

        /// <summary>
        /// Prepare the Email body
        /// </summary>
        /// <param name="tags">User info from the csv file</param>
        /// <returns>User specific message</returns>
        string PrepareEmailBody(Dictionary<string, string> tags)
        {
            string message = emailBodyTemplate;
            foreach (KeyValuePair<string, string> keyValue in tags)
            {
                message = message.Replace($"${keyValue.Key}$", keyValue.Value);
            }

            return message;
        }

        /// <summary>
        /// Send email with pdf attachment 
        /// </summary>
        /// <param name="users">Users list</param>
        /// <param name="appSettings">App settings</param>
        void SendEmail(CSVTable users)
        {
            Dictionary<string, string> messages = GetMessages(users);
            string attachmentFolder = appSettings["OutputFolder"];

            foreach (CSVRecord user in users.Records)
            {
                if (!user.ContainsKey("Email") || string.IsNullOrWhiteSpace(user["Email"]))
                {
                    continue;
                }

                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(appSettings["EmailFrom"], appSettings["EmailSenderName"]);
                    mail.To.Add(user["Email"]);
                    mail.Subject = appSettings["EmailSubject"];

                    mail.Body = messages[user[uniqueKey]];
                    mail.IsBodyHtml = true;

                    if (Convert.ToBoolean(appSettings["SendAttachment"]))
                    {
                        string attachmentPath = Path.Combine(attachmentFolder, user[uniqueKey] + FileExtension);
                        if (File.Exists(attachmentPath))
                        {
                            Attachment attachment = new Attachment(attachmentPath);
                            mail.Attachments.Add(attachment);
                        }
                    }

                    //mail.Attachments.Add(new Attachment("D:\\TestFile.txt"));//--Uncomment this to send any attachment  
                    using (SmtpClient smtp = new SmtpClient(appSettings["EmailSMTPHost"], Convert.ToInt32(appSettings["EmailSMTPPort"])))
                    {
                        smtp.Credentials = new NetworkCredential(appSettings["EmailFrom"], appSettings["EmailPassword"]);
                        smtp.EnableSsl = Convert.ToBoolean(appSettings["EmailEnableSSL"]);
                        Console.WriteLine($"Sending Email to {user["Email"]}");
                        smtp.Send(mail);
                        Console.WriteLine($"Sent Email to {user["Email"]}");
                    }
                }
            }
        }
    }
}
