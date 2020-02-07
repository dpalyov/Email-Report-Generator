using Microsoft.Extensions.Configuration;
using MailService;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using Extensions;
using System.Collections.Generic;

namespace MailReportGenerator
{
    class Start
    {
        private static IConfigurationRoot _appConfig;

        /// <summary>
        /// Runs a program that configures and sends by email a dynamic report defined by the user (appConfig.js)
        /// </summary>
        /// <param name="args">
        /// Params as follow:
        /// args[0] = App configuration location
        /// args[1] = Script file location OR Plain script
        /// </param>
        static void Main(string[] args)
        {
            //checks if we have the 2 mandatory arguments
            if (args.Length != 2)
            {
                throw new ArgumentException("Missing argument. Please provide app settings configuration and sql script file/command");
            }
            var appConfig = args[0];
            var sqlScript = args[1];

            if (!File.Exists(appConfig))
            {
                throw new FileNotFoundException("Specified app configuration file couldn`t be located");
            }

            string sqlCommand;

            //check if we have a text file to read or direct script
            if (!File.Exists(sqlScript))
            {
                sqlCommand = sqlScript;
            }
            else
            {
                using (var reader = new StreamReader(sqlScript))
                {
                    sqlCommand = reader.ReadToEnd();
                }
            }

            //setting up the app config file
            _appConfig = new ConfigurationBuilder()
              .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
              .AddJsonFile(appConfig,false,true)
              .Build();

            //retrieving data from the server...
            var dt = RetrieveFromDb(sqlCommand);

            bool.TryParse(_appConfig.GetSection("SendAsAttachment").Value, out var useXl);


            //binds to Email from config
            var email = _appConfig.GetSection("Email").Get<Email>();


            //check if attachment is enabled else sends email without
            if (useXl)
            {
                //binds to ExcelOptions from config
                var xlOpts = _appConfig.GetSection("ExcelOptions").Get<ExcelOptions>();
                var attachment = dt.ExportXL(_appConfig.GetSection("AttachmentLocation").Value, xlOpts);

                email.Attachments = new List<string> { attachment };

            }
            else
            {
                //constructs dynamically the html table based on the DataTable provided
                var htmlBody = new StringBuilder();
                htmlBody.Append(dt.ParseHtml());
                email.MessageBody += htmlBody;
            }

            //dispatches the message and disposes
            using (var mailService = new MailClient())
            {
                mailService.ConfigureMessage(email);
                mailService.SendMessage();
            }



        }
        /// <summary>
        /// Retrieves data from a db server
        /// </summary>
        /// <param name="sql">
        /// Script command to execute against the DB
        /// </param>
        /// <returns>DataTable object containing the result from the query</returns>
        static DataTable RetrieveFromDb(string sql)
        {
            var dt = new DataTable();
            using (var connection = new SqlConnection(_appConfig.GetSection("ConnectionStr").Value))
            {
                var command = connection.CreateCommand();
                command.CommandText = sql;

                connection.Open();
                using (var reader = command.ExecuteReader())
                {
                    dt.Load(reader);
                }
                connection.Close();
            }

            return dt;
        }

    }
}
