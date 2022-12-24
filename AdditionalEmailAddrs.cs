// Copyright © 2020-2022  ASM-SW
//asmeyers@outlook.com  https://github.com/asm-sw

using FileHelpers;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Text;
//using System.Windows;

namespace EmailWithAttachedFile
{
    /// <summary>
    /// This Class defines the columns in the IO file.  FileHelpers uses it to read and write the CSV file.
    /// </summary>
    [DelimitedRecord(","), IgnoreFirst]
    class CustomerRecord
    {
        // NameLastFirst, Name,FileName, Email
        [FieldQuoted('"', QuoteMode.OptionalForBoth, MultilineMode.NotAllow), FieldTrimAttribute(TrimMode.Both)]
        public string NameLastFirst { get; set; }
        [FieldQuoted('"', QuoteMode.OptionalForBoth, MultilineMode.NotAllow), FieldTrimAttribute(TrimMode.Both)]
        public string Name { get; set; }
        [FieldQuoted('"', QuoteMode.OptionalForBoth, MultilineMode.NotAllow), FieldTrimAttribute(TrimMode.Both)]
        public string FileName { get; set; }
        [FieldQuoted('"', QuoteMode.OptionalForBoth, MultilineMode.NotAllow), FieldTrimAttribute(TrimMode.Both)]
        public string Email { get; set; }
    }

    /// <summary>
    /// This is the class that does all of the work of reading in files, parsing and modifying the IO file.
    /// The IO file email column will be modified to hold a semicolon seprated list of email addresses.
    /// These addtional addresses are from the addition eamil addres file. 
    /// The email address found in the IO file is looked up in the main email column of the addtional email address file.
    /// </summary>
    class AdditionalEmailAddrs
    {
        readonly Dictionary<string, EmailInfo> m_additionEmails = new Dictionary<string, EmailInfo>();

        /// <summary>
        /// This class is used to hold one line in the addition emails input file.
        /// </summary>
        class EmailInfo
        {
            public string Name { get; set; }
            public string EmailAddrs { get; set; }

            public EmailInfo()
            {
                Name = string.Empty;
                EmailAddrs = string.Empty;
            }
        }


        /// <summary>
        /// Reads in the addtional email file
        /// </summary>
        /// <param name="errMessage"></param>
        /// <param name="filenameAdditionEmaillAddr">filename to parse for email addresses</param>
        /// <returns>false on error with message in errMessage</returns>
        public bool ReadAdditionalEmailAddrFile(string filenameAdditionEmaillAddr, ref StringBuilder errMessage)
        {
            bool res = true;
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(filenameAdditionEmaillAddr))
                {
                    csvReader.SetDelimiters(new string[] { "," });
                    csvReader.HasFieldsEnclosedInQuotes = true;
                    string[] colFields = csvReader.ReadFields();

                    int nameIdx = -1;  // Customer name
                    int mainEmailIdx = -1;  // this is the main email address and is used as the key in the dictionaray
                    res = FindColumn(errMessage, colFields, "Customer", ref nameIdx);
                    res &= FindColumn(errMessage, colFields, "Main Email", ref mainEmailIdx);

                    // find the columns in the data that contain additional email address
                    List<int> emailIdxs = new List<int>();
                    for (int idx = 0; idx < colFields.Length; idx++)
                    {
                        if (idx == mainEmailIdx)
                            continue;
                        if (colFields[idx].Contains("Email"))
                            emailIdxs.Add(idx);
                    }
                    if (emailIdxs.Count < 1)
                    {
                        res = false;
                        errMessage.AppendFormat("No email addresses found in {0}\n", filenameAdditionEmaillAddr);
                    }
                    if (!res)
                        return res;

                    while (!csvReader.EndOfData)
                    {
                        EmailInfo info = new EmailInfo();
                        string[] fieldData = csvReader.ReadFields();
                        if (string.IsNullOrWhiteSpace(fieldData[mainEmailIdx]))
                            continue;

                        StringBuilder additionalAddrs = new StringBuilder(128);
                        additionalAddrs.Append(fieldData[mainEmailIdx]);
                        bool bFoundAdditionalEmails = false;
                        foreach (int index in emailIdxs)
                        {
                            if (!string.IsNullOrWhiteSpace(fieldData[index]))
                            {
                                bFoundAdditionalEmails = true;
                                additionalAddrs.AppendFormat(";{0}", fieldData[index]);
                            }
                        }
                        if (bFoundAdditionalEmails)
                        {
                            info.Name = fieldData[nameIdx];
                            info.EmailAddrs = additionalAddrs.ToString();

                            try
                            {
                                m_additionEmails.Add(fieldData[mainEmailIdx], info);
                            }
                            catch (ArgumentException)
                            {
                                errMessage.AppendFormat("Duplicate email found for: {0} - {1}\n", info.Name, info.EmailAddrs);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                errMessage.AppendLine(ex.ToString());
                return false;
            }
            return res;
        }

        /// <summary>
        /// Finds the column colName in colFields array.
        /// </summary>
        /// <param name="errMessage"></param>
        /// <param name="colFields">array to find colName in</param>
        /// <param name="colName">column to find</param>
        /// <param name="index">set to index of found column.</param>
        /// <returns>true if colName is found</returns>
        private bool FindColumn(StringBuilder errMessage, string[] colFields, string colName, ref int index)
        {
            index = Array.IndexOf(colFields, colName);
            if (index < 0)
            {
                errMessage.AppendFormat("Cannot find required column {0} in additional email addresses file\n", colName);
                return false;
            }
            return true;
        }

        /// <summary>
        /// gets a list of email addresses 
        /// </summary>
        /// <param name="mainEmailAddr">the main email address</param>
        /// <param name="additionalEmailAddresses">semicolon separated email addresses including the main email address</param>
        /// <returns>true if additonal email addresses are found, false if not</returns>
        public bool GetAddtionalEmailAddresses(string mainEmailAddr, out string additionalEmailAddresses)
        {
            if (m_additionEmails.TryGetValue(mainEmailAddr, out EmailInfo emailInfo))
            {
                additionalEmailAddresses = emailInfo.EmailAddrs;
                return true;
            }
            additionalEmailAddresses = string.Empty;
            return false;
        }

    }
}
