using NLog;
using NLog.Fluent;
using NLog.Targets;
using NTST.ScriptLinkService.Objects;
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System.Collections;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Services;
using System.Security.Policy;
using System.Net.Mail;
using System.Security.Cryptography.X509Certificates;

namespace TxPlanHardStop
{
    // Configuration File
    class TxPlanHardStopConfig
    {
        private static TxPlanHardStopConfig instance = null;
        private static Logger log = LogManager.GetCurrentClassLogger();
        // mailing list of poeple to notify
        public string mailList { get; set; }

        // Does a particular note service type require a treatment plan?
        public HashSet<string> RequiresTreatmentPlan = new HashSet<string>();

        private static void init()
        {
            // NOTE - DATA DRIVE THIS - PUT INTO A TABLE AND USE A SELECT TO
            //        RETRIEVE THESE VALUES
            // Read Configuration
            using (FileStream fs = File.OpenRead("C:\\inetpub\\wwwroot\\TxPlanHardStop\\TxPlanHardStopService.json"))
            {
                using (StreamReader configStreamReader = new StreamReader(fs))
                {
                    string configStr = configStreamReader.ReadToEnd();
                    log.Debug("SDJL: configStr = '" + configStr + "'");
                    instance = JsonConvert.DeserializeObject<TxPlanHardStopConfig>(configStr);
                    log.Debug("SDJL - RequiresTreatmentPlan = " + TxPlanHardStopConfig.GetTxPlanHardStopConfig().RequiresTreatmentPlan);
                    log.Debug("Config = " + TxPlanHardStopConfig.GetTxPlanHardStopConfig());
                    log.Debug("Config.RequiresHardStop = " + TxPlanHardStopConfig.GetTxPlanHardStopConfig().RequiresTreatmentPlan);
                    log.Debug("Config.mailList = '" + TxPlanHardStopConfig.GetTxPlanHardStopConfig().mailList + "'");
                }
            }
        }

        public static TxPlanHardStopConfig GetTxPlanHardStopConfig()
        {
            if (instance == null)
            {
                init();                
            }
            return instance;
        }

        public static void SetTxPlanHardStopConfig(TxPlanHardStopConfig config)
        {
            instance = config;
        }
    }

    /// <summary>
    /// Summary description for WebService1
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class TxPlanHardStopService : System.Web.Services.WebService
    {
        String client = "";
        String serviceStartTimeStr = "";
        String serviceEndTimeStr = "";
        String practitioner = "";
        String coPractitioner = "";
        String incidentToPractitioner = "";
        String claimType = "";
        String needToFindOut7002_12 = "";
        String modeOfDelivery = "";
        String specialStudiesCode = "";
        String primaryDiagnosis = "";
        String primaryDiagnoisisArrayStr = "";
        String[] primaryDiagnoses;
        String secondaryDiagnosis = "";
        String secondaryDiagnosisArrayStr;
        String[] secondaryDiagnoses;
        String progress = "";
        String noteType = "";
        String typeOfService = "";
        String dateOfServiceStr = "";
        String draftOrFinal = "";
        DateTime dateOfService;
        // If no end date, assume end date is today, (ASSUMPTION -> ASK TERRI)
        DateTime planStartDate = DateTime.Now;
        DateTime planEndDate = DateTime.Now;
        String planStatusValue = "DEFAULT";

        Logger log = LogManager.GetCurrentClassLogger();


        [WebMethod]
        public string GetVersion()
        {
            return "1.0";
        }

        [WebMethod]
        public OptionObject2015 RunScript(OptionObject2015 inputObject, String scriptParameter)
        {
            OptionObject2015 returnObject = CopyObject(inputObject);// = new OptionObject2015();
            try
            {
                // For testing purposes
                //returnObject.ErrorMesg = "IGNORE - Test Message from David Lloyd";
                //returnObject.ErrorMesg = "TEST ERROR MESSAGE";
                //sendNotificationToMailingList(returnObject);

                returnObject.ErrorCode = 0;  // Default value
                returnObject.ErrorMesg = ""; // Default value


                log.Debug("-----------------------------------------");
                log.Debug("SDJL - BEGIN TxPlanHardStop RunScript 13 '" + scriptParameter + "'");

                // Actually perform the action
                //Add your script call(s) here
                switch (scriptParameter)
                {
                    case "HS_TxPlanHardStop NoteOptionLoad":
                        log.Debug("SDJL - HS_TxPlanHardStop NoteOptionLoad");
                        break;

                    case "HS_TxPlanHardStop SelectType":
                        foreach (FormObject form in inputObject.Forms)
                        {
                            log.Debug("Form ID: " + form.FormId);
                            foreach (FieldObject field in form.CurrentRow.Fields)
                            {
                                log.Debug("SDJL FieldNumber '" + field.FieldNumber + "'");
                                log.Debug("SDJL FieldValue '" + field.FieldValue + "'");
                                // NOTE - move to switch statement once you find out what the 
                                // correct values to serach for are here.
                                switch (field.FieldNumber)
                                {
                                    case "50010":
                                        draftOrFinal = field.FieldValue;
                                        break;
                                    case "51200":
                                        // Client                                 
                                        client = field.FieldValue;
                                        break;

                                    case "51001":
                                        typeOfService = field.FieldValue;
                                        break;

                                    case "51011":
                                        dateOfServiceStr = field.FieldValue;
                                        try
                                        {
                                            dateOfService = DateTime.Parse(dateOfServiceStr);
                                        }
                                        catch (FormatException e)
                                        {
                                            returnObject.ErrorCode = 3;
                                            returnObject.ErrorMesg = "Invalid Date Formate: Date Of Service";
                                            log.Debug("SDJL Exception Caught: " + e.StackTrace());
                                        }
                                        break;

                                    case "3003":
                                        serviceStartTimeStr = field.FieldValue;
                                        break;
                                    case "3004":
                                        serviceEndTimeStr = field.FieldValue;
                                        break;
                                    case "7000":
                                        practitioner = field.FieldValue;
                                        break;
                                    case "7000.2":
                                        coPractitioner = field.FieldValue;
                                        break;
                                    case "7000.12":
                                        incidentToPractitioner = field.FieldValue;
                                        break;
                                    case "7000.3":
                                        claimType = field.FieldValue;
                                        break;
                                    case "7052":
                                        modeOfDelivery = field.FieldValue;
                                        break;
                                    case "7053":
                                        specialStudiesCode = field.FieldValue;
                                        break;
                                    case "10034":
                                        primaryDiagnoisisArrayStr = field.FieldValue;
                                        break;
                                    case "10050":
                                        primaryDiagnoisisArrayStr = field.FieldValue;
                                        break;
                                    case "10035":
                                        secondaryDiagnosis = field.FieldValue;
                                        break;
                                    case "10051":
                                        secondaryDiagnosisArrayStr = field.FieldValue;
                                        break;
                                    case "10750":
                                        progress = field.FieldValue;
                                        break;
                                    case "10751":
                                        noteType = field.FieldValue;
                                        break;

                                    default:
                                        log.Debug("Unknown Field Number: '" + field.FieldNumber + "'");
                                        break;


                                }
                            }
                        }
                        log.Debug("SDJL - Client  = '" + client + "'");
                        log.Debug("Date Of Service Str = '" + dateOfServiceStr + "'");
                        log.Debug("Date Of Service = '" + dateOfService.ToString("MM-dd-yyyy"));
                        log.Debug("SDJL - serviceStartTimeStr'" + serviceStartTimeStr + "'");
                        log.Debug("SDJL - serviceEndTimeStr'" + serviceEndTimeStr + "'");
                        log.Debug("SDJL - practitioner'" + practitioner + "'");
                        log.Debug("SDJL - coPractitioner'" + coPractitioner + "'");
                        log.Debug("SDJL - incidentToPractitioner'" + incidentToPractitioner + "'");
                        log.Debug("SDJL - claimType'" + claimType + "'");
                        log.Debug("SDJL - modeOfDelivery'" + modeOfDelivery + "'");
                        log.Debug("SDJL - modeOfDelivery'" + modeOfDelivery + "'");
                        log.Debug("SDJL - primaryDiagnoisisArrayStr '" + primaryDiagnoisisArrayStr + "'");
                        log.Debug("SDJL - secondaryDiagnosis '" + secondaryDiagnosis + "'");
                        log.Debug("SDJL - secondaryDiagnosisArrayStr '" + secondaryDiagnosisArrayStr + "'");
                        log.Debug("SDJL - progress '" + progress + "'");
                        log.Debug("SDJL - Note Type '" + noteType + "'");
                        log.Debug("SDJL - typeOfService  '" + typeOfService + "'");
                        log.Debug("SDJL - currentDate = '" + DateTime.Now.ToString("MM-dd-yyyy"));
                        log.Debug("SDJL - draftOrFinal = '" + draftOrFinal + "'");

                        // PERFORM SQL query - Retrieve plan end date, last updated date                   
                        //    SELECT patid, last_updated, plan_date, plan_end_date, tx_plan_number 
                        //        FROM tx_plan WHERE patid='<client>' ORDER BY last_updated DESC

                        /*
                         *   if most recent treatment plan.end date < current date
                         *           hard stop
                         *
                         *
                         *     if most recent treatment plan.end_date >= date_of_service
                         *           if (treatment plan status != final)
                         *               hard stop
                         */
                        // TODO - Data drive connection string in configuration file.
                        // TODO - connection string username/password needs to be in secure (encrypted) storage.
                        // TODO - This should be in an init method.
                        String connectionString = "DRIVER={InterSystems ODBC};SERVER=10.50.10.148;PORT=1972;DATABASE=AVCWS;UID=UATHS:HSData1;PWD=hsdata!001;";

                        // Retrieve the latest updated treatment plan for the client.
                        string queryString =
                            "SELECT TOP 1 patid, last_updated, plan_date, plan_end_date, tx_plan_number, plan_status_value FROM tx_plan WHERE patid=? ORDER BY last_updated DESC";
                        log.Debug("Before OdbcConnection");
                        using (OdbcConnection connection = new OdbcConnection(connectionString))
                        {
                            log.Debug("SDJL - 23c - Before create command");
                            OdbcCommand command = new OdbcCommand(queryString, connection);
                            command.Parameters.Add("@client", OdbcType.VarChar).Value = client.Trim();
                            log.Debug("SDJL command='" + command.CommandText + "'");
                            log.Debug("SDJL Paramters = '" + command.Parameters.ToString() + "'");


                            log.Debug("SDJL - After Open");
                            connection.Open();
                            log.Debug("SDJL - After Open");

                            // The row will contain the latest updated treatment plan.
                            OdbcDataReader reader = command.ExecuteReader();
                            log.Debug("SDJL - Before Reader");
                            while (reader.Read())
                            {
                                planStartDate = (DateTime)reader[2];
                                planEndDate = (DateTime)reader[3];
                                planStatusValue = (String)reader[5];
                                log.Debug("SDJL - Treatment planStartDate = '" + planStartDate + "'");
                                log.Debug("SDJL - Treatment planEndDate = '" + planEndDate + "'");
                                log.Debug("SDJL - Treatment planStatusValue = '" + planStatusValue + "'");
                            }
                            log.Debug("SDJL - After Reader");
                            // Close the reader - done.
                            reader.Close();
                        }

                        // ACTUAL BUSINESS LOGIC
                        // Avatar return codes
                        // 1 - Returns an error message and stops further processing of scripts
                        // 2 - Returns a message with OK/Cancel buttons
                        // 3 - Returns a message with an OK button
                        // 4 - Returns a message with Yes/No buttons
                        // 5 - Returns a URL to be opened in a new browser window
                        // 6 - Returns a list of FormIDs to launch Avatar forms
                        // 0 - Process Option Object and show no message

                        if (requiresTP(typeOfService))
                        {
                            log.Debug("SDJL - Hard Stop Point 0");
                            if (planEndDate < DateTime.Now)
                            {
                                log.Debug("SDJL - Hard Stop 1");
                                returnObject.ErrorCode = 1;
                                returnObject.ErrorMesg = "Hard Stop: Plan End Date: '" + planEndDate + "' is before today";
                            }
                            else
                            {
                                log.Debug("SDJL - Hard Stop 1 - Point 2");
                            }
                            if (planEndDate >= dateOfService)
                            {
                                bool planStatusValueEqualsFinal = String.Equals(planStatusValue, "Final", StringComparison.CurrentCultureIgnoreCase);
                                log.Debug("SDJL - Hard Stop 2, Point 1.2 planStatusValue='" + planStatusValue + "' testCompare='" + planStatusValueEqualsFinal + "'");
                                if (!planStatusValueEqualsFinal)
                                {
                                    log.Debug("SDJL - Hard Stop 2, Point 2 - '" + planStatusValue + "'");
                                    returnObject.ErrorCode = 1;
                                    returnObject.ErrorMesg = "Hard Stop: Plan End Date: '" + planEndDate + "' is before the date of service '" + dateOfService + "'.  Treatment plan is NOT FINAL.";
                                }
                                else
                                {
                                    log.Debug("SDJL - Hard Stop 2, Point 3.  Treatment plan is final.");
                                }
                            }
                            else
                            {
                                log.Debug("SDJL - Hard Stop 2 - Point 4 - planEndDate < dateOfService");
                            }
                            if (planStartDate > dateOfService)
                            {
                                log.Debug("SDJL - Hard Stop 3, Point 1");
                                returnObject.ErrorCode = 1;
                                returnObject.ErrorMesg = "Hard Stop: Plan Start Date: '" + planStartDate + "' is after the date of service '" + dateOfService + "'";
                            }
                            else
                            {
                                log.Debug("SDJL - Hard Stop 3, Point 2 - planStartDate <= dateOfService");
                            }
                        }
                        else
                        {
                            log.Debug("Service '" + typeOfService + "' does not require require hard stop");
                        }
                        break;

                }
                log.Debug("SDJL - END TxPlanHardStop RunScript 13 '" + scriptParameter + "'");
                log.Debug("SDJL - Before Return RETVAL = '" + returnObject.ErrorCode + "' message='" + returnObject.ErrorMesg + "'");

                // Send notification to mailing list
                if (returnObject.ErrorCode != 0)
                {
                    sendNotificationToMailingList(returnObject);
                }
                log.Debug("---------------------------------------");
            }
            catch (Exception e)
            {
                log.Debug("SDJL - Exception: "+e.ToString());
            }
            return returnObject;
        }

        private static OptionObject2015 CopyObject(OptionObject2015 inputObject)
        {
            OptionObject2015 returnObject = new OptionObject2015();
            returnObject.OptionId = inputObject.OptionId;
            returnObject.Facility = inputObject.Facility;
            returnObject.SystemCode = inputObject.SystemCode;
            returnObject.NamespaceName = inputObject.NamespaceName;
            returnObject.ParentNamespace = inputObject.ParentNamespace;
            returnObject.ServerName = inputObject.ServerName;
            return returnObject;
        }


        private bool requiresTP(String service)
        {
            bool retVal = false;
            log.Debug("SDJL: serviceChosen: Service '" + service + "'");
            log.Debug("SDJL: serviceConfiguration'" + TxPlanHardStopConfig.GetTxPlanHardStopConfig().ToString() + "'");
            log.Debug("SDJL: serviceConfiguration.RequiresTreatmentPlan'" + TxPlanHardStopConfig.GetTxPlanHardStopConfig().RequiresTreatmentPlan + "'");
            log.Debug("SDJL: RequiresTreatmentPlan.Contains('" + service + "') returnValue= '" + TxPlanHardStopConfig.GetTxPlanHardStopConfig().RequiresTreatmentPlan.Contains(service) + "'");
            if (TxPlanHardStopConfig.GetTxPlanHardStopConfig().RequiresTreatmentPlan.Contains(service))
            {
                retVal = true;
            }
            log.Debug("SDJL: END TxPlanHardStop requiresTP RETVAL = '" + retVal + "'");
            return retVal;
        }

        private void sendNotificationToMailingList(OptionObject2015 hardStopToReturn)
        {
            // Data drive configuration
            try
            {
                MailMessage mail = new MailMessage("davidl@spmhc.org",TxPlanHardStopConfig.GetTxPlanHardStopConfig().mailList);
                SmtpClient SmtpServer = new SmtpClient("10.20.20.214");

                mail.Subject = "HARD STOP: " + System.DateTime.Now.ToString("MM/dd/yyyy h:mm");
                mail.Body = "A hard stop occurred at " + System.DateTime.Now.ToString("MM/dd/yyyy h:mm") + " for client: " + client + ".  ";
                mail.Body += "\n  Terri, this an contain whatever you like.\n";
                mail.Body += "Value:  " + hardStopToReturn.ErrorMesg;

                SmtpServer.Port = 25;
                SmtpServer.Credentials = new System.Net.NetworkCredential("davidl@spmhc.org", "3edcVFR$3edc");

                SmtpServer.Send(mail);
                log.Debug("SDJL: Mail Sent at: " + DateTime.Now.ToString("MM/dd/yyyy h:mm"));
            }
            catch (Exception e)
            {
                log.Debug("SDJL Send Email Exception: " + e.ToString());
            }
        }
    }
}
