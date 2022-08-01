using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TMSKRAintegration
{
    public class Utilities
    {
        static string path = AppDomain.CurrentDomain.BaseDirectory + @"\Config.xml";
        static string NAVWebService = null;
        static string IsPasswordEncrypted = null;      
        public static void WriteLog(string text)
        {
            try
            {
                //set up a filestream
                string strPath = @"C:\Logs\TMSKRALOGS";
                string fileName = DateTime.Now.ToString("MMddyyyy") + "_logs.txt";
                string filenamePath = strPath + '\\' + fileName;
                Directory.CreateDirectory(strPath);
                FileStream fs = new FileStream(filenamePath, FileMode.OpenOrCreate, FileAccess.Write);
                //set up a streamwriter for adding text
                StreamWriter sw = new StreamWriter(fs);
                //find the end of the underlying filestream
                sw.BaseStream.Seek(0, SeekOrigin.End);
                //add the text
                sw.WriteLine(DateTime.Now.ToString() + " : " + text);
                //add the text to the underlying filestream
                sw.Flush();
                //close the writer
                sw.Close();
            }
            catch (Exception ex)
            {
                //throw;
                ex.Data.Clear();
            }
        }
        private static string responseString = null;
        public static string CallWebService(string req)
        {
            string action = "";
           
            var _url = GetConfigData("NavWebService");
            var _action = action;
            try
            {
                XmlDocument soapEnvelopeXml = CreateSoapEnvelope(req);
                HttpWebRequest webRequest = CreateWebRequest(_url, _action);
                try
                {
                    InsertSoapEnvelopeIntoWebRequest(soapEnvelopeXml, webRequest);

                    // begin async call to web request.
                    IAsyncResult asyncResult = webRequest.BeginGetResponse(null, null);

                    // suspend this thread until call is complete. You might want to
                    // do something usefull here like update your UI.
                    asyncResult.AsyncWaitHandle.WaitOne();

                    // get the response from the completed web request.
                    string soapResult;

                    using (WebResponse webResponse = webRequest.EndGetResponse(asyncResult))
                    {
                        using (StreamReader rd = new StreamReader(webResponse.GetResponseStream()))
                        {
                            soapResult = rd.ReadToEnd();
                        }
                        responseString = soapResult;
                    }
                }
                catch (WebException ex)
                {
                    StreamReader responseReader = null;

                    string exMessage = ex.Message;

                    if (ex.Response != null)
                    {
                        using (responseReader = new StreamReader(ex.Response.GetResponseStream()))
                        {
                            responseString = responseReader.ReadToEnd();
                        }
                    }
                }
            }
            catch (WebException ex) when (ex.Status == WebExceptionStatus.ProtocolError)
            {
                //code specifically for a WebException ProtocolError
                ex.Message.ToString();
            }
            //catch (WebException ex) when ((ex.Response as HttpWebResponse)?.StatusCode == HttpStatusCode.NotFound)
            //{
            //    //code specifically for a WebException NotFound
            //    responseString = ParseExceptionRespose(ex);
            //}
            //catch (WebException ex) when ((ex.Response as HttpWebResponse)?.StatusCode == HttpStatusCode.InternalServerError)
            //{
            //    //code specifically for a WebException InternalServerError
            //    responseString = ParseExceptionRespose(ex);
            //}
            //catch (WebException ex) when (ex.Status == WebExceptionStatus.Timeout)
            //{
            //    //code specifically for a WebException InternalServerError
            //    responseString = ParseExceptionRespose(ex);
            //}
            //finally
            //{
            //    //call this if exception occurs or not
            //    //wc?.Dispose();
            //}

            return responseString;
        }
        private static HttpWebRequest CreateWebRequest(string url, string action)
        {
            string username = GetConfigData("Username");
            string password = GetConfigData("Password");
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(url);
            webRequest.Headers.Add("SOAPAction", action);
            webRequest.ContentType = "text/xml;charset=\"utf-8\"";
            webRequest.Accept = "text/xml";
            webRequest.Method = "POST";
            /////////////////webRequest.Credentials = CredentialCache.DefaultCredentials;
            webRequest.Timeout = 10000; //time-out value in milliseconds
            NetworkCredential creds = new System.Net.NetworkCredential(username, password);
            webRequest.Credentials = creds;
            webRequest.PreAuthenticate = true;
            webRequest.UseDefaultCredentials = true;
            return webRequest;
        }
        private static XmlDocument CreateSoapEnvelope(string req)
        {
            XmlDocument soapEnvelopeDocument = new XmlDocument();
            soapEnvelopeDocument.LoadXml(req);
            return soapEnvelopeDocument;
        }
        private static void InsertSoapEnvelopeIntoWebRequest(XmlDocument soapEnvelopeXml, HttpWebRequest webRequest)
        {
            try
            {
                using (Stream stream = webRequest.GetRequestStream())
                {
                    soapEnvelopeXml.Save(stream);
                }
            }
            catch (Exception e)
            {
                responseString = e.ToString();
                Utilities.WriteLog(e.ToString());
            }
        }
        public static string GetJSONResponse(string str)
        {
            string resp = "";

            if (!string.IsNullOrEmpty(str) && str.TrimStart().StartsWith("<"))
            {
                XmlDocument xmlSoapRequest = new XmlDocument();
                xmlSoapRequest.LoadXml(str);

                XmlNode faultstringNode;
                XmlNode return_valueNode;
                XmlElement root = xmlSoapRequest.DocumentElement;

                // Selects all the title elements that have an attribute named lang
                faultstringNode = xmlSoapRequest.GetElementsByTagName("faultstring")[0];
                return_valueNode = xmlSoapRequest.GetElementsByTagName("return_value")[0];

                if (faultstringNode != null)
                {
                    // It was found, manipulate it.
                    resp = faultstringNode.InnerText;

                }
                if (return_valueNode != null)
                {
                    // It was found, manipulate it.
                    resp = return_valueNode.InnerText;
                }
            }
            else
            { }

            return resp;
        }
        public static void GetServiceConstants()
        {
            NAVWebService = GetConfigData("NavWebService");
            var UserName = GetConfigData("Username");
            var Password = GetConfigData("Password");
            IsPasswordEncrypted = GetConfigData("IsEncrypted");
            //JobQueueCategory = GetConfigData("JobQueueCategory");
            //  connectionstring = GetConfigData("ConnectionString");

            if (IsPasswordEncrypted == "N")
            {
                string EncryptedPassword = EncryptDecrypt.Encrypt(Password, true);
                //updateConfig
                UpDateConfig(NAVWebService, "Settings/NavWebService");
                // UpDateConfig(connectionstring, "Settings/ConnectionString");
                UpDateConfig("Y", "Settings/IsEncrypted");
                UpDateConfig(EncryptedPassword, "Settings/Password");
            }
            else if (IsPasswordEncrypted == "Y")
            {
                Password = EncryptDecrypt.Decrypt(Password, true);
            }
        }
        private static void UpDateConfig(string Value, string XMLNode)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(path);
                doc.SelectSingleNode(XMLNode).InnerText = Value;

                doc.Save(path); //This will save the changes to the file.
            }
            catch (Exception es)
            {
                WriteLog(es.Message);
            }
        }
        public static string GetConfigData(string XMLNode)
        {
            string value = "";
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(path);
                XmlNode WebServiceNameNode = doc.GetElementsByTagName(XMLNode)[0];

                value = WebServiceNameNode.InnerText;
            }
            catch (Exception es)
            {
                WriteLog(es.Message);
            }
            return value;
        }
    }
}
