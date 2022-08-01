using IronBarCode;
using Newtonsoft.Json;
using QRCoder;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using TMSKRAintegration;

namespace KRATMSInter
{
    public partial class TMSKRA : ServiceBase
    {
        public TMSKRA()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Utilities.WriteLog("Service Started");

            Timer timer = new Timer();

            timer.Elapsed += new System.Timers.ElapsedEventHandler(this._timer_Tick);

            timer.Enabled = true;

            timer.Interval = 10000;

            //timer.Start();

            Utilities.GetServiceConstants();
        }
        static string path = AppDomain.CurrentDomain.BaseDirectory + @"\Config.xml";
        static string localFilePath = Path.GetFullPath(Utilities.GetConfigData("Folderpath"));
        static string Uploaded = Path.GetFullPath(Utilities.GetConfigData("Uploaded"));
        static string imagepath = Path.GetFullPath(Utilities.GetConfigData("Imagepath"));
        static string qrpath = Path.GetFullPath(Utilities.GetConfigData("QRpath"));
        private void _timer_Tick(object sender, ElapsedEventArgs e)
        {
            string filename = null;
            try
            {
                //get all text files in the folder
                string[] filePaths = Directory.GetFiles(localFilePath, "*.txt");
                List<string> value = filePaths.ToList();

                if (value != null)
                {
                    //update one by one
                    foreach (var doc in value)
                    {
                        filename = Path.GetFileName(doc);
                        ////get only filename                   
                        //var file = invoicelines(doc);
                        try
                        {
                            string[] lines = System.IO.File.ReadAllLines(doc);
                            if (lines != null)
                            {
                                foreach (string line in lines)
                                {
                                    // Use a tab to indent each line of the file.
                                    if (line.Contains("https:"))
                                    {
                                        Link = line;

                                    }
                                    else if (line.Contains("TSIN:"))
                                    {
                                        TSIN = line.Substring(line.IndexOf(':') + 1).TrimEnd();

                                    }
                                    else if (line.Contains("DATE:"))
                                    {
                                        Date = line.Substring(line.IndexOf(':') + 1).TrimEnd();

                                    }
                                    else if (line.Contains("CUSN:"))
                                    {
                                        CUSN = line.Substring(line.IndexOf(':') + 1).TrimEnd();

                                    }
                                    else if (line.Contains("CUIN:"))
                                    {
                                        CUIN = line.Substring(line.IndexOf(':') + 1).TrimEnd();
                                    }

                                }

                                var FiscalSeal = invoicelines(Link);
                                var resTSIN = invoicelines(TSIN);
                                var resDate = invoicelines(Date);
                                var resCUSN = invoicelines(CUSN);
                                var resCUIN = invoicelines(CUIN);
                                //get details from invoice

                                QRCodeGenerator qrcode = new QRCodeGenerator();
                                QRCodeData data = qrcode.CreateQrCode(FiscalSeal, QRCodeGenerator.ECCLevel.Q);
                                QRCode qR = new QRCode(data);
                                var valueda = qR.ToString();
                                System.Web.UI.WebControls.Image imgBarCode = new System.Web.UI.WebControls.Image();
                                imgBarCode.Height = 150;
                                imgBarCode.Width = 150;
                                using (Bitmap bitMap = qR.GetGraphic(20))
                                {
                                    using (MemoryStream ms = new MemoryStream())
                                    {
                                        bitMap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                        byte[] byteImage = ms.ToArray();
                                        System.Drawing.Image img = System.Drawing.Image.FromStream(ms);
                                        img.Save(imagepath + resCUIN + ".Jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);
                                    }
                                }

                                string[] image = Directory.GetFiles(imagepath, "*.Jpeg");
                                List<string> qr = image.ToList();
                                foreach (var im in qr)
                                {
                                    var imagename = Path.GetFileName(im);
                                    dqbase64 = GetBase64StringForImage(im);
                                    //Utilities.WriteLog(dqbase64);
                                    dynamic json1 = updateinfor(dqbase64, resTSIN, resDate, resCUSN, resCUIN);
                                    Utilities.WriteLog(json1);
                                    File.Copy(im, Path.Combine(@"C:\FBtemp\Imagepath\Uploaded\", imagename), true);
                                }

                            }
                            else
                            {
                                Utilities.WriteLog("The file is empty");
                                File.Delete(filename);
                            }
                            Directory.GetFiles(imagepath).ToList().ForEach(File.Delete);
                            if (File.Exists(doc))
                            {
                                File.Copy(doc, Path.Combine(@"C:\FBtemp\invoicebackup\", filename), true);
                                File.Delete(doc);
                            }
                        }
                        catch (Exception ex)
                        {
                            Utilities.WriteLog(ex.Message);
                            if (ex.Message == " Object reference not set to an instance of an object")
                            {
                                File.Delete(filename);
                            }

                        }
                    }
                }
                else
                {
                    Utilities.WriteLog("there is no File available");
                }


            }
            catch (Exception ex)
            {
                Utilities.WriteLog(ex.Message);
               
            }

        }
        public static string invoicelines(string textfile)
        {
            var position = textfile.Replace("|", string.Empty);
            var values = position.Replace(" ", string.Empty);
            return values;
        }
        public static string GetBase64StringForImage(string imgPath)
        {
            byte[] imageBytes = System.IO.File.ReadAllBytes(imgPath);
            string base64String = Convert.ToBase64String(imageBytes);
            return base64String;
        }
        public static string Date { get; private set; }
        public static string CUSN { get; private set; }
        public static string CUIN { get; private set; }
        public static string TSIN { get; private set; }
        public static string Link { get; private set; }
        public string dqbase64 { get; private set; }

        ///function for getting  updating to nav
        public static string updateinfor(string FiscalSeal, string TSIN, string TXDate, string CUSN, string CUIN)
        {
            string itemlist = @"<Envelope xmlns=""http://schemas.xmlsoap.org/soap/envelope/"">
                                    <Body>
                                        <UpdateSalesInvoiceWithTIMSDetails xmlns=""urn:microsoft-dynamics-schemas/codeunit/KRAEInvoicingIntegration"">
                                            <prFiscalSeal>" + FiscalSeal + @"</prFiscalSeal>
                                            <prTSIN>" + TSIN + @"</prTSIN>
                                            <prTXDate>" + TXDate + @"</prTXDate>
                                            <prCUSN>" + CUSN + @"</prCUSN>
                                            <prCUIN>" + CUIN + @"</prCUIN>
                                        </UpdateSalesInvoiceWithTIMSDetails>
                                    </Body>
                                </Envelope>";
            string response = Utilities.CallWebService(itemlist);
            return Utilities.GetJSONResponse(response);
        }
        protected override void OnStop()
        {
            Utilities.WriteLog("Service Stopped.");
        }
    }
}
