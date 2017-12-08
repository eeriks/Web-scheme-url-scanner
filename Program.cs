using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Windows.Forms;
using WIA;

namespace RbfloteScanner {
    static class Program {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args) {
            string url = "localhost/test"; // replace with empty string for production
            string schema_name = "rbf-scan:";
            foreach (string a in args) {
                if (a.Length > schema_name.Length) {
                    url = a.Remove(0, schema_name.Length);  // [schema_name:]link
                    break;
                }
            }
            if (url.Length > 0) {
                Scanner scanner = new Scanner(url);  
                scanner.StartScan();
            }
        }
    }
    class Scanner {
        static string url;

        public Scanner(string sLink)
        {
            url = sLink;
            //url = "192.168.16.249:38000/lietvediba/1/d/371/";
        }
        public void StartScan() {
            CommonDialogClass commonDialogClass = new CommonDialogClass();
            Device scannerDevice = null;
            try {
                scannerDevice = commonDialogClass.ShowSelectDevice(WIA.WiaDeviceType.UnspecifiedDeviceType, false, false);
            } catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("Skeneris nav atrasts.Sazinieties ar sistēmas administratoru", "Nav skenera");
                System.Environment.Exit(1);
            }
            List <Image> images = Scan(scannerDevice.DeviceID);
            DialogResult answer;
            answer = MessageBox.Show(string.Format("Ieskenētas {0} lapas. Skenēt vēl?", images.Count), "Ieskenēts", MessageBoxButtons.YesNo);
            while (answer == DialogResult.Yes) {
                images.AddRange(Scan(scannerDevice.DeviceID));
                answer = MessageBox.Show(string.Format("Ieskenētas {0} lapas. Skenēt vēl?", images.Count), "Ieskenēts", MessageBoxButtons.YesNo);
            }
            if (images.Count > 0) {
                string temp_pdf = Path.GetTempFileName() + ".pdf";

                iTextSharp.text.Document document = new iTextSharp.text.Document();
                //iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 0f, 0f, 0f, 0f);
                using (var stream = new FileStream(temp_pdf, FileMode.Create, FileAccess.Write, FileShare.None)) {
                    iTextSharp.text.pdf.PdfWriter.GetInstance(document, stream);
                    
                    foreach (Image image in images) {
                        iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance((Image)image, (iTextSharp.text.BaseColor)null);
                        document.SetPageSize(new iTextSharp.text.Rectangle(image.Width, image.Height));
                        document.SetMargins(0f, 0f, 0f, 0f);
                        if (document.IsOpen()) {
                            document.NewPage();
                        } else {
                            document.Open();
                        }
                        // Fit image to A4 size - margins
                        img.ScaleToFit(document.PageSize.Width, document.PageSize.Height);
                        document.Add(img);
                        //File.Delete(image);
                    }
                    document.Close();
                    HttpUploadFile(url, temp_pdf);
                    File.Delete(temp_pdf);
                }
            }
        }

        public static List<Image> Scan(string scannerId) {
            List<Image> images = new List<Image>();
            bool hasMorePages = true;
            bool bFlatbed = true;
            while (hasMorePages) {
                // select the correct scanner using the provided scannerId parameter
                WIA.DeviceManager manager = new WIA.DeviceManager();
                WIA.Device device = null;
                foreach (WIA.DeviceInfo info in manager.DeviceInfos) {
                    if (info.DeviceID == scannerId) {
                        // connect to scanner
                        device = info.Connect();
                        break;
                    }
                }
                // device was not found
                if (device == null) {
                    // enumerate available devices
                    string availableDevices = "";
                    foreach (WIA.DeviceInfo info in manager.DeviceInfos) {
                        availableDevices += info.DeviceID + "\n";
                    }
                    // show error with available devices
                    throw new Exception("The device with provided ID could not be found. Available Devices:\n" + availableDevices);
                }
                if (!bFlatbed) {
                    IProperties properties = device.Properties;
                    foreach (Property property in properties)
                    {
                        if (property.PropertyID == 3096)
                        {
                            Property prop = property;
                            prop.set_Value(1);
                            break;
                        }
                    }
                }

                Item item = device.Items[1] as Item;
                int dpi = 200;
                Property horizontal_dpi = item.Properties.get_Item("6147");
                horizontal_dpi.set_Value(dpi);
                Property vertical_dpi = item.Properties.get_Item("6148");
                vertical_dpi.set_Value(dpi);

                /* 200dpi A4 page in pixels 1654 x 2339 pix */
                Property horizontal_size_px = item.Properties.get_Item("6151");
                horizontal_size_px.set_Value(1654);
                Property vertical_size_px = item.Properties.get_Item("6152");
                vertical_size_px.set_Value(2339);

                try
                {
                    // scan image
                    WIA.CommonDialog wiaCommonDialog = new WIA.CommonDialog();
                    ImageFile image = (ImageFile)wiaCommonDialog.ShowTransfer(item, FormatID.wiaFormatPNG, true);
                    // save to temp file
                    string fileName = Path.GetTempFileName();
                    File.Delete(fileName);
                    image.SaveFile(fileName);
                    image = null;
                    // add file to output list
                    images.Add(Image.FromFile(fileName));
                } catch (ArgumentException) {
                    bFlatbed = false;
                } finally {
                    item = null;
                    //determine if there are any more pages waiting
                    WIA.Property documentHandlingSelect = null;
                    WIA.Property documentHandlingStatus = null;
                    foreach (WIA.Property prop in device.Properties) {
                        if (prop.PropertyID == WIA_PROPERTIES.WIA_DPS_DOCUMENT_HANDLING_SELECT)
                            documentHandlingSelect = prop;
                        if (prop.PropertyID == WIA_PROPERTIES.WIA_DPS_DOCUMENT_HANDLING_STATUS)
                            documentHandlingStatus = prop;
                    }
                    // assume there are no more pages
                    hasMorePages = false;
                    // may not exist on flatbed scanner but required for feeder
                    if (documentHandlingSelect != null)
                    {
                        // check for document feeder
                        if ((Convert.ToUInt32(documentHandlingSelect.get_Value()) & WIA_DPS_DOCUMENT_HANDLING_SELECT.FEEDER) != 0)
                        {
                            hasMorePages = ((Convert.ToUInt32(documentHandlingStatus.get_Value()) & WIA_DPS_DOCUMENT_HANDLING_STATUS.FEED_READY) != 0);
                        }
                    }
                }
            }
            return images;
        }

        class WIA_DPS_DOCUMENT_HANDLING_SELECT {
            public const uint FEEDER = 0x00000001;
            public const uint FLATBED = 0x00000002;
        }
        class WIA_DPS_DOCUMENT_HANDLING_STATUS {
            public const uint FEED_READY = 0x00000001;
        }
        class WIA_PROPERTIES {
            public const uint WIA_RESERVED_FOR_NEW_PROPS = 1024;
            public const uint WIA_DIP_FIRST = 2;
            public const uint WIA_DPA_FIRST = WIA_DIP_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
            public const uint WIA_DPC_FIRST = WIA_DPA_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
            //
            // Scanner only device properties (DPS)
            //
            public const uint WIA_DPS_FIRST = WIA_DPC_FIRST + WIA_RESERVED_FOR_NEW_PROPS;
            public const uint WIA_DPS_DOCUMENT_HANDLING_STATUS = WIA_DPS_FIRST + 13;
            public const uint WIA_DPS_DOCUMENT_HANDLING_SELECT = WIA_DPS_FIRST + 14;
        }

        public static void HttpUploadFile(string url, string file)
        {
            //MessageBox.Show(string.Format("Uploading {0} to {1}", file, url));

            // newpath used if file failed to upload
            string newpath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop) + "\\" + DateTime.Now.ToString("yyyy-MM-dd_HHmmss") + ".pdf";

            string boundary = "---------------------------" + DateTime.Now.Ticks.ToString("x");
            byte[] boundarybytes = System.Text.Encoding.ASCII.GetBytes("\r\n--" + boundary + "\r\n");

            HttpWebRequest wr = (HttpWebRequest)WebRequest.Create("https:" + url);
            wr.ContentType = "multipart/form-data; boundary=" + boundary;
            wr.Method = "POST";
            wr.KeepAlive = true;
            wr.Credentials = System.Net.CredentialCache.DefaultCredentials;

            Stream rs = null;

            try
            {
                rs = wr.GetRequestStream();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kļūda sūtot uz sistēmu.\r\nSaglabāts: " + newpath, "Kļūda");
                File.Copy(file, newpath);
                System.Environment.Exit(1);
            }

            rs.Write(boundarybytes, 0, boundarybytes.Length);
            byte[] formitembytes = System.Text.Encoding.UTF8.GetBytes("Content-Disposition: form-data; name=\"action\"\r\n\r\nupload");
            rs.Write(formitembytes, 0, formitembytes.Length);
            rs.Write(boundarybytes, 0, boundarybytes.Length);

            string headerTemplate = "Content-Disposition: form-data; name=\"file\"; filename=\"{0}\"\r\nContent-Type: application/pdf\r\n\r\n";
            string header = string.Format(headerTemplate, string.Format("{0}.pdf", DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss")));
            byte[] headerbytes = System.Text.Encoding.UTF8.GetBytes(header);
            rs.Write(headerbytes, 0, headerbytes.Length);

            FileStream fileStream = new FileStream(file, FileMode.Open, FileAccess.Read);
            byte[] buffer = new byte[4096];
            int bytesRead = 0;
            while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) != 0)
            {
                rs.Write(buffer, 0, bytesRead);
            }
            fileStream.Close();

            byte[] trailer = System.Text.Encoding.ASCII.GetBytes("\r\n--" + boundary + "--\r\n");
            rs.Write(trailer, 0, trailer.Length);
            rs.Close();

            WebResponse wresp = null;
            try
            {
                wresp = wr.GetResponse();
                Stream stream2 = wresp.GetResponseStream();
                StreamReader reader2 = new StreamReader(stream2);
                System.Environment.Exit(0);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Kļūda sūtot uz sistēmu.\r\nSaglabāts: " + newpath, "Kļūda");
                File.Copy(file, newpath);
                System.Environment.Exit(1);
                if (wresp != null)
                {
                    wresp.Close();
                    wresp = null;
                }
            }
            finally
            {
                wr = null;
            }
        }
    }
}
