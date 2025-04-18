using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Timers;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Specialized;
using System.Configuration;
using Microsoft.Win32;
using System.Xml;
using System.IO.Compression;
using System.Security.Cryptography;


namespace SpreadWorkDesktop
{
    public partial class Form1 : Form
    {
        private NameValueCollection cfgGlobal;
        System.Timers.Timer TimerCaptura = new System.Timers.Timer();
        int Interval = 0;
        DateTime StartDateTime = new DateTime();
        DateTime EndDateTime = new DateTime() ;
        String PicWidth = "";
        String PicHeight = "";
        String Percent = "";
        float scale = 1;
        int scaleWidth =1;
        int scaleHeight = 1;
        string Path;
        string RemoteConfigPath;
        string RemoteFilesPath;
        string Event;
        private static log4net.ILog log; 
        public Form1()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            LoadProgramConfig();
            SetProgramConfig();
            LoadRegisterConfig();
            SetLogConfig();
            SetRemoteConfig();
            this.MouseClick += delegate (object sender, MouseEventArgs e)
            {
                log.Info("Program hide after click");
                this.WindowState = FormWindowState.Minimized;
                this.Hide();
            };

            this.Activated += delegate (object sender, System.EventArgs e)
            {
                log.Info("Program hide after activation");
                this.WindowState = FormWindowState.Minimized;
                this.Hide();
            };

            this.Load += delegate (object sender, System.EventArgs e)
            {
                log.Info("Program hide after load");
                this.WindowState = FormWindowState.Minimized;
                this.Hide();
            };
        }

        private void Form1_Load(object sender, EventArgs e)
        {
         // Escribe log
         log.Info("SpreadWorkDesktop Program Started"); 
         // Configurar Timer para ejecución períodica de capturas
        TimerCaptura.Elapsed += new ElapsedEventHandler(OnElapsedTimeForCapture);
        TimerCaptura.Interval = Interval;
        TimerCaptura.Enabled = true;
        }

        private void OnElapsedTimeForCapture(object sender, ElapsedEventArgs e)
        {
           // Escribe log
           log.Info("Capture check on " + Interval.ToString() + " ms elapsed."); 
            // Detener Timer para hacer el proceso
           TimerCaptura.Enabled = false;
           // Si está en el periodo indicado en la configuración
           if (DateTime.Now >=StartDateTime && DateTime.Now <=EndDateTime)
            { 
            log.Info("Time to take a screen capture configurated"); 
             if (String.IsNullOrWhiteSpace(PicWidth) && String.IsNullOrWhiteSpace(PicHeight) && String.IsNullOrWhiteSpace(Percent))
                { 
                 log.Info("Taking screen capture without resize"); 
                 // Captura imagen a tamaño real
                 Size size = GetDisplayResolution();
                 using (Bitmap bitmap = new Bitmap(size.Width, size.Height, PixelFormat.Format32bppArgb))
                 {
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
                        g.CopyFromScreen(Point.Empty, Point.Empty, size);
                    }
                    SaveScreen(bitmap);
                 }
             } else
             { 
              log.Info("Taking screen capture with " + Percent + " scale"); 
              // Aplicar reescalado 
              Size size = GetDisplayResolution();
               // Si Percent tiene valor se usa ese valor en lugar del Width y el Height. 
               // Pero los valores se hacen siempre en base a Width y Height
               if (!String.IsNullOrEmpty(Percent)) { 
                    PicHeight = (size.Height * float.Parse(Percent)).ToString();
                    PicWidth = (size.Width * float.Parse(Percent)).ToString();
                }
              using (MemoryStream ms = new MemoryStream())
               {
                using (Bitmap bitmap = new Bitmap(size.Width, size.Height, PixelFormat.Format32bppArgb))
                 {
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
                        g.CopyFromScreen(Point.Empty, Point.Empty, size);
                    }
                    bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg); 
                 }
               using (Bitmap bitmap = new Bitmap(Int32.Parse(PicWidth), Int32.Parse(PicHeight), PixelFormat.Format32bppArgb))
               {
                using (Graphics g = Graphics.FromImage(bitmap))
                 {
                  scale = Math.Min((float)Int32.Parse(PicWidth) / (float)size.Width, (float)Int32.Parse(PicHeight) / (float)size.Height);
                  scaleWidth = (int)(size.Width * scale);
                  scaleHeight = (int)(size.Height * scale);
                  // Para mejorar la calidad de la imagen escalada
                  g.InterpolationMode = InterpolationMode.High;
                  g.CompositingQuality = CompositingQuality.HighQuality;
                  g.SmoothingMode = SmoothingMode.AntiAlias;
                  // Fondo negro
                  using (Brush baseColor = new SolidBrush(Color.Black))
                     g.FillRectangle(baseColor, new Rectangle(0, 0, Int32.Parse(PicWidth), Int32.Parse(PicHeight)));
                  
                  // Reescalar el buffer con la imagen a tamaño normal.
                  g.DrawImage(Image.FromStream(ms), ((int)Int32.Parse(PicWidth) - scaleWidth) / 2, ((int)Int32.Parse(PicHeight) - scaleHeight) / 2, scaleWidth, scaleHeight);
                  }
                 SaveScreen(bitmap);
               }
               }
             }
            } 
           else { 
            log.Info("Program not configurated to take a screen capture");
                // Comprobar si hay que enviar datos 
                // Si se sobrepasan las fechas de comienzo y fin
                if (DateTime.Now > StartDateTime && DateTime.Now > EndDateTime)
                {
                    // Si hay datos en el directorio local 
                    int numFiles = CheckForFilesToSend();
                    if (numFiles != 0)
                    {
                        if (CheckRemoteFilesPath())
                        {

                            log.Info("Preparing files to Send: " + numFiles.ToString());
                            string zipName = GetZipFileName();
                            CreateZipFile(zipName);
                            bool result = SendZipFile(zipName);
                            if (result)
                            {
                                log.Info("Zip File sent: " + zipName);
                                if (DeleteZipAndFiles(zipName))
                                    log.Info("Zip " + zipName + " and Files deleted");
                            }
                            else
                            {
                                log.Info("Error sendig Zip File " + zipName);
                            }
                        }
                        else
                        {
                            log.Info("Exists " +  numFiles.ToString() + " files to send but remote Path " + RemoteFilesPath + " is not accesible");
                        }
                    }
                    else
                        log.Info("No files pending to send");
                }

            }
            // Comprobar cambios remotos de configuración y aplicarlos
            SetRemoteConfig();
            // Inicializa Timer
            TimerCaptura.Enabled = true;

        }

        #region Configuration
    private void LoadProgramConfig() { 
            // Leer todos los parámetros del fichero de configuración
            ConfigurationManager.RefreshSection("appSettings");
            cfgGlobal = ConfigurationManager.AppSettings;
        }

    private void SetProgramConfig() { 
            // Establecer valores cargados de App.Config
            Interval = Int32.Parse(cfgGlobal.Get("Interval"));
            StartDateTime = DateTime.Parse(cfgGlobal.Get("StartDateTime"));
            EndDateTime = DateTime.Parse(cfgGlobal.Get("EndDateTime"));
            PicWidth = cfgGlobal.Get("Width");
            PicHeight =cfgGlobal.Get("Height");
            Percent =cfgGlobal.Get("Percent");
            Event = cfgGlobal.Get("Event");
        }

        private void LoadRegisterConfig() {
            Path = GetKeyFromRegistry("Folder");
            RemoteConfigPath = GetKeyFromRegistry("RemoteConfigPath");
            RemoteFilesPath = GetKeyFromRegistry("RemoteFilesPath");
            // Si no existe crearlo oculto
            try
            {
                if (!Directory.Exists(Path))
                {
                    DirectoryInfo di = Directory.CreateDirectory(Path);
                    di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                }
            }
            catch (Exception ex)
            {
                log.Error("Error creation local path " + Path + " - " + ex.Message);
            }
                    
        }

        private void SetLogConfig()
        {
            string path = Path + "logs\\";
            if (!Directory.Exists(path))
            {
                DirectoryInfo di = Directory.CreateDirectory(path);
                di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
            }
            string logFile = path + "\\SpreadWorkDesktopLog.log";
            log4net.GlobalContext.Properties["LogFileName"] = logFile;
            log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        }


        private void SetRemoteConfig()
        {
            //Obtener Grupo del registro del 
            string group = GetKeyFromRegistry("Group");
            if (group == "NONE")
                log.Error("No group found in the register key ");
            //Comprobar acceso a ruta remota de configuración
             if (Directory.Exists(RemoteConfigPath))
                {
                LoadRemoteValues(RemoteConfigPath + "\\" + group + ".xml"); 
                }
                else {
                    log.Error("No access to remote folder " + RemoteConfigPath);
                }
            }

        private string GetKeyFromRegistry(string regkey) {
            // Ruta de la clave de registro
            string regkeyname = @"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\SpreadWork\Config";
            string keyvalue = (string) Registry.GetValue(regkeyname, regkey, "NONE");
            return keyvalue;
        }

        private void LoadRemoteValues(string xmlFile) {
            XmlDocument doc = new XmlDocument();
            try
            {
                doc.Load(xmlFile);
                //Actualizar claves
                if (!String.IsNullOrEmpty(doc.DocumentElement.SelectSingleNode("/root/Interval").InnerText))
                    Interval = Int32.Parse(doc.DocumentElement.SelectSingleNode("/root/Interval").InnerText);
                if (!String.IsNullOrEmpty(doc.DocumentElement.SelectSingleNode("/root/StartDateTime").InnerText))
                    StartDateTime = DateTime.Parse(doc.DocumentElement.SelectSingleNode("/root/StartDateTime").InnerText);
                if (!String.IsNullOrEmpty(doc.DocumentElement.SelectSingleNode("/root/EndDateTime").InnerText))
                    EndDateTime = DateTime.Parse(doc.DocumentElement.SelectSingleNode("/root/EndDateTime").InnerText);
                if (!String.IsNullOrEmpty(doc.DocumentElement.SelectSingleNode("/root/Width").InnerText))
                    PicWidth = doc.DocumentElement.SelectSingleNode("/root/Width").InnerText;
                if (!String.IsNullOrEmpty(doc.DocumentElement.SelectSingleNode("/root/Height").InnerText))
                    PicHeight = doc.DocumentElement.SelectSingleNode("/root/Height").InnerText;
                if (!String.IsNullOrEmpty(doc.DocumentElement.SelectSingleNode("/root/Percent").InnerText))
                    Percent = doc.DocumentElement.SelectSingleNode("/root/Percent").InnerText;
                if (!String.IsNullOrEmpty(doc.DocumentElement.SelectSingleNode("/root/Event").InnerText))
                    Event = doc.DocumentElement.SelectSingleNode("/root/Event").InnerText;
                // Actualizar intervalo
                TimerCaptura.Interval = Interval;
            }
            catch (Exception ex) {
                log.Error("Error reading remote XML config file:  " + ex.Message);
            }
        }

        #endregion

        #region file sending
        private int CheckForFilesToSend() {
            string directorio = Path + "files\\"; 
            DirectoryInfo directorioInfo = new DirectoryInfo(directorio);
            if (directorioInfo.Exists)
            {
                FileInfo[] archivos = directorioInfo.GetFiles();
                if (archivos.Length > 0)
                    return archivos.Length;
                else
                    return 0;
            }
            else
            {
                Console.WriteLine("El directorio no existe.");
                return 0;
            }
        }

        private string GetZipFileName() {
            // Obtener nombre de cuenta logeada en el equipo. 
            string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            int pos = userName.IndexOf('\\');
            userName =  pos != -1 ? userName.Substring(pos + 1) : userName;
            // Obtener nombre del equipo (hostname)
            string hostName = System.Windows.Forms.SystemInformation.ComputerName;
            // Concatenar con fecha 
            string fecha = DateTime.Now.ToString("ddMMyyyy");
            // Añadir Random de 1 a 1000
            Random random = new Random();
            string rString = random.Next(0, 1000).ToString();
            return (userName + "_" + hostName+"_"+fecha+"_"+ rString);
        }

        private void CreateZipFile(string zipName) {
            string OrigPath = Path + "files\\";
            string DestPath = Path + "tmp\\";
            string DestFile = DestPath + zipName + ".zip";
            
            // Crear directorio temporal para el fichero .zip si no existe
            if (!Directory.Exists(DestPath))
            {
                DirectoryInfo di = Directory.CreateDirectory(DestPath);
                di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
            }
            // Comprimir el contenido de Path\files
            try
            {
                ZipFile.CreateFromDirectory(OrigPath, DestFile, CompressionLevel.Fastest, false);
            }
            catch (Exception ex) {
                log.Error("Error creating zip file: " + ex.Message);
            }

        }

        private bool CheckRemoteFilesPath()
        {
            string DestPath = RemoteFilesPath;
            // Crear directorio destino
            try
            {
                if (Directory.Exists(DestPath))
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                log.Error("Can't acces destination folder for leaving zip file" + ex.Message);
                return false;
            }
        }

        private bool SendZipFile(string zipName)
        {
            string DestPath = RemoteFilesPath + Event + "\\";
            // Crear directorio destino
            try
            {
                if (!Directory.Exists(DestPath))
                    Directory.CreateDirectory(DestPath);
            }
            catch (Exception ex)
            {
                log.Error("Can't create destination folder for zip file" + ex.Message);
                return false;
            }
            // Enviar fichero 
            string OrigPath = Path + "tmp\\";
            string OrigFile = OrigPath + zipName + ".zip";
            string DestFile = DestPath  + zipName + ".zip";
            try {
                File.Copy(OrigFile, DestFile, true);
            }
            catch (Exception ex)
            {
                log.Error("Can't send zip file to destination folder: " + ex.Message);
                return false;
            }
            // Comprobar que ha llegado usando MD5 Hash
            string origFileMD5 = getMD5(OrigFile);
            string destFileMD5 = getMD5(DestFile);
            if (!string.IsNullOrEmpty(origFileMD5) && !string.IsNullOrEmpty(destFileMD5) && (origFileMD5 == destFileMD5))
                return true;
            else
                return false;
        }

        private string getMD5(string filePath) {
            try
            {
                using (var md5 = MD5.Create())
                {
                    using (var stream = File.OpenRead(filePath))
                    {
                        var hash = md5.ComputeHash(stream);
                        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                    }
                }
            }
            catch (Exception ex) {
                log.Error("Error calculating MD5 for file " + filePath + " - " + ex.Message);
                return string.Empty;
            }
        }

        private bool DeleteZipAndFiles(string zipName)
        {
            try
            {
                // Borrar el fichero zip en el path local 
                string OrigZipPath = Path + "tmp\\";
                string OrigZipFile = OrigZipPath + zipName + ".zip";
                if (File.Exists(OrigZipFile))
                    File.Delete(OrigZipFile);
                // Borrar el contenido del directorio files
                string OrigFilesPath = Path + "files\\";
                if (Directory.GetFiles(OrigFilesPath).Length > 0)
                    Array.ForEach(Directory.GetFiles(OrigFilesPath), File.Delete);
                // Debolver ok
                return true;
            }
            catch (Exception ex)
            {
                log.Error("Files sent can't be deleted: " + ex.Message);
                return false;
            }
        }


        #endregion

        #region Save captures
        private void SaveScreen(Bitmap bitmap)
        {
            // Crear directirio para guardar fotos si no existe
            //var path = AppDomain.CurrentDomain.BaseDirectory + "files\\";
            var path = Path + "files\\";
            if (!Directory.Exists(path)) { 
            DirectoryInfo di = Directory.CreateDirectory(path);
            di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
            }
            // Nombre de la foto usando tiempo unix
            // var horaunix = DateTime.Now.ToFileTimeUtc().ToString();
            var horaunix =  DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            var fechahora = DateTime.Now.ToString("ddMM_HHmm");
            var name = "pic" + horaunix +"_" +fechahora +".jpg";
            var file = path + name ;
            try { 
                bitmap.Save(file, System.Drawing.Imaging.ImageFormat.Jpeg);
                }
            catch(Exception e) {
                log.Error("Error saving screen capture" + e.Message);
            }
        }

        #endregion

        #region Display Resolution

        [DllImport("gdi32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern int GetDeviceCaps(IntPtr hDC, int nIndex);

        public enum DeviceCap
        {
            VERTRES = 10,
            DESKTOPVERTRES = 117
        }


        public static double GetWindowsScreenScalingFactor(bool percentage = true)
        {
            //Create Graphics object from the current windows handle
            Graphics GraphicsObject = Graphics.FromHwnd(IntPtr.Zero);
            //Get Handle to the device context associated with this Graphics object
            IntPtr DeviceContextHandle = GraphicsObject.GetHdc();
            //Call GetDeviceCaps with the Handle to retrieve the Screen Height
            int LogicalScreenHeight = GetDeviceCaps(DeviceContextHandle, (int)DeviceCap.VERTRES);
            int PhysicalScreenHeight = GetDeviceCaps(DeviceContextHandle, (int)DeviceCap.DESKTOPVERTRES);
            //Divide the Screen Heights to get the scaling factor and round it to two decimals
            double ScreenScalingFactor = Math.Round(PhysicalScreenHeight / (double)LogicalScreenHeight, 2);
            //If requested as percentage - convert it
            if (percentage)
            {
                ScreenScalingFactor *= 100.0;
            }
            //Release the Handle and Dispose of the GraphicsObject object
            GraphicsObject.ReleaseHdc(DeviceContextHandle);
            GraphicsObject.Dispose();
            //Return the Scaling Factor
            return ScreenScalingFactor;
        }

        public static Size GetDisplayResolution()
        {
            var sf = GetWindowsScreenScalingFactor(false);
            var screenWidth = Screen.PrimaryScreen.Bounds.Width * sf;
            var screenHeight = Screen.PrimaryScreen.Bounds.Height * sf;
            return new Size((int)screenWidth, (int)screenHeight);
        }

        #endregion



    }
}
