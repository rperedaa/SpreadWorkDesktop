using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Timers;
using System.Drawing.Imaging;
using System.IO;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Collections.Specialized;
using System.Configuration;

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


        
        public Form1()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            LoadProgramConfig();
            SetProgramConfig();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
         // Ejecución períodica
        TimerCaptura.Elapsed += new ElapsedEventHandler(OnElapsedTimeForCapture);
        TimerCaptura.Interval = Interval;
        TimerCaptura.Enabled = true;
        }

    private void OnElapsedTimeForCapture(object sender, ElapsedEventArgs e)
    {
           
           // Escribe log
           // WriteLog("{0} ms elapsed."); 
           // Detener Timer para hacer el proceso
           TimerCaptura.Enabled = false;
           // Si está en el periodo indicado en la configuración
           if (DateTime.Now >=StartDateTime && DateTime.Now <=EndDateTime)
            { 
             if (String.IsNullOrWhiteSpace(PicWidth) && String.IsNullOrWhiteSpace(PicHeight) && String.IsNullOrWhiteSpace(Percent))
                { 
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
            // recargar configuración
            LoadProgramConfig();
            SetProgramConfig();

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
            // Establecer el intervalo
            Interval = Int32.Parse(cfgGlobal.Get("Interval"));
            StartDateTime = DateTime.Parse(cfgGlobal.Get("StartDateTime"));
            EndDateTime = DateTime.Parse(cfgGlobal.Get("EndDateTime"));
            PicWidth = cfgGlobal.Get("Width");
            PicHeight =cfgGlobal.Get("Height");
            Percent =cfgGlobal.Get("Percent");
        }

        #endregion

    private void SaveScreen(Bitmap bitmap)
        {
            // Crear directirio para guardar fotos si no existe
            var path = AppDomain.CurrentDomain.BaseDirectory + "files\\";
            if (!Directory.Exists(path))
                Directory.CreateDirectory (path);
            // Nombre de la foto usando tiempo unix
            // var horaunix = DateTime.Now.ToFileTimeUtc().ToString();
            var horaunix =  DateTimeOffset.UtcNow.ToUnixTimeSeconds();
            var fechahora = DateTime.Now.ToString("ddMM_HHmm");
            var name = "pic" + horaunix +"_" +fechahora +".jpg";
            var file = path + name ;
            bitmap.Save(file, System.Drawing.Imaging.ImageFormat.Jpeg);
        }

    #region Log 
    private void WriteLog(string logMessage, bool addTimeStamp = true)
    {
       
          try
    {
            
        var path = AppDomain.CurrentDomain.BaseDirectory;
        if (!Directory.Exists(path))
            Directory.CreateDirectory(path);

        var filePath = String.Format("{0}\\{1}_{2}.txt",
            path,
            System.AppDomain.CurrentDomain.FriendlyName,
            DateTime.Now.ToString("yyyyMMdd", CultureInfo.CurrentCulture)
            );

        if (addTimeStamp)
            logMessage = String.Format("[{0}] - {1}",
                DateTime.Now.ToString("HH:mm:ss", CultureInfo.CurrentCulture),
                logMessage);

        
       File.AppendAllText(filePath, logMessage);
     }
    catch(Exception e)
    {
      // do nothing
      WriteLog ("{0} exception writing log:" + e.ToString());
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
