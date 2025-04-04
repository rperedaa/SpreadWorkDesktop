using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Timers;
using System.Drawing.Imaging;
using System.IO;
using System.Globalization;
using System.Runtime.InteropServices;

namespace SpreadWorkDesktop
{
    public partial class Form1 : Form
    {
        System.Timers.Timer Timer = new System.Timers.Timer();
        int Interval = 10000;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {



            WriteLog("Servicio SpreadWorkService arrancado");
        // Ejecución períodica del servicio
        Timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
        Timer.Interval = Interval;
        Timer.Enabled = true;
        }

    private void OnElapsedTime(object sender, ElapsedEventArgs e)
    {
        Timer.Enabled = false;
        // Escribe log
        WriteLog("{0} ms elapsed.");
            // Captura imagen

     
           Size size = GetDisplayResolution();
         
        using (Bitmap bitmap = new Bitmap(size.Width, size.Height, PixelFormat.Format32bppArgb))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.CopyFromScreen(Point.Empty, Point.Empty, size);
            }
            bitmap.Save("test.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
        }
     



        // Inicializa Timer
        Timer.Enabled = true;
    }

    private void WriteLog(string logMessage, bool addTimeStamp = true)
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
