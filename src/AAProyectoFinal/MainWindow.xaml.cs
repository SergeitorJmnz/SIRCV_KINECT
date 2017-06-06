// -----------------------------------------------------------------------
//Sergio Jímenez Salmerón y Gabriel Bonilla Ruiz
// -----------------------------------------------------------------------

namespace FaceTrackingBasics
{
    using System;
    using System.IO;
    using System.Windows;
    using System.Windows.Data;
    using System.Windows.Media;
    using System.Windows.Media.Imaging;
    using Microsoft.Kinect;
    using Microsoft.Kinect.Toolkit;

    using KinectMouseController;                                        //Raton
    using Coding4Fun.Kinect.Wpf;                                        //Raton

    using Microsoft.Speech.AudioFormat;                                 //Reconocimiento de voz
    using Microsoft.Speech.Recognition;                                 //Reconocimiento de voz

    using System.Runtime.InteropServices;
    //using System.Windows.Forms;

    public static class Globals                                         //Variables globales para uso conjunto con FaceTrackingViewer
                                                                        
    {
        public static bool cambio = false;                              //Variable booleana para cambio de modo de control del cursor.
                                                                        //Control con cabeza -> Cambio = false
                                                                        //Control con mano -> Cambio = true

        public static int CursorX=800;                                  //Posición inicial del cursor en eje X
        public static int CursorY=300;                                  //Posición inicial del cursor en eje Y
        public static int YAW;                                          //Variable para control del eje Y del ratón
        public static int PITCH;                                        //Variable para control del eje X del ratón
    }

    public partial class MainWindow : Window                            //Programa principal
    {
        /// Variables para uso de nuestro programa 
        private static readonly int Bgr32BytesPerPixel = (PixelFormats.Bgr32.BitsPerPixel + 7) / 8;
        private readonly KinectSensorChooser sensorChooser = new KinectSensorChooser();
        private WriteableBitmap colorImageWritableBitmap;
        private byte[] colorImageData;
        private ColorImageFormat currentColorImageFormat = ColorImageFormat.Undefined;


        KinectSensor myKinect;

        /// Speech recognition engine using audio data from Kinect.-----------
        private SpeechRecognitionEngine speechEngine;                   //Variable para reconocimiento de voz

        private bool click = false;                                     //Variable booleana para el click del ratón
        private bool foto = false;                                      //Variable booleana para hacer una captura de la cámara
        bool centraCursor = true;                                       //Variable booleana para centrar el raton en el modo cabeza

        ///
        /// Active Kinect sensor
        /// 
        private KinectSensor sensor;

        /// 
        /// USO DEL RATON
        /// 
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]  // Necesario para usar los eventos del ratón
        static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        public enum tipo
        {                              // Direcciones de memoria que contienen los eventos del ratón
            MOUSEEVENTF_LEFTDOWN = 0x00000002,         // Botón izquierdo abajo
            MOUSEEVENTF_LEFTUP = 0x00000004,           // Botón izquierdo arriba
            MOUSEEVENTF_RIGHTDOWN = 0x00000008,        // Botón derecho abajo
            MOUSEEVENTF_RIGHTUP = 0x00000010,          // Botón derecho arriba
            MOUSEEVENTF_MIDDLEDOWN = 0x00000020,       // Botón central abajo
            MOUSEEVENTF_MIDDLEUP = 0x00000040,         // Botón central arriba
            MOUSEEVENTF_WHEEL = 0x00000800             // Rueda ratón
        };

        public MainWindow()
        {
            InitializeComponent();

            if (Globals.cambio == false)
            {
                /// Activamos FaceTrackingViewer
                var faceTrackingViewerBinding = new Binding("Kinect") { Source = sensorChooser };
                faceTrackingViewer.SetBinding(FaceTrackingViewer.KinectProperty, faceTrackingViewerBinding);
            
            sensorChooser.KinectChanged += SensorChooserOnKinectChanged;
            sensorChooser.Start();
            }
        }

        ///
        /// VOZ
        ///

        // Funcion que se encarga del reconocimiento de voz
        private static RecognizerInfo GetKinectRecognizer()             
        {
            foreach (RecognizerInfo recognizer in SpeechRecognitionEngine.InstalledRecognizers())
            {
                string value;
                recognizer.AdditionalInfo.TryGetValue("Kinect", out value);
                if ("True".Equals(value, StringComparison.OrdinalIgnoreCase) && "es-ES".Equals(recognizer.Culture.Name, StringComparison.OrdinalIgnoreCase))
                {
                    return recognizer;
                }
            }

            return null;
        }

        ///
        /// VOZ
        /// 
        private void WindowLoaded(object sender, RoutedEventArgs e)
        {

            // Look through all sensors and start the first connected one.
            // This requires that a Kinect is connected at the time of app startup.
            // To make your app robust against plug/unplug, 
            // it is recommended to use KinectSensorChooser provided in Microsoft.Kinect.Toolkit (See components in Toolkit Browser).
            foreach (var potentialSensor in KinectSensor.KinectSensors)
            {
                if (potentialSensor.Status == KinectStatus.Connected)
                {
                    this.sensor = potentialSensor;
                    break;
                }
            }

            if (null != this.sensor)
            {
                // Turn on the skeleton stream to receive skeleton frames
                this.sensor.SkeletonStream.Enable();

                // Add an event handler to be called whenever there is new color frame data
                this.sensor.SkeletonFrameReady += this.SensorSkeletonFrameReady;

                myKinect = KinectSensor.KinectSensors[0];
                myKinect.ColorStream.Enable();
                myKinect.ColorFrameReady += myKinect_ColorFrameReady;
                // Start the sensor!
                try
                {
                    this.sensor.Start();
                }
                catch (IOException)
                {
                    this.sensor = null;
                }
            }
            //---------------------VOZ

            //Gramatica correspondiente a los comandos de voz
            RecognizerInfo ri = GetKinectRecognizer();
            if (null != ri)
            {

                this.speechEngine = new SpeechRecognitionEngine(ri.Id);

                // Use this code to create grammar programmatically rather than froma grammar file.
                var directions = new Choices();
                directions.Add(new SemanticResultValue("izquierdo", "CLICKIZQ"));
                directions.Add(new SemanticResultValue("derecho", "CLICKDER"));
                directions.Add(new SemanticResultValue("foto", "CAPTURA"));
                directions.Add(new SemanticResultValue("arriba", "ARRIBA"));
                directions.Add(new SemanticResultValue("abajo", "ABAJO"));
                directions.Add(new SemanticResultValue("cambio", "CAMBIO"));
                directions.Add(new SemanticResultValue("sube", "SUBE"));
                directions.Add(new SemanticResultValue("baja", "BAJA"));
                directions.Add(new SemanticResultValue("salir", "SALIR"));

                var gb = new GrammarBuilder { Culture = ri.Culture };
                gb.Append(directions);

                var g = new Grammar(gb);
                /* 
                ****************************************************************/

                // Create a grammar from grammar definition XML file.
                //using (var memoryStream = new MemoryStream(Encoding.ASCII.GetBytes(Properties.Resources.SpeechGrammar)))
                //{
                //    var g = new Grammar(memoryStream);
                speechEngine.LoadGrammar(g);
                //}

                speechEngine.SpeechRecognized += SpeechRecognized;
                //speechEngine.SpeechRecognitionRejected += SpeechRejected;

                // For long recognition sessions (a few hours or more), it may be beneficial to turn off adaptation of the acoustic model. 
                // This will prevent recognition accuracy from degrading over time.
                ////speechEngine.UpdateRecognizerSetting("AdaptationOn", 0);

                speechEngine.SetInputToAudioStream(
                    sensor.AudioSource.Start(), new SpeechAudioFormatInfo(EncodingFormat.Pcm, 16000, 16, 1, 32000, 2, null));
                speechEngine.RecognizeAsync(RecognizeMode.Multiple);
            }
            //VOZ
        }

        //------- VOZ
        //En esta funciona se realizan las acciones provenientes de los comandos de voz
        private void SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            // Speech utterance confidence below which we treat speech as if it hadn't been heard
            const double ConfidenceThreshold = 0.3;



            if (e.Result.Confidence >= ConfidenceThreshold)
            {
                switch (e.Result.Semantics.Value.ToString())
                {
                    case "CLICKIZQ":                                     //Realiza un click izquierdo de ratón
                        mouse_event((int)(tipo.MOUSEEVENTF_LEFTDOWN), 0, 0, 0, 0);
                        mouse_event((int)(tipo.MOUSEEVENTF_LEFTUP), 0, 0, 0, 0);
                        break;

                    case "CLICKDER":                                     //Realiza un click derecho de ratón
                        mouse_event((int)(tipo.MOUSEEVENTF_RIGHTDOWN), 0, 0, 0, 0);
                        mouse_event((int)(tipo.MOUSEEVENTF_RIGHTUP), 0, 0, 0, 0);
                        break;

                    case "SUBE":
                        mouse_event((int)tipo.MOUSEEVENTF_WHEEL, 0, 0, 300, 0); // 300 es el valor que se desplaza al hacer scroll
                        break;

                    case "BAJA":
                        mouse_event((int)tipo.MOUSEEVENTF_WHEEL, 0, 0, -300, 0); // -300 porque es hacia abajo
                        break;

                    case "CAPTURA":                                     //Realiza una captura de la cámara
                        this.kinectVideo1.Visibility = Visibility.Visible;
                        foto = true;
                        break;

                    case "ARRIBA":                                      //Sube el sensor
                        this.sensor.ElevationAngle += 3;
                        break;

                    case "ABAJO":                                       //Baja el sensor
                        this.sensor.ElevationAngle -= 3;
                        break;
                    case "SALIR":
                        Close();
                        break;

                    case "CAMBIO":                                      //Cambia entre modos
                        if (Globals.cambio == false)
                        {
                            Globals.cambio = true;
                            centraCursor = true;
                        }
                        else
                        {
                            Globals.cambio = false;
                        }
                        break;
                }
            }
        }//voz

        void myKinect_ColorFrameReady(object sender, ColorImageFrameReadyEventArgs e)
        {
            using (ColorImageFrame colorFrame = e.OpenColorImageFrame())
            {
                if (colorFrame == null) return;
                byte[] colorData = new byte[colorFrame.PixelDataLength];
                colorFrame.CopyPixelDataTo(colorData);
                //camara que siempre se muestra
                kinectVideo.Source = BitmapSource.Create(
                        colorFrame.Width, colorFrame.Height, 96, 96, PixelFormats.Bgr32, null, colorData, colorFrame.Width * colorFrame.BytesPerPixel
                                            );
                if (foto)//si se activa el comando de voz correspondiente se realiza una captura
                {
                    kinectVideo1.Source = BitmapSource.Create(
                     colorFrame.Width, colorFrame.Height,
                     96, 96,
                     PixelFormats.Bgr32,
                     null,
                     colorData,
                     colorFrame.Width * colorFrame.BytesPerPixel
                     );
                    foto = false;
                }
            }
        }

 
        /// Event handler for Kinect sensor's SkeletonFrameReady event llama a la funcion raton(la encargada del movimiento).
        private void SensorSkeletonFrameReady(object sender, SkeletonFrameReadyEventArgs e)
        {
            Skeleton[] skeletons = new Skeleton[0];

            using (SkeletonFrame skeletonFrame = e.OpenSkeletonFrame())
            {
                if (skeletonFrame != null)
                {
                    skeletons = new Skeleton[skeletonFrame.SkeletonArrayLength];
                    skeletonFrame.CopySkeletonDataTo(skeletons);
                }
            }

            foreach (Skeleton skel in skeletons)
            {
                if (skel.TrackingState == SkeletonTrackingState.Tracked)
                {
                    this.raton(skel);//se realiza la llamada a raton
                }
            }
        }


        private void raton(Skeleton skeleton) // Movimiento del raton y click.
        {
            // Render Joints
            foreach (Joint joint in skeleton.Joints)
            {
                if (joint.JointType == JointType.HandRight)
                {
                    if (Globals.cambio == true)
                    {
                        Joint scaledRight = skeleton.Joints[JointType.HandRight].ScaleTo(1600, 900);
                        Globals.CursorX = (int)scaledRight.Position.X;//tomamos la posicion X de la mano derecha
                        Globals.CursorY = (int)scaledRight.Position.Y;//tomamos la posicion Y de la mano derecha
                        //llamamos a la funcion con los parametros de la posicion y el booleano click
                        KinectMouseController.KinectMouseMethods.SendMouseInput((Globals.CursorX - 300), (Globals.CursorY - 150), 1000, 350, click);//se ha configurado segun los parametros del ordenador en el que ha sido desarrollado. Debería ser 1600x900
                    }
                    if (Globals.cambio == false)
                    {
                        if (centraCursor == true) //lo centramos
                        {
                            Globals.CursorX = 800;
                            Globals.CursorY = 300;
                            centraCursor = false;
                        }
                        if (Globals.YAW == 1) //izquierda
                        {
                            if (Globals.CursorX > 10)
                            Globals.CursorX -= 10;
                        }
                        if (Globals.YAW == -1) //derecha
                        {
                            if (Globals.CursorX < 1550)
                            Globals.CursorX += 10;
                        }
                        if (Globals.PITCH == 1) //arriba
                        {
                            if (Globals.CursorY > 10)
                            Globals.CursorY -= 10;
                        }
                        if (Globals.PITCH == -1) //abajo
                        {
                            if (Globals.CursorY < 875)
                            Globals.CursorY += 10;
                        }
                        KinectMouseController.KinectMouseMethods.SendMouseInput(Globals.CursorX, Globals.CursorY, 1600, 900, click);//llamamos a la funcion con los parametros de la posicion y el booleano click
                    }
                    click = false;//una vez realizado el click se pone a falso
                    
                }
            }
        }



        private void SensorChooserOnKinectChanged(object sender, KinectChangedEventArgs kinectChangedEventArgs) //trackeo para el facetracking con otras variables del sensor
        {
                KinectSensor oldSensor = kinectChangedEventArgs.OldSensor;
                KinectSensor newSensor = kinectChangedEventArgs.NewSensor;

                if (oldSensor != null)
                {
                    oldSensor.AllFramesReady -= KinectSensorOnAllFramesReady;
                    oldSensor.ColorStream.Disable();
                    oldSensor.DepthStream.Disable();
                    oldSensor.DepthStream.Range = DepthRange.Default;
                    oldSensor.SkeletonStream.Disable();
                    oldSensor.SkeletonStream.EnableTrackingInNearRange = false;
                    oldSensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Default;
                }

                if (newSensor != null)
                {
                    try
                    {
                        newSensor.ColorStream.Enable(ColorImageFormat.RgbResolution640x480Fps30);
                        newSensor.DepthStream.Enable(DepthImageFormat.Resolution320x240Fps30);
                        try
                        {
                            // This will throw on non Kinect For Windows devices.
                            newSensor.DepthStream.Range = DepthRange.Near;
                            newSensor.SkeletonStream.EnableTrackingInNearRange = true; 
                        }
                        catch (InvalidOperationException)
                        {
                            newSensor.DepthStream.Range = DepthRange.Default;
                            newSensor.SkeletonStream.EnableTrackingInNearRange = false;
                        }

                        newSensor.SkeletonStream.TrackingMode = SkeletonTrackingMode.Seated; 
                        newSensor.SkeletonStream.Enable();
                        newSensor.AllFramesReady += KinectSensorOnAllFramesReady;
                    }
                    catch (InvalidOperationException)
                    {
                        // This exception can be thrown when we are trying to
                        // enable streams on a device that has gone away.  This
                        // can occur, say, in app shutdown scenarios when the sensor
                        // goes away between the time it changed status and the
                        // time we get the sensor changed notification.
                        //
                        // Behavior here is to just eat the exception and assume
                        // another notification will come along if a sensor
                        // comes back.
                    }
                }
        }

        private void WindowClosed(object sender, EventArgs e)//se para el sensor 
        {
            sensorChooser.Stop();
            faceTrackingViewer.Dispose();
        }

        //funcion que muestra el trackeo de la cara en otra cámara RGB
        private void KinectSensorOnAllFramesReady(object sender, AllFramesReadyEventArgs allFramesReadyEventArgs)
        {
            if (Globals.cambio == false)
            {
                this.faceTrackingViewer.Visibility = Visibility.Visible;//no se esconde el track de la cara
                this.ColorImage.Visibility = Visibility.Visible;//no se esconde la camara de la cara
                using (var colorImageFrame = allFramesReadyEventArgs.OpenColorImageFrame())
                {
                    if (colorImageFrame == null)
                    {
                        return;
                    }

                    // Make a copy of the color frame for displaying.
                    var haveNewFormat = this.currentColorImageFormat != colorImageFrame.Format;
                    if (haveNewFormat)
                    {
                        this.currentColorImageFormat = colorImageFrame.Format;
                        this.colorImageData = new byte[colorImageFrame.PixelDataLength];
                        this.colorImageWritableBitmap = new WriteableBitmap(
                            colorImageFrame.Width, colorImageFrame.Height, 96, 96, PixelFormats.Bgr32, null);
                        ColorImage.Source = this.colorImageWritableBitmap;
                    }

                    colorImageFrame.CopyPixelDataTo(this.colorImageData);
                    this.colorImageWritableBitmap.WritePixels(
                        new Int32Rect(0, 0, colorImageFrame.Width, colorImageFrame.Height),
                        this.colorImageData,
                        colorImageFrame.Width * Bgr32BytesPerPixel,
                        0);
                }
            }
            else{
                this.faceTrackingViewer.Visibility = Visibility.Hidden;//se esconde el track de la cara
                this.ColorImage.Visibility = Visibility.Hidden;//se esconde la camara de la cara
            }
        }
    }
}
