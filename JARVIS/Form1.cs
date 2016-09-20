using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Threading;
using System.Xml.Linq;
using System.Xml;
using System.IO;
using System.Runtime.InteropServices;
using System.Net.Mail;
using System.Diagnostics;
using System.Net;
using System.Configuration;
using JARVIS;
using Newtonsoft.Json;
using System.Timers;
using Newtonsoft.Json.Linq;
using ChatterBotAPI;
using CUETools.Codecs;
using CUETools.Codecs.FLAKE;
using NAudio.Wave;

namespace JARVISV2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class Form1 : Form
    {
        //RootObject and Datum are required to capture google speech Api's response
        public class Rootobject
        {
            [JsonProperty("result")]
            public List<Datum> data { get; set; }

            [JsonProperty("result_index")]
            public int result_index;
        }

        public class Datum
        {
            [JsonProperty("alternative")]
            public Dictionary<string, string>[] alternative { get; set; }

            [JsonProperty("final")]
            public string final;

        }


        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", EntryPoint = "SendMessage", SetLastError = true)]
        static extern IntPtr SendMessage(IntPtr hWnd, Int32 Msg, IntPtr wParam, IntPtr lParam);
        const int WM_COMMAND = 0x111;
        const int MIN_ALL = 419;
        const int MIN_ALL_UNDO = 416;

        SpeechRecognitionEngine _recognizer = new SpeechRecognitionEngine();
        SpeechRecognitionEngine _recognizer2 = new SpeechRecognitionEngine();

        static SpeechSynthesizer JARVIS = new SpeechSynthesizer();
        String Temperature;
        String Condition;
        String Humidity;
        String WinSpeed;
        String TFCond;
        String TFHigh;
        String TFLow;
        String WOEID = "2295420"; //<<<<-------- CHANGE THE WOEID CODE HERE
        String Town;
        String QEvent;
        DateTime now = DateTime.Now;
        String userName = Environment.UserName;
        Random rnd = new Random();
        int timer = 11;
        int count = 1;
        List<string> listOfWords;
        string[] socialNetworkingsites = { "http://www.facebook.com", "http://twitter.com", "http://youtube.com" };
        string[] knowledgeSites = { "http://quora.com" };
        string[] jobSites = { "http://linkedin.com", "http://monster.com", "http://naukri.com" };

        bool isTalkingModeOn = false;
        private static System.Timers.Timer aTimer;
        static string currentSpeechWavFilename = "recordings\\wav\\currentSpeechwav.wav";
        static string currentSpeechWav16kFilename = "recordings\\wav\\currentSpeechwav16k.wav";
        static string currentSpeechFlacFilename = "recordings\\flac\\currentSpeechflac.flac";
        static string googleSpeechApiURL = "https://www.google.com/speech-api/v2/recognize?output=json&lang=en-us&";
        static string googleSpeechApiKey = "key=AIzaSyBkSa30ExjnF9219rZqMyIfH94JrgTu2iY";

        public Form1()
        {
            InitializeComponent();
            _recognizer.SetInputToDefaultAudioDevice();
            _recognizer.LoadGrammar(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"Commands.txt")))));
            
            listOfWords = new List<string>();
            
            _recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(_recognizer_SpeechRecognized);
            _recognizer.RecognizeAsync(RecognizeMode.Multiple);

            _recognizer2.SetInputToDefaultAudioDevice();
            _recognizer2.LoadGrammar(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines("Commands2.txt")))));
            _recognizer2.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(_recognizer_SpeechRecognized2);
            _recognizer2.RecognizeAsync(RecognizeMode.Multiple);
        }


        [DllImport("winmm.dll", EntryPoint = "mciSendStringA", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        private static extern int record(string lpstrCommand, string lpstrReturnString, int uReturnLength, int hwndCallback);

        void _recognizer_SpeechRecognized2(object sender, SpeechRecognizedEventArgs e)
        {
            string speech = e.Result.Text;
            if (isTalkingModeOn)
            {
                if (speech.Contains("Dear Friday"))
                {
                    //Record,send to google,Chatbot and make Jarvis speak
                    JARVIS.Speak("Tin Tin");
                    record("open new Type waveaudio Alias recsound", "", 0, 0);
                    record("record recsound", "", 0, 0);
                    System.Threading.Thread.Sleep(7000);
                    record("save recsound " + currentSpeechWavFilename, "", 0, 0);
                    record("close recsound", "", 0, 0);
                    JARVIS.Speak("Tin Tin");
                    

                    using (var reader = new WaveFileReader(currentSpeechWavFilename))
                    {
                        var newFormat = new WaveFormat(16000, 16, 1);
                        using (var conversionStream = new WaveFormatConversionStream(newFormat, reader))
                        {
                            WaveFileWriter.CreateWaveFile(currentSpeechWav16kFilename, conversionStream);
                        }
                    }

                    try
                    {
                        File.Delete(currentSpeechFlacFilename);
                    }
                    catch (Exception ex)
                    {
                        //handle it
                    }


                    using (var outputflacstream = File.Create(currentSpeechFlacFilename))
                    {
                        var inputwavstream = new FileStream(currentSpeechWav16kFilename, FileMode.Open);

                        ConvertToFlac(inputwavstream, outputflacstream);

                    }

                    string transcript = "";
                    try
                    {
                        //Reading the flac audio file into memory stream 
                        FileStream fileStream = File.OpenRead(currentSpeechFlacFilename);
                        MemoryStream memoryStream = new MemoryStream();
                        memoryStream.SetLength(fileStream.Length);
                        fileStream.Read(memoryStream.GetBuffer(), 0, (int)fileStream.Length);
                        byte[] BA_AudioFile = memoryStream.GetBuffer();
                        fileStream.Close();

                        //Creating a Http request to google speech API
                        HttpWebRequest _GoogleSpeechToText = null;
                        _GoogleSpeechToText =
                                    (HttpWebRequest)HttpWebRequest.Create(googleSpeechApiURL + googleSpeechApiKey
                                        );
                        _GoogleSpeechToText.Credentials = CredentialCache.DefaultCredentials;
                        _GoogleSpeechToText.Method = "POST";
                        _GoogleSpeechToText.ContentType = "audio/x-flac; rate=16000";
                        _GoogleSpeechToText.ContentLength = BA_AudioFile.Length;

                        //Write the flac audio file to the HttpWebrequest stream
                        Stream stream = _GoogleSpeechToText.GetRequestStream();
                        stream.Write(BA_AudioFile, 0, BA_AudioFile.Length);
                        stream.Close();

                        //Get the response from google
                        HttpWebResponse HWR_Response = (HttpWebResponse)_GoogleSpeechToText.GetResponse();
                        if (HWR_Response.StatusCode == HttpStatusCode.OK)
                        {
                            StreamReader SR_Response = new StreamReader(HWR_Response.GetResponseStream());

                            //Get the JSON response
                            var rawJson = SR_Response.ReadToEnd();
                            rawJson = rawJson.Replace("{\"result\":[]}\n", "");
                            rawJson = rawJson.Replace("\\", "");
                            var json = JObject.Parse(rawJson);  //Turns the raw string into a key value lookup
                            string jsonString = json.ToString();

                            //pushing the JSON data into Rootobject
                            Rootobject item = JsonConvert.DeserializeObject<Rootobject>(json.ToString());
                            List<Datum> datumList = item.data;
                            Datum datumsingle = datumList.FirstOrDefault();

                            //Final speech as a string
                            transcript = datumsingle.alternative.FirstOrDefault().ElementAt(0).Value;
                        }

                        //Use the chatterbot API to get the reply for the user
                        ChatterBotFactory factory = new ChatterBotFactory();
                        ChatterBot pandorabot = factory.Create(ChatterBotType.PANDORABOTS, "b0dafd24ee35a477");
                        ChatterBotSession pandorabotsession = pandorabot.CreateSession();

                       
                        string s = pandorabotsession.Think(transcript);
                        

                        //Make Jarvis speak
                        JARVIS.SetOutputToDefaultAudioDevice();
                        JARVIS.Speak(s);

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                }
            }
        }


        private static void ConvertToFlac(Stream sourceStream, Stream destinationStream)
        {
            var audioSource = new WAVReader(null, sourceStream);
            try
            {
                if (audioSource.PCM.SampleRate != 16000)
                {
                    throw new InvalidOperationException("Incorrect frequency - WAV file must be at 16 KHz.");
                }
                var buff = new CUETools.Codecs.AudioBuffer(audioSource, 0x10000);
                var flakeWriter = new FlakeWriter(null, destinationStream, audioSource.PCM);
                flakeWriter.CompressionLevel = 8;
                while (audioSource.Read(buff, -1) != 0)
                {
                    flakeWriter.Write(buff);
                }
                flakeWriter.Close();
            }
            finally
            {
                audioSource.Close();
            }
        }


        void _recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string time = "The time is " + now.TimeOfDay.Hours +" " +now.TimeOfDay.Minutes +" Hours";
            RecognizedAudio rad = e.Result.Audio;
            ProcessStartInfo sInfo;

            int ranNum;
            string speech = e.Result.Text;
            string speech1 = speech.ToLower();
            

            if (speech.Contains("Activate talking mode"))
            {
                isTalkingModeOn = true;
                JARVIS.Speak("Talking mode Activated. You can talk into the microphone.");
            }
            else if (speech.Contains("Deactivate talking mode"))
            {
                isTalkingModeOn = false;
                JARVIS.Speak("Talking mode Deactivated.");
            }


            if (!isTalkingModeOn)
            {
                switch (speech)
                {
                    //GREETINGS
                    case "Hey":
                    case "Hey Jarvis":
                    case "Hello":
                    case "Hello Jarvis":
                    case "Hi":
                    case "Hi Jarvis":
                        if (now.Hour >= 5 && now.Hour < 12)
                        { JARVIS.Speak("Goodmorning " + userName); }
                        if (now.Hour >= 12 && now.Hour < 18)
                        { JARVIS.Speak("Good afternoon " + userName); }
                        if (now.Hour >= 18 && now.Hour < 24)
                        { JARVIS.Speak("Good evening " + userName); }
                        if (now.Hour < 5)
                        { JARVIS.Speak("Hello, it is getting late " + userName); }
                        break;

                    //Bye
                    case "Bye":
                    case "Bye Jarvis":
                    case "Goodbye":
                    case "Goodbye Jarvis":
                    case "Close":
                    case "Close Jarvis":
                        JARVIS.Speak("Farewell");
                        Close();
                        break;


                    //Call Jarvis
                    case "Jarvis":
                        ranNum = rnd.Next(1, 4);
                        if (ranNum == 1) { QEvent = ""; JARVIS.Speak("Yes sir"); }
                        else if (ranNum == 2) { QEvent = ""; JARVIS.Speak("Yes?"); }
                        else if (ranNum == 3) { QEvent = ""; JARVIS.Speak("Yes Sir. I'm listening"); }
                        break;

                    //Date,Time, Weather etc.
                    case "What time is it":
                        JARVIS.Speak(time);
                        break;
                    case "What day is it":
                        JARVIS.Speak(DateTime.Today.ToString("dddd"));
                        break;
                    case "What's the date":
                    case "What's todays date":
                        JARVIS.Speak(DateTime.Today.ToString("dd-MM-yyyy"));
                        break;
                    case "How's the weather":
                    case "What's the weather like":
                    case "What's it like outside":
                    case "What's the temperature outside":
                        GetWeatherJSON();
                        break;
                   
                    //WINDOW COMMANDS
                    case "Switch Window":
                        SendKeys.Send("%{TAB " + count + "}");
                        count += 1;
                        break;
                    case "Hide All Windows":
                        JARVIS.Speak("Yes Sir , your wish is my command");
                        IntPtr lHwnd1 = FindWindow("Shell_TrayWnd", null);
                        SendMessage(lHwnd1, WM_COMMAND, (IntPtr)MIN_ALL, IntPtr.Zero);
                        break;
                    case "Show All Windows":
                        JARVIS.Speak("Yes Sir , your wish is my command. Here you go!");
                        IntPtr lHwnd = FindWindow("Shell_TrayWnd", null);
                        SendMessage(lHwnd, WM_COMMAND, (IntPtr)MIN_ALL_UNDO, IntPtr.Zero);
                        break;
                    case "Reset":
                        count = 1;
                        int timer = 11;
                        lblTimer.Visible = false;
                        ShutdownTimer.Enabled = false;
                        lstCommands.Visible = false;
                        break;
                    case "Out of the way":
                        if (WindowState == FormWindowState.Normal)
                        {
                            WindowState = FormWindowState.Minimized;
                            JARVIS.Speak("My apologies");
                        }
                        break;
                    case "Come back":
                        if (WindowState == FormWindowState.Minimized)
                        {
                            JARVIS.Speak("Alright");
                            WindowState = FormWindowState.Normal;
                        }
                        break;
                    case "Show commands":
                        string[] commands = File.ReadAllLines("Commands.txt");
                        JARVIS.Speak("Yes Sir. Here it is");
                        lstCommands.Items.Clear();
                        lstCommands.SelectionMode = SelectionMode.None;
                        lstCommands.Visible = true;
                        foreach (string command in commands)
                        {
                            lstCommands.Items.Add(command);
                        }
                        break;
                    case "Hide commands":
                        lstCommands.Visible = false;
                        break;


                    //COMPUTER SHUTDOWN RESTART LOG OFF
                    case "Shutdown":
                        if (ShutdownTimer.Enabled == false)
                        {
                            QEvent = "shutdown";
                            JARVIS.Speak("I will shutdown shortly");
                            lblTimer.Visible = true;
                            ShutdownTimer.Enabled = true;
                        }
                        break;
                    case "Log off":
                        if (ShutdownTimer.Enabled == false)
                        {
                            QEvent = "logoff";
                            JARVIS.Speak("Logging off");
                            lblTimer.Visible = true;
                            ShutdownTimer.Enabled = true;
                        }
                        break;
                    case "Restart":
                        if (ShutdownTimer.Enabled == false)
                        {
                            QEvent = "restart";
                            JARVIS.Speak("I'll be back shortly");
                            lblTimer.Visible = true;
                            ShutdownTimer.Enabled = true;
                        }
                        break;
                    case "Abort":
                        if (ShutdownTimer.Enabled == true)
                        {
                            timer = 11;
                            lblTimer.Text = timer.ToString();
                            ShutdownTimer.Enabled = false;
                            lblTimer.Visible = false;
                        }
                        break;
                    case "JARVIS I want Roma":
                        Speak(VoiceGender.Female);
                        break;
                    case "Roma I want JARVIS":
                        Speak(VoiceGender.Male);
                        break;
                    
                    
                    //Browser Websites
                    case "I'm getting bored":
                        JARVIS.Speak("Sir, Let's do some social networking");
                        JARVIS.Speak("Here you go");
                        string site = socialNetworkingsites[new Random().Next(0, socialNetworkingsites.Length)];
                        sInfo = new ProcessStartInfo(site);
                        Process.Start(sInfo);
                        break;

                    case "I want some knowledge":
                        JARVIS.Speak("Here you go sir, read this");
                        site = knowledgeSites[new Random().Next(0, knowledgeSites.Length)];
                        sInfo = new ProcessStartInfo(site);
                        Process.Start(sInfo);
                        break;

                    case "I dont like my job":
                        JARVIS.Speak("Sir, Don't be worried. There are tons of jobs out there! Let me take you there.");
                        site = jobSites[new Random().Next(0, jobSites.Length)];
                         sInfo = new ProcessStartInfo(site);
                        Process.Start(sInfo);
                        break;
                }
            }
            
        }


        private string get_web_content(string url)
        {
            string output = "";
            try
            {
                Uri uri = new Uri(url);
                HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(uri);
                request.Method = WebRequestMethods.Http.Get;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                StreamReader reader = new StreamReader(response.GetResponseStream());
                output = reader.ReadToEnd();
                response.Close();
            }
            catch
            {
                output = "";
            }

            return output;
        }

        private void GetWeatherJSON()
        {
            string url = "http://api.openweathermap.org/data/2.5/weather?q={0}&appid={1}";
            string city = "Bangalore,IN";
            string apikey = "bf1e7df813732a75613bef23aeeea743";
            url = string.Format(url, city, apikey);
            string json = get_web_content(url);
            if (!string.IsNullOrEmpty(json))
            {
                WeatherRootObject weatherReport = JsonConvert.DeserializeObject<WeatherRootObject>(json);
                JARVIS.Speak("The weather in " + weatherReport.name + " is " + weatherReport.weather.FirstOrDefault().description + " at " + weatherReport.main.temp + " degrees. There is a humidity of " + weatherReport.main.humidity + " and a windspeed of " + weatherReport.wind.speed + " miles per hour");
            }
            else
            {
                JARVIS.Speak("I'm sorry sir. It is difficult for me to predict the weather due to poor internet connection");
            }
           
        }

       
        private void ShutdownTimer_Tick(object sender, EventArgs e)
        {
            if (timer == 0)
            {
                lblTimer.Visible = false;
                ComputerTermination();
                ShutdownTimer.Enabled = false;
            }
            else
            {
                timer = timer - 1;
                lblTimer.Text = timer.ToString();
            }
        }
        private void ComputerTermination()
        {
            switch (QEvent)
            {
                case "shutdown":
                    System.Diagnostics.Process.Start("shutdown", "-s");
                    break;
                case "logoff":
                    System.Diagnostics.Process.Start("shutdown", "-l");
                    break;
                case "restart":
                    System.Diagnostics.Process.Start("shutdown", "-r");
                    break;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void Speak(VoiceGender voiceGender)
        {
            var genderVoices = JARVIS.GetInstalledVoices().Where(arg => arg.VoiceInfo.Gender == voiceGender).ToList();
            var firstVoice = genderVoices.FirstOrDefault();
            if (firstVoice == null)
                return;
            JARVIS.SelectVoice(firstVoice.VoiceInfo.Name);
            JARVIS.Speak("How are you today?");
        }
    }
}