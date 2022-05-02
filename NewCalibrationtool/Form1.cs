using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

#if NET4
using System.Threading.Tasks; //code only for .NET 4
#endif
using System.Net.Sockets;
using System.Globalization;
using AC3R2NetUtil;
using System.IO.Ports;
using System.IO;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
//using range = Microsoft.Office.Interop.Excel.Range;
using System.Runtime.InteropServices;
//using ClosedXML.Excel;
//using System.Web.UI.DataVisualization.Charting;
//using Series = System.Windows.Forms.DataVisualization.Charting.Series;
//using seriesCollection = Microsoft.Office.Interop.Excel.SeriesCollection.seriesCollection;

//using System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
using ZedGraph;
using static System.Math;

namespace NewCalibrationtool
{
    public struct ComPortSettings   // COM port
    {
        public string PortName;
        public string BaudRate;
        public string Parity;
        public string DataBits;
        public string StopBits;
        public string Handshake;

    };
    public struct RS232CmdList
    {
        public static string cmdRemote = "REMOTEMODE";
        public static string outRemote = "SYST:REM\r\n";
        public static string cmdPowerOnOff = "SELONOFF";
        public static string outPowerON = "OUTP ON\r\n";
        public static string outPowerOFF = "OUTP OFF\r\n";
        public static string cmdsensVal = "SELSENSORVALUE";
        public static string outsensValPlat = "PLAT:STAN PT385A\r\n";
        public const string cmdSetTemp = "SELTEMP";
        public static string outPlat = "PLAT ";
        public static string outCRLR = "\r\n";
        public static string Version = ":SYSTem:";
        public const string requestTemp = "A?\r\n";
        public const string SetSensorValue = "A"; // Set Temperature add on the Temperature. 
        public const string SystemID = "*IDN?\r\n";

    }
    public enum REMOTE_COMD
    {
        REMOTEMODE,
        SYS_ID,
        SYSREMOTE,
        POWERON,
        POWEROFF,
        SENSORVAL,
        SENSORVALPLAT,
        SETTEMP,
        PLAT,
    };


    public partial class Calibrationtool : Form
    {
        private const char LF = '\n';
        private const char CR = '\r';
        delegate void SetTextCallback(System.Windows.Forms.Control ctr, string text);
        delegate void SetBgColorCallback(System.Windows.Forms.Control ctr, Color color);
        private delegate void SetTextDeleg(string text);
        const int RECV_BUF_LEN = 128;
        NetUtil ac3r2Net = new NetUtil();
        private byte[] m_recv_buf = new byte[RECV_BUF_LEN];
        byte[] m_send_buf = new byte[RECV_BUF_LEN];
        REMOTE_COMD m_cmd = REMOTE_COMD.SYSREMOTE;
        string m_Response = string.Empty;
        bool m_Initialized = false;
        string m_messages = string.Empty;
        private bool timerEnablePlot = false;
        List<Double> tempvalues = new List<double>();
        List<Double> diffvalues = new List<double>();
        List<double> Avglist = new List<double>(10);
        int list = 0;

        //-------------------------XL sheet related -----------------------

        Microsoft.Office.Interop.Excel.Application xlapp = new Microsoft.Office.Interop.Excel.Application();

        Microsoft.Office.Interop.Excel.Application xlApp;
        Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
        Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;



        // Excel object references.
        private Excel.Application m_objExcel = null;
        private Excel.Workbooks m_objBooks = null;
        private Excel._Workbook m_objBook = null;
        private Excel.Sheets m_objSheets = null;
        private Excel._Worksheet m_objSheet = null;
        private Excel.Range m_objRange = null;
        private Excel.Range Range = null;
        private Excel.Font m_objFont = null;
        private Excel.QueryTables m_objQryTables = null;
        private Excel._QueryTable m_objQryTable = null;
        //----------------------------------------------------------

        SerialPort serialport;

        PointPairList pointpairlist = new PointPairList();
        PointPairList pointpairlist1 = new PointPairList();  
        
        PointPairList pointpairlist2 = new PointPairList();
        PointPairList pointpairlist3 = new PointPairList();
        PointPairList pointpairlist4 = new PointPairList();
        PointPairList pointpairlist5 = new PointPairList();

        ComPortSettings comportSettings;
        private string recv_buf = string.Empty;

        public Calibrationtool()
        {
            
            InitializeComponent();
            
            this.Text += Application.ProductVersion;
            this.comportSettings.PortName = string.Empty;
            this.InitializeComPortsList();
            graphPane = zedGraphControl1.GraphPane;
            Creategraph();
        }

//---------------------------------Only for graph settings---------------------------------------------------------------------------------------------------------------------
        GraphPane graphPane;
        private void Creategraph()
        {
           
            LineItem lineItem = graphPane.AddCurve("0 Cycle", pointpairlist, Color.Red, SymbolType.Circle);
            LineItem lineItem1 = graphPane.AddCurve("1 Cycle", pointpairlist1, Color.Blue, SymbolType.XCross);
            LineItem lineItem2 = graphPane.AddCurve("2 Cycle", pointpairlist2, Color.Green, SymbolType.Triangle); 
            LineItem lineItem3 = graphPane.AddCurve("3 Cycle", pointpairlist3, Color.Yellow, SymbolType.Square); 
            LineItem lineItem4 = graphPane.AddCurve("4 Cycle", pointpairlist4, Color.Orange, SymbolType.Plus);
            LineItem lineItem5 = graphPane.AddCurve("5 Cycle", pointpairlist5, Color.Violet, SymbolType.Star);
            


            zedGraphControl1.AxisChange();
            

            graphPane.YAxis.Title.Text = "Difference of temp. of controller and simulator";
            graphPane.XAxis.Title.Text = "Temperature";
            graphPane.Title.Text = " Diff. temperature VS Temperature";
            zedGraphControl1.GraphPane.XAxis.MajorGrid.IsVisible = true;
            zedGraphControl1.GraphPane.YAxis.MajorGrid.IsVisible = true;
            zedGraphControl1.GraphPane.YAxis.Scale.FontSpec.FontColor = Color.Red;
            zedGraphControl1.GraphPane.YAxis.Title.FontSpec.FontColor = Color.Red;
            zedGraphControl1.GraphPane.XAxis.Scale.FontSpec.FontColor = Color.Red;
            zedGraphControl1.GraphPane.XAxis.Title.FontSpec.FontColor = Color.Red;



        }
    
    //--------------------------xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx--------------------------    
        private void label8_Click(object sender, EventArgs e)
        {

        }
        // ------------------------For initialization of comportlist------------------------------------------------------------------------------
        private void InitializeComPortsList()
        {
            foreach (var port in SerialPort.GetPortNames())
            {

                this.CBoxComport.Items.Add(port);
            }
            if (CBoxComport.Items.Count > 0)
            {
                serialport = new SerialPort();
                serialport.DataReceived += new SerialDataReceivedEventHandler(this.port_DataReceived);
                this.MeasConnectBtn.Enabled = true;
                if (CBoxComport.Items.Count == 1)
                    this.CBoxComport.SelectedIndex = 0;
                this.CBoxBaudrate.SelectedIndex = 5;
                this.CBoxStopbits.SelectedIndex = 0;
                this.CBoxParitybits.SelectedIndex = 0;
                this.CBoxDatabits.SelectedIndex = 1;
                this.comportSettings.PortName = this.CBoxComport.Text;
                this.comportSettings.BaudRate = this.CBoxBaudrate.Text;
                this.comportSettings.DataBits = this.CBoxDatabits.Text;
                this.comportSettings.Parity = this.CBoxParitybits.Text;
                this.comportSettings.StopBits = this.CBoxStopbits.Text;
                this.comportSettings.Handshake = "None";
            }

        }

        //-------------------Below from here is for when connect button is pressed for communicating with simulator ------------------------------------------------------------------------
        #region When connect button is pressed for communicating with simulator 
        private void MeasConnectBtn_Click_1(object sender, EventArgs e)
        {
            if (serialport != null)
            {
                if (MeasConnectBtn.Text == "Connect")
                {
                    serialport.PortName = this.comportSettings.PortName;
                    serialport.BaudRate = int.Parse(this.comportSettings.BaudRate);
                    serialport.DataBits = int.Parse(this.comportSettings.DataBits);
                    serialport.Parity = (Parity)Enum.Parse(typeof(Parity), this.comportSettings.Parity);
                    serialport.StopBits = (StopBits)Enum.Parse(typeof(StopBits), this.comportSettings.StopBits);
                    serialport.Handshake = (Handshake)Enum.Parse(typeof(Handshake), this.comportSettings.Handshake);
                    if (!serialport.IsOpen)
                    {
                        try
                        {
                            serialport.Open();
                            if (serialport.IsOpen)
                            {
                                MeasConnectBtn.Text = "Disconnect";
                                serialport.DataReceived += new SerialDataReceivedEventHandler(this.port_DataReceived);
                                serialport.Handshake = Handshake.None;
                                m_cmd = REMOTE_COMD.SYSREMOTE;
                               
                              
                                SendRemoteCommand();
                                

                            }
                        }
                        catch (IOException ex)
                        {
                            MessageBox.Show(string.Format("The selected COM Port {0} is being used by another Application\n Please close the COM Port and try again. \nInner exception: {1}", comportSettings.PortName, ex.Message));
                        }
                        catch (UnauthorizedAccessException ex)
                        {
                            MessageBox.Show(string.Format("The selected COM Port {0} is being used by another Application\n Please the COM Port and try again.  \nInner exception: {1}", comportSettings.PortName, ex.Message));
                        }
                        catch (System.Exception se)
                        {
                            MessageBox.Show(string.Format("The selected COM Port {0} is being used by another Application\n Please the COM Port and try again.  \nInner exception: {1}", comportSettings.PortName, se.Message));

                        }
                    }
                    else
                    {
                        MessageBox.Show(string.Format("The selected COM Port {0} is being used by another Application\n Please close the COM Port and try again.}", comportSettings.PortName));

                    }
                }
                else
                {
                    this.DisConnectPort();
                }
            }

        }
        private void DisConnectPort()
        {
            try
            {
                this.ReadComTimer.Stop();
                this.TimeoutTimer.Enabled = false;
                if (serialport.IsOpen)
                {
                    serialport.Close();
                    this.MeasConnectBtn.Text = "Connect";
                }

            }
            catch (IOException)
            {
            }
        }
        #endregion
        //-------------------------------------ONLY FOR SETTEXT FUNCTION--------------------------------------------------------------
        private void SetText(System.Windows.Forms.Control ctr, string text)
        {
            // InvokeRequired required compares the thread ID of the 
            // calling thread to the thread ID of the creating thread. 
            // If these threads are different, it returns true. 
            try
            {
                if (ctr.InvokeRequired)
                {
                    SetTextCallback d = new SetTextCallback(SetText);
                    this.Invoke(d, new object[] { ctr, text });
                }
                else
                {
                    ctr.Text = text;
                }
            }
            catch (ObjectDisposedException) { }

        }

        //-----------------------------FOR CHANGE BACKGROUND COLOR FUNCTION-------------------------------------------------------------
        private void ChangBgColor(System.Windows.Forms.Control ctr, Color val)
        {
            // InvokeRequired required compares the thread ID of the 
            // calling thread to the thread ID of the creating thread. 
            // If these threads are different, it returns true. 
            try
            {
                if (ctr.InvokeRequired)
                {
                    SetBgColorCallback d = new SetBgColorCallback(ChangBgColor);
                    this.Invoke(d, new object[] { ctr, val });
                }
                else
                {
                    ctr.BackColor = val;
                }
            }
            catch (ObjectDisposedException) { }

        }
        private void IPAddrTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void ReadConfigurationData()
        {
            E_DUMP_PERSISTENT_CMDS cmd = E_DUMP_PERSISTENT_CMDS.E_DUMP_PTCALIB_PERSISTENT_STORAGE;

            byte[] buf = AC3R2NetUtil.CommandSet.GetDumpParamSetPersistentCmd(cmd);
            buf[buf.Length - 1] = CommandSet.CalcCRCSum(buf, buf.Length - 1);

            int count = ac3r2Net.TCPSendMessage(buf, buf.Length);
            if (count > 0 && buf[3] == 0x02)
            {
                ac3r2Net.GetBytesRead(m_recv_buf, count);
            
                UpdateTimer.Enabled = true;
            }
        }

        //------------------------------------------BELOW FROM HERE IS FOR CONNECT BUTTON FOR IP ADDRESS CONNECTION WITH CONTROLLER-------------------------------------------
        #region Connect button for connection with controller IP Address
        private void ConnectBtn_Click(object sender, EventArgs e)
        {
            if (ConnectBtn.Text.Equals("Connect"))
            {
                
                TcpClient client = new TcpClient();
                try
                {
                    if (this.ac3r2Net.Connect2Server(this.IPAddrTextBox.Text, this.PortTextBox.Text))
                    {
                        SetText(ConnectBtn, "Disconnect");
                        SetText(this.NetStatusTextBox, "Connected");
                        ChangBgColor(this.NetStatusTextBox, Color.Green);
                        ReadConfigurationData();

                    }
                }
                catch (AC3R2NetException ae)
                {
                    string m = ae.Message +
                                "Could not connect to the given Host.\r\n\r\n" +
                                    "Please make sure:\r\n" +
                                    " - The Host is on and your PC is connected to it\r\n\r\n" +
                                    " - Your PC has a valid IP-Address i.e. same Network as the Controller.\r\n" +
                                    "   e.g. if the controller has 172.30.240.xxx, your PC's IP-Address should be similar to 172.30.xxx.xxx\r\n" +
                                    "   Your can Enable the DHCP Server on the Controller to get a dynamic IP-Address.\r\n\r\n" +
                                    " - Your PC's Firewall is not blocking this application.\r\n" +
                                    "   If that is the case, add this application (Port 2002) to your Firewall rules";

                    MessageBox.Show("Information", m);
                }
            }
            else
            {

                this.ac3r2Net.DisconectFromServer();

                this.ConnectBtn.Text = "Connect";
                this.NetStatusTextBox.Text = "Offline";
                this.NetStatusTextBox.BackColor = Color.Red;
                this.UpdateTimer.Enabled = false;
            }
        }
        #endregion 
        string simulatorTemp1;
        double simtemp;
        private void InterpreteResponse()
        {
           
            TimeoutTimer.Stop();

            switch (m_cmd)
            {
                case REMOTE_COMD.SYSREMOTE:
                    m_cmd = REMOTE_COMD.SYS_ID;
                    break;
                case REMOTE_COMD.SYS_ID:
                    SetText(SysIDText, m_Response);
                    m_cmd = REMOTE_COMD.POWERON;
                    SendRemoteCommand();
                    m_cmd = REMOTE_COMD.SENSORVAL;
                    SendRemoteCommand();
                   
                    break;
                case REMOTE_COMD.PLAT:
                    break;
                case REMOTE_COMD.POWEROFF:
                    break;
                case REMOTE_COMD.POWERON:
                    break;
                case REMOTE_COMD.REMOTEMODE:
                    break;
                case REMOTE_COMD.SENSORVAL:
                    if (m_Initialized)
                    {
                        SetText(rectext, m_messages);
                        simulatorTemp1 = m_Response.ToString();                                     //<<<<<<<<<<<<<<<<-----------------Here I added this to read simulator tempreture .
                
                        simtemp = Double.Parse(simulatorTemp1.Replace('.', ','));                  //<<<<<<<<<<<<<<<<-----------------Here I added this to read simulator tempreture .



                    }
                    else
                    {
                        m_cmd = REMOTE_COMD.SETTEMP;
                        m_Initialized = true;
                        SendRemoteCommand();
                        m_cmd = REMOTE_COMD.SENSORVAL;
                        SendRemoteCommand();
                    }
                    break;
                case REMOTE_COMD.SENSORVALPLAT:
                    break;
                case REMOTE_COMD.SETTEMP:
                    break;
             

            }
            // SendRemoteCommand();
        }
        private void port_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            string buf = string.Empty;

            int bytesToRead = serialport.BytesToRead;
            for (int index = 0; index < bytesToRead; index++)
            {
                char ch = (char)serialport.ReadChar();
                if (ch != CR && ch != LF)
                {
                    recv_buf += ch;
                }
                else if (ch == LF)
                {
                    buf = recv_buf + "\r\n";
                    m_Response = buf;
                    InterpreteResponse();
                    m_messages += m_Response;
                    SetText(rectext, buf);
                    recv_buf = string.Empty;
                }

            }
        }
        #region Combobox for comport ,baudrate , databits,stopbits ,parity bits 
        private void CBoxComport_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.comportSettings.PortName = this.CBoxComport.Text;
        }

        private void CBoxBaudrate_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.comportSettings.BaudRate = this.CBoxBaudrate.Text;
        }

        private void CBoxDatabits_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.comportSettings.DataBits = this.CBoxDatabits.Text;
        }

        private void CBoxStopbits_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.comportSettings.StopBits = this.CBoxStopbits.Text;
        }

        private void CBoxParitybits_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.comportSettings.Parity = this.CBoxParitybits.Text;
        }
        #endregion
        //----------------------When we press the REMOTE button this happens---------------------------------------------------------
        #region when REMOTE button is pressed 
        private void init_Click(object sender, EventArgs e)
        {
            if (init.Text.Equals("REMOTE"))
            {

                this.init.Text = "LOCAL";
                m_cmd = REMOTE_COMD.SYSREMOTE;
                SendRemoteCommand();

            }
            else
            {
                Disconnectinit();
            }

        }
        //-------------------For sending the REMOTE command to the simulator ---------------------------------------------------------
        private void SendRemoteCommand()
        {
            string cmdstr = GetRemoteCommand(m_cmd);
            serialport.WriteLine(cmdstr);
            if (m_cmd == REMOTE_COMD.SYSREMOTE)
            {
                m_cmd = REMOTE_COMD.SYS_ID;
                cmdstr = GetRemoteCommand(m_cmd);
                serialport.WriteLine(cmdstr);

            }
            if (m_cmd == REMOTE_COMD.SYS_ID || m_cmd == REMOTE_COMD.SENSORVAL)
            {
                TimeoutTimer.Start();
            }

        }
        private string GetRemoteCommand(REMOTE_COMD rcmd)
        {
            string cmd = string.Empty;

            switch (rcmd)
            {
                case REMOTE_COMD.REMOTEMODE:
                    cmd = RS232CmdList.cmdRemote;
                    break;
                case REMOTE_COMD.SYS_ID:
                    cmd = RS232CmdList.SystemID;
                    break;
                case REMOTE_COMD.SYSREMOTE:
                    cmd = RS232CmdList.outRemote;
                    break;
                case REMOTE_COMD.PLAT:
                    cmd = RS232CmdList.outPlat;
                    break;
                case REMOTE_COMD.POWEROFF:
                    cmd = RS232CmdList.outPowerOFF;
                    break;
                case REMOTE_COMD.POWERON:
                    cmd = RS232CmdList.outPowerON;
                    break;
                case REMOTE_COMD.SETTEMP:
                    cmd = RS232CmdList.cmdSetTemp + ExternalTemp.Text + "\r\n";
                    break;
                case REMOTE_COMD.SENSORVAL:
                    cmd = RS232CmdList.requestTemp;
                    break;
                case REMOTE_COMD.SENSORVALPLAT:
                    cmd = RS232CmdList.outsensValPlat;
                    break;


            }
            return cmd;

        }

        private void Disconnectinit()
        {
            try
            {
                poweroff();
                this.init.Text = "Init On";
            }
            catch (IOException)
            {
            }
        }
        private void poweroff()
        {
            m_cmd = REMOTE_COMD.POWEROFF;
            SendRemoteCommand();
        }

        #endregion

        private void GetPT100OffsetScale()
        {
            m_send_buf.Initialize();
            byte[] sb;
            sb = CommandSet.GetDumpParamSetVolatileCmd(E_DUMP_VOLATILE_CMDS.E_DUMP_PTCALIB_VOLATILE_STORAGE);
            Array.Copy(sb, m_send_buf, 5);
            m_send_buf[5] = CommandSet.CalcCRCSum(m_send_buf, 5);
            int count = ac3r2Net.TCPSendMessage(m_send_buf, 6);
            if (count > 0)
            {
                ac3r2Net.GetBytesRead(m_recv_buf, count);
                UpdateOffsetScale();
            }


        }

        // -------------------For setting the floating points values in the linearity and offset textboxes--------------------------------------------------------------
        private void UpdateOffsetScale()
        {
            
            SetText(TxtBoxOffset0, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 5)));
            SetText(TxtBoxOffset1, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 9)));
            SetText(TxtBoxOffset2, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 13)));
            SetText(TxtBoxOffset3, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 17)));
            SetText(TxtBoxOffset4, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 21)));
            SetText(TxtBoxOffset5, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 25)));
            SetText(TxtBoxOffset6, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 29)));
            SetText(TxtBoxOffset7, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 33)));

            SetText(TxtBoxLiner0, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 37)));
            SetText(TxtBoxLiner1, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 41)));
            SetText(TxtBoxLiner2, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 45)));
            SetText(TxtBoxLiner3, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 49)));
            SetText(TxtBoxLiner4, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 53)));
            SetText(TxtBoxLiner5, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 57)));
            SetText(TxtBoxLiner6, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 61)));
            SetText(TxtBoxLiner7, string.Format("{0:f4}", BitConverter.ToSingle(m_recv_buf, 65)));

        }
        //------------------>>>>>>>>>>>>>>>>>----------------------"UPDATE TIMER STARTS" IN Below code --------------------------------<<<<<<<<<<<<<<<<<<<<<---------------
        int j = 0;
        double AVGTEMP;

        enum channelType
        {
            Channel1,
            Channel2,
            Channel3,
            Channel4,
            Channel5,
            Channel6,
            Channel7,
            Channel8,
            Undef,

        }
        channelType ChannelType = channelType.Undef;

        private void UpdateTimer_Tick_1(object sender, EventArgs e)
        {
            double dChuckTemp=0;
         

            if ((ac3r2Net != null) && (ac3r2Net.isConnected))
            {

                try
                {
                    GetOffsetTemp();
                    

                    GetPT100OffsetScale();
                    Chucktemp1 = Convert.ToDouble(GetChuckTemp("GET CPT\n"));
                    Chucktemp2 = Convert.ToDouble(GetChuckTemp("GET CPTE1\n"));
                    Chucktemp3= Convert.ToDouble(GetChuckTemp("GET CPTE2\n"));
                    Chucktemp4 = Convert.ToDouble(GetChuckTemp("GET CPTE3\n"));
                    Chucktemp5 = Convert.ToDouble(GetChuckTemp("GET CPTE4\n"));
                    Chucktemp6 = Convert.ToDouble(GetChuckTemp("GET CPTE5\n"));
                    Chucktemp7 = Convert.ToDouble(GetChuckTemp("GET CPTE6\n"));
                    Chucktemp8 = Convert.ToDouble(GetChuckTemp("GET CPTE7\n"));
             
                    switch (ChannelType)
                    {
                        case channelType.Channel1:
                            {
                          
                                dChuckTemp = Chucktemp1;
                            }
                            break;
                        case channelType.Channel2:
                            {
                                dChuckTemp = Chucktemp2;
                            }
                            break;
                        case channelType.Channel3:
                            {
                                dChuckTemp = Chucktemp3;
                           
                            }
                            break;
                        case channelType.Channel4:
                            {
                                dChuckTemp = Chucktemp4;
                              
                            }
                            break;
                        case channelType.Channel5:
                            {
                                dChuckTemp = Chucktemp5;
                                
                            }
                            break;
                        case channelType.Channel6:
                            {
                                dChuckTemp = Chucktemp6;
                              
                            }
                            break;
                        case channelType.Channel7:
                            {
                                dChuckTemp = Chucktemp7;
                             
                            }
                            break;
                        case channelType.Channel8:
                            {
                                dChuckTemp = Chucktemp8;
                               
                            }
                            break;
                    }
                    //dChuckTemp = Convert.ToDouble(temp1);
                    Averagefunction(dChuckTemp);

                  

                }
                catch (AC3R2NetException)
                {
                    Reconnect2Server();

                }
            }
        }

    

        Double Chucktemp1;
        Double Chucktemp2;
        Double Chucktemp3;
        Double Chucktemp4;
        Double Chucktemp5;
        Double Chucktemp6;
        Double Chucktemp7;
        Double Chucktemp8;

        //-----------------------------------------Average function for the average of chuck tempreture -------------------------------------------------------
        #region Average function 
        private void Averagefunction(double ct)
        {
            if (j < 10)
            {
                Avglist.Insert(j, ct);
                //Avgtemp.Add(r);
                if (j == 9)
                {
                    AVGTEMP = Avglist.Average();
                    label37.Text = AVGTEMP.ToString();
                    label37.Text = string.Format("{0:f3}", AVGTEMP);
                    j = 0;

                    Avglist.Clear();
                }
                else
                {
                    j++;
                }
            }
        }
        #endregion
        // --------------------------------------------------------------For controller temperature display--------------------------------------------------
        #region Display and read temperature values of controller
        private string ExtractControllerVariables(string val)
        {
            CultureInfo culture = CultureInfo.InvariantCulture;

            List<string> param = val.Split(':').ToList();
            List<string> vals = param[1].Split('\n').ToList();
            string ret_val = string.Empty;
            float chuck_temp = 0.0F;

            string st_var = val;
            if ((param.Count > 0) && (vals.Count > 1) && (vals[1].Equals("OK")))
            {
                string cmd = param[0];
                st_var = vals[0];

                switch (cmd)
                {
                    case "CPT": // Chuck Temperature
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f3}", chuck_temp);
                        this.PT0TextBox.Text = ret_val;
                      
                        Chucktemp1 = Convert.ToSingle(ret_val);
                        
                      
                 
                        break;
                    case "CPTE1": // Extra temp 1
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f3}", chuck_temp);
                        this.PT1TextBox.Text = ret_val;
                        Chucktemp2 = Convert.ToSingle(ret_val);
                        break;
                    case "CPTE2": // Extra temp 2
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f3}", chuck_temp);
                        this.PT2TextBox.Text = ret_val;
                        Chucktemp3 = Convert.ToSingle(ret_val);
                        break;
                    case "CPTE3": // Extra temp 3
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f3}", chuck_temp);
                        this.PT3TextBox.Text = ret_val;
                        Chucktemp4 = Convert.ToSingle(ret_val);
                        break;
                    case "CPTE4": // Extra temp 4
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f3}", chuck_temp);
                        this.PT4TextBox.Text = ret_val;
                        Chucktemp5 = Convert.ToSingle(ret_val);
                        break;
                    case "CPTE5": // Extra temp 5
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f3}", chuck_temp);
                        this.PT5TextBox.Text = ret_val;
                        Chucktemp6 = Convert.ToSingle(ret_val);
                        break;
                    case "CPTE6": // Extra temp 6
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f3}", chuck_temp);
                        this.PT6TextBox.Text = ret_val;
                        Chucktemp7 = Convert.ToSingle(ret_val);
                        break;
                    case "CPTE7": // Extra temp 7
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f3}", chuck_temp);
                        this.PT7TextBox.Text = ret_val;
                        Chucktemp8 = Convert.ToSingle(ret_val);
                        break;
                    default:
                        break;

                }
            }
            return ret_val;
        }
        #endregion
        //--------------------------------------------xxxxxxxxxxxxxxxxxxxxxxxx------------------------------xxxxxxxxxxxxxxxxxxxxxxx

        private string GetChuckTemp(string cmd)
        {
            string resp = string.Empty;
            ASCIIEncoding encoding = new ASCIIEncoding();
            byte[] sendbuf = new byte[16];
            sendbuf = encoding.GetBytes(cmd);
            int count = ac3r2Net.TCPSendMessage(sendbuf, cmd.Length);
            if (count > 0)
            {
                ac3r2Net.GetBytesRead(m_recv_buf, count);
                string s = System.Text.Encoding.ASCII.GetString(m_recv_buf, 0, 32);
                resp = ExtractControllerVariables(s);
            }
            return resp;
        }


        //------------------------------------ For displaying the offset values on screen --------------------------------------------------------- 
        #region For reading and displaying offset values on screen 

        float Offsetscale;
        private string SetOffset(string Offset)
        {
            CultureInfo culture = CultureInfo.InvariantCulture;

            List<string> param = Offset.Split(':').ToList();
            List<string> vals = param[1].Split('\n').ToList();
            string ret_val = string.Empty;
            float chuck_temp = 0.0F;

            string st_var = Offset;
            if ((param.Count > 0) && (vals.Count > 1) && (vals[1].Equals("OK")))
        
            {
                string cmd = param[0];
                st_var = vals[0];

                switch (cmd)
                {
                    case "PTOF.1": // Offset for sensor 1
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);
                    
                        this.TxtBoxOffset0.Text = ret_val;

                        break;

                    case "PTOF.2": // Offset for sensor 1
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxOffset1.Text = ret_val;

                        break;

                    case "PTOF.3": // Offset for sensor 1
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxOffset2.Text = ret_val;

                        break;

                    case "PTOF.4": // Offset for sensor 1
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxOffset3.Text = ret_val;

                        break;

                    case "PTOF.5": // Offset for sensor 1
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxOffset4.Text = ret_val;

                        break;

                    case "PTOF.6": // Offset for sensor 1
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxOffset5.Text = ret_val;

                        break;

                    case "PTOF.7": // Offset for sensor 1
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxOffset6.Text = ret_val;

                        break;

                    case "PTOF.8": // Offset for sensor 1
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxOffset7.Text = ret_val;

                        break;


                    default:
                        break;

                }
            }
            return ret_val;
        }
        private string GetOffsetTemp()
        {
            string resp = string.Empty;
            ASCIIEncoding encoding = new ASCIIEncoding();
            byte[] sendbuf = new byte[16];
            string s = "GET PTO\n";
            sendbuf = encoding.GetBytes(s);
            int count = ac3r2Net.TCPSendMessage(sendbuf, s.Length);
            if (count > 0)
            {
                ac3r2Net.GetBytesRead(m_recv_buf, count);
                s = System.Text.Encoding.ASCII.GetString(m_recv_buf, 0, 32);
                resp = SetOffset(s);
            }
            resp = SetOffset(s);
            return resp;
        }
        private string SetOffsetTempEx(channelType Chan, float val)
        {
            string s = "GET PTOF.1\n";
            string Sval = val.ToString();
            Sval = Sval.Replace(',', '.');//------For scale variable-----
            switch (Chan)
            {
                case channelType.Channel1:
                    s = "SET PTOF.1 " + Sval + "\n";
                    break;
                case channelType.Channel2:
                    s = "SET PTOF.2 " + Sval + "\n";
                    break;
                case channelType.Channel3:
                    s = "SET PTOF.3 " + Sval + "\n";
                    break;
                case channelType.Channel4:
                    s = "SET PTOF.4 " + Sval + "\n";
                    break;
                case channelType.Channel5:
                    s = "SET PTOF.5 " + Sval + "\n";
                    break;
                case channelType.Channel6:
                    s = "SET PTOF.6 " + Sval + "\n";
                    break;
                case channelType.Channel7:
                    s = "SET PTOF.7 " + Sval + "\n";
                    break;
                case channelType.Channel8:
                    s = "SET PTOF.8 " + Sval + "\n";
                    break;



            }
            string resp = GetOffTemp(s);
            return resp;
        }




        private string GetOffTemp(string s)
        {
            string resp = string.Empty;
            ASCIIEncoding encoding = new ASCIIEncoding();
            byte[] sendbuf = new byte[16];
            // string s = "GET CPTE1\n";
            //string s = "GET PTOF0\n";
            sendbuf = encoding.GetBytes(s);
            int count = ac3r2Net.TCPSendMessage(sendbuf, s.Length);
            if (count > 0)
            {
                ac3r2Net.GetBytesRead(m_recv_buf, count);
                s = System.Text.Encoding.ASCII.GetString(m_recv_buf, 0, 32);
                resp = SetOffset(s);
            }
            resp = SetOffset(s);
            return resp;
        }
        #endregion
        //-------------------------------------Function for reconnecting to the server-------------------------------------------------
        #region Reconnect to server function
        private void Reconnect2Server()
        {

            ac3r2Net.DisconectFromServer();

            this.ConnectBtn.Text = "Connect";
            this.NetStatusTextBox.Text = "Offline";
            this.NetStatusTextBox.BackColor = Color.Red;
            UpdateTimer.Enabled = false;

            if (ac3r2Net.isConnected)
            {
                SetText(ConnectBtn, "Disconnect");
                SetText(this.NetStatusTextBox, "Connected");
                ChangBgColor(this.NetStatusTextBox, Color.Green);
                UpdateTimer.Enabled = true;
            }

        }
        #endregion
        //------------------------------------xxxxxxxxxxxxxxxxxxxxxxxxxxx-----------------------------------------------------------------
        //-----------------------When offset and linearity textbox value is changed----------------------------------------
        #region When offset and linearity values are changed in textbox
        private void TxtBoxOffset0_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxOffset0.Text;
            float yval = parseFloat(s);
            SetOffsetTempEx(channelType.Channel1, yval);
        }
        private void TxtBoxLiner0_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxLiner0.Text;
            float xval = parseFloat(s);
            SetLinearityTempEx(channelType.Channel1, xval);
        }

        private void TxtBoxOffset1_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxOffset1.Text;
            float yval = parseFloat(s);
            SetOffsetTempEx(channelType.Channel2, yval);
        }

        private void TxtBoxLiner1_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxLiner1.Text;
            float xval = parseFloat(s);
            SetLinearityTempEx(channelType.Channel2, xval);
        }

        private void TxtBoxOffset2_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxOffset2.Text;
            float yval = parseFloat(s);
            SetOffsetTempEx(channelType.Channel3, yval);
        }

        private void TxtBoxLiner2_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxLiner2.Text;
            float xval = parseFloat(s);
            SetLinearityTempEx(channelType.Channel3, xval);
        }

        private void TxtBoxOffset3_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxOffset3.Text;
            float yval = parseFloat(s);
            SetOffsetTempEx(channelType.Channel4, yval);
        }

        private void TxtBoxLiner3_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxLiner3.Text;
            float xval = parseFloat(s);
            SetLinearityTempEx(channelType.Channel4, xval);

        }

        private void TxtBoxOffset4_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxOffset4.Text;
            float yval = parseFloat(s);
            SetOffsetTempEx(channelType.Channel5, yval);

        }

        private void TxtBoxLiner4_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxLiner4.Text;
            float xval = parseFloat(s);
            SetLinearityTempEx(channelType.Channel5, xval);
        }

        private void TxtBoxOffset5_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxOffset5.Text;
            float yval = parseFloat(s);
            SetOffsetTempEx(channelType.Channel6, yval);

        }

        private void TxtBoxLiner5_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxLiner5.Text;
            float xval = parseFloat(s);
            SetLinearityTempEx(channelType.Channel6, xval);
        }

        private void TxtBoxOffset6_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxOffset6.Text;
            float yval = parseFloat(s);
            SetOffsetTempEx(channelType.Channel7, yval);

        }

        private void TxtBoxLiner6_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxLiner6.Text;
            float xval = parseFloat(s);
            SetLinearityTempEx(channelType.Channel7, xval);
        }

        private void TxtBoxOffset7_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxOffset7.Text;
            float yval = parseFloat(s);
            SetOffsetTempEx(channelType.Channel8, yval);
        }

        private void TxtBoxLiner7_TextChanged(object sender, EventArgs e)
        {
            string s = TxtBoxLiner7.Text;
            float xval = parseFloat(s);
            SetLinearityTempEx(channelType.Channel8, xval);
        }

        private float parseFloat(string s)
        {
            float val = 0;
            float.TryParse(s, out val);

            return val;
        }
        #endregion

        //---------------------------------- Function code for setting linearity values----------------------------------------------------
        #region Function code for setting linearity values 
        private string SetLinearityTempEx(channelType chan, float rval)
        {
            
            string s = "GET PTSC.1\n";
            string Sval = rval.ToString();
            Sval= Sval.Replace(',', '.');//------For scale variable-----
            //Sval = string.Format("{0:f4}", Sval);


            switch (chan)
            {
                case channelType.Channel1:
                    s = "SET PTSC.1 " + Sval + "\n";
                    break;
                case channelType.Channel2:
                    s = "SET PTSC.2 " + Sval + "\n";
                    break;
                case channelType.Channel3:
                    s = "SET PTSC.3 " + Sval + "\n";
                    break;
                case channelType.Channel4:
                    s = "SET PTSC.4 " + Sval + "\n";
                    break;
                case channelType.Channel5:
                    s = "SET PTSC.5 " + Sval + "\n";
                    break;
                case channelType.Channel6:
                    s = "SET PTSC.6 " + Sval + "\n";
                    break;
                case channelType.Channel7:
                    s = "SET PTSC.7 " + Sval + "\n";
                    break;
                case channelType.Channel8:
                    s = "SET PTSC.8 " + Sval + "\n";
                    break;



            }
            string resp = GetLinearTemp(s);
            return resp;
        }
        //---------------------------------- Function code for getting linearity values----------------------------------------------------

        private string GetLinearityTempEx(int num)
        {
            string s = "GET PTSC.1\n";
            switch (num)
            {
                case 1:
                    s = "GET PTSC.1\n";
                    break;
                case 2:
                    s = "GET PTSC.2\n";
                    break;
                case 3:
                    s = "GET PTSC.3\n";
                    break;
                case 4:
                    s = "GET PTSC.4\n";
                    break;
                case 5:
                    s = "GET PTSC.5\n";
                    break;
                case 6:
                    s = "GET PTSC.6\n";
                    break;
                case 7:
                    s = "GET PTSC.7\n";
                    break;
                case 8:
                    s = "GET PTSC.8\n";
                    break;



            }
            string resp = GetLinearTemp(s);
            return resp;
        }

        private string GetLinearTemp(string s)
        {
            string resp = string.Empty;
            ASCIIEncoding encoding = new ASCIIEncoding();
            byte[] sendbuf = new byte[16];
            // string s = "GET CPTE1\n";
            //string s = "GET PTOF0\n";
            sendbuf = encoding.GetBytes(s);
            int count = ac3r2Net.TCPSendMessage(sendbuf, s.Length);
            if (count > 0)
            {
                ac3r2Net.GetBytesRead(m_recv_buf, count);
                s = System.Text.Encoding.ASCII.GetString(m_recv_buf, 0, 32);
                resp = SetLinearity(s);
            }
            resp = SetLinearity(s);
            return resp;
        }

        #endregion

        //--------------------------------------Below code For displaying the linearity value on screen----------------------------------------------------------
        #region Displaying and reading linearity value 
        float scale;
        private string SetLinearity(string val)
        {
            CultureInfo culture = CultureInfo.InvariantCulture;

            List<string> param = val.Split(':').ToList();
            List<string> vals = param[1].Split('\n').ToList();
            string ret_val = string.Empty;
            float chuck_temp = 0.0F;

            string st_var = val;
            if ((param.Count > 0) && (vals.Count > 1) && (vals[1].Equals("OK")))
            {
                string cmd = param[0];
                st_var = vals[0];

                switch (cmd)
                {
                    case "PTSC.1":
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxLiner0.Text = ret_val;
                       
                     
                        break;
                    case "PTSC.2":
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxLiner1.Text = ret_val;
                      
                        break;

                    case "PTSC.3":
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxLiner2.Text = ret_val;
                  

                        break;
                    case "PTSC.4":
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxLiner3.Text = ret_val;
                      

                        break;
                    case "PTSC.5":
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxLiner4.Text = ret_val;
                     

                        break;
                    case "PTSC.6":
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxLiner5.Text = ret_val;
                        break;
                    case "PTSC.7":
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxLiner6.Text = ret_val;
                        break;
                    case "PTSC.8":
                        chuck_temp = Single.Parse(st_var, culture);
                        ret_val = string.Format("{0:f4}", chuck_temp);

                        this.TxtBoxLiner7.Text = ret_val;
                        

                        break;
                    default:
                        break;

                }
            }
            return ret_val;
        }
        #endregion
        //----------------------------xxxxxxxxxxxxxxxxxxxxxxx------------------------xxxxxxxxxxxxxxxxx---------------------------------------------------


        private void TimeoutTimer_Tick(object sender, EventArgs e)
        {
            TimeoutTimer.Stop();

            MessageBox.Show("No response from remote device. \r\nPlease check your connection.");
        }

        private void Timerlimit_Tick(object sender, EventArgs e)
        {
            if (bestvalue == false)
            {
                MessageBox.Show("We are not able to set this channel automatically. Please set this channel manually. ");
                loopTimer.Stop();
                Timerlimit.Stop();

            }


        }

        //-------------------When temperature cycle button is turned on --------------------------------------------------------------------
        #region Temperature cycle button

        private void cycleTemp_Click(object sender, EventArgs e)
        {
            Timerlimit.Start();
            #region Set maximum time limit
            if (maxtimerlimit.Text == "")
            {
                Timerlimit.Interval = 3600000;
            }
            else if(maxtimerlimit.Text != "")
            {
                int x;
                x = Convert.ToInt32(maxtimerlimit.Text);
                Timerlimit.Interval = x ;
            }
            #endregion

            max = 0;                                                          //<<<<<<<<----------- For reset the max temp value
            min = 6;                                                      //<<<<<<<<<<<-------------- For reset the min temp value

        
            if (ChannelType==channelType.Channel1)
            {
                Offsetscale = Convert.ToSingle(TxtBoxOffset0.Text);
                scale = Convert.ToSingle(TxtBoxLiner0.Text);

            }
            else if (ChannelType == channelType.Channel2)
            {
                Offsetscale = Convert.ToSingle(TxtBoxOffset1.Text);
                scale = Convert.ToSingle(TxtBoxLiner1.Text);
            }
            else if (ChannelType == channelType.Channel3)
            {
                Offsetscale = Convert.ToSingle(TxtBoxOffset2.Text);
                scale = Convert.ToSingle(TxtBoxLiner2.Text);
            }
            else if(ChannelType== channelType.Channel4)
            {
                Offsetscale = Convert.ToSingle(TxtBoxOffset3.Text);
                scale = Convert.ToSingle(TxtBoxLiner3.Text);
            }
            else if (ChannelType == channelType.Channel5)
            {
                Offsetscale = Convert.ToSingle(TxtBoxOffset4.Text);
                scale = Convert.ToSingle(TxtBoxLiner4.Text);
            }
            else if(ChannelType== channelType.Channel6)
            {
                Offsetscale = Convert.ToSingle(TxtBoxOffset5.Text);
                scale = Convert.ToSingle(TxtBoxLiner5.Text);
            }
            else if (ChannelType == channelType.Channel7)
            {
                Offsetscale = Convert.ToSingle(TxtBoxOffset6.Text);
                scale = Convert.ToSingle(TxtBoxLiner6.Text);
            }
            else if (ChannelType == channelType.Channel8)
            {
                Offsetscale = Convert.ToSingle(TxtBoxOffset7.Text);
                scale = Convert.ToSingle(TxtBoxLiner7.Text);
            }
           
            
            timerEnablePlot = true;
            loopTimer.Interval = 30000;
            switch(ChannelType)
            {
                case channelType.Channel1:
                case channelType.Channel2:
                case channelType.Channel3:
                case channelType.Channel4:
                case channelType.Channel5:
                case channelType.Channel6:
                case channelType.Channel7:
                case channelType.Channel8:
                simulatorTemp = -80;
                ExternalTemp.Text = simulatorTemp.ToString();
                sendTemp(simulatorTemp.ToString());

                loopTimer.Enabled = true;
                loopTimer.Start();
                break;
                case channelType.Undef:
                loopTimer.Stop();
                MessageBox.Show("Please select channel type.");
                break;
            }
            
            


        }
        #endregion
        //----------------------------------------xxxxxxxxxxxxxxxxxx-------------------------------------------------------

        //---------------This is for stopping the temperature cycle----------------------------------------------------------
        #region Stop button
        private void Stopbtn_Click(object sender, EventArgs e)
        {
            loopTimer.Stop();
            MessageBox.Show("Temperature cycle cancelled.");
        }
        #endregion
        //------------------------xxxxxxxxxxxxxxxxxxxx----------------------------------------------------------
        bool bestvalue=false;
        bool stopbestloop = false;
        bool bestloop = false;

        Double simulatorTemp;
        Double difference;
        private void measureTempDifference()
        {
            difference = Math.Abs(AVGTEMP - simulatorTemp);
           
        }

        #region Update simulator temperature
        private void updateSimulatorTemp()
        {
            if (simulatorTemp < 320 )
            {
                if (bestvalue == true && counter <= 3)  // if best value equals to true and counter is less than 3 then we will take small temperature steps of 20.
                {
                    simulatorTemp += 20;
                    bestloop = true;                     
                   
                }
                else if (counter < 3)                  //If counter is less than 3 then we will take temperature steps of 60.
                {
                    simulatorTemp += 60;
                }
          
                else                                  // If counter is more than 3 then we will take temperature steps of 80.   
                {
                    simulatorTemp += 80;
                }
                ExternalTemp.Text = simulatorTemp.ToString();
                label34.Text = string.Format("{0:f3}", difference);
            }
       
            else
            {

                if (bestvalue == true && bestloop == true)   // If best values and best loop both are true then we will stops timers .
                {

                    loopTimer.Stop();

                    Add_data_excelsheet();
                    stopbestloop = true;

                }
                pointpairlist.Clear();
                loopTimer.Stop();
               // updateGraphList();
                loopTimer.Enabled = false;
                //list++;
                autoLinearity();

                if (bestvalue == false )
                {
                   
                    tempvalues.Clear();
                    diffvalues.Clear();
                    max = 0;                                                          //<<<<<<<<----------- For reset the max temp value
                    min = 6;                                                      //<<<<<<<<<<<-------------- For reset the min temp value
                    simulatorTemp = -80;
                    ExternalTemp.Text = simulatorTemp.ToString();
                    sendTemp(simulatorTemp.ToString());
                    timerEnablePlot = true;
                    loopTimer.Enabled = true;
                    loopTimer.Start();
                }
              
          
                
         


            }


        }
        #endregion

        private void Add_data_excelsheet()
        {
            xlWorkBook = xlapp.Workbooks.Open(@"C:\Users\apanchal\Desktop\New folder\TEMPDATA.xls");
            //  xlWorkBook = xlapp.Workbooks.Add(misvalue);
            xlWorkSheet = xlWorkBook.Worksheets[1];
            xlWorkSheet.Cells[8, column] = scale.ToString();
            xlWorkSheet.Cells[9, column] = Offsetscale.ToString();
            for (int j = 0; j < diffvalues.Count; j++)
            {
                xlWorkSheet.Cells[12 + j, 1] = tempvalues[j];
                xlWorkSheet.Cells[12 + j, column] = diffvalues[j];
            }





            xlWorkSheet.Columns.AutoFit();
            xlWorkBook.SaveAs(@"C:\Users\apanchal\Desktop\New folder\TEMPDATA.xls");
            xlWorkBook.Close(true, misvalue, misvalue);
            xlapp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlapp);

            MessageBox.Show("Excel file created , you can find the file @C:\\Users\\apanchal\\Desktop\\New folder\\TEMPDATA.xls");
        }
        #region Update graph list function
        private void updateGraphList()
        {

            if (list == 0)
            {
               // for (int i = 0; i < diffvalues.Count; i++)
                {
                    // PointPair pointpair = new PointPair(tempvalues[i],diffvalues[i]);
                    PointPair pointpair = new PointPair(simulatorTemp, difference);
                    pointpairlist.Add(pointpair);
                }
            
            }
            else if (list == 1)
            {
                PointPair pointpair1 = new PointPair(simulatorTemp, difference);
                pointpairlist1.Add(pointpair1);
            }
            else if (list == 2)
            {
                PointPair pointpair2 = new PointPair(simulatorTemp, difference);
                pointpairlist2.Add(pointpair2);
            }
            else if (list == 3)
            {
                PointPair pointpair3 = new PointPair(simulatorTemp, difference);
                pointpairlist3.Add(pointpair3);
            }
            else if (list == 4)
            {
                PointPair pointpair4 = new PointPair(simulatorTemp, difference);
                pointpairlist4.Add(pointpair4);
            }
            else if (list == 5)
            {
                PointPair pointpair5 = new PointPair(simulatorTemp, difference);
                pointpairlist5.Add(pointpair5);
            }
   

            this.zedGraphControl1.AxisChange();
            this.zedGraphControl1.Invalidate();          

        }
        #endregion
        //-----------------Function given below is for setting the linearity automatically--------------------------------------------
        #region Autolinearity function

        float step;
        float offsetstep;
    
        //int channelType ;
        
        
        private void autoLinearity()
        {
          
             
            CultureInfo culture = CultureInfo.InvariantCulture;
            

            
           
            {
                if ( counter<=3 )   
                {
                   // MessageBox.Show("Updating offset value.");
                    if (max <0.009)
                    {
                        offsetstep = +0.001F;
                        Offsetscale = Offsetscale + offsetstep;
                        SetOffsetTempEx(ChannelType, Offsetscale);
                        bestvalue = false;
                    }
                    else if (min >0.009)                 // Here we used 0.009 for the making condition min >0.00
                    {
                       offsetstep = -0.001F;
                        Offsetscale = Offsetscale + offsetstep;
                        SetOffsetTempEx(ChannelType, Offsetscale);
                        bestvalue = false;

                    }
                    else if (max>=0.009 && min<=0.009)
                    { 
                 
                            bestvalue = true;
                            MessageBox.Show("We got best Offset and lineartity values ");
                        if (stopbestloop == false)
                        {
                            tempvalues.Clear();
                            diffvalues.Clear();
                            max = 0;                                                          //<<<<<<<<----------- For reset the max temp value
                            min = 6;                                                      //<<<<<<<<<<<-------------- For reset the min temp value
                            simulatorTemp = -80;
                            ExternalTemp.Text = simulatorTemp.ToString();
                            sendTemp(simulatorTemp.ToString());
                            timerEnablePlot = true;
                            loopTimer.Enabled = true;
                            loopTimer.Start();

                        }

                    }
                  
                    
                    counter = 0;
               
                }
                else
                {
                    
                    if ((max - min) > 0.03 && counter >3)
                    {
                      if( comboBox2.Text == "Channel 8")
                      {
                            step = -0.001F;
                      }
                      else
                      {
                            step = -0.0001F;
                      }
                        
                        scale = scale + step;
                        SetLinearityTempEx(ChannelType, scale);
                        bestvalue = false;
                        counter = 0;
                    }
                    if ((max - min) < 0.03 && counter>3)
                    {
                        if ( comboBox2.Text == "Channel 8")
                        {
                            step = +0.001F;
                        }
                        else
                        {
                            step = +0.0001F;
                        }
                       
                        scale = scale + step;
                        SetLinearityTempEx(ChannelType, scale);
                        bestvalue = false;
                        counter = 0;
                    }
                }
            

            }
        }
        #endregion
        //-----------------------------xxxxxxxxxxxxxxxxxxxxxxxxxxxx---------------------------------------
        int counter =0;
        
        private void loopTimer_Tick(object sender, EventArgs e)
        {
            try
            {
               
                loopTimer.Stop();
                
                measureTempDifference();
                if (difference > 0.05) { counter++; }

              
                maxDifferenceinTemp();
                minDifferenceinTemp();
       
                updateGraphList();
                tempvalues.Add(simulatorTemp);
                diffvalues.Add(difference);
                updateSimulatorTemp();
          
                
            }
            catch
            {

            }
            finally
            {
                if (timerEnablePlot == true && stopbestloop==false )
                {
                    loopTimer.Start();
                }

            }
        }

        //-----------------------Function for finding the maximum difference in temperature -------------------------------------------------------------
        #region Function for finding the maximum and minimum values
        double max =0;
        private void maxDifferenceinTemp()
        {
            if (difference > max)
            {
                max = difference;
                labelmax.Text = max.ToString();
                labelmax.Text = string.Format("{0:F3}", max);
            }
        }
     
        //------------------------------------------xxxxxxxxxxxxxxx---------------------------------------------------------------
        //--------------------------Function for finding the minimum difference in temperature ----------------------------------------------
        double min = 6;
        private void minDifferenceinTemp()
        {
            if (difference < min)
            {
                min = difference;
                labelmin.Text = min.ToString();
                labelmin.Text = string.Format("{0:F3}",min);
            }

        }
        #endregion
        //------------------------------------------xxxxxxxxxxxxxxx---------------------------------------------------------------


        //---------------------------If we insert temperature values in given textbox ---------------------------------
        #region when external temperature is given as input 
        private void ExternalTemp_TextChanged(object sender, EventArgs e)
        {
            sendTemp(ExternalTemp.Text);
            simulatorTemp = Convert.ToSingle(ExternalTemp.Text);
        }
        #endregion
        //------------------------------------------xxxxxxxxxxxxxxx---------------------------------------------------------------



        //-----------------Below function is for sending the temperature value on simulator---------------------------------------------
        #region Function for sending temperature to simulator 
        private void sendTemp(string temp)
        {
            validateTempNumber(temp);
            ValidateTempLimits(temp);

            serialport.Write(RS232CmdList.outPlat + temp + RS232CmdList.outCRLR);
            

        }
        #endregion
        //---------------------------------xxxxxxxxxxxxxx--------------------------------------------------------

        //------------------------------For verifying if the tempreture value is valid or not -----------------------------------
        #region Verify temperature values are valid or not 
        private static bool ValidateTempLimits(string setlimit)
        {
            bool limit = true;
            foreach (char l in setlimit)
            {
                if (l >= -200 && l >= 850)

                    Console.WriteLine("Please Enter the valid input !! Value is between -200 to +850");
            }
            return limit;
        }

        private static bool validateTempNumber(string inputstring)
        {
            bool validNumber = true;

            foreach (char c in inputstring)
            {
                if (!Char.IsDigit(c))
                    validNumber = false;
                Console.WriteLine("Please Enter the valid input !!");
            }
            return validNumber;
        }
        #endregion
        //--------------------------------xxxxxxxxxxxxxxxxxxxxxx-----------------------------------------------
        private void PT0TextBox_TextChanged(object sender, EventArgs e)
        {

        }

        //---------------------------------------For selecting channel type checkbox-------------------------------------------
        #region Channel selection combobox
        private bool checkChannel(double temp)
        {
            bool channelCorrect = false;
            if (Round(temp - simtemp) < 5)
            {
            
                channelCorrect = true;
            }
            else
            {
                MessageBox.Show("Wrong channel selected .");
                ChannelType = channelType.Undef;
               
            }
            return channelCorrect;
        }
        int column=0;
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text== "Channel 1")
            {
                if(checkChannel(Chucktemp1))
                {
                    ChannelType = channelType.Channel1;
                    column = 2;
                }
               
            }
            else if(comboBox2.Text=="Channel 2")
            {
                if (checkChannel(Chucktemp2))
                {
                    ChannelType= channelType.Channel2;
                    column = 3;
                }
            
            }
            else if (comboBox2.Text == "Channel 3")
            {
                if (checkChannel(Chucktemp3))
                {
                    ChannelType = channelType.Channel3;
                    column = 4;
                }
        
            }
            else if (comboBox2.Text == "Channel 4")
            {
                if (checkChannel(Chucktemp4))
                {
                    ChannelType = channelType.Channel4;
                    column = 5;
                }
             
            }
            else if (comboBox2.Text == "Channel 5")
            {
                if (checkChannel(Chucktemp5))
                {
                    ChannelType = channelType.Channel5;
                    column = 6;
                }
           
            }
            else if (comboBox2.Text == "Channel 6")
            {
                if (checkChannel(Chucktemp6))
                {
                    ChannelType = channelType.Channel6;
                    column = 7;
                }
       
            }
            else if (comboBox2.Text == "Channel 7")
            {
                if (checkChannel(Chucktemp7))
                {
                    ChannelType = channelType.Channel7;
                    column = 8;
                }
              
            }
            else if (comboBox2.Text == "Channel 8")
            {
                if (checkChannel(Chucktemp8))
                {
                    ChannelType = channelType.Channel8;
                    column = 9;
                }
              
            }
            else if (comboBox2.Text== "None")
            {
                ChannelType = channelType.Undef;
                MessageBox.Show("Please select the channel.");
            }


            
        }
        #endregion
        #region Excelsheet save button
        object misvalue = System.Reflection.Missing.Value;
        private void Save_datasheet_Click(object sender, EventArgs e)
        {

            if (xlapp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }
            xlWorkBook = xlapp.Workbooks.Open(@"C:\Users\apanchal\Desktop\New folder\TEMPDATA.xls");
            //  xlWorkBook = xlapp.Workbooks.Add(misvalue);
            xlWorkSheet = xlWorkBook.Worksheets[1];
          
            //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Excel.Range xlRange = xlWorkSheet.UsedRange;

           



            xlWorkSheet.Cells[1, 1] = "Controller Type:";
            xlWorkSheet.Cells[1, 1].Font.Bold = true;
            xlWorkSheet.Cells[1, 2] = textBox1.Text;
            xlWorkSheet.Cells[2, 1] = "Serial Number:";
            xlWorkSheet.Cells[2, 1].Font.Bold = true;
            xlWorkSheet.Cells[2, 2] = textBox2.Text;
            xlWorkSheet.Cells[3, 1] = "ID-Number:";
            xlWorkSheet.Cells[3, 1].Font.Bold = true;
            xlWorkSheet.Cells[3, 2] = textBox3.Text;
            xlWorkSheet.Cells[4, 1] = "Inspector:";
            xlWorkSheet.Cells[4, 1].Font.Bold = true;
            xlWorkSheet.Cells[4, 2] = textBox4.Text;
            xlWorkSheet.Cells[5, 1] = "Date:";
            xlWorkSheet.Cells[5, 1].Font.Bold = true;
            xlWorkSheet.Cells[5, 2] = textBox5.Text;
            xlWorkSheet.Cells[6, 1] = "ChargeNr:";
            xlWorkSheet.Cells[6, 1].Font.Bold = true;
            xlWorkSheet.Cells[6, 2] = textBox6.Text;
            xlWorkSheet.Cells[7, 1] = "System ID:";
            xlWorkSheet.Cells[7, 1].Font.Bold = true;
            xlWorkSheet.Cells[7, 2] = SysIDText.Text;
            xlWorkSheet.Cells[8, 1] = "Linearity:";
            xlWorkSheet.Cells[8, 1].Font.Bold = true;
           
            xlWorkSheet.Cells[8, column] = scale.ToString();
               
            
         
            xlWorkSheet.Cells[9, 1] = "Offset:";
            xlWorkSheet.Cells[9, 1].Font.Bold = true;
            xlWorkSheet.Cells[9, column] = Offsetscale.ToString();
            object[] objHeaders = { "Mesuared Value CH0", "Mesuared Value CH1", "Mesuared Value CH2", "Mesuared Value CH3", "Mesuared Value CH4", "Mesuared Value CH5", "Mesuared Value CH6", "Mesuared Value CH7" };
            m_objRange = xlWorkSheet.get_Range("B11", "I11");
            m_objRange.set_Value(misvalue, objHeaders);
            xlWorkSheet.Cells[11, 1].entirerow.Font.Bold = true;
            xlWorkSheet.Cells[11, 1] = "Set Tempreture";
            xlWorkSheet.Cells[34, 1] = "AC3 TS0xx/SP1xx HTU:";
            xlWorkSheet.Cells[34, 1].Font.Bold = true;
            xlWorkSheet.Cells[36, 1] = "AC3 TS0xx/SP1xx :";
            xlWorkSheet.Cells[36, 1].Font.Bold = true;
            xlWorkSheet.Cells[34, 2] = "tolerance: ± 0,05°C -50°C bis 150°C";
            xlWorkSheet.Cells[35, 2] = "± 0,10°C <-50°C und >150°C";
            xlWorkSheet.Cells[36, 2] = "tolerance:     ± 0,10°C";
            xlWorkSheet.Cells[37, 1] = "Passed:";
            xlWorkSheet.Cells[37, 1].Font.Bold = true;
            xlWorkSheet.Cells[37, 2] = "Yes____";
            xlWorkSheet.Cells[37, 3] = "No____";

            xlWorkSheet.Cells[39, 1] = "Comment:";
            xlWorkSheet.Cells[39, 1].Font.Bold = true;
            xlWorkSheet.Cells[42, 1] = "Signature:";
            xlWorkSheet.Cells[42, 1].Font.Bold = true;
            for (int j = 0; j < diffvalues.Count; j++)
            {
                xlWorkSheet.Cells[12 + j, 1] = tempvalues[j];
                xlWorkSheet.Cells[12 + j, column] = diffvalues[j];
            }





            xlWorkSheet.Columns.AutoFit();
            xlWorkBook.SaveAs(@"C:\Users\apanchal\Desktop\New folder\TEMPDATA.xls");
            xlWorkBook.Close(true, misvalue, misvalue);
            xlapp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlapp);

            MessageBox.Show("Excel file created , you can find the file @C:\\Users\\apanchal\\Desktop\\New folder\\TEMPDATA.xls");

        }
        #endregion
   

        private void Clear_datasheet_Click(object sender, EventArgs e)
        {
            xlWorkBook = xlapp.Workbooks.Open(@"C:\Users\apanchal\Desktop\New folder\TEMPDATA.xls");
            xlWorkSheet = xlWorkBook.Worksheets[1];
            xlWorkSheet.Cells.ClearContents();
            Excel.Range xlRange = xlWorkSheet.UsedRange;


            xlWorkSheet.Cells[1, 1] = "Controller Type:";
            xlWorkSheet.Cells[1, 1].Font.Bold = true;
            xlWorkSheet.Cells[1, 2] = textBox1.Text;
            xlWorkSheet.Cells[2, 1] = "Serial Number:";
            xlWorkSheet.Cells[2, 1].Font.Bold = true;
            xlWorkSheet.Cells[2, 2] = textBox2.Text;
            xlWorkSheet.Cells[3, 1] = "ID-Number:";
            xlWorkSheet.Cells[3, 1].Font.Bold = true;
            xlWorkSheet.Cells[3, 2] = textBox3.Text;
            xlWorkSheet.Cells[4, 1] = "Inspector:";
            xlWorkSheet.Cells[4, 1].Font.Bold = true;
            xlWorkSheet.Cells[4, 2] = textBox4.Text;
            xlWorkSheet.Cells[5, 1] = "Date:";
            xlWorkSheet.Cells[5, 1].Font.Bold = true;
            xlWorkSheet.Cells[5, 2] = textBox5.Text;
            xlWorkSheet.Cells[6, 1] = "ChargeNr:";
            xlWorkSheet.Cells[6, 1].Font.Bold = true;
            xlWorkSheet.Cells[6, 2] = textBox6.Text;
            xlWorkSheet.Cells[7, 1] = "System ID:";
            xlWorkSheet.Cells[7, 1].Font.Bold = true;
            xlWorkSheet.Cells[7, 2] = SysIDText.Text;
            xlWorkSheet.Cells[8, 1] = "Linearity:";
            xlWorkSheet.Cells[8, 1].Font.Bold = true;
            if (column != 0)
            {
                xlWorkSheet.Cells[8, column] = scale.ToString();
                xlWorkSheet.Cells[9, column] = Offsetscale.ToString();
            }
          
            xlWorkSheet.Cells[9, 1] = "Offset:";
            xlWorkSheet.Cells[9, 1].Font.Bold = true;
         
            object[] objHeaders = { "Mesuared Value CH0", "Mesuared Value CH1", "Mesuared Value CH2", "Mesuared Value CH3", "Mesuared Value CH4", "Mesuared Value CH5", "Mesuared Value CH6", "Mesuared Value CH7" };
            m_objRange = xlWorkSheet.get_Range("B11", "I11");
            m_objRange.set_Value(misvalue, objHeaders);
            xlWorkSheet.Cells[11, 1].entirerow.Font.Bold = true;
            xlWorkSheet.Cells[11, 1] = "Set Tempreture";
            xlWorkSheet.Cells[34, 1] = "AC3 TS0xx/SP1xx HTU:";
            xlWorkSheet.Cells[34, 1].Font.Bold = true;
            xlWorkSheet.Cells[36, 1] = "AC3 TS0xx/SP1xx :";
            xlWorkSheet.Cells[36, 1].Font.Bold = true;
            xlWorkSheet.Cells[34, 2] = "tolerance: ± 0,05°C -50°C bis 150°C";
            xlWorkSheet.Cells[35, 2] = "± 0,10°C <-50°C und >150°C";
            xlWorkSheet.Cells[36, 2] = "tolerance:     ± 0,10°C";
            xlWorkSheet.Cells[37, 1] = "Passed:";
            xlWorkSheet.Cells[37, 1].Font.Bold = true;
            xlWorkSheet.Cells[37, 2] = "Yes____";
            xlWorkSheet.Cells[37, 3] = "No____";

            xlWorkSheet.Cells[39, 1] = "Comment:";
            xlWorkSheet.Cells[39, 1].Font.Bold = true;
            xlWorkSheet.Cells[42, 1] = "Signature:";
            xlWorkSheet.Cells[42, 1].Font.Bold = true;
            if (column != 0)
            {
                for (int j = 0; j < diffvalues.Count; j++)
                {
                    xlWorkSheet.Cells[12 + j, 1] = tempvalues[j];
                    xlWorkSheet.Cells[12 + j, column] = diffvalues[j];
                }
            }
         


            xlWorkSheet.Columns.AutoFit();
            xlWorkBook.SaveAs(@"C:\Users\apanchal\Desktop\New folder\TEMPDATA.xls");
            xlWorkBook.Close(true, misvalue, misvalue);
            xlapp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlapp);

            MessageBox.Show("Excel file created , you can find the file @ C:\\Users\\panchal\\source\\repos\\Excel files.xlsx");

        }

   



        //----------------------------------xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx-------------------------
    }
}
