using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp_PCcontrolMachine
{
    public partial class Form1 : Form
    {
        private SerialPort serialPort;
        
        // Y接點對應
        private string RL_Y_Name = "Y15"; 
        private string YL_Y_Name = "Y16"; 
        private string GL_Y_Name = "Y17"; 
        
        private Panel[] X_Lights = new Panel[16];
        private Panel[] Y_Lights = new Panel[16];

        private string[] Y_Names = { "Y0", "Y1", "Y2", "Y3", "Y4", "Y5", "Y6", "Y7", "Y10", "Y11", "Y12", "Y13", "Y14", "Y15", "Y16", "Y17" };
        private string[] X_Names = { "X0", "X1", "X2", "X3", "X4", "X5", "X6", "X7", "X10", "X11", "X12", "X13", "X14", "X15", "X16", "X17" };

        public static readonly string ReadYCommand = "02 30 30 30 41 30 30 32 03 36 36";
        public static readonly string ReadXCommand = "02 30 30 30 38 30 30 32 03 35 44";

        public static readonly Dictionary<string, string> ForceOnCommands = new Dictionary<string, string> {
            {"Y0", "02 37 30 30 30 35 03 46 46"},
            {"Y1", "02 37 30 31 30 35 03 30 30"},
            {"Y2", "02 37 30 32 30 35 03 30 31"},
            {"Y3", "02 37 30 33 30 35 03 30 32"},
            {"Y4", "02 37 30 34 30 35 03 30 33"},
            {"Y5", "02 37 30 35 30 35 03 30 34"},
            {"Y6", "02 37 30 36 30 35 03 30 35"},
            {"Y7", "02 37 30 37 30 35 03 30 36"},
            {"Y10", "02 37 30 38 30 35 03 30 37"},
            {"Y11", "02 37 30 39 30 35 03 30 38"},
            {"Y12", "02 37 30 41 30 35 03 31 30"},
            {"Y13", "02 37 30 42 30 35 03 31 31"},
            {"Y14", "02 37 30 43 30 35 03 31 32"},
            {"Y15", "02 37 30 44 30 35 03 31 33"},
            {"Y16", "02 37 30 45 30 35 03 31 34"},
            {"Y17", "02 37 30 46 30 35 03 31 35"}
        };

        public static readonly Dictionary<string, string> ForceOffCommands = new Dictionary<string, string> {
            {"Y0", "02 38 30 30 30 35 03 30 30"},
            {"Y1", "02 38 30 31 30 35 03 30 31"},
            {"Y2", "02 38 30 32 30 35 03 30 32"},
            {"Y3", "02 38 30 33 30 35 03 30 33"},
            {"Y4", "02 38 30 34 30 35 03 30 34"},
            {"Y5", "02 38 30 35 30 35 03 30 35"},
            {"Y6", "02 38 30 36 30 35 03 30 36"},
            {"Y7", "02 38 30 37 30 35 03 30 37"},
            {"Y10", "02 38 30 38 30 35 03 30 38"},
            {"Y11", "02 38 30 39 30 35 03 30 39"},
            {"Y12", "02 38 30 41 30 35 03 31 31"},
            {"Y13", "02 38 30 42 30 35 03 31 32"},
            {"Y14", "02 38 30 43 30 35 03 31 33"},
            {"Y15", "02 38 30 44 30 35 03 31 34"},
            {"Y16", "02 38 30 45 30 35 03 31 35"},
            {"Y17", "02 38 30 46 30 35 03 31 36"}
        };

        private Queue<string> forceQueue = new Queue<string>();
        private bool waitingForResponse = false;
        private int timeoutCounter = 0;
        private int pollStep = 0;
        private string currentCommand = "";
        private List<byte> rxBuffer = new List<byte>();

        public Form1()
        {
            InitializeComponent();
            timerPoll.Interval = 10; // 10ms (Default value)
// 100ms (我在這裡明確指定了！)
            serialPort = new SerialPort();
            serialPort.BaudRate = 9600;
            serialPort.DataBits = 7;
            serialPort.StopBits = StopBits.One;
            serialPort.Parity = Parity.Even;
            serialPort.DataReceived += SerialPort_DataReceived;
            
            cmbPorts.Items.AddRange(SerialPort.GetPortNames());
            if (cmbPorts.Items.Count > 0) cmbPorts.SelectedIndex = 0;

            X_Lights = new Panel[] { lightX0, lightX1, lightX2, lightX3, lightX4, lightX5, lightX6, lightX7, lightX8, lightX9, lightX10, lightX11, lightX12, lightX13, lightX14, lightX15 };
            Y_Lights = new Panel[] { lightY0, lightY1, lightY2, lightY3, lightY4, lightY5, lightY6, lightY7, lightY8, lightY9, lightY10, lightY11, lightY12, lightY13, lightY14, lightY15 };

            for (int i = 0; i < 16; i++)
            {
                System.Drawing.Drawing2D.GraphicsPath pathX = new System.Drawing.Drawing2D.GraphicsPath();
                pathX.AddEllipse(0, 0, X_Lights[i].Width, X_Lights[i].Height);
                X_Lights[i].Region = new Region(pathX);

                System.Drawing.Drawing2D.GraphicsPath pathY = new System.Drawing.Drawing2D.GraphicsPath();
                pathY.AddEllipse(0, 0, Y_Lights[i].Width, Y_Lights[i].Height);
                Y_Lights[i].Region = new Region(pathY);
            }

            System.Drawing.Drawing2D.GraphicsPath pathRL = new System.Drawing.Drawing2D.GraphicsPath();
            pathRL.AddEllipse(0, 0, lightRL.Width, lightRL.Height);
            lightRL.Region = new Region(pathRL);

            System.Drawing.Drawing2D.GraphicsPath pathYL = new System.Drawing.Drawing2D.GraphicsPath();
            pathYL.AddEllipse(0, 0, lightYL.Width, lightYL.Height);
            lightYL.Region = new Region(pathYL);

            System.Drawing.Drawing2D.GraphicsPath pathGL = new System.Drawing.Drawing2D.GraphicsPath();
            pathGL.AddEllipse(0, 0, lightGL.Width, lightGL.Height);
            lightGL.Region = new Region(pathGL);
        }

        private void BtnConnect_Click(object sender, EventArgs e)
        {
            if (serialPort.IsOpen)
            {
                timerPoll.Stop();
                serialPort.Close();
                btnConnect.Text = "Connect";
                lblStatus.Text = "Status: Disconnected";
                lblStatus.ForeColor = Color.Red;
            }
            else
            {
                try
                {
                    if (cmbPorts.SelectedItem != null)
                    {
                        serialPort.PortName = cmbPorts.SelectedItem.ToString();
                        serialPort.Open();
                        btnConnect.Text = "Disconnect";
                        lblStatus.Text = "Status: Connected";
                        lblStatus.ForeColor = Color.Green;
                        
                        forceQueue.Clear();
                        waitingForResponse = false;
                        timeoutCounter = 0;
                        pollStep = 0;
                        rxBuffer.Clear();
                        
                        timerPoll.Start();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening port: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnY_Click(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen) return;

            Button btn = (Button)sender;
            int index = (int)btn.Tag;
            string yName = Y_Names[index];
            
            // Determine state based on light color
            bool isCurrentlyOn = Y_Lights[index].BackColor == Color.Lime;
            
            string cmdHex = isCurrentlyOn ? ForceOffCommands[yName] : ForceOnCommands[yName];
            forceQueue.Enqueue(cmdHex);
        }

        private void TimerPoll_Tick(object sender, EventArgs e)
        {
            if (!serialPort.IsOpen) return;

            if (waitingForResponse)
            {
                timeoutCounter++;
                if (timeoutCounter > 5) // 500ms timeout
                {
                    waitingForResponse = false;
                    timeoutCounter = 0;
                    rxBuffer.Clear();
                }
                else
                {
                    return; // wait for response
                }
            }

            if (forceQueue.Count > 0)
            {
                string cmd = forceQueue.Dequeue();
                SendCommand(cmd, "FORCE");
            }
            else
            {
                if (pollStep == 0)
                {
                    SendCommand(ReadXCommand, "READ_X");
                    pollStep = 1;
                }
                else
                {
                    SendCommand(ReadYCommand, "READ_Y");
                    pollStep = 0;
                }
            }
        }

        private byte[] ParseHex(string hexString)
        {
            hexString = hexString.Replace(" ", "").Trim();
            byte[] data = new byte[hexString.Length / 2];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
            }
            return data;
        }

        private void SendCommand(string hexCmd, string cmdType)
        {
            currentCommand = cmdType;
            byte[] buffer = ParseHex(hexCmd);
            try
            {
                if (serialPort.IsOpen)
                {
                    serialPort.Write(buffer, 0, buffer.Length);
                    waitingForResponse = true;
                    timeoutCounter = 0;
                    rxBuffer.Clear();
                }
            }
            catch { }
        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                int bytes = serialPort.BytesToRead;
                byte[] buf = new byte[bytes];
                serialPort.Read(buf, 0, bytes);

                this.Invoke(new Action(() =>
                {
                    rxBuffer.AddRange(buf);
                    ProcessBuffer();
                }));
            }
            catch { }
        }

        private void ProcessBuffer()
        {
            while (rxBuffer.Count > 0)
            {
                int stxIndex = rxBuffer.IndexOf(0x02);
                if (stxIndex < 0)
                {
                    if(rxBuffer.IndexOf(0x06) >= 0) // ACK for force command
                    {
                        waitingForResponse = false;
                        rxBuffer.Clear();
                    }
                    else if (rxBuffer.IndexOf(0x15) >= 0) // NAK
                    {
                        waitingForResponse = false;
                        rxBuffer.Clear();
                    }
                    return; 
                }

                if (stxIndex > 0)
                {
                    rxBuffer.RemoveRange(0, stxIndex);
                }

                int etxIndex = rxBuffer.IndexOf(0x03);
                if (etxIndex > 0)
                {
                    if (rxBuffer.Count >= etxIndex + 3) // Need 2 bytes for checksum
                    {
                        byte[] frame = rxBuffer.Take(etxIndex + 3).ToArray();
                        rxBuffer.RemoveRange(0, etxIndex + 3);
                        HandleFrame(frame);
                        waitingForResponse = false;
                    }
                    else
                    {
                        return; // wait for checksum
                    }
                }
                else
                {
                    return; // wait for ETX
                }
            }
        }

        private void HandleFrame(byte[] frame)
        {
            int dataLength = frame.Length - 4; // exclude STX, ETX, CHK1, CHK2
            if (dataLength == 4) // Read response is 4 chars (16 bits)
            {
                string hexData = System.Text.Encoding.ASCII.GetString(frame, 1, 4);
                try
                {
                    byte b1 = Convert.ToByte(hexData.Substring(0, 2), 16);
                    byte b2 = Convert.ToByte(hexData.Substring(2, 2), 16);

                    if (currentCommand == "READ_X")
                    {
                        UpdateLights(X_Lights, b1, b2);
                    }
                    else if (currentCommand == "READ_Y")
                    {
                        UpdateLights(Y_Lights, b1, b2);
                    }
                }
                catch { }
            }
        }

        private bool prevX16State = false;

        private void UpdateLights(Panel[] lights, byte b1, byte b2)
        {
            for (int i = 0; i < 8; i++)
            {
                bool isOn = (b1 & (1 << i)) != 0;
                lights[i].BackColor = isOn ? Color.Lime : Color.Gray;
            }
            for (int i = 0; i < 8; i++)
            {
                bool isOn = (b2 & (1 << i)) != 0;
                lights[8 + i].BackColor = isOn ? Color.Lime : Color.Gray;
            }

            if (lights == X_Lights)
            {
                bool currentX16State = lights[14].BackColor == Color.Lime; // X16 is index 14
                if (currentX16State && !prevX16State) // rising edge
                {
                    bool isWorkMode = lights[11].BackColor == Color.Lime; // X13 (COS1) = 工作模式
                    // 假設單一循環模式時 COS2_L(X14) 和 COS2_R(X15) 都是 OFF (置中)
                    bool isSingleCycle = lights[12].BackColor == Color.Gray && lights[13].BackColor == Color.Gray; 
                    
                    if (isWorkMode && isSingleCycle && !isRunningCycle) 
                    {
                        this.Invoke(new Action(() => { StartSingleCycleProcess(); }));
                    }
                }
                prevX16State = currentX16State;
            }
            CheckStandbyState();
        }

        private bool isStandbyGLOn = false;
        private void CheckStandbyState()
        {
            if (X_Lights[1] == null || Y_Lights[7] == null) return;
            
            bool isAAtBack = X_Lights[1].BackColor == Color.Lime; // X1 (a0)
            bool isMStopped = Y_Lights[7].BackColor != Color.Lime; // Y7 (M+)
            bool isDAtTop = X_Lights[6].BackColor == Color.Lime; // X6 (d0)
            bool isCOff = Y_Lights[5].BackColor != Color.Lime; // Y5 (C+)
            
            bool currentStandby = isAAtBack && isMStopped && isDAtTop && isCOff && !isRunningCycle;

            if (currentStandby && !isStandbyGLOn)
            {
                SetY(GL_Y_Name, true);
                isStandbyGLOn = true;
            }
            else if (!currentStandby && isStandbyGLOn)
            {
                SetY(GL_Y_Name, false);
                isStandbyGLOn = false;
            }
        }
        
        private void ShowMessage(string msg)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => { lblMessage.Text = msg; }));
            }
            else
            {
                lblMessage.Text = msg;
            }
        }
        
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (serialPort != null && serialPort.IsOpen)
            {
                serialPort.Close();
            }
            base.OnFormClosing(e);
        }

        private bool isRunningCycle = false;

        private async Task WaitForSensor(int xIndex, bool targetState, int timeoutMs = 15000)
        {
            int elapsed = 0;
            while ((X_Lights[xIndex].BackColor == Color.Lime) != targetState)
            {
                await Task.Delay(50);
                elapsed += 50;
                if (elapsed > timeoutMs) break; // timeout 保護
            }
        }

        private void SetY(string yName, bool state)
        {
            if (state) forceQueue.Enqueue(ForceOnCommands[yName]);
            else forceQueue.Enqueue(ForceOffCommands[yName]);
        }

        private void btnStartProcess_Click(object sender, EventArgs e)
        {
            StartSingleCycleProcess();
        }

        private async void StartSingleCycleProcess()
        {
            if (isRunningCycle) return;
            if (serialPort == null || !serialPort.IsOpen)
            {
                MessageBox.Show("請先連接 COM Port！", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            isRunningCycle = true;
            btnStartProcess.Enabled = false;
            ShowMessage(""); // clear message

            try
            {
                // RL ON
                SetY(RL_Y_Name, true);
                lightRL.BackColor = Color.Red;
                // GL is handled automatically by CheckStandbyState

                // 1. 推料氣壓缸推出 (A 伸出 Y0)
                ShowMessage("1. 推料氣壓缸推出 (A 伸出 Y0)");
                SetY("Y0", true);

                // 2. 判別出凹槽偏置方向 (X10 = s1, index 8)
                ShowMessage("2. 判別出凹槽偏置方向");
                int s1PulseCount = 0;
                bool prevS1State = X_Lights[8].BackColor == Color.Lime;
                int elapsed = 0;
                
                // 等待 a1 推料缸前端點 (X0) 亮起，期間計算 s1 從 off 轉 on 的次數
                while (X_Lights[0].BackColor != Color.Lime)
                {
                    bool currentS1State = X_Lights[8].BackColor == Color.Lime;
                    if (currentS1State && !prevS1State) s1PulseCount++;
                    prevS1State = currentS1State;

                    await Task.Delay(20); // 快速輪詢
                    elapsed += 20;
                    if (elapsed > 15000) break; 
                }

                bool isLeft = false;
                if (s1PulseCount == 2)
                {
                    isLeft = false; // 右側
                }
                else if (s1PulseCount == 1)
                {
                    isLeft = true; // 左側
                }
                else
                {
                    ShowMessage("方向感測器非左非右");
                    SetY("Y0", false); // 退回推料缸
                    return; // 停止動作
                }

                // 3. 迴轉缸轉至可吸取料件方向
                ShowMessage("3. 迴轉缸轉至可吸取料件方向");
                if (isLeft)
                {
                    // 凹槽在左側，轉至 +90°
                    SetY("Y3", true);
                    SetY("Y4", false);
                    SetY("Y2", false);
                    await WaitForSensor(3, true); // wait for b1 (X3)
                }
                else
                {
                    // 凹槽在右側，轉至 -90°
                    SetY("Y4", true);
                    SetY("Y3", false);
                    SetY("Y2", false);
                    await WaitForSensor(4, true); // wait for b2 (X4)
                }

                // 4. 配合垂直升降模組將料件吸取
                ShowMessage("4. 配合垂直升降模組將料件吸取");
                // 垂直升降下降 (馬達 M+ Y7)
                ShowMessage("垂直升降下降 (馬達 M+ Y7)");
                SetY("Y7", true);
                SetY("Y5", true);
                SetY("Y6", false);
                await WaitForSensor(7, true); // wait for d1 (X7 下端點)
                SetY("Y7", false);

                // 真空吸盤吸 (Y5 ON, Y6 OFF)
                ShowMessage("真空吸盤吸 (Y5 ON, Y6 OFF)");
                await WaitForSensor(5, true); // wait for ps1 (X5 真空壓力開關)

                // 垂直升降上升
                ShowMessage("垂直升降上升");
                SetY("Y7", true);
                await WaitForSensor(6, true); // wait for d0 (X6 上端點)
                SetY("Y7", false);

                // 推料缸退回
                ShowMessage("推料缸退回");
                SetY("Y0", false);

                // 5. 將料件轉至 0°(凹槽在後)
                ShowMessage("5. 將料件轉至 0°(凹槽在後)");
                SetY("Y3", true);
                SetY("Y4", true);
                SetY("Y2", true); // 依照指示三者皆ON
                await WaitForSensor(2, true); // wait for b0 (X2 0°端點)
                await Task.Delay(500); // 再等待0.5秒

                // 確保推料缸已退回
                ShowMessage("確保推料缸已退回");
                await WaitForSensor(1, true); // wait for a0 (X1 後端點)

                // 6. 由出料斜坡排料
                ShowMessage("6. 由出料斜坡排料");
                // 垂直升降下降
                ShowMessage("垂直升降下降");
                SetY("Y7", true);
                await WaitForSensor(7, true); // wait for d1 (X7 下端點)
                SetY("Y7", false);

                // 放開料件
                ShowMessage("放開料件");
                SetY("Y5", false);
                SetY("Y6", true); // 破真空
                await Task.Delay(500); // 吹氣0.5秒
                SetY("Y6", false);

                // 垂直升降上升
                ShowMessage("垂直升降上升");
                SetY("Y7", true);
                await WaitForSensor(6, true); // wait for d0 (X6 上端點)
                SetY("Y7", false);
                
                ShowMessage("排料完成，回到待機狀態");
            }
            catch (Exception ex)
            {
                ShowMessage("發生錯誤: " + ex.Message);
            }
            finally
            {
                // 完成排料後，機構回到待機狀態，運轉燈(RL)滅。待機燈(GL)將由 CheckStandbyState 處理。
                SetY(RL_Y_Name, false);
                lightRL.BackColor = Color.Gray;

                isRunningCycle = false;
                btnStartProcess.Enabled = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
