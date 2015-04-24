using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using System.IO;
using System.Net.Mail;
using System.Web;
using System.Net;

namespace spider_electro56
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////
        //                           Запрос CRC кода к выданным байтам                                      //
        //////////////////////////////////////////////////////////////////////////////////////////////////////
        public static byte[] GetCRC(byte[] btMass)                                                          // 
        {                                                                                                   // 
            byte[] newBtMass;                                                                               //
            newBtMass = new byte[btMass.Length + 2];                                                        // 
            //
            for (int i = 0; i < btMass.Length; i++)                                                         // 
            {                                                                                               //
                newBtMass[i] = btMass[i];                                                                   //
            }                                                                                               //
            //                      
            //
            int Registr = 0xFFFF;                                                                           //
            for (int i = 0; i < btMass.Length; i++)                                                         //
            {                                                                                               //
                Registr = (Registr ^ btMass[i]);                                                            //
                //
                for (int j = 0; j < 8; j++)                                                                 //
                {                                                                                           //
                    if ((Registr & 0x1) == 1)                                                               //
                    {                                                                                       //
                        Registr = Registr >> 1;                                                             //
                        Registr = (Registr ^ 0xA001);                                                       //
                    }                                                                                       //
                    //
                    else                                                                                    //
                    {                                                                                       //
                        Registr = Registr >> 1;                                                             //
                    }                                                                                       //
                }                                                                                           //
            }                                                                                               //
            byte lCRC = (byte)(Registr & 0xff);                                                             //
            byte hCRC = (byte)(Registr >> 8);                                                               //
            newBtMass[newBtMass.Length - 1] = hCRC;                                                         //    
            newBtMass[newBtMass.Length - 2] = lCRC;                                                         //    
            return newBtMass;                                                                               //
        }                                                                                                   //
        //////////////////////////////////////////////////////////////////////////////////////////////////////






        public class Crc16
        {
            private static ushort[] CrcTable = {
            0X0000, 0XC0C1, 0XC181, 0X0140, 0XC301, 0X03C0, 0X0280, 0XC241,
            0XC601, 0X06C0, 0X0780, 0XC741, 0X0500, 0XC5C1, 0XC481, 0X0440,
            0XCC01, 0X0CC0, 0X0D80, 0XCD41, 0X0F00, 0XCFC1, 0XCE81, 0X0E40,
            0X0A00, 0XCAC1, 0XCB81, 0X0B40, 0XC901, 0X09C0, 0X0880, 0XC841,
            0XD801, 0X18C0, 0X1980, 0XD941, 0X1B00, 0XDBC1, 0XDA81, 0X1A40,
            0X1E00, 0XDEC1, 0XDF81, 0X1F40, 0XDD01, 0X1DC0, 0X1C80, 0XDC41,
            0X1400, 0XD4C1, 0XD581, 0X1540, 0XD701, 0X17C0, 0X1680, 0XD641,
            0XD201, 0X12C0, 0X1380, 0XD341, 0X1100, 0XD1C1, 0XD081, 0X1040,
            0XF001, 0X30C0, 0X3180, 0XF141, 0X3300, 0XF3C1, 0XF281, 0X3240,
            0X3600, 0XF6C1, 0XF781, 0X3740, 0XF501, 0X35C0, 0X3480, 0XF441,
            0X3C00, 0XFCC1, 0XFD81, 0X3D40, 0XFF01, 0X3FC0, 0X3E80, 0XFE41,
            0XFA01, 0X3AC0, 0X3B80, 0XFB41, 0X3900, 0XF9C1, 0XF881, 0X3840,
            0X2800, 0XE8C1, 0XE981, 0X2940, 0XEB01, 0X2BC0, 0X2A80, 0XEA41,
            0XEE01, 0X2EC0, 0X2F80, 0XEF41, 0X2D00, 0XEDC1, 0XEC81, 0X2C40,
            0XE401, 0X24C0, 0X2580, 0XE541, 0X2700, 0XE7C1, 0XE681, 0X2640,
            0X2200, 0XE2C1, 0XE381, 0X2340, 0XE101, 0X21C0, 0X2080, 0XE041,
            0XA001, 0X60C0, 0X6180, 0XA141, 0X6300, 0XA3C1, 0XA281, 0X6240,
            0X6600, 0XA6C1, 0XA781, 0X6740, 0XA501, 0X65C0, 0X6480, 0XA441,
            0X6C00, 0XACC1, 0XAD81, 0X6D40, 0XAF01, 0X6FC0, 0X6E80, 0XAE41,
            0XAA01, 0X6AC0, 0X6B80, 0XAB41, 0X6900, 0XA9C1, 0XA881, 0X6840,
            0X7800, 0XB8C1, 0XB981, 0X7940, 0XBB01, 0X7BC0, 0X7A80, 0XBA41,
            0XBE01, 0X7EC0, 0X7F80, 0XBF41, 0X7D00, 0XBDC1, 0XBC81, 0X7C40,
            0XB401, 0X74C0, 0X7580, 0XB541, 0X7700, 0XB7C1, 0XB681, 0X7640,
            0X7200, 0XB2C1, 0XB381, 0X7340, 0XB101, 0X71C0, 0X7080, 0XB041,
            0X5000, 0X90C1, 0X9181, 0X5140, 0X9301, 0X53C0, 0X5280, 0X9241,
            0X9601, 0X56C0, 0X5780, 0X9741, 0X5500, 0X95C1, 0X9481, 0X5440,
            0X9C01, 0X5CC0, 0X5D80, 0X9D41, 0X5F00, 0X9FC1, 0X9E81, 0X5E40,
            0X5A00, 0X9AC1, 0X9B81, 0X5B40, 0X9901, 0X59C0, 0X5880, 0X9841,
            0X8801, 0X48C0, 0X4980, 0X8941, 0X4B00, 0X8BC1, 0X8A81, 0X4A40,
            0X4E00, 0X8EC1, 0X8F81, 0X4F40, 0X8D01, 0X4DC0, 0X4C80, 0X8C41,
            0X4400, 0X84C1, 0X8581, 0X4540, 0X8701, 0X47C0, 0X4680, 0X8641,
            0X8201, 0X42C0, 0X4380, 0X8341, 0X4100, 0X81C1, 0X8081, 0X4040 };

            public static UInt16 ComputeCrc(byte[] data)
            {
                ushort crc = 0xFFFF;

                foreach (byte datum in data)
                {
                    crc = (ushort)((crc >> 8) ^ CrcTable[(crc ^ datum) & 0xFF]);
                }

                return crc;
            }
        }

        void ans_explode(byte[] _data, out string _adress, out string _memory, out string _date, out string _time, out string _answer, out string _crc)
        {
            string mem1, mem2;
            string year, month, day;
            string hour, minute;
            string crc1, crc2;

            if (Convert.ToString(_data[0].ToString("X")).Length == 1)
            {
                _adress = "0" + Convert.ToString(_data[0].ToString("X"));
            }
            else
            {
                _adress = Convert.ToString(_data[0].ToString("X"));
            }

            if (Convert.ToString(_data[1].ToString("X")).Length == 1)
            {
                mem1 = "0" + Convert.ToString(_data[1].ToString("X"));
            }
            else
            {
                mem1 = Convert.ToString(_data[1].ToString("X"));
            }

            if (Convert.ToString(_data[2].ToString("X")).Length == 1)
            {
                mem2 = "0" + Convert.ToString(_data[2].ToString("X"));
            }
            else
            {
                mem2 = Convert.ToString(_data[2].ToString("X"));
            }

            _memory = mem1 + mem2;

            if (Convert.ToString(_data[6].ToString("X")).Length == 1)
            {
                year = "0" + Convert.ToString(_data[6].ToString("X"));
            }
            else
            {
                year = Convert.ToString(_data[6].ToString("X"));
            }

            if (Convert.ToString(_data[7].ToString("X")).Length == 1)
            {
                month = "0" + Convert.ToString(_data[7].ToString("X"));
            }
            else
            {
                month = Convert.ToString(_data[7].ToString("X"));
            }

            if (Convert.ToString(_data[8].ToString("X")).Length == 1)
            {
                day = "0" + Convert.ToString(_data[8].ToString("X"));
            }
            else
            {
                day = Convert.ToString(_data[8].ToString("X"));
            }

            _date = year + "." + month + "." + day;

            if (Convert.ToString(_data[4].ToString("X")).Length == 1)
            {
                hour = "0" + Convert.ToString(_data[4].ToString("X"));
            }
            else
            {
                hour = Convert.ToString(_data[4].ToString("X"));
            }

            if (Convert.ToString(_data[5].ToString("X")).Length == 1)
            {
                minute = "0" + Convert.ToString(_data[5].ToString("X"));
            }
            else
            {
                minute = Convert.ToString(_data[5].ToString("X"));
            }

            _time = hour + ":" + minute;
            _answer = "";

            if (Convert.ToString(_data[10].ToString("X")).Length == 1)
            {
                crc1 = "0" + Convert.ToString(_data[10].ToString("X"));
            }
            else
            {
                crc1 = Convert.ToString(_data[10].ToString("X"));
            }

            if (Convert.ToString(_data[11].ToString("X")).Length == 1)
            {
                crc2 = "0" + Convert.ToString(_data[11].ToString("X"));
            }
            else
            {
                crc2 = Convert.ToString(_data[11].ToString("X"));
            }

            _crc = crc1 + crc2;
        }

        void prof_explode(byte[] _data, out string _adress_prof, out string _time_prof, out string _date_prof, out string _period_prof, out string _active_prof, out string _reactive_prof, out string _crc)
        {
            _adress_prof = "";
            _time_prof = "";
            _date_prof = "";
            _period_prof = "";
            _active_prof = "";
            _reactive_prof = "";
            _crc = "";

            string year, month, day;
            string hour, minute;
            string _active_prof1, _active_prof2;
            string _reactive_prof1, _reactive_prof2;

            if (Convert.ToString(_data[6].ToString("X")).Length == 1)
            {
                year = "0" + Convert.ToString(_data[6].ToString("X"));
            }
            else
            {
                year = Convert.ToString(_data[6].ToString("X"));
            }

            if (Convert.ToString(_data[5].ToString("X")).Length == 1)
            {
                month = "0" + Convert.ToString(_data[5].ToString("X"));
            }
            else
            {
                month = Convert.ToString(_data[5].ToString("X"));
            }

            if (Convert.ToString(_data[4].ToString("X")).Length == 1)
            {
                day = "0" + Convert.ToString(_data[4].ToString("X"));
            }
            else
            {
                day = Convert.ToString(_data[4].ToString("X"));
            }

            _date_prof = "20" + year + month + day;


            if (Convert.ToString(_data[2].ToString("X")).Length == 1)
            {
                hour = "0" + Convert.ToString(_data[2].ToString("X"));
            }
            else
            {
                hour = Convert.ToString(_data[2].ToString("X"));
            }

            if (Convert.ToString(_data[3].ToString("X")).Length == 1)
            {
                minute = "0" + Convert.ToString(_data[3].ToString("X"));
            }
            else
            {
                minute = Convert.ToString(_data[3].ToString("X"));
            }

            _time_prof = hour + minute + "00";

            if (Convert.ToString(_data[9].ToString("X")).Length == 1)
            {
                _active_prof1 = "0" + Convert.ToString(_data[9].ToString("X"));
            }
            else
            {
                _active_prof1 = Convert.ToString(_data[9].ToString("X"));
            }

            if (Convert.ToString(_data[8].ToString("X")).Length == 1)
            {
                _active_prof2 = "0" + Convert.ToString(_data[8].ToString("X"));
            }
            else
            {
                _active_prof2 = Convert.ToString(_data[8].ToString("X"));
            }

            byte[] act_prof = { Convert.ToByte(Convert.ToInt32(_active_prof2, 16)), Convert.ToByte(Convert.ToInt32(_active_prof1, 16)) };

            UInt32 fl1 = 0;

            fl1 = BitConverter.ToUInt16(act_prof, 0);

            float fl2 = Convert.ToSingle(fl1) / 1000;

            _active_prof = Convert.ToString(Convert.ToSingle(fl1) / 1000);


            if (Convert.ToString(_data[13].ToString("X")).Length == 1)
            {
                _reactive_prof1 = "0" + Convert.ToString(_data[13].ToString("X"));
            }
            else
            {
                _reactive_prof1 = Convert.ToString(_data[13].ToString("X"));
            }

            if (Convert.ToString(_data[12].ToString("X")).Length == 1)
            {
                _reactive_prof2 = "0" + Convert.ToString(_data[12].ToString("X"));
            }
            else
            {
                _reactive_prof2 = Convert.ToString(_data[12].ToString("X"));
            }

            byte[] react_prof = { Convert.ToByte(Convert.ToInt32(_reactive_prof2, 16)), Convert.ToByte(Convert.ToInt32(_reactive_prof1, 16)) };

            fl1 = 0;
            fl2 = 0;
            fl1 = BitConverter.ToUInt16(react_prof, 0);

            fl2 = Convert.ToSingle(fl1) / 1000;

            _reactive_prof = Convert.ToString(Convert.ToSingle(fl1) / 1000);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            int interval = 16;
            //string adr_req = "42";
            //string cod_req = "0603";
            //string num_req = "0F";

            string conn_str = "Database=resources;Data Source=10.1.1.50;User Id=user;Password=password";
            MySqlLib.MySqlData.MySqlExecute.MyResult result = new MySqlLib.MySqlData.MySqlExecute.MyResult();

            result = MySqlLib.MySqlData.MySqlExecute.SqlScalar("SELECT electro56_datetime FROM resources.`electro56` ORDER BY electro56_datetime DESC LIMIT 0,1", conn_str);
            string date_from_mysql = result.ResultText;

            string[] portnames = SerialPort.GetPortNames();
            SerialPort port = new SerialPort("COM4", 9600, Parity.None, 8, StopBits.One);
            byte[] buff = { 0x00, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x77, 0x81 };
            //byte[] buff = { 0x42, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01, 0x01 };
            byte[] buff2 = { 0x56, 0x08, 0x13, 0xD7, 0xDD };
            byte[] req = { };

            DateTime date1 = DateTime.Now;
            DateTime date2 = DateTime.Now;

            if (date_from_mysql.Length == 19)
            {
                date2 = new DateTime(Convert.ToInt32(date_from_mysql.Substring(6, 4)), Convert.ToInt32(date_from_mysql.Substring(3, 2)), Convert.ToInt32(date_from_mysql.Substring(0, 2)), Convert.ToInt32(date_from_mysql.Substring(11, 2)), Convert.ToInt32(date_from_mysql.Substring(14, 2)), Convert.ToInt32(date_from_mysql.Substring(17, 2)));
            }
            if (date_from_mysql.Length == 18)
            {
                date2 = new DateTime(Convert.ToInt32(date_from_mysql.Substring(6, 4)), Convert.ToInt32(date_from_mysql.Substring(3, 2)), Convert.ToInt32(date_from_mysql.Substring(0, 2)), Convert.ToInt32(date_from_mysql.Substring(11, 1)), Convert.ToInt32(date_from_mysql.Substring(13, 2)), Convert.ToInt32(date_from_mysql.Substring(16, 2)));
            }



            TimeSpan time_interval = date1 - date2;

            int k = 0;



            port.Open();

            port.Write(buff, 0, 11);
            Thread.Sleep(1000);
            int byteRecieved = port.BytesToRead;
            byte[] messByte = new byte[byteRecieved];
            port.Read(messByte, 0, byteRecieved);
            Thread.Sleep(1000);

            port.Write(buff2, 0, 5);
            Thread.Sleep(1000);
            int byteRecieved2 = port.BytesToRead;
            byte[] messByte2 = new byte[byteRecieved2];
            port.Read(messByte2, 0, byteRecieved2);
            Thread.Sleep(1000);


            string adress, memory, date, time, answer, crc;
            ans_explode(messByte2, out adress, out memory, out date, out time, out answer, out crc);


            int num = Int32.Parse(memory, System.Globalization.NumberStyles.HexNumber);
            string hexValue = num.ToString("X");


            port.Close();



            byte[] byte_mem = { byte.Parse(hexValue.Substring(2, 2), System.Globalization.NumberStyles.HexNumber), byte.Parse(hexValue.Substring(0, 2), System.Globalization.NumberStyles.HexNumber) };

            int m = 0;
            int req_m = 0;
            m = BitConverter.ToUInt16(byte_mem, 0);


            port.Open();
            for (k = Convert.ToInt32(Math.Floor(time_interval.TotalMinutes / 30)); k >= 1; k--)
            {


                req_m = m - interval * (k + 1);

                if (req_m <= 0)
                {
                    req_m = 65536 + m - interval * (k + 1);
                }
                byte[] intBytes = BitConverter.GetBytes(req_m);
                Array.Reverse(intBytes);
                byte[] result_mem = intBytes;

                byte[] data111 = { 0x56, 0x06, 0x03, result_mem[2], result_mem[3], 0x0F };
                req = GetCRC(data111);

                port.Write(req, 0, req.Length);
                Thread.Sleep(1000);
                int byteRecieved4 = port.BytesToRead;
                byte[] messByte4 = new byte[byteRecieved4];
                port.Read(messByte4, 0, byteRecieved4);
                Thread.Sleep(1000);

                string adress_prof, time_prof, date_prof, period_prof, active_prof, reactive_prof, crc_prof;
                prof_explode(messByte4, out adress_prof, out time_prof, out date_prof, out period_prof, out active_prof, out reactive_prof, out crc_prof);


                char[] chars = active_prof.ToCharArray();
                for (int n = 0; n < active_prof.Length; n++)
                {
                    if (chars[n] == ',')
                    {
                        chars[n] = '.';
                    }
                }
                string active_prof_str = new string(chars);

                char[] chars2 = reactive_prof.ToCharArray();
                for (int n = 0; n < reactive_prof.Length; n++)
                {
                    if (chars2[n] == ',')
                    {
                        chars2[n] = '.';
                    }
                }
                string reactive_prof_str = new string(chars2);

                result = MySqlLib.MySqlData.MySqlExecute.SqlNoneQuery("INSERT INTO resources.`electro56` (`electro56_datetime`,`electro56_active`,`electro56_reactive`) VALUES (" + "'" + date_prof + time_prof + "'" + "," + "'" + active_prof_str + "'" + "," + "'" + reactive_prof_str + "'" + ")", conn_str);


            }
            port.Close();

            Application.Exit();
        }
    }
}


namespace MySqlLib
{
    /// <summary>
    /// Набор компонент для простой работы с MySQL базой данных.
    /// </summary>
    public class MySqlData
    {

        /// <summary>
        /// Методы реализующие выполнение запросов с возвращением одного параметра либо без параметров вовсе.
        /// </summary>
        public class MySqlExecute
        {

            /// <summary>
            /// Возвращаемый набор данных.
            /// </summary>
            public class MyResult
            {
                /// <summary>
                /// Возвращает результат запроса.
                /// </summary>
                public string ResultText;
                /// <summary>
                /// Возвращает True - если произошла ошибка.
                /// </summary>
                public string ErrorText;
                /// <summary>
                /// Возвращает текст ошибки.
                /// </summary>
                public bool HasError;
            }

            /// <summary>
            /// Для выполнения запросов к MySQL с возвращением 1 параметра.
            /// </summary>
            /// <param name="sql">Текст запроса к базе данных</param>
            /// <param name="connection">Строка подключения к базе данных</param>
            /// <returns>Возвращает значение при успешном выполнении запроса, текст ошибки - при ошибке.</returns>
            public static MyResult SqlScalar(string sql, string connection)
            {
                MyResult result = new MyResult();
                try
                {
                    MySql.Data.MySqlClient.MySqlConnection connRC = new MySql.Data.MySqlClient.MySqlConnection(connection);
                    MySql.Data.MySqlClient.MySqlCommand commRC = new MySql.Data.MySqlClient.MySqlCommand(sql, connRC);
                    connRC.Open();
                    try
                    {
                        result.ResultText = commRC.ExecuteScalar().ToString();
                        result.HasError = false;
                    }
                    catch (Exception ex)
                    {
                        result.ErrorText = ex.Message;
                        result.HasError = true;

                    }
                    connRC.Close();
                }
                catch (Exception ex)//Этот эксепшн на случай отсутствия соединения с сервером.
                {
                    result.ErrorText = ex.Message;
                    result.HasError = true;
                }
                return result;
            }


            /// <summary>
            /// Для выполнения запросов к MySQL без возвращения параметров.
            /// </summary>
            /// <param name="sql">Текст запроса к базе данных</param>
            /// <param name="connection">Строка подключения к базе данных</param>
            /// <returns>Возвращает True - ошибка или False - выполнено успешно.</returns>
            public static MyResult SqlNoneQuery(string sql, string connection)
            {
                MyResult result = new MyResult();
                try
                {
                    MySql.Data.MySqlClient.MySqlConnection connRC = new MySql.Data.MySqlClient.MySqlConnection(connection);
                    MySql.Data.MySqlClient.MySqlCommand commRC = new MySql.Data.MySqlClient.MySqlCommand(sql, connRC);
                    connRC.Open();
                    try
                    {
                        commRC.ExecuteNonQuery();
                        result.HasError = false;
                    }
                    catch (Exception ex)
                    {
                        result.ErrorText = ex.Message;
                        result.HasError = true;
                    }
                    connRC.Close();
                }
                catch (Exception ex)//Этот эксепшн на случай отсутствия соединения с сервером.
                {
                    result.ErrorText = ex.Message;
                    result.HasError = true;
                }
                return result;
            }

        }
        /// <summary>
        /// Методы реализующие выполнение запросов с возвращением набора данных.
        /// </summary>
        public class MySqlExecuteData
        {
            /// <summary>
            /// Возвращаемый набор данных.
            /// </summary>
            public class MyResultData
            {
                /// <summary>
                /// Возвращает результат запроса.
                /// </summary>
                public DataTable ResultData;
                /// <summary>
                /// Возвращает True - если произошла ошибка.
                /// </summary>
                public string ErrorText;
                /// <summary>
                /// Возвращает текст ошибки.
                /// </summary>
                public bool HasError;
            }
            /// <summary>
            /// Выполняет запрос выборки набора строк.
            /// </summary>
            /// <param name="sql">Текст запроса к базе данных</param>
            /// <param name="connection">Строка подключения к базе данных</param>
            /// <returns>Возвращает набор строк в DataSet.</returns>
            public static MyResultData SqlReturnDataset(string sql, string connection)
            {
                MyResultData result = new MyResultData();
                try
                {
                    MySql.Data.MySqlClient.MySqlConnection connRC = new MySql.Data.MySqlClient.MySqlConnection(connection);
                    MySql.Data.MySqlClient.MySqlCommand commRC = new MySql.Data.MySqlClient.MySqlCommand(sql, connRC);
                    connRC.Open();
                    try
                    {
                        MySql.Data.MySqlClient.MySqlDataAdapter AdapterP = new MySql.Data.MySqlClient.MySqlDataAdapter();
                        AdapterP.SelectCommand = commRC;
                        DataSet ds1 = new DataSet();
                        AdapterP.Fill(ds1);
                        result.ResultData = ds1.Tables[0];
                    }
                    catch (Exception ex)
                    {
                        result.HasError = true;
                        result.ErrorText = ex.Message;
                    }
                    connRC.Close();
                }
                catch (Exception ex)//Этот эксепшн на случай отсутствия соединения с сервером.
                {
                    result.ErrorText = ex.Message;
                    result.HasError = true;
                }
                return result;
            }
        }
    }
}