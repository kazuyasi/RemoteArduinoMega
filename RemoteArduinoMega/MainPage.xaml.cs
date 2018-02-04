using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Devices.Enumeration;
using Windows.Devices.SerialCommunication;
using Windows.Storage.Streams;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using System.Text;


// 空白ページの項目テンプレートについては、https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x411 を参照してください

namespace RemoteArduinoMega
{
    /// <summary>
    /// それ自体で使用できる空白ページまたはフレーム内に移動できる空白ページ。
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// COMポートで指定するシリアルデバイス（StandardFirmataPlusを書き込んだArduino Mega）
        /// </summary>
        SerialDevice device;

        /// <summary>
        /// Firmataのプロトコルバージョン
        /// </summary>
        double version;

        /// <summary>
        /// Arduinoに書き込まれたスケッチ名。
        /// </summary>
        string name;

        /// <summary>
        /// テスト場所。
        /// 末尾にブレイクをかけると各変数の返り値がわかってよさげ。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void OnLoaded(object sender, RoutedEventArgs e)
        {
            await InitilizeFirmata();

            //起動時送信文字列の読み出し。
            var task = await ReadAsync();
            var msg = ExtractSysexMessage(task);
            name = ExtractStringFromSysexMessage(3, msg);
            version = msg[1] + msg[2] / 10.0;

            var fuga = await ReadDigital(3);

            //UART初期化と読み込み状態の定義を行う。
            InitializedUART(1, 19200);
            ReadUART(1, 256);           //これがないと受信しない。最初はreadToBytesが0

            //UART書き込みを行う。
            WriteUART(1, new byte[] { 0x12, 0x34, 0x56, 0x78, 0x9A, 0xBC, 0xDE, 0xF0 });
            FlushUART(1);

            //TX1とRX1をショートさせて強制的にループバックした値を読み出す。
            byte[] hoge = ExtractUARTReplay(await ReadAsync());

            //文字列の送受信はこう。
            WriteUART(1, Encoding.ASCII.GetBytes("TANOSI-"));
            FlushUART(1);
            byte[] piyo = ExtractUARTReplay(await ReadAsync());
            var foo = Encoding.ASCII.GetString(piyo);
        }


        /// <summary>
        /// 以下は移植可能なコアコード。
        /// </summary>
        #region CORE
        /// <summary>
        /// UARTからの受信データを抽出する。
        /// </summary>
        /// <param name="rawData">生の受信データ</param>
        /// <returns>UARTにて受信されたデータ</returns>
        private byte[] ExtractUARTReplay(byte[] rawData)
        {
            List<byte> list = new List<byte>();
            List<byte> subList = new List<byte>();

            int idx = 0;
            while (idx < rawData.Length)
            {
                var start = Array.FindIndex(rawData, idx, s => s == 0xF0);
                if(start != -1)
                {
                    idx = start + 3;
                    var end = Array.FindIndex(rawData, start, s => s == 0xF7);

                    if (end != -1)
                    {
                        for(; idx < end; idx += 2)
                        {
                            subList.Add(rawData[idx]);
                            subList.Add(rawData[idx + 1]);
                        }
                    }
                    else break;
                }
                else break;
            }

            for (int i = 0; i < subList.Count; i += 2) list.Add((byte)(subList[i] + (subList[i + 1] << 7)));

            return list.ToArray();
        }

        /// <summary>
        /// UARTにバイト配列を書き込む。
        /// フラッシュしないと確定しないので注意。
        /// </summary>
        /// <param name="port">実行するポート</param>
        /// <param name="data">送信するデータ</param>
        private async void WriteUART(byte port, byte[] data)
        {
            List<byte> list = new List<byte>();

            if(data != null)
            {
                foreach (var d in data)
                {
                    list.Add(Bits(d, 0, 6));
                    list.Add(Bits(d, 7, 7));
                }
            }

            await WriteAsync(FormatUART(list.ToArray(), 0x20, port));
        }

        /// <summary>
        /// UARTを初期化し、オープンする。
        /// </summary>
        /// <param name="port">実行するポート</param>
        /// <param name="baud">ボーレート</param>
        private async void InitializedUART(byte port, int baud)
        {
            //ボーレートは正の数(上限も決めたほうがいいかも)
            if (baud < 0) throw new FormatException();

            await WriteAsync(FormatUART(new byte[] { Bits(baud, 0, 6), Bits(baud, 7, 13), Bits(baud, 14, 20) }, 0x10, port));
        }

        /// <summary>
        /// UARTのフラッシュを実行する。
        /// </summary>
        /// <param name="port">実行するポート</param>
        private async void FlushUART(byte port)
        {
            await WriteAsync( FormatUART(null, 0x60, port) );
        }

        /// <summary>
        /// UARTを読み込み状態にし、読み込み最大バイトを指定する。
        /// 実際に値を読み出すわけではないので注意。
        /// </summary>
        /// <param name="port">実行するポート</param>
        /// <param name="maxBytes">読み込み最大バイト</param>
        private async void ReadUART(byte port, ushort maxBytes)
        {
            //最大読み取りバイト数は最大0x3FFF
            if (maxBytes >= 0x4000) throw new FormatException();

            await WriteAsync(FormatUART(new byte[] { 0, Bits(maxBytes, 0, 6), Bits(maxBytes, 7, 13) }, 0x30, port));
        }

        /// <summary>
        /// UARTコマンド用にフォーマットする。
        /// </summary>
        /// <param name="buf">データ</param>
        /// <param name="cmd">コマンドコード</param>
        /// <param name="port">指定ポート</param>
        /// <returns>フォーマット後のデータ</returns>
        private byte[] FormatUART(byte[] buf, byte cmd, byte port)
        {
            List<byte> list = new List<byte> { MakeCode(cmd, port) }; //コマンドコード

            if(buf != null) list.AddRange(buf.AsEnumerable());

            return FortmatSysexMessage(list.ToArray(), 0x60);
        }

        /// <summary>
        /// 指定のデジタルポートを読む。
        /// </summary>
        /// <param name="port">指定ポート</param>
        /// <returns>デジタル</returns>
        private async Task<byte> ReadDigital(byte port)
        {
            await WriteAsync(new byte[] { MakeCode(0xD0, port), 1 });
            var buf = await ReadAsync();

            return ComposeByte(buf, 1);
        }

        /// <summary>
        /// LSBとMSBを結合して一つのbyteにする。
        /// </summary>
        /// <param name="buf">結合元</param>
        /// <param name="idx">結合開始インデックス</param>
        /// <returns>結合後のバイト</returns>
        private byte ComposeByte(byte[] buf, int idx)
        {
            return (byte)(buf[idx] + (buf[idx + 1] & 1) << 7);
        }

        /// <summary>
        /// LSBとMSBを結合して一つのwordにする。
        /// </summary>
        /// <param name="buf"></param>
        /// <param name="idx"></param>
        /// <returns></returns>
        private ushort ComposeWord(byte[] buf, int idx)
        {
            return (ushort)(buf[idx] + Bits(buf[idx + 1], 0, 6) << 7);
        }

        /// <summary>
        /// コードとポートを融合する。
        /// </summary>
        /// <param name="code">コード</param>
        /// <param name="port">ポート</param>
        /// <returns>実際のコマンドコード</returns>
        private byte MakeCode(byte code, byte port)
        {
            //コードは16進数で表したとき下１桁0
            if (code % 0x10 != 0) throw new FormatException();

            //ポートは16進数で表したとき0x10未満
            if (port >= 0x10 ) throw new FormatException();

            return (byte)(code | port);
        }

        /// <summary>
        /// 生データをSYSEXメッセージ形式にする。
        /// </summary>
        /// <param name="rawdata"></param>
        /// <returns></returns>
        private byte[] FortmatSysexMessage(byte[] rawdata)
        {
            List<byte> list = new List<byte> { 0xF0 }; //開始バイト

            //中に含まれる生データは0x80未満
            if (rawdata.Any(s => s >= 0x80)) throw new FormatException();

            if (rawdata != null) list.AddRange(rawdata.AsEnumerable());

            list.Add(0xF7); //終了バイト

            return list.ToArray();
        }

        /// <summary>
        /// 生データをコマンドコード付きでSYSEXメッセージ形式にする。
        /// </summary>
        /// <param name="rawdata"></param>
        /// <param name="code"></param>
        /// <returns></returns>
        private byte[] FortmatSysexMessage(byte[] rawdata, byte code)
        {
            List<byte> list = new List<byte> { code }; //先頭にコード埋め込み

            list.AddRange(rawdata.AsEnumerable());

            return FortmatSysexMessage(list.ToArray());
        }

        /// <summary>
        /// 受信したSYSEXメッセージから文字列を抽出する。
        /// </summary>
        /// <param name="idx">読み取り開始インデックス</param>
        /// <param name="msg">SYSEXメッセージ</param>
        /// <returns>抽出した文字列</returns>
        private string ExtractStringFromSysexMessage(int idx, byte[] msg)
        {
            if (msg == null) return null;
            if (msg.Length < idx) return null;
            if ((msg.Length - idx) % 2 != 0) return null;

            string nm = string.Empty;

            for (int i = idx; i < msg.Length; i += 2)
            {
                nm += (char)(msg[i] + ((1 & msg[i + 1]) << 7));
            }

            return nm;
        }

        /// <summary>
        /// 特定の値から指定ビットを取り出す。
        /// </summary>
        /// <param name="value">値</param>
        /// <param name="startBit">開始ビット</param>
        /// <param name="endBit">終了ビット</param>
        /// <returns>ぬきだしたビット</returns>
        private byte Bits(int value, byte startBit, byte endBit)
        {
            if (startBit > endBit) throw new ArgumentException();
            if ((endBit - startBit) > 7) throw new ArgumentException();

            byte ret = (byte)(value >> startBit);

            byte mask = 0;
            switch (endBit - startBit)
            {
                case 0: mask = 0b00000001; break;
                case 1: mask = 0b00000011; break;
                case 2: mask = 0b00000111; break;
                case 3: mask = 0b00001111; break;
                case 4: mask = 0b00011111; break;
                case 5: mask = 0b00111111; break;
                case 6: mask = 0b01111111; break;
                case 7: mask = 0b11111111; break;
            }

            ret &= mask;

            return ret;
        }

        /// <summary>
        /// Firmataの通信からメッセージを受信する。
        /// </summary>
        /// <param name="rawData">生の受信データ</param>
        /// <returns>開始バイトと終了バイトで挟まれたデータ配列</returns>
        private byte[] ExtractSysexMessage(byte[] rawData)
        {
            if (rawData == null) return null;

            var start = Array.FindIndex(rawData, s => s == 0xF0);
            if (start != -1)
            {
                var end = Array.FindIndex(rawData, start, s => s == 0xF7);

                if (end != -1) return rawData.Skip(start + 1).Take(end - start - 1).ToArray();
                return null;
            }

            return null;
        }

        /// <summary>
        /// Firmataからのデータを受信する。
        /// </summary>
        /// <returns></returns>
        private async Task<byte[]> ReadAsync()
        {
            DataReader reader = new DataReader(device.InputStream);
            
            await reader.LoadAsync(0x80);
            List<byte> list = new List<byte>();
            while (reader.UnconsumedBufferLength != 0)
            {
                list.Add(reader.ReadByte());
            }

            return list.ToArray();
        }

        /// <summary>
        /// Firmataへバイト配列を送信する。
        /// </summary>
        /// <param name="data">送信するバイト配列</param>
        /// <returns>タスク</returns>
        private async Task WriteAsync(byte[] data)
        {
            DataWriter writer = new DataWriter(device.OutputStream);
            writer.WriteBytes(data);
            await writer.StoreAsync();
        }

        /// <summary>
        /// Firmataとの通信を確立し、リセットをかける。
        /// 確立できない場合、アプリケーションを終了する。
        /// </summary>
        /// <param name="portName">通信したいポートネーム</param>
        /// <returns></returns>
        private async Task InitilizeFirmata(string portName)
        {
            var myDevices = await DeviceInformation.FindAllAsync(SerialDevice.GetDeviceSelector(portName));

            if (myDevices.Count == 0) Application.Current.Exit();
            await Initialize(myDevices);
        }

        /// <summary>
        /// initializeFirmataのサブメソッド。
        /// デバイスを指定してからの処理を行う。
        /// </summary>
        /// <param name="myDevices"></param>
        /// <returns></returns>
        private async Task Initialize(DeviceInformationCollection myDevices)
        {
            device = await SerialDevice.FromIdAsync(myDevices[0].Id);
            device.BaudRate = 57600;
            device.DataBits = 8;
            device.StopBits = SerialStopBitCount.One;
            device.Parity = SerialParity.None;
            device.Handshake = SerialHandshake.None;
            device.ReadTimeout = TimeSpan.FromMilliseconds(1000);
            device.WriteTimeout = TimeSpan.FromMilliseconds(1000);

            //リセットをHW的にかけ、起動するまで待つ。
            device.IsDataTerminalReadyEnabled = true;   //ArduinoはDTR:OFF->ONでリセットがかかる。
            await Task.Delay(3000);
            device.IsDataTerminalReadyEnabled = false;
        }

        /// <summary>
        /// Firmataとの通信を確立し、リセットをかける。
        /// 確立できない場合、アプリケーションを終了する。検索対象はArduino Mega
        /// </summary>
        /// <returns></returns>
        private async Task InitilizeFirmata()
        {
            var myDevices = await DeviceInformation.FindAllAsync(SerialDevice.GetDeviceSelectorFromUsbVidPid(0x2341, 0x0042));

            if (myDevices.Count == 0) Application.Current.Exit();

            await Initialize(myDevices);
        }
        #endregion 
    }
}
