using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SQLite;
using NDde.Client;


namespace RSSCacheSaver2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private Dictionary<string, string> _item;
        private SQLiteConnection _connection;
        private DdeClient client;

        private long prevVolume = 0;
        private double preBid = 0;
        private double preAsk = 0;
        private long tick = 0;

        // 起動オプション
        private Options opts = new Options();
        private long _count = 0; // 秒間リクエストの計測
        private string _name = "";

        internal class CurrentTimeGetter
        {
            private static int lastTicks = -1;
            private static DateTime lastDateTime = DateTime.MinValue;

            /// <summary>        
            /// Gets the current time in an optimized fashion.        
            /// </summary>        
            /// <value>Current time.</value>        

            public static DateTime Now
            {
                get
                {
                    int tickCount = Environment.TickCount;
                    if (tickCount == lastTicks)
                    {
                        return lastDateTime;
                    }
                    DateTime dt = DateTime.Now;
                    lastTicks = tickCount;
                    lastDateTime = dt;
                    return dt;
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            txtCode.Text = "3697";
            lblCount.Text = "";
            Options();
            if (opts.Code != "" && opts.Code != null) { txtCode.Text = opts.Code; }
        }

        /// <summary>
        /// DDEClientからのコールバック関数
        /// </summary>
        private void OnAdvise(object sender, DdeAdviseEventArgs args)
        {
            string raw = Encoding.Default.GetString(args.Data).Trim('\0', ' ').ToString();
            if (raw == "") return;

            if (_item.TryGetValue(args.Item, out string value))
            {
                using (SQLiteCommand cmd = _connection.CreateCommand())
                {
                    string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    //string now = CurrentTimeGetter.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    cmd.CommandText =
                        $"insert into rss(Time,Item,Value) values ('{now}','{args.Item}','{raw}')";
                    cmd.ExecuteNonQuery();

                    string k = $"{args.Item}";
                    string v = $"{raw}";
                    _item[k] = v;

                    switch (k)
                    {
                        case "出来高":
                            long volume = long.Parse(v);
                            tick = volume - prevVolume;
                            prevVolume = volume;
                            break;
                        case "出来高加重平均":
                            if (tick > 0)
                            {
                                double price = 0;
                                if (_item["現在値"] != "")
                                {
                                    price = double.Parse(_item["現在値"]);
                                }
                                double bid = price;
                                double ask = price;

                                if (_item["最良売気配値１"] != "")
                                {
                                    bid = double.Parse(_item["最良売気配値１"]);
                                }
                                if (_item["最良買気配値１"] != "")
                                {
                                    ask = double.Parse(_item["最良買気配値１"]);
                                }

                                string kind = "";
                                if (preBid - price < price - preAsk)
                                {
                                    kind = "R";
                                }
                                else
                                {
                                    kind = "G";
                                }
                                preBid = bid;
                                preAsk = ask;

                                var t = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds();
                                //cmd.CommandText = $"insert into tick (No, Time,Tick,Kind,Vwap,Bid,Ask,Price) values (" +
                                //    $"'{t}','{now}','{tick}','{kind}','{v}','{bid}','{ask}','{price}')";
                                cmd.CommandText = $"insert into tick (Time,Tick,Kind,Vwap,Bid,Ask,Price) values (" +
                                    $"'{now}','{tick}','{kind}','{v}','{bid}','{ask}','{price}')";
                                cmd.ExecuteNonQuery();
                            }
                            break;
                    }

                    _count += 1;
                }
            }
        }

        /// <summary>
        /// DDEClientからのコールバック関数
        /// </summary>
        private void OnDisconnected(object sender, DdeDisconnectedEventArgs args)
        {
            _connection.Close();
            Console.WriteLine("OnDisconnected: " + args.IsServerInitiated.ToString() + " " +
                "" + args.IsDisposed.ToString());
        }

        /// <summary>
        /// 日毎のデータベースを作成する
        /// </summary>
        /// <param name="code"></param>
        private void CreateDatabase(string code)
        {
            //TODO

            string path = $"{DateTime.Now.ToString("yyyyMMdd")}_RSS_{code}.db"; ;
            System.Diagnostics.Debug.WriteLine($"{path}");

            using (SQLiteConnection connection = new SQLiteConnection("Data Source=" + path))
            {
                connection.Open();
                using (SQLiteCommand cmd = connection.CreateCommand())
                {
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS RSS(" +
                        "No INTEGER NOT NULL PRIMARY KEY," +
                        "Time TEXT NOT NULL," +
                        "Item TEXT NOT NULL," +
                        "Value REAL NOT NULL);";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS Tick(" +
                        "No INTEGER NOT NULL PRIMARY KEY," +
                        "Time TEXT NOT NULL," +
                        "Tick INTEGER NOT NULL," +
                        "Kind TEXT NOT NULL," + // Yellow, Red, Green
                        "Vwap REAL NOT NULL," +
                        "Bid REAL NOT NULL," +
                        "Ask REAL NOT NULL," +
                        "Price REAL NOT NULL);";
                    cmd.ExecuteNonQuery();
                }
            }

            //_connection = new SQLiteConnection("Data Source=" + path);
            CreateConnection(path);
        }

        /// <summary>
        /// 2019.10.19
        /// SQLite3 へのコネクションを開く。同期オフ、ジャーナルモードWAL 
        /// </summary>
        /// <param name="path"></param>
        private void CreateConnection(string path)
        {
            var builder = new SQLiteConnectionStringBuilder()
            {
                DataSource = path,
                Version = 3,
                LegacyFormat = false,
                SyncMode = SynchronizationModes.Normal, // off は途中で落ちる
                JournalMode = SQLiteJournalModeEnum.Wal
            };
            _connection = new SQLiteConnection(builder.ToString());
        }

        /// <summary>
        /// 自動終了処理
        /// </summary>
        private void IsRestTime()
        {
            // 現在時を取得
            DateTime d = DateTime.Now;
            //時間を設定
            DateTime d_0900 = new DateTime(d.Year, d.Month, d.Day, 08, 50, 30); //年, 月, 日, 時間, 分, 秒
            DateTime d_1130 = new DateTime(d.Year, d.Month, d.Day, 11, 30, 30); //年, 月, 日, 時間, 分, 秒
            DateTime d_1230 = new DateTime(d.Year, d.Month, d.Day, 12, 29, 30); //年, 月, 日, 時間, 分, 秒
            DateTime d_1500 = new DateTime(d.Year, d.Month, d.Day, 15, 00, 30); //年, 月, 日, 時間, 分, 秒
            DateTime d_1505 = new DateTime(d.Year, d.Month, d.Day, 15, 05, 00); //年, 月, 日, 時間, 分, 秒

            //現在の時間が設定の時間になった時の処理
            if ((d_1230 >= d & d >= d_1130) | (d_0900 >= d | d >= d_1500))
            {
                //背景を黒くする
                this.BackColor = Color.Blue;
            }
            else
            {
                this.BackColor = SystemColors.Control;
            }
            if (d >= d_1505) { Application.Exit(); }
        }

        /// <summary>
        /// 監視開始関数
        /// </summary>
        private void Start()
        {
            _item = new Dictionary<string, string>()
            {
                //{ "最良売気配値１", ""},
                //{ "最良買気配値１", ""},
                //{ "最良買気配数量１", ""},
                //{ "最良売気配数量１", ""},
                { "現在値", ""},
                { "出来高", ""},
                { "出来高加重平均", ""},
                { "売買代金", ""},
                
                {"最良売気配値１" ,""},
                {"最良売気配値２" ,""},
                {"最良売気配値３" ,""},
                {"最良売気配値４" ,""},
                {"最良売気配値５" ,""},
                {"最良売気配値６" ,""},
                {"最良売気配値７" ,""},
                {"最良売気配値８" ,""},
                {"最良売気配値９" ,""},
                {"最良売気配値１０" ,""},
                {"最良買気配値１" ,""},
                {"最良買気配値２" ,""},
                {"最良買気配値３" ,""},
                {"最良買気配値４" ,""},
                {"最良買気配値５" ,""},
                {"最良買気配値６" ,""},
                {"最良買気配値７" ,""},
                {"最良買気配値８" ,""},
                {"最良買気配値９" ,""},
                {"最良買気配値１０" ,""},
                {"最良売気配数量１" ,""},
                {"最良売気配数量２" ,""},
                {"最良売気配数量３" ,""},
                {"最良売気配数量４" ,""},
                {"最良売気配数量５" ,""},
                {"最良売気配数量６" ,""},
                {"最良売気配数量７" ,""},
                {"最良売気配数量８" ,""},
                {"最良売気配数量９" ,""},
                {"最良売気配数量１０" ,""},
                {"最良買気配数量１" ,""},
                {"最良買気配数量２" ,""},
                {"最良買気配数量３" ,""},
                {"最良買気配数量４" ,""},
                {"最良買気配数量５" ,""},
                {"最良買気配数量６" ,""},
                {"最良買気配数量７" ,""},
                {"最良買気配数量８" ,""},
                {"最良買気配数量９" ,""},
                {"最良買気配数量１０" ,""},
                {"OVER気配数量" ,""},
                {"UNDER気配数量" ,""},
            };

            // DB作成
            CreateDatabase(txtCode.Text);

            // 監視開始
            client = new DdeClient("RSS", $"{txtCode.Text}.T", this);
            client.Disconnected += new EventHandler<DdeDisconnectedEventArgs>(OnDisconnected);
            client.Connect();
            foreach (string key in _item.Keys)
            {
                client.StartAdvise(key, 1, true, 60000);
                _count += 1;
            }
            client.Advise += new EventHandler<DdeAdviseEventArgs>(OnAdvise);
            _connection.Open();
        }

        /// <summary>
        /// コマンドライン引数を解析する
        /// </summary>
        public void Options()
        {
            string[] i_args = System.Environment.GetCommandLineArgs();
            var result = CommandLine.Parser.Default.ParseArguments<Options>(i_args) as CommandLine.Parsed<Options>;

            if (result != null)
            {
                opts.Code = result.Value.Code;
                opts.AutoStart = result.Value.AutoStart;
                opts.Debug = result.Value.Debug;

                //解析に成功した時は、解析結果を表示
                Console.WriteLine(string.Format("code: {0}\r\nstart: {1}\r\n",
                        opts.Code, opts.AutoStart));
            }
            else
            {
                //解析に失敗
                Console.WriteLine("コマンドライン引数の解析に失敗");
            }
        }

        /// <summary>
        /// 監視開始ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void timer1_Tick(object sender, EventArgs e)
        {
            Console.WriteLine(opts.Debug);
            if (!opts.Debug)
                IsRestTime();

            lblCount.Text = $"({_count}/秒)";
            _count = 0;
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            client.Disconnect();
            btnStart.Enabled = true;
            btnStop.Enabled = false;
            timer1.Enabled = false;
            _connection.Close();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            CreateDatabase(txtCode.Text);

            Start();

            Byte[] b = client.Request("銘柄名称", 1, 6000);
            _name = Encoding.Default.GetString(b);
            this.Text = _name;

            btnStart.Enabled = false;
            btnStop.Enabled = true;
            timer1.Enabled = true;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            _connection.Close();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            if (opts.AutoStart)
            {
                btnStart_Click(sender, e);
            }

        }
    }
}
