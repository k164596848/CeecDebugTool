using System;
using System.Collections.Generic;

using System.Drawing;
using Emgu.CV;
using Emgu.CV.Structure;
using Emgu.CV.CvEnum;
using Emgu.CV.Util;
using System.IO;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Data.OleDb;

using System.Text;

using System.Data;

namespace ExaminationCenter_02
{
    public class Question
    {
        public int group_num = -1;
        public List<Rectangle> item;
    }
    public class Answer //記錄判斷的答案
    {
        public int group_num = -1;
        public List<String> item;
    }
    public class JObject
    {
        public string area { get; set; }
        public string Filename { get; set; }
        public int Qnumber { get; set; }
        public List<List<int>> position { get; set; }
        public List<string> OptionID { get; set; }
        public double RedWhiteThreshold { get; set; } //取得空白平均灰度值
        public double avgPixelValue { get; set; }
        public double RedPixelValue { get; set; }
        public double WhitePixelValue { get; set; }
    }
    public class Record_Diff //記錄不一樣的答案
    {
        public String RecordList;
    }
    public class Record_Same //記錄相同的答案
    {
        public String RecordList;
    }
    class Program
    {
        public static String fullparameter = "";
        
        public static OleDbConnection conn;
        public static bool method13IsNotNull = false;
        public static int methodFalse = 0;
        public static int nullpaper = 0;

        static void Main(string[] args)
        {
            Console.WriteLine("hello world");

            System.Diagnostics.Stopwatch sw1 = new System.Diagnostics.Stopwatch();//引用stopwatch物件
            sw1.Reset();//碼表歸零
            sw1.Start();//碼表開始計時
            System.Diagnostics.Stopwatch sw2 = new System.Diagnostics.Stopwatch();//引用stopwatch物件
            sw2.Reset();//碼表歸零
            /*Json檔讀入*/
            //PaperList(Option list) 整張考卷的每個題組
            List<Question> PaperList = new List<Question>();
            String jsonfile = @"parameter";//存放三個參數的位置(資料庫 / 圖片 / json)
            /*讀從root中的txt檔中取得json路徑*/
            fullparameter = "";
            fullparameter = Path.GetFullPath(jsonfile);
            String fulljsonpath = fullparameter + "//json.txt";
            string path;
            string ReadLine1;
            path = Path.GetFullPath(fulljsonpath);
            if (!File.Exists(path))
            {
                Console.WriteLine("沒有讀取json的根目錄");
                ReadLine1 = "";
            }
            else
            {
                Console.WriteLine("有讀取json的根目錄");
                Console.WriteLine("path = " + path);

                string  target_jsonpath = fullparameter + @"\json\英聽";
                System.IO.File.WriteAllText(path, target_jsonpath);
                
                //目標json的所在位置
                Console.WriteLine("josn.txt content = " + target_jsonpath);

                StreamReader str = new StreamReader(path);               
                ReadLine1 = str.ReadLine();                                         
                str.Close();
                
                //System.IO.File.WriteAllText(path, string.Empty);

            }
            //歷遍目錄下的所有檔案(position.json)
            List<String> file_json_position = new List<String>();
            try
            {
                file_json_position = Directory.EnumerateFiles(ReadLine1, "*position.json", SearchOption.AllDirectories).ToList();
                Console.WriteLine("*position.json work");

            }
            catch (System.Exception except)
            {
                Console.WriteLine(except.Message);
            }

            //歷遍目錄下的所有檔案(avgPixel.json)
            List<String> file_json_avgPixel = new List<String>();
            try
            {
                file_json_avgPixel = Directory.EnumerateFiles(ReadLine1, "*avgPixel.json", SearchOption.AllDirectories).ToList();
                Console.WriteLine("*avgPixel.json work");
            }
            catch (System.Exception except)
            {
                Console.WriteLine(except.Message);
            }

            /*OMR答案卷處理*/
            //Answer List 用來記錄學生答案
            List<Answer> AnswerList_One = new List<Answer>();   //方法一List
            List<Answer> AnswerList_Two = new List<Answer>();   //方法二List
            List<Answer> AnswerList_Three = new List<Answer>(); //方法三List


            //從root中的txt檔中取得路徑並歷遍目錄下的所有檔案(.jpg)
            string result_path = fullparameter + "\\picture_root.txt";
            string ReadLine2;
            path = Path.GetFullPath(result_path);
            if (!File.Exists(path))
            {
                Console.WriteLine("沒有讀取校正圖片的根目錄");
                ReadLine2 = "";
            }
            else
            {
                Console.WriteLine("有讀取picture的根目錄");
                string target_pic_path = @"C:\Users\User\Desktop\CeecDebugTool\輸出檔\haha";
                System.IO.File.WriteAllText(path,target_pic_path);

                StreamReader str = new StreamReader(path);
                ReadLine2 = str.ReadLine();
                str.Close();
                //System.IO.File.WriteAllText(path, string.Empty);
            }

            List<String>[] ImagePath = new List<String>[4];

            //讀取每一頁圖片
            int j = 0;
            for (var i = 1; i < 5; i++)
            {
                ImagePath[j] = Directory.EnumerateFiles(ReadLine2, "*_P" + i.ToString() + ".jpg", SearchOption.AllDirectories).ToList();
                if (ImagePath[j].Count != 0)
                {
                    j++;
                }
            }
            if (j == 0)//代表是英聽/指考三科(因為圖片檔名沒有頁數)
            {
                ImagePath[j] = Directory.EnumerateFiles(ReadLine2, "*.jpg", SearchOption.AllDirectories).ToList();
            }

            //製作mask遮罩 s
            //讀取第一張答案卷以獲取影像尺寸
            Image<Gray, Byte>[] mask = new Image<Gray, byte>[4];
            for (var i = 0; i < file_json_position.Count; i++)
            {
                if (ImagePath[i].Count != 0)
                {
                    Image<Gray, Byte> temp_mask = new Image<Gray, byte>(ImagePath[i].First());
                    mask[i] = new Image<Gray, byte>(temp_mask.Width, temp_mask.Height);
                }
            }

            
            /*對debugtool.CSV 做前置處理開始*/
            String tragetFilePath = @"..\OMR記錄檔\DebugTool.csv";
            string DebugToolCsv_file;
            DebugToolCsv_file = Path.GetFullPath(tragetFilePath);
            Console.WriteLine("DebugToolCsv_file 位置= " + DebugToolCsv_file);
            System.IO.File.WriteAllText(DebugToolCsv_file, string.Empty);//先清空，避免被上次留下的資料汙染
            FileStream fs = null;
            fs = new FileStream(DebugToolCsv_file, FileMode.Open);
            using (StreamWriter sw = new StreamWriter(fs, Encoding.Default))
            {
                //寫入標題
                sw.Write("卷號" + "," + "題號" + "," + "選項A" + "," + "選項B" + "," + "選項C" + "," + "選項D" + "," + "選項E");
                sw.WriteLine(" ");
                sw.Close();
                sw.Dispose();
            }
            Console.WriteLine("DebugToolCsv_file working ");
            /*對debugtool.CSV 做前置處理結束*/


            try
            {
                int Qcount = 0;
                for (var i = 0; i < file_json_position.Count; i++)
                {
                    //清空各Answer List
                    AnswerList_Three.Clear();
                   
                    PaperList.Clear();
                    //讀取position.json的內容
                    JsonProcessing(file_json_position[i], PaperList);
                    //讀取avgPixel.json的內容
                    AvgJsonProcessing(file_json_avgPixel[i]);                 

                    foreach (String dChild in ImagePath[i])
                    {
                        method13IsNotNull = false;
                        //切割檔名，去除附檔名(.jpg)
                        string[] File_Name = dChild.Split('\\');
                        //Console.WriteLine("Filfename：" + File_Name[4]);
                        File_Name = File_Name[File_Name.Length - 1].Split('_');
                        if (i < 2)//是第幾面
                        {
                            File_Name = File_Name[0].Split('A');
                            File_Name = File_Name[0].Split('.');
                        }
                        else
                        {
                            File_Name = File_Name[0].Split('B');
                        }

                        int error = 0;//變數，判斷有無底色過黑的卷號
                        if (true)//如果error = 1則就不做答案辨識
                        {
                            //生成以dChild為指引的image object
                            Image<Gray, Byte> AnsPaper   = new Image<Gray, byte>(dChild);
                            //列印檔名(卷號)
                            Console.WriteLine("卷號：" + File_Name[0]);
                            //呼叫方法三
                            GrayPixel(AnswerList_Three, AnsPaper, PaperList);
                            Qcount = PaperList[0].group_num;
                            //寫入CSV
                            Decide_Answer(File_Name[0], PaperList, AnswerList_Three, Qcount, i);

                        }
                    }
                }
                
            }
            catch (System.Exception except)
            {
                Console.WriteLine(except.Message);
            }


            sw1.Stop();//碼錶停止
            Console.WriteLine("執行完畢，按任意鍵以關閉主控台視窗");
            //按任意鍵以關閉主控台視窗
            Console.ReadKey(true);

        }
        /*判讀答案前處理*/
        //設置答案字典
        public static string[] AnsDictionary = new string[] { };

        //Json檔案格式轉換,把每個選項格(OptionID)的位置放到 AnsDictionary[]陣列中記錄
        public static void JsonProcessing(String file_path, List<Question> PaperList)
        {
            string DeJsonData = File.ReadAllText(file_path);
            JArray JsonArray = JArray.Parse(DeJsonData);

            int mix = 0;
            for (int i = 0; i < JsonArray.Count; i++)//題目
            {
                String tmp = JsonArray[i].ToString();
                
                Question tmp_package = new Question();
                JObject model = JsonConvert.DeserializeObject<JObject>(tmp);
                //Console.WriteLine("tmp(position.json的)  "+ tmp);

                //題號
                tmp_package.group_num = model.Qnumber;
                //選項物件生成
                tmp_package.item = new List<Rectangle>();

                //把正方形(選項範圍)放入
                for (int j = 0; j < model.position.Count; j++)//選項
                {
                    Rectangle rec = new Rectangle();
                    rec.X = model.position[j][0];
                    rec.Y = model.position[j][1];


                    rec.Width = model.position[j][2];
                    rec.Height = model.position[j][3];
                    tmp_package.item.Add(rec);
                }
                PaperList.Add(tmp_package);

                //動態設置答案字典
                //調整答案字典陣列的大小
                if (mix < model.OptionID.Count)
                {
                    mix = model.OptionID.Count;
                    System.Array.Resize(ref AnsDictionary, model.OptionID.Count);
                    //取出model.OptionID裡的選項代號                 
                    for (int k = 0; k < AnsDictionary.Length; k++)
                    {
                        //指定新的陣列值
                        AnsDictionary[k] = model.OptionID[k];
                    }
                }
            }
        }
        //取得空白平均灰度值
        public static double AvgGrayCount;
        public static int TestingThreshold=225;//用來測試用得threshold
        public static void AvgJsonProcessing(String file_path)
        {
            string DeJsonData = File.ReadAllText(file_path);
            JObject model = JsonConvert.DeserializeObject<JObject>(DeJsonData);
            AvgGrayCount = model.avgPixelValue;//

            Console.WriteLine("model.avgPixelValue" + model.avgPixelValue);
            Console.WriteLine("model.RedWhiteThreshold" + model.RedWhiteThreshold);

        }
        /*判讀答案前處理  end*/

        /*方法三：使用答案平均灰度總和進行答案辨識*/
        //區塊計算
        public static void GrayPixel(List<Answer> AnswerList, Image<Gray, Byte> AnsPaper, List<Question> PaperList)
        {
            //宣告Answer List用來記錄學生答案
            Answer ans;
            int Height = PaperList[0].item[0].Height;
            int Width = PaperList[0].item[0].Width;
           
            for (int i = 0; i < PaperList.Count; i++)
            {
                //宣告new class
                ans = new Answer();
                ans.item = new List<String>();

                //建立參數                
                
                
               
                //輸入option內的whhite pixel個數                
                for (int j = 0; j < PaperList[i].item.Count; j++)
                {
                    Rectangle target = PaperList[i].item[j]; //取框的範圍
                    AnsPaper.ROI = target; //給予ROI矩形範圍 
                    int count_white = Count_Gray(AnsPaper, TestingThreshold);
                   
                    ans.group_num = PaperList[i].group_num; //加入題號                  
                    ans.item.Add(count_white.ToString());//加入option white numbers    
                    AnsPaper.ROI = Rectangle.Empty; //清除ROI範圍
                    
                      
                    
                }
                //將答案寫入AnsList中
                AnswerList.Add(ans);
            }
        }

        //計算灰度值
        public static int Count_Gray(Image<Gray, Byte> AnsPaper, int threshold)
        {
            int sum = 0; //計算總和
            for (int x = 0; x < AnsPaper.Height; x++)
            {
                for (int y = 0; y < AnsPaper.Width; y++)
                {
                    int intensity = (int)AnsPaper[x, y].Intensity; //讀取指定位置的像素值
                    if (intensity >= threshold) sum++;
                    //sum = sum + intensity; //加總和
                }
            }
            return sum;
        }
        //判斷三個方法的結果是否一致
        public static void Decide_Answer(String File_Name, List<Question> PaperList, List<Answer> AnswerList_Three, int Qcount, int page)
        {
            List<Record_Same> Record = new List<Record_Same>();//寫入用答案
            Record_Same record;//小段放入

            for (int i = Qcount; i < Qcount + PaperList.Count; i++)
            {
                
                String Merge_Three = String.Join(",", AnswerList_Three[i - Qcount].item.ToArray()); //將方法三得到的List轉為String              
                record = new Record_Same();
                record.RecordList = Merge_Three;//option_white_num put in array list
                Record.Add(record);

            }
            string result_path2 = @"..\OMR記錄檔\DebugTool.csv";
            string path2;
            path2 = Path.GetFullPath(result_path2);
            using (StreamWriter sw2 = File.AppendText(path2))
            {
                //寫入答案卷判別答案
                if (Record.Count != 0)
                {
                    for (int i = 0; i < Record.Count; i++)
                    {
                        //寫入卷號
                        sw2.Write(File_Name + ",");
                        // 寫入題目編號
                        sw2.Write(i+1 + ",");
                        //將判別答案寫入
                        sw2.Write(Record[i].RecordList);
                        sw2.WriteLine(" ");
                    }
                }
                sw2.Close();
                sw2.Dispose();
            }
            Console.WriteLine("CSV 寫入成功");

        }

        public static void CreateFile(String File_Name, List<Record_Diff> Record_Diff)
        {
            String result_path2 = @"..\OMR記錄檔\DebugTool.csv";
            string path2;
            path2 = Path.GetFullPath(result_path2);
            if (!File.Exists(path2))
            {
                using (System.IO.StreamWriter sw2 = new System.IO.StreamWriter(path2))
                {
                    //寫入答案卷判別答案
                    if (Record_Diff.Count != 0)
                    {
                        for (int i = 0; i < Record_Diff.Count; i++)
                        {
                            //寫入卷號
                            sw2.Write(File_Name + ",");
                            //將判別答案寫入
                            sw2.Write(Record_Diff[i].RecordList);
                            sw2.WriteLine(" ");
                        }
                    }
                    sw2.Close();
                    sw2.Dispose();
                }
            }
            else
            {
                //保留原本資料
                using (StreamWriter sw2 = File.AppendText(path2))
                {
                    //寫入答案卷判別答案
                    if (Record_Diff.Count != 0)
                    {
                        for (int i = 0; i < Record_Diff.Count; i++)
                        {
                            //寫入卷號
                            sw2.Write(File_Name + ",");
                            //將判別答案寫入
                            sw2.Write(Record_Diff[i].RecordList);
                            sw2.WriteLine(" ");
                        }
                    }
                    sw2.Close();
                    sw2.Dispose();
                }
            }
        }
    }
}