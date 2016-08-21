using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ShiraSayaListSystemVer._5
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("しらさや名簿システムを起動します");
            Console.WriteLine("STEP1/4 Configファイルを読み込みます");
            const string CONFIGPATH = @"J:\H28体育大会選手選出表\しらさや名簿システム\Config.ini";
            var ConfigFile = new ConfigFile(CONFIGPATH);//TODO:パスの変更
            var ExcelSheet = new ShiranuiSayaka.Libraries.ExcelControl();
            var AccessTable = new ShiranuiSayaka.Libraries.AccessControl(ConfigFile.ReadAccessPath());

            string[] ClassStudentArray;
            int Grade = 0; int Class = 0;
            bool IsManSheet = false;

            if (ExcelSheet.ReadCell(ConfigFile.ReadClassNumberCell()).IndexOf("M") < 0)
            {
                Console.WriteLine("STEP2/4 クラス識別番号を読み込みます");
                int ClassInfomationNumber = int.Parse(ExcelSheet.ReadCell(ConfigFile.ReadClassNumberCell()));
                Grade = ClassInfomationNumber / 10; Class = ClassInfomationNumber % 10;

                Console.WriteLine("STEP3/4 クラスの名簿を読み込みます");
                string Sql = $"SELECT StudentNumber,StudentName FROM 生徒名簿 WHERE Grade = {Grade} AND ClassNumber = {Class}";
                var ClassStudentList = AccessTable.SelectQuery(Sql);
                ClassStudentArray = ClassStudentList["StudentName"].ToArray();
            }
            else
            {
                Console.WriteLine("男子シートである事を検出しました。クラスの識別を省略します");
                string ClassInfomation = ExcelSheet.ReadCell(ConfigFile.ReadClassNumberCell());
                Grade = int.Parse(ClassInfomation.Substring(0, 1));
                ClassStudentArray = new string[0];
                IsManSheet = true;
            }
            Console.WriteLine("STEP4/4 出席番号を名前に変換します");
            var Area = ConfigFile.ReadWriteCellArea();
            foreach (string Range in Area)
            {
                int InData;
                if (IsManSheet == false && int.TryParse(ExcelSheet.ReadCell(Range) == "" ? "0" : ExcelSheet.ReadCell(Range), out InData) != false && InData != 0)
                {
                    string OutData = ClassStudentArray[InData - 1];
                    ExcelSheet.WriteCell(Range, OutData);
                    Console.WriteLine($"{Grade}{Class}{InData.ToString()} = {OutData}");
                }
                if(IsManSheet == true && ExcelSheet.ReadCell(Range) != "")
                {
                    string InTextData = ExcelSheet.ReadCell(Range);
                    string ClassName = InTextData.Substring(0, 1).ToUpper();
                    int StudentNumber = int.Parse(InTextData.Substring(1));
                    string Sql = $@"SELECT StudentName FROM 生徒名簿 WHERE Grade = {Grade} AND ClassName = ""{ClassName}"" AND StudentNumber = {StudentNumber}";
                    var SqlResult = AccessTable.SelectQuery(Sql);
                    string OutData = SqlResult.Select(x => x.Value).First().First();
                    ExcelSheet.WriteCell(Range, OutData);
                    Console.WriteLine($"{Grade}{ClassName}{StudentNumber} = {OutData}");
                }
            }
            AccessTable.Close();
            Console.WriteLine("処理が正常に終了しました。10秒後に終了します");
            Console.WriteLine("名簿は内容を確認の後、印刷して生徒会に提出して下さい");
            System.Threading.Thread.Sleep(10000);
        }
    }
    internal class ConfigFile
    {
        private List<string> ConfigData;
        internal ConfigFile(string Path)
        {
            ConfigData = new List<string>();
            using (var Reader = new System.IO.StreamReader(Path, Encoding.Default))
            {
                while (Reader.Peek() != -1)
                {
                    ConfigData.Add(Reader.ReadLine());
                }
            }
        }
        internal string ReadAccessPath()
        {
            var PathLine = ConfigData.Where(x => x.IndexOf("Path") >= 0).Select(x => x).First();
            string Path = PathLine.Split('=')[1];
            return Path;
        }
        internal List<string> ReadWriteCellArea()
        {
            var CellArea = new List<string>();
            var WriteCellAreaLines = ConfigData.Where(x => x.IndexOf("Area") >= 0).Select(x => x).ToList();
            foreach (string Line in WriteCellAreaLines)
            {
                string Area = Line.Split('=')[1];
                string StartCell = Area.Split('.')[0];
                string StopCell = Area.Split('.')[1];
                List<string> TempReturnsList = WritingStartToStopCell(StartCell, StopCell);
                CellArea.AddRange(TempReturnsList);
            }
            return CellArea;
        }
        private List<string> WritingStartToStopCell(string StartCell, string StopCell)
        {
            int StartCellLeftCharNumber = StartCell.Substring(0, 1).ToCharArray()[0];
            int StartCellRightNumber = int.Parse(StartCell.Substring(1));
            int StopCellLeftCharNumber = StopCell.Substring(0, 1).ToCharArray()[0];
            int StopCellRightNumber = int.Parse(StopCell.Substring(1));

            var Cells = new List<string>();
            var LeftRange = Enumerable.Range(StartCellLeftCharNumber, (StopCellLeftCharNumber - StartCellLeftCharNumber) + 1).ToList();
            var RightRange = Enumerable.Range(StartCellRightNumber, (StopCellRightNumber - StartCellRightNumber) + 1).ToList();
            foreach (var Left in LeftRange)
            {
                string LeftString = Convert.ToChar(Left).ToString();
                foreach (var Right in RightRange)
                {
                    Cells.Add(LeftString + Right.ToString());
                }
            }
            return Cells;
        }
        internal string ReadClassNumberCell()
        {
            string ClassNumberCellArea = ConfigData.Where(x => x.IndexOf("ClassNumberCell") >= 0).Select(x => x).First();
            string ClassNumberCell = ClassNumberCellArea.Split('=')[1];
            return ClassNumberCell;
        }
    }
}
