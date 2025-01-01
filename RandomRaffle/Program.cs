using NPOI.XSSF.UserModel;

namespace RandomRaffle;

public class RaffleInit
{
    public static void Main()
    {
        //获取表格数据
        var data = GetRaffleData("Source.xlsx");
        //去重,构建索引，LuckNumber重复就按照ID排序选取一个
        var index = data.GroupBy(x => x.LuckyNumber).ToDictionary(x => x.Key, x => x.OrderBy(y => y.Id).First());
        //产生随机种子
        var seed = Guid.NewGuid().GetHashCode();
        Console.WriteLine($"种子:{seed}");
        //打乱字典元素
        var random = new Random(seed);
        var shuffled = index.OrderBy(x => random.Next()).ToDictionary(x => x.Key, x => x.Value);
        //初始化抽奖
        List<PrizeData> prizes = new List<PrizeData>()
        {
            new()
            {
                PrizeName = "三等奖",
                Count = 2
            },
            new()
            {
                PrizeName = "二等奖",
                Count = 1
            },
            new()
            {
                PrizeName = "一等奖",
                Count = 1
            },
            new()
            {
                PrizeName = "特等奖",
                Count = 1
            },
        };
        //抽奖
        Raffle(prizes, shuffled);
    }

    public static List<RaffleData> GetRaffleData(string filePath)
    {
        var data = new List<RaffleData>();
        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            XSSFWorkbook workbook = new XSSFWorkbook(fs);
            var sheet = workbook.GetSheetAt(0);

            var headerRow = sheet.GetRow(0);
            if (headerRow == null) return data;
            
            var columnIndexes = GetColumnIndexes(headerRow);
            
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null) continue;

                var record = new RaffleData()
                {
                    Id = i,
                    Name = GetCellValue(row, columnIndexes["Submiter"]),
                    Number = GetCellValue(row, columnIndexes["Number"]),
                    LuckyNumber = float.TryParse(GetCellValue(row, columnIndexes["LuckyNumber"]), out float lucky) ? lucky : 0,
                    Date = DateTime.TryParse(GetCellValue(row, columnIndexes["SubmitTime"]), out DateTime date) ? date : DateTime.MinValue
                };

                data.Add(record);
            }
        }

        return data;
    }
    
    public static Dictionary<string, int> GetColumnIndexes(NPOI.SS.UserModel.IRow headerRow)
    {
        var columnIndexes = new Dictionary<string, int>();

        for (int i = 0; i < headerRow.LastCellNum; i++)
        {
            var cellValue = headerRow.GetCell(i)?.ToString().Trim();
            if (!string.IsNullOrEmpty(cellValue))
            {
                columnIndexes[cellValue] = i;
            }
        }

        return columnIndexes;
    }
    
    public static string GetCellValue(NPOI.SS.UserModel.IRow row, int columnIndex)
    {
        var cell = row.GetCell(columnIndex);
        return cell?.ToString() ?? string.Empty;
    }
    
    public static void Raffle(List<PrizeData> prizes, Dictionary<float, RaffleData> shuffled)
    {
        var luckyNumbers = shuffled.Keys.ToList();
        //每输入一次回车确定一个奖
        foreach (var prize in prizes)
        {
            for (int i = 0; i < prize.Count; i++)
            {
                var index = new Random(Guid.NewGuid().GetHashCode()).Next(0, luckyNumbers.Count);
                var luckyNumber = luckyNumbers[index];
                var winner = shuffled[luckyNumber];
                Console.WriteLine($"获得 {prize.PrizeName} 的数字是 {luckyNumber}");
                luckyNumbers.RemoveAt(index);
            }
            Console.ReadLine();
        }
    }
}

