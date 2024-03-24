using IronXL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;

namespace Emil.Models
{
    public class TrendNode : TreeNode
    {
        public string ColumnId;
        public string RowId;
        public string PeriodId;

        public WorkSheet workSheet;

        public TrendNode(string title)
        {
            this.Text = title;
        }

        /// <summary>
        /// Проверка на валидность периода
        /// </summary>
        /// <param name="periodStr"></param>
        /// <returns></returns>
        bool PeriodIsValid(string periodStr)
        {
            if(string.IsNullOrEmpty(periodStr) == true)
            {
                return false;
            }

            if (PeriodId == periodStr)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Поиск даты
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        string FindData(int row)
        {
            for(int i = row; i >= 0; i--)
            {
                var dateRaw = workSheet[$"D{i}"].Value.ToString();

                if(string.IsNullOrEmpty(dateRaw) == true)
                {
                    continue;
                }

                return dateRaw;

            }

            return string.Empty;
            
        }

        /// <summary>
        /// Поиск наименования команды
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        string FindCommand(int row)
        {
            for (int i = row; i >= 0; i--)
            {
                var dateRaw = workSheet[$"E{i}"].Value.ToString();

                if (string.IsNullOrEmpty(dateRaw) == true)
                {
                    continue;
                }

                return dateRaw;

            }

            return string.Empty;

        }

        public string GetTitle()
        {
            string title = $"Период: {PeriodId} Показатель: {this.Text}";

            return title;
        }

        /// <summary>
        /// Чтение данных
        /// </summary>
        /// <returns></returns>
        public PointPairList Read()
        {
            // Создадим список точек
            PointPairList list = new PointPairList();

            //Постройка уровня дат матчей
            var rows = workSheet.Rows.Count;
            var columns = workSheet.Columns.Count;

            //Наполнение данными
            for (int i = 4; i < rows; i++)
            {
                //Проверяем, чтоб период совпадал
                var periodId = workSheet[$"F{i}"].Value.ToString();

                var commandId = FindCommand(i);

                if(commandId.ToUpper().Contains("ТРАКТОР") == false)
                {
                    continue;
                }

                var tp = this.PeriodId;
                if (PeriodIsValid(periodId))
                {
                    var valueRaw = workSheet[$"{ColumnId}{i}"].Value.ToString();
                    var dateRaw = FindData(i);

                    double value = 0.0;
                    var res = double.TryParse(valueRaw, out value);

                    DateTime dateTime = new DateTime();
                    res = DateTime.TryParse(dateRaw, out dateTime);

                    if(res == true)
                    {
                        list.Add(new PointPair(new XDate(dateTime), value));
                    }
                    
                }
            }

            return list;

        }
    }
}
