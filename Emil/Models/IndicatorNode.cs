using IronXL;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;

namespace Emil.Models
{
    public class IndicatorNode : TreeNode
    {
        public string Key;

        public double Value;

        public IndicatorNode(string IndicatorName)
        {
            this.Text = IndicatorName;
        }

        /// <summary>
        /// Получение заголовка
        /// </summary>
        /// <returns></returns>
        public string GetTitle()
        {
            var teamName = this.Parent.Parent.Text;
            var period = this.Parent.Text;
            var indicator = this.Text;

            return $"{teamName}/{period}/{indicator}";


        }


                

        /// <summary>
        /// Чтение данных
        /// </summary>
        /// <returns></returns>
        public PointPairList Read()
        {
            // Создадим список точек
            PointPairList list = new PointPairList();

            OleDbConnection DbConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DataBase.accdb;Persist Security Info=False;");
            DbConnection.Open();

            var teamName = this.Parent.Parent.Text;
            var period = this.Parent.Text;
            var indicator = this.Text;

            // текст запроса
            string query = $"SELECT [DateTimeStamp], [{indicator}] FROM KHL_23_24 WHERE TeamName = \"{teamName}\" AND Period = \"{period}\";";

            // создаем объект OleDbCommand для выполнения запроса к БД MS Access
            OleDbCommand command = new OleDbCommand(query, DbConnection);

            // получаем объект OleDbDataReader для чтения табличного результата запроса SELECT
            OleDbDataReader reader = command.ExecuteReader();

            while (reader.Read()) 
            {
                try
                {
                    var dateTime = Convert.ToDateTime(reader[0]);
                    var value = Convert.ToDouble(reader[1]);

                    list.Add(new PointPair(new XDate(dateTime), value));
                }
                catch 
                {
                
                }
            }

            var orderedList = list.OrderBy(o => o.X).ToList();
            list.Clear();

            foreach ( var p in orderedList )
            {
                list.Add(p);
            }

            return list;
        }
    }
}
