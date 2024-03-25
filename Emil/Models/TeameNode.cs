using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Emil.Models
{
    public class TeameNode : TreeNode
    {
        public TeameNode(string title, string[] indicators)
        {
            this.Text = title;

            //Создание списка периодов
            this.Nodes.Add(new PeriodNode(title = "1", indicators));
            this.Nodes.Add(new PeriodNode(title = "2", indicators));
            this.Nodes.Add(new PeriodNode(title = "3", indicators));
            this.Nodes.Add(new PeriodNode(title = "Final", indicators));
            this.Nodes.Add(new PeriodNode(title = "OverTime", indicators));
        }
    }
}
