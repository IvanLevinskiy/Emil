using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Emil.Models
{
    public class PeriodNode : TreeNode
    {
        public PeriodNode(string header)
        {
            this.Text = header;
        }
    }
}
