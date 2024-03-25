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
        public PeriodNode(string header, string[] indicators)
        {
            this.Text = header;

            //Наполнение индикаторами
            foreach (string indicator in indicators) 
            {
                //Добавляем только те индикаторы, которые начинаются с i
                if (indicator[0] == 'i')
                {
                    this.Nodes.Add(new IndicatorNode(indicator));
                }
                
            }

           
        }
    }
}
