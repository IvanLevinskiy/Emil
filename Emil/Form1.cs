using IronXL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ZedGraph;

namespace Emil
{
    public partial class Form1 : Form
    {

        WorkBook Book;
        WorkSheet Sheet;

        #region ЦВЕТА ГРАФИКА

        /// <summary>
        /// Цвета
        /// </summary>
        List<Color> BrushesList
        {
            get;
            set;
        }

        /// <summary>
        /// Цвета сетки
        /// </summary>
        public Color GridColor
        {
            get
            {
                return gridColor;
            }
            set
            {
                gridColor = value;
                //InitPane();
            }
        }
        Color gridColor = Color.Gray;

        /// <summary>
        /// Цвет заднего фона
        /// </summary>
        public Color BackColor
        {
            get
            {
                return backColor;
            }
            set
            {
                backColor = value;
                //InitPane();
            }
        }
        Color backColor = Color.White;

        /// <summary>
        /// Цвет перекрестия
        /// </summary>
        public Color CrosHairColor
        {
            get
            {
                return crosHairColor;
            }
            set
            {
                crosHairColor = value;
                //InitPane();
            }
        }
        Color crosHairColor = Color.DarkGray;

        /// <summary>
        /// Панель
        /// </summary>
        GraphPane pane;


        #endregion

        /// <summary>
        /// Метод для инициализации кистей
        /// </summary>
        void InitBrushes()
        {
            BrushesList = new List<Color>()
            {
                //Синий
                Color.Blue,

                Color.Red,

                Color.Green,

                Color.Violet,

                Color.DarkOrange,

                Color.Aqua,

                Color.LightGreen,

                Color.Gray,
            };
        }

        public Form1()
        {
            InitializeComponent();
            InitBrushes();

            InitZedgraphControl();
            InitPane();
            OpenExcelFile("Book.xlsx");

            tree.DoubleClick += (s, e) =>
            {
                if(tree.SelectedNode is Models.TrendNode)
                {
                    Models.TrendNode trendNode = (Models.TrendNode)tree.SelectedNode;
                    DrawTrend(trendNode);
                }
            };

            this.FormClosing += Form1_FormClosing;

            //tree.AllowDrop = true;
            //tree.DragEnter += (s, e) =>
            //{
            //    e.Effect = DragDropEffects.Copy;
            //};

            //tree.DragEnter += (s, e) =>
            //{
            //    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            //    foreach (string file in files) Console.WriteLine(file);
            //};

        }

        //Закрытие книги
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //this.Book.Close();
        }

        /// <summary>
        /// Метод инициализации ZedgraphControl
        /// </summary>
        private void InitZedgraphControl()
        {
            //Получение MasterPane
            var masterpane = zedGraphControl.GraphPane;

            //Разрешение синхронизации оси X
            zedGraphControl.IsSynchronizeXAxes = true;

            zedGraphControl.PanButtons = System.Windows.Forms.MouseButtons.Left;
            zedGraphControl.PanModifierKeys = System.Windows.Forms.Keys.None;

            //Кнопки увеличения графика
            zedGraphControl.ZoomButtons = System.Windows.Forms.MouseButtons.Left;
            zedGraphControl.ZoomModifierKeys = System.Windows.Forms.Keys.Shift;
        }

        /// <summary>
        /// Инициализация графика
        /// </summary>
        void InitPane()
        {
            pane = zedGraphControl.GraphPane;

            //Подписка на перемещение курсора по графику
            //zedGraph.MouseMove += (s, e) =>
            //{

            //};

            //Подписка на изменение масштаба
            //zedGraph.ZoomEvent += (zgControl, oldState, newState) =>
            //{

            //};

            //Подписка на показ сообщения о точке
            //zedGraph.PointValueEvent += (s, e) =>
            //{

            //};

            //Подписка на клик по графику
            //zedGraph.MouseClick += (s, e) =>
            //{

            //};


            // Точки можно перемещать, как по горизонтали,...
            //zedGraph.IsEnableHEdit = false;

            // ... так и по вертикали.
            //zedGraph.IsEnableVEdit = false;

            //Убираем легенду
            pane.Legend.IsVisible = true;
            pane.Legend.Position = LegendPos.Float;
            // Координаты будут отсчитываться в системе координат окна графика
            //pane.Legend.Location.CoordinateFrame = CoordType.ChartFraction;
            // Задаем выравнивание, относительно которого мы будем задавать координаты
            // В данном случае мы будем располагать легенду справа внизу
            pane.Legend.Border.Color = Color.Transparent;
            pane.Legend.Fill.Color = Color.Red;
            pane.Legend.FontSpec.Family = "Tahoma";
            pane.Legend.FontSpec.Size = 10.0f;

            // Задаем координаты легенды
            // Вычитаем 0.02f, чтобы был небольшой зазор между осями и легендой
            //pane.Legend.Location.TopLeft = new PointF(1.0f - 0.02f, 1.0f - 0.02f);

            //Запрет менять масштаб шрифтов
            pane.IsFontsScaled = false;

            // !!! Установим значение параметра IsBoundedRanges как true.
            // !!! Это означает, что при автоматическом подборе масштаба
            // !!! нужно учитывать только видимый интервал графика
            pane.IsBoundedRanges = true;

            //Кнопки для перемещения графика
            //zedGraph.PanButtons = System.Windows.Forms.MouseButtons.Left;
            //zedGraph.PanModifierKeys = System.Windows.Forms.Keys.None;

            //Кнопки увеличения графика
            //zedGraph.ZoomButtons = System.Windows.Forms.MouseButtons.Left;
            //zedGraph.ZoomModifierKeys = System.Windows.Forms.Keys.Shift;

            // !!! Изменим угол наклона меток по осям. 
            // Углы задаются в градусах
            pane.XAxis.Scale.FontSpec.Angle = 0;

            //Настройка обрамления
            pane.Chart.Border.Color = GridColor;

            //Установка прозрачного цвета для Pane
            pane.Border.Color = Color.Transparent;

            //Масштам по умолчанию
            pane.XAxis.Title = new AxisLabel("", "Times New Roman", 12, System.Drawing.Color.Blue, false, false, false);
            pane.XAxis.Scale.FontSpec.Size = 8.0f;
            pane.XAxis.Scale.FontSpec.FontColor = GridColor;
            pane.XAxis.Title.FontSpec.FontColor = GridColor;
            pane.XAxis.MinorTic.Color = GridColor;
            pane.XAxis.MajorTic.Color = GridColor;
            pane.XAxis.Type = AxisType.Date;
            pane.XAxis.Scale.Format = "dd.MM.yyyy\nHH:mm:ss";
            pane.XAxis.Scale.Align = AlignP.Center;


            pane.YAxis.Title = new AxisLabel("", "Times New Roman", 12, Color.Black, false, false, false);
            pane.YAxis.Scale.FontSpec.Size = 8.0f;
            pane.YAxis.Scale.FontSpec.FontColor = Color.Black;
            pane.YAxis.Title.FontSpec.FontColor = Color.Black ;
            pane.YAxis.MinorTic.Color = GridColor;
            pane.YAxis.MajorTic.Color = GridColor;



            // Заголовок
            pane.Title.IsVisible = false;
            pane.IsBoundedRanges = true;

            //СЕТКА
            // Включаем отображение сетки напротив крупных рисок по оси X
            pane.XAxis.MajorGrid.IsVisible = true;
            pane.XAxis.MajorGrid.Color = GridColor;
            pane.XAxis.MajorGrid.PenWidth = 0.2f;

            pane.YAxis.MajorGrid.IsVisible = true;
            pane.YAxis.MajorGrid.Color = GridColor;
            pane.YAxis.MajorGrid.PenWidth = 0.2f;

            // Задаем вид пунктирной линии для крупных рисок по оси X:
            // Длина штрихов равна 10 пикселям, ... 
            pane.XAxis.MajorGrid.DashOn = 2;

            // затем 5 пикселей - пропуск
            pane.XAxis.MajorGrid.DashOff = 2;


            // Включаем отображение сетки напротив мелких рисок по оси Y
            pane.XAxis.MinorGrid.IsVisible = false;
            pane.XAxis.MinorGrid.Color = GridColor;
            pane.XAxis.MinorGrid.PenWidth = 0.1f;

            // Аналогично задаем вид пунктирной линии для крупных рисок по оси Y
            pane.XAxis.MinorGrid.DashOn = 2;
            pane.XAxis.MinorGrid.DashOff = 2;

            pane.XAxis.Color = GridColor;

            //Настраиваем остальные свойства пане
            pane.Fill = new Fill(BackColor);
            pane.Margin.Right = 5;
            pane.Margin.Left = 10;

            //Цвет графика
            pane.Chart.Fill = pane.Fill;

            //UpdateGraphPane();
        }

        /// <summary>
        /// Загрузка Excel
        /// </summary>
        void LoadBook(string file)
        {
            //Приминение ключа
            var key = "IRONSTUDIO-1367548932-609745-4E91C77-3BEEDAB01-21DDB5-UExE0BBC0859E6A119-2056197345";
            IronXL.License.LicenseKey = key;

            Book = IronXL.WorkBook.Load(file);

            Sheet = Book.WorkSheets[1];
        }

        /// <summary>
        /// Получение или создание нового узла PeriodNode
        /// </summary>
        /// <param name="node"></param>
        /// <param name="title"></param>
        /// <param name="newNode"></param>
        /// <returns></returns>
        TreeNode GetOrCreatePeriodNode(TreeNode node, string title, out bool newNode)
        {
            newNode = false;

            foreach (TreeNode n in node.Nodes)
            {
                if(n.Text == title)
                {
                    return n;
                }
            }

            newNode = true;

            var treenode = new Models.PeriodNode(title);
            node.Nodes.Add(treenode);
            return treenode;
        }

        /// <summary>
        /// Открытие файла Excel
        /// </summary>
        /// <param name="file"></param>
        void OpenExcelFile(string file)
        {
            LoadBook(file);

            //Наполнение дерева
            TreeNode rootNode = new TreeNode("Трактор");
            tree.Nodes.Add(rootNode);

            //Постройка уровня дат матчей
            var rows = Sheet.Rows.Count;
            var columns = Sheet.Columns.Count;

            //Наполнение получение периодов
            for (int i = 4; i < rows; i++)
            {
                var cellRaw = Sheet[$"F{i}"].Value.ToString();

                if (string.IsNullOrEmpty(cellRaw) == true)
                {
                    continue;
                }

                bool newNode = false;
                var periodNode = GetOrCreatePeriodNode(rootNode, cellRaw, out newNode);
                var periodId = cellRaw;

                //Получение сигналов
                for (byte k = 0; k < columns; k++)
                {
                    //G3
                    var columnIndex = AddSymbol("G", k);

                    cellRaw = Sheet[$"{columnIndex}3"].Value.ToString();

                    if (string.IsNullOrEmpty(cellRaw) == true)
                    {
                        continue;
                    }

                    if(newNode == false)
                    {
                        continue;
                    }

                    var trendNode = new Models.TrendNode(cellRaw)
                    {
                        ColumnId = columnIndex,
                        RowId = $"{i}",
                        workSheet = this.Sheet,
                        PeriodId = periodId
                    };
                    periodNode.Nodes.Add(trendNode);
                }
            }
        }

        /// <summary>
        /// Метод для долучения следующего
        /// столбца
        /// </summary>
        /// <param name="symbol"></param>
        /// <param name="addative"></param>
        /// <returns></returns>
        string AddSymbol(string symbol, byte addative)
        {
            var result = "";
            byte s = (byte)symbol[symbol.Length - 1];

            //Если символов больше 1 -
            //удаляем последний
            if (result.Length == 2)
            {
                result.Remove(1);
            }

            //Добавляем к символу добавку
            s = (byte)(s + addative);


            if (s > 90)
            {
                result = "A";
                s = (byte)(s - 26);
            }

            result += (char)s;
            return result;
        }

        /// <summary>
        /// Метод для поиска листа Excel
        /// по имени
        /// </summary>
        /// <param name="book"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        WorkSheet SearchWorkSheet(WorkBook book, string name)
        {
            var sheets = book.WorkSheets;

            foreach (var sh in sheets)
            {
                if (sh.Name == name)
                {
                    return sh;
                }
            }

            return null;
        }
    
        bool TrendIsExist(string title)
        {
            foreach(ListViewItem item in listView.Items)
            {
                if(item.Text == title)
                {
                    return true;
                }
            }

            return false;
        }

        void DrawTrend(Models.TrendNode trendNode)
        {
            if(BrushesList.Count == 0)
            {
                return;
            }

            //
            var trenttitle = trendNode.GetTitle();

            //Проверка на то, что тренд уже добавлен
            var trendExist = TrendIsExist(trenttitle);
            if (trendExist) 
            { 
                return; 
            }

            //Чтение данных из Excel
            var list = trendNode.Read();


            var color = BrushesList[0];
            BrushesList.RemoveAt(0);

            // Создаем кривую,
            // которая будет рисоваться голубым цветом (Color.Blue),
            // опорные точки выделяться не будут (SymbolType.None)
            // Создание кривой с использованием дат ничем не отличается от создания других кривых
            LineItem myCurve = pane.AddCurve("", list, color, SymbolType.None);

            var lv = new ListViewItem()
            {
                Text = trendNode.GetTitle(),
                ForeColor = color
            };

            listView.Items.Add(lv);

            // Вызываем метод AxisChange (), чтобы обновить данные об осях.
            zedGraphControl.AxisChange();

            // Обновляем график
            zedGraphControl.Invalidate();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            //Возвращение цветов
            BrushesList.Clear();
            InitBrushes();

            //Очистка листбокса
            listView.Items.Clear();

            //Очистка содержимого
            pane.CurveList.Clear();

            // Вызываем метод AxisChange (), чтобы обновить данные об осях.
            zedGraphControl.AxisChange();

            // Обновляем график
            zedGraphControl.Invalidate();
        }
    }
}
