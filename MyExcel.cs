using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Laboratorna1_Excel_
{
    public partial class MyExcel : Form
    {
        public MyExcel()
        {
            InitializeComponent();
        }
        Table table = new Table();
        private void InitializeDataGridView(int rows, int columns)
        {
            dataGridView1.ColumnHeadersVisible = true;
            dataGridView1.RowHeadersVisible = true;
            dataGridView1.ColumnCount = columns;
            for(int i = 0; i<columns; i++)
            {
                string columnName = NumberConverter.To26System(i);
                dataGridView1.Columns[i].Name = columnName;
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            for(int i=0; i<rows; i++)
            {
                dataGridView1.Rows.Add("");
                dataGridView1.Rows[i].HeaderCell.Value = (i).ToString();
            }
            dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            table.setTable(columns, rows);
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int col = dataGridView1.SelectedCells[0].ColumnIndex;
            int row = dataGridView1.SelectedCells[0].RowIndex;
            string expression = Table.grid[row][col].expression;
            string value = Table.grid[row][col].value;
            textBox.Text = expression;
            textBox.Focus();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int col = dataGridView1.SelectedCells[0].ColumnIndex;
            int row = dataGridView1.SelectedCells[0].RowIndex;
            string expression = textBox.Text;
            if (expression == "") return;
            table.ChangeCellWithAllPointers(row, col, expression, dataGridView1);
            dataGridView1[col, row].Value = Table.grid[row][col].value;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataGridViewRow row = new DataGridViewRow();
            if(dataGridView1.Columns.Count == 0)
            {
                MessageBox.Show("there are no colums!");
                return;
            }
            dataGridView1.Rows.Add(row);
            dataGridView1.Rows[table.rowCount].HeaderCell.Value = (table.rowCount).ToString();
            table.AddRow(dataGridView1);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!table.DeleteRow(dataGridView1))
            {
                return;
            }
            dataGridView1.Rows.RemoveAt(table.rowCount);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string name = NumberConverter.To26System(table.colCount);
            dataGridView1.Columns.Add(name, name);
            table.AddColumn(dataGridView1);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (!table.DeleteColumn(dataGridView1))
                return;
            dataGridView1.Columns.RemoveAt(table.colCount);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "TableFile|*.txt";
            openFileDialog.Title = "Open Table File";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;
            StreamReader streamReader = new StreamReader(openFileDialog.FileName);
            table.Clear();
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            int row;
            int column;
            Int32.TryParse(streamReader.ReadLine(), out row);
            Int32.TryParse(streamReader.ReadLine(), out column);
            InitializeDataGridView(row, column);
            table.Open(row, column, streamReader, dataGridView1);
            streamReader.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "TableFile|*.txt";
            saveFileDialog.Title = "Save table file";
            saveFileDialog.ShowDialog();
            if(saveFileDialog.FileName != "")
            {
                FileStream fileStream = (FileStream)saveFileDialog.OpenFile();
                StreamWriter streamWriter = new StreamWriter(fileStream);
                table.Save(streamWriter);
                streamWriter.Close();
                fileStream.Close();
            }
        }
    }
}
