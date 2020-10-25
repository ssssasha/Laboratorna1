using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;

namespace Laboratorna1_Excel_
{
    public class Table
    {
        private const int defaultCol = 30;
        private const int defaultRow = 10;
        public int colCount;
        public int rowCount;
        public static List<List<Cell>> grid = new List<List<Cell>>();
        public Dictionary<string, string> dictionary = new Dictionary<string, string>();
        Table table = new Table();
        public Table()
        {
            setTable(defaultCol, defaultRow);
        }
        public void setTable(int col, int row)
        {
            Clear();
            colCount = col;
            rowCount = row;
            for(int i=0; i<rowCount; i++)
            {
                List<Cell> newRow = new List<Cell>();
                for(int j=0; j< colCount; j++)
                {
                    newRow.Add(new Cell(i, j));
                    dictionary.Add(newRow.Last().getName(), "");
                }
                grid.Add(newRow);
            }
        }
        public void Clear()
        {
            foreach(List<Cell>list in grid)
            {
                list.Clear();
            }
            grid.Clear();
            dictionary.Clear();
            rowCount = 0;
            colCount = 0;
        }
        public void ChangeCellWithAllPointers(int row, int col, string expression, System.Windows.Forms.DataGridView dataGridView1)
        {
            grid[row][col].DeletePointersAndReferences();
            grid[row][col].expression = expression;
            grid[row][col].newReferencesFromThis.Clear();
            if (expression != "")
            {
                if (expression[0] != "=")
                {
                    grid[row][col].value = expression;
                    dictionary[FullName(row, col)] = expression;
                    foreach (Cell cell in grid[row][col].pointersToThis)
                    {
                        RefreshCellAndPointers(cell, dataGridView1);
                    }
                    return;
                }
            }
            string newExpression = ConvertReferences(row, col, expression);
            if (newExpression != "")
            {
                newExpression = newExpression.Remove(0, 1);
            }
            if (!grid[row][col].CheckLoop(grid[row][col].newReferencesFromThis))
            {
                System.Windows.Forms.MessageBox.Show("There is a loop! Change the expression");
                grid[row][col].expression = "";
                grid[row][col].value = "0";
                dataGridView1[col, row].Value = "0";
                return;
            }
            grid[row][col].AddPointersAndReferences();
            string val = Calculate(newExpression);
            if (val == "Error")
            {
                System.Windows.Forms.MessageBox.Show("Error in cell" + FullName(row, col));
                grid[row][col].expression = "";
                grid[row][col].value = "0";
                dataGridView1[col, row].Value = "0";
                return;
            }
            grid[row][col].value = val;
            dictionary[FullName(row, col)] = val;
            foreach (Cell cell in grid[row][col].pointersToThis)
            {
                RefreshCellAndPointers(cell, dataGridView1);
            }
        }
        private string FullName(int row, int col)
        {
            Cell cell = new Cell(row, col);
            return cell.getName();
        }
        public bool RefreshCellAndPointers(Cell cell, DataGridView dataGridView1)
        {
            cell.newReferencesFromThis.Clear();
            string newExpression = ConvertReferences(cell.row, cell.column, cell.expression);
            newExpression = newExpression.Remove(0, 1);
            string Value = Calculate(newExpression);
            if(Value == "Error")
            {
                MessageBox.Show("Error in cell" + cell.getName());
                cell.expression = "";
                cell.value = "0";
                dataGridView1[cell.column, cell.row].Value = "0";
                return false;
            }
            grid[cell.row][cell.column].value = Value;
            dictionary[FullName(cell.row, cell.column)] = Value;
            dataGridView1[cell.column, cell.row].Value = Value;
            foreach(Cell point in cell.pointersToThis)
            {
                if(!RefreshCellAndPointers(point, dataGridView1))
                {
                    return false;
                }
            }
            return true;
        }
        public void RefreshReferences()
        {
            foreach(List<Cell> list in grid)
            {
                foreach (Cell cell in list)
                {
                    if (cell.referencesFromThis != null)
                    {
                        cell.referencesFromThis.Clear();
                    }
                    if(cell.newReferencesFromThis != null)
                    {
                        cell.newReferencesFromThis.Clear();
                    }
                    if(cell.expression == "")
                        continue;
                    string newExpression = cell.expression;
                    if(cell.expression[0] == "=")
                    {
                        newExpression = ConvertReferences(cell.row, cell.column, cell.expression);
                        cell.referencesFromThis.AddRange(cell.newReferencesFromThis);
                    }
                }
            }
        }
        public string ConvertReferences(int row, int col, string expr)
        {
            string cellPattern = @"[A-Z]+[0-9]+";
            Regex regex = new Regex(cellPattern, RegexOptions.IgnoreCase);
            Index nums;
            foreach(Match match in regex.Matches(expr))
            {
                if (dictionary.ContainsKey(match.Value))
                {
                    nums = NumberConverter.From26System(match.Value);
                    grid[row][col].newReferencesFromThis.Add(grid[nums.row][nums.column]);
                }
            }
            MatchEvaluator evaluator = new MatchEvaluator(referenceToValue);
            string newExpression = regex.Replace(expr, evaluator);
            return newExpression;
        }
        public string referenceToValue(Match m)
        {
            if (dictionary.ContainsKey(m.Value))
                if(dictionary[m.Value] == "")
                {
                    return "0";
                }
                else
                {
                    return dictionary[m.Value];
                }
                return m.Value;
            
        }
        public string Calculate(string expression)
        {
            string result = null;
            try
            {
                result = Convert.ToString(Calculator.Evaluate(expression));
                if(result == "∞")
                {
                    result = "Dividion by zerro error";
                }
                return result;
            }
            catch
            {
                return "Error";
            }
        }
        public void AddRow(DataGridView dataGridView1)
        {
            List<Cell> newRow = new List<Cell>();
            for(int j = 0; j < colCount; j++)
            {
                newRow.Add(Cell(rowCount, j));
                dictionary.Add(newRow.Last().getName(), "");
            }
            grid.Add(newRow);
            RefreshReferences();
            foreach(List<Cell> list in grid)
            {
                foreach(Cell cell in list)
                {
                    if(cell.referencesFromThis != null)
                    {
                        foreach(Cell cellRef in cell.referencesFromThis)
                        {
                            if(cellRef.row == rowCount)
                            {
                                if (!cellRef.pointersToThis.Contains(cell))
                                {
                                    cellRef.pointersToThis.Add(cell);
                                }
                            }
                        }
                    }
                }
            }
            for(int j = 0; j < colCount; j++)
            {
                ChangeCellWithAllPointers(rowCount, j, "", dataGridView1);
            }
            rowCount++;
        }
        public void AddColumn(DataGridView dataGridView)
        {
            List<Cell> newCol = new List<Cell>();

            for (int j = 0; j < rowCount; j++)
            {
                string name = FullName(j, colCount);
                //newCol.Add(new Cell(ColCount, j));
                grid[j].Add(new Cell(j, colCount));
                dictionary.Add(name, "");
            }

            RefreshReferences();
            foreach (List<Cell> list in grid)
            {
                foreach (Cell cell in list)
                {
                    if (cell.referencesFromThis != null)
                    {
                        foreach (Cell cellRef in cell.referencesFromThis)
                        {
                            if (cellRef.column == colCount)
                            {
                                if (!cellRef.pointersToThis.Contains(cell))
                                    cellRef.pointersToThis.Add(cell);
                            }
                        }
                    }
                }
            }
            for (int j = 0; j < rowCount; j++)
            {
                ChangeCellWithAllPointers(j, colCount, "", dataGridView);
            }
            colCount++;
        }
        public bool DeleteRow(DataGridView dataGridView1)
        {
            List<Cell> lastRow = new List<Cell>();
            List<string> notEmptyCells = new List<string>();
            if(rowCount == 0)
            {
                return false;
            }
            int currentCount = rowCount - 1;
            for(int i=0; i < colCount; i++)
            {
                string name = FullName(currentCount, i);
                if(dictionary[name] != "0" && dictionary[name] != "" && dictionary[name] !=" ")
                {
                    notEmptyCells.Add(name);
                }
                if (grid[currentCount][i].pointersToThis.Count != 0)
                {
                    lastRow.AddRange(grid[currentCount][i].pointersToThis);
                }
            }
            if (lastRow.Count != 0 || notEmptyCells.Count != 0)
            {
                string errorMessage = "";
                if(notEmptyCells.Count != 0)
                {
                    errorMessage = "There are not empty cells:";
                    errorMessage += string.Join(";", notEmptyCells.ToArray());
                    errorMessage += ' ';
                }
                if(lastRow.Count != 0)
                {
                    errorMessage += "There are cells that point to cells from current row : ";
                    foreach(Cell cell in lastRow)
                    {
                        errorMessage += string.Join(";", cell.getName());
                        errorMessage += " ";
                    }
                }
                errorMessage += "Are you sure you want to delete this row?";
                DialogResult result = MessageBox.Show(errorMessage, "Warning", MessageBoxButtons.YesNo);
                if(result == DialogResult.No)
                {
                    return false;
                }
            }
            for(int i = 0; i < colCount; i++)
            {
                string name = FullName(currentCount, i);
                dictionary.Remove(name);
            }
            foreach(Cell cell in lastRow)
            {
                RefreshCellAndPointers(cell, dataGridView1);
            }
            grid.RemoveAt(currentCount);
            rowCount--;
            return true;
        }
        public bool DeleteColumn(DataGridView dataGridView1)
        {
            List<Cell> lastCol = new List<Cell>();
            List<string> notEmptyCells = new List<string>();
            if(colCount == 0)
            {
                return false;
            }
            int currentCount = colCount - 1;
            for(int i = 0; i< rowCount; i++)
            {
                string name = FullName(i, currentCount);
                if (dictionary[name] != "0" && dictionary[name] !=  "" && dictionary[name] != " ")
                {
                    notEmptyCells.Add(name);
                }
                if(grid[i][currentCount].pointersToThis.Count != 0)
                {
                    lastCol.AddRange(grid[i][currentCount].pointersToThis);
                }
            }
            if (lastCol.Count != 0 || notEmptyCells.Count != 0)
            {
                string errorMessage = "";
                if (notEmptyCells.Count != 0)
                {
                    errorMessage = "There are not empty cells: ";
                    errorMessage += string.Join(";", notEmptyCells.ToArray());
                }
                if (lastCol.Count != 0)
                {
                    errorMessage += "There are cells that point to cells from current column :";
                    foreach (Cell cell in lastCol)
                    {
                        errorMessage += string.Join(";", cell.getName());
                    }
                }
                errorMessage += " Are you sure you want to delete this column?";
                DialogResult result = MessageBox.Show(errorMessage, "Warning", MessageBoxButtons.YesNo);
                if (result == DialogResult.No)
                {
                    return false;
                }
            }
            for( int i = 0; i < rowCount; i++)
            {
                string name = FullName(i, currentCount);
                dictionary.Remove(name);
            }
            foreach (Cell cell in lastCol)
            {
                RefreshCellAndPointers(cell, dataGridView1);
            }
            for(int i = 0; i < rowCount; i++)
            {
                grid[i].RemoveAt(currentCount);
            }
            colCount--;
            return true;
        }
        public void Open(int r, int c, StreamReader streamReader, DataGridView dataGridView)
        {
            for (int i = 0; i < r; i++)
            {
                for(int j = 0; j < c; j++)
                {
                    string index = streamReader.ReadLine();
                    string expression = streamReader.ReadLine();
                    string value = streamReader.ReadLine();
                    if(expression != "")
                    {
                        dictionary[index] = value;
                    }
                    else
                    {
                        dictionary[index] = "";
                    }
                    int refCount = Convert.ToInt32(streamReader.ReadLine());
                    List<Cell> newRef = new List<Cell>();
                    string refer;
                    for(int k=0; k < refCount; k++)
                    {
                        refer = streamReader.ReadLine();
                        if(NumberConverter.From26System(refer).row < rowCount &&
                            NumberConverter.From26System(refer).column < colCount)
                        {
                            newRef.Add(grid[NumberConverter.From26System(refer).row][NumberConverter.From26System(refer).column]);
                        }
                    }
                    int pointCount = Convert.ToInt32(streamReader.ReadLine());
                    List<Cell> newPoint = new List<Cell>();
                    string point;
                    for(int k = 0; k < pointCount; k++)
                    {
                        point = streamReader.ReadLine();
                        newPoint.Add(grid[NumberConverter.From26System(point).row][NumberConverter.From26System(point).column]);
                    }
                    grid[i][j].setCell(expression, value, newRef, newPoint);
                    int currentCol = grid[i][j].column;
                    int currentRow = grid[i][j].row;
                    dataGridView[currentCol, currentRow].Value = dictionary[index];
                }
            }
        }
        public void Save(StreamWriter streamWriter)
        {
            streamWriter.WriteLine(rowCount);
            streamWriter.WriteLine(colCount);
            foreach(List<Cell> list in grid)
            {
                foreach(Cell cell in list)
                {
                    streamWriter.WriteLine(cell.getName());
                    streamWriter.WriteLine(cell.expression);
                    streamWriter.WriteLine(cell.value);
                    if (cell.referencesFromThis == null)
                    {
                        streamWriter.WriteLine("0");
                    }
                    else
                    {
                        streamWriter.WriteLine(cell.referencesFromThis.Count);
                        foreach (Cell refCell in cell.referencesFromThis)
                        {
                            streamWriter.WriteLine(refCell.getName());
                        }
                    }
                    if(cell.pointersToThis == null)
                    {
                        streamWriter.WriteLine("0");
                    }
                    else
                    {
                        streamWriter.WriteLine(cell.pointersToThis.Count);
                        foreach(Cell pointCell in cell.pointersToThis)
                        {
                            streamWriter.WriteLine(pointCell.getName());
                        }
                    }
                }
            }
        }
    }
}
