using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Laboratorna1_Excel_
{
    class Cell
    {
        public string expression { get; set; }
        public string value { get; set; }
        public int row { get; set; }
        public int column { get; set; }
        string name { get; set; }
        public List<Cell> pointersToThis = new List<Cell>();
        public List<Cell> referencesFromThis = new List<Cell>();
        public List<Cell> newReferencesFromThis = new List<Cell>();
        public Cell(int row, int column)
        {
            this.row = row;
            this.column = column;
            name = NumberConverter.To26System(column) + Convert.ToString(row);
            value = "0";
            expression = "";
        }
        public void setCell(string expr, string val, List<Cell> references, List<Cell> pointers)
        {
            value = val;
            expression = expr;
            referencesFromThis.Clear();
            referencesFromThis.AddRange(references);
            pointersToThis.Clear();
            pointersToThis.AddRange(pointers);
        }
        public string getName()
        {
            return name;
        }
        public bool CheckLoop(List<Cell> list)
        {
            foreach(Cell cell in list){
                if (cell.name == name)
                    return false;
            }
            foreach(Cell point in pointersToThis)
            {
                foreach(Cell cell in list)
                {
                    if(cell.name == point.name)
                    {
                        return false;
                    }
                }
                if (!point.CheckLoop(list))
                {
                    return false;
                }
            }
            return true;
        }
        public void AddPointersAndReferences()
        {
            foreach(Cell point in newReferencesFromThis)
            {
                point.pointersToThis.Add(this);
            }
            referencesFromThis = newReferencesFromThis;
        }
        public void DeletePointersAndReferences()
        {
            if (referencesFromThis != null)
            {
                foreach (Cell point in referencesFromThis)
                {
                    point.pointersToThis.Remove(this);
                }
                referencesFromThis = null;
            }
        }
    }
}
