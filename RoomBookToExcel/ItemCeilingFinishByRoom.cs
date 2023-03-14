using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace RoomBookToExcel
{
    public class ItemCeilingFinishByRoom
    {
        public string RoomNumber { get; set; }
        public string RoomName { get; set; }
        public List<CeilingType> CeilingTypesList { get; set; }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is ItemCeilingFinishByRoom))
            {
                return false;
            }

            ItemCeilingFinishByRoom other = (ItemCeilingFinishByRoom)obj;
            if (CeilingTypesList.Count != other.CeilingTypesList.Count)
            {
                return false;
            }

            for (int i = 0; i < CeilingTypesList.Count; i++)
            {
                if (CeilingTypesList[i].Id != other.CeilingTypesList[i].Id)
                {
                    return false;
                }
            }

            return true;
        }

        public override int GetHashCode()
        {
            int hash = 19;
            foreach (CeilingType item in CeilingTypesList)
            {
                hash = hash * 31 + item.Id.GetHashCode();
            }
            return hash;
        }
    }
}
