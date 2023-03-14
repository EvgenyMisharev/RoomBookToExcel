using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace RoomBookToExcel
{
    class ItemWallFinishByRoom
    {
        public string RoomNumber { get; set; }
        public string RoomName { get; set; }
        public List<WallType> WallTypesList { get; set; }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is ItemWallFinishByRoom))
            {
                return false;
            }

            ItemWallFinishByRoom other = (ItemWallFinishByRoom)obj;
            if (WallTypesList.Count != other.WallTypesList.Count)
            {
                return false;
            }

            for (int i = 0; i < WallTypesList.Count; i++)
            {
                if (WallTypesList[i].Id != other.WallTypesList[i].Id)
                {
                    return false;
                }
            }

            return true;
        }

        public override int GetHashCode()
        {
            int hash = 19;
            foreach (WallType item in WallTypesList)
            {
                hash = hash * 31 + item.Id.GetHashCode();
            }
            return hash;
        }
    }
}
