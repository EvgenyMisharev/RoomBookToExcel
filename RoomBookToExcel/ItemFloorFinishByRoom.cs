using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoomBookToExcel
{
    class ItemFloorFinishByRoom
    {
        public string RoomNumber { get; set; }
        public string RoomName { get; set; }
        public List<FloorType> FloorTypesList { get; set; }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is ItemFloorFinishByRoom))
            {
                return false;
            }

            ItemFloorFinishByRoom other = (ItemFloorFinishByRoom)obj;
            if (FloorTypesList.Count != other.FloorTypesList.Count)
            {
                return false;
            }

            for (int i = 0; i < FloorTypesList.Count; i++)
            {
                if (FloorTypesList[i].Id != other.FloorTypesList[i].Id)
                {
                    return false;
                }
            }

            return true;
        }

        public override int GetHashCode()
        {
            int hash = 19;
            foreach (FloorType item in FloorTypesList)
            {
                hash = hash * 31 + item.Id.GetHashCode();
            }
            return hash;
        }
    }
}
