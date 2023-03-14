using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoomBookToExcel
{
    class ItemFloorFinishByRoomComparer : IEqualityComparer<ItemFloorFinishByRoom>
    {
        public bool Equals(ItemFloorFinishByRoom x, ItemFloorFinishByRoom y)
        {
            return x.Equals(y);
        }

        public int GetHashCode(ItemFloorFinishByRoom obj)
        {
            return obj.GetHashCode();
        }
    }
}
