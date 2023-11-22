using System.Collections.Generic;

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
