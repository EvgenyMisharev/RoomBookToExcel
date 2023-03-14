using System.Collections.Generic;

namespace RoomBookToExcel
{
    class ItemWallFinishByRoomComparer : IEqualityComparer<ItemWallFinishByRoom>
    {
        public bool Equals(ItemWallFinishByRoom x, ItemWallFinishByRoom y)
        {
            return x.Equals(y);
        }

        public int GetHashCode(ItemWallFinishByRoom obj)
        {
            return obj.GetHashCode();
        }
    }
}
