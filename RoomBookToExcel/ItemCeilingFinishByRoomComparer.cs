using System.Collections.Generic;

namespace RoomBookToExcel
{
    class ItemCeilingFinishByRoomComparer : IEqualityComparer<ItemCeilingFinishByRoom>
    {
        public bool Equals(ItemCeilingFinishByRoom x, ItemCeilingFinishByRoom y)
        {
            return x.Equals(y);
        }

        public int GetHashCode(ItemCeilingFinishByRoom obj)
        {
            return obj.GetHashCode();
        }
    }
}
