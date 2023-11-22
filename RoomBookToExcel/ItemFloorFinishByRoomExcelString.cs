using System.Collections.Generic;

namespace RoomBookToExcel
{
    class ItemFloorFinishByRoomExcelString
    {
        public string RoomNumber { get; set; }
        public string RoomName { get; set; }
        public Dictionary<string, double> ItemData { get; set; }
    }
}
