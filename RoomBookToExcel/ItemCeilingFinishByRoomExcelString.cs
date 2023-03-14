using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoomBookToExcel
{
    class ItemCeilingFinishByRoomExcelString
    {
        public string RoomNumber { get; set; }
        public string RoomName { get; set; }
        public Dictionary<string, double> ItemData { get; set; }
    }
}
