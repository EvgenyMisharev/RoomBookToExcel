using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using OfficeOpenXml;
using System.Threading;

namespace RoomBookToExcel
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class RoomBookToExcelCommand : IExternalCommand
    {
        RoomBookToExcelProgressBarWPF roomBookToExcelProgressBarWPF;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            Guid roombookRoomNumber = new Guid("22868552-0e64-49b2-b8d9-9a2534bf0e14");
            Guid roombookRoomName = new Guid("b59a22a9-7890-45bd-9f93-a186341eef58");
            Guid elemData = new Guid("659c3180-6565-41bc-a332-d82502953510");

            List<Room> roomList = new FilteredElementCollector(doc)
                .OfClass(typeof(SpatialElement))
                .WhereElementIsNotElementType()
                .Where(r => r.GetType() == typeof(Room))
                .Cast<Room>()
                .Where(r => r.Area > 0)
                .OrderBy(r => r.Number, new AlphanumComparatorFastString())
                .ToList();
            if(roomList.Count == 0)
            {
                TaskDialog.Show("Revit", "Проект не содержит помещения!");
                return Result.Cancelled;
            }

            RoomBookToExcelWPF roomBookToExcelWPF = new RoomBookToExcelWPF();
            roomBookToExcelWPF.ShowDialog();
            if (roomBookToExcelWPF.DialogResult != true)
            {
                return Result.Cancelled;
            }
            string exportOptionName = roomBookToExcelWPF.ExportOptionName;

            if(exportOptionName == "rbt_FinishingForEachRoom")
            {
                Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();
                int step = 0;
                Thread.Sleep(100);
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0);
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = roomList.Count);

                // создаем новый пакет Excel
                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (ExcelPackage package = new ExcelPackage())
                    {
                        // создаем новый лист
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("RoomBook");

                        // задаем ширину столбцов
                        worksheet.Column(1).Width = 10;
                        worksheet.Column(2).Width = 30;
                        worksheet.Column(3).Width = 10;
                        worksheet.Column(4).Width = 10;
                        worksheet.Column(5).Width = 50;
                        worksheet.Column(6).Width = 10;
                        worksheet.Column(7).Width = 15;
                        worksheet.Column(8).Width = 20;

                        // объединяем ячейки заголовка 1
                        worksheet.Cells[1, 1, 1, 8].Merge = true;
                        // вписываем текст
                        worksheet.Cells[1, 1].Value = "Таблица вид 2";
                        worksheet.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[1, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[1, 1].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[1, 1].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[1, 1].Style.Font.Size = 10;

                        // объединяем ячейки заголовка 2
                        worksheet.Cells[2, 1, 2, 8].Merge = true;
                        // вписываем текст
                        worksheet.Cells[2, 1].Value = "Румбук - Спецификация помещений";
                        worksheet.Cells[2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, 1].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[2, 1].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[2, 1].Style.Font.Size = 14;
                        worksheet.Cells[2, 1].Style.Font.Bold = true;

                        // объединяем ячейки заголовка 3
                        worksheet.Cells[3, 1, 3, 8].Merge = true;
                        // вписываем текст
                        worksheet.Cells[3, 1].Value = "Ссылки на листы документации";
                        worksheet.Cells[3, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        worksheet.Cells[3, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[3, 1].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[3, 1].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[3, 1].Style.Font.Size = 10;

                        // добавляем заголовки
                        worksheet.Cells[4, 1].Value = "Номер помещения";
                        worksheet.Cells[4, 2].Value = "Имя помещения";
                        worksheet.Cells[4, 3].Value = "Тип элемента";
                        worksheet.Cells[4, 4].Value = "Марка элемента";
                        worksheet.Cells[4, 5].Value = "Наименование элемента";
                        worksheet.Cells[4, 6].Value = "Ед. изм";
                        worksheet.Cells[4, 7].Value = "Кол-во";
                        worksheet.Cells[4, 8].Value = "Примечание";
                        worksheet.Cells[4, 1, 4, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[4, 1, 4, 8].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[4, 1, 4, 8].Style.WrapText = true;
                        worksheet.Cells[4, 1, 4, 8].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[4, 1, 4, 8].Style.Font.Size = 10;

                        int row = 5;
                        foreach (Room room in roomList)
                        {
                            step++;
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = room.Name);

                            int startRow = row;
                            //Полы в помещении
                            List<Floor> floorList = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Floors)
                                .OfClass(typeof(Floor))
                                .WhereElementIsNotElementType()
                                .Cast<Floor>()
                                .Where(f => f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                                .Where(f => f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Пол"
                                || f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Полы")
                                .Where(f => f.get_Parameter(roombookRoomNumber) != null)
                                .Where(f => f.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                                .OrderBy(f => f.FloorType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
                                .ToList();

                            //Стены в помещении
                            List<Wall> wallList = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Walls)
                                .OfClass(typeof(Wall))
                                .WhereElementIsNotElementType()
                                .Cast<Wall>()
                                .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                                .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Отделка стен")
                                .Where(w => w.get_Parameter(roombookRoomNumber) != null)
                                .Where(w => w.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                                .OrderBy(w => w.WallType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
                                .ToList();

                            //Потолки в помещении
                            List<Ceiling> ceilingList = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Ceilings)
                                .OfClass(typeof(Ceiling))
                                .WhereElementIsNotElementType()
                                .Cast<Ceiling>()
                                .Where(c => doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                                .Where(c => doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолок"
                                || doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолки")
                                .Where(w => w.get_Parameter(roombookRoomNumber) != null)
                                .Where(w => w.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                                .OrderBy(c => doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
                                .ToList();

                            if (floorList.Count == 0 && wallList.Count == 0 && ceilingList.Count == 0)
                            {
                                continue;
                            }

                            //Обработка полов
                            List<FloorType> floorTypesList = new List<FloorType>();
                            List<ElementId> floorTypesIdList = new List<ElementId>();
                            foreach (Floor floor in floorList)
                            {
                                if (!floorTypesIdList.Contains(floor.FloorType.Id))
                                {
                                    floorTypesList.Add(floor.FloorType);
                                    floorTypesIdList.Add(floor.FloorType.Id);
                                }
                            }

                            floorTypesList = floorTypesList.OrderBy(ft => ft.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString()).ToList();
                            foreach (FloorType floorType in floorTypesList)
                            {
                                ItemForEachRoomExcelString floorItemForExcelString = new ItemForEachRoomExcelString();
                                floorItemForExcelString.RoomNumber = room.Number;
                                floorItemForExcelString.RoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                                floorItemForExcelString.ItemTypeDescription = "Пол";
                                floorItemForExcelString.ItemTypMark = floorType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString();
                                floorItemForExcelString.ItemData = floorType.get_Parameter(elemData).AsString();
                                floorItemForExcelString.ItemUnits = "м2";
                                double floorArea = 0;
                                List<Floor> tmpFloorList = floorList.Where(w => w.FloorType.Id == floorType.Id).ToList();
                                foreach (Floor floor in tmpFloorList)
                                {
#if R2019 || R2020 || R2021
                                    floorArea += UnitUtils.ConvertFromInternalUnits(floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                            floorArea += UnitUtils.ConvertFromInternalUnits(floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), UnitTypeId.SquareMeters);
#endif
                                }
                                floorItemForExcelString.ItemArea = Math.Round(floorArea, 2);

                                worksheet.Cells[row, 1].Value = floorItemForExcelString.RoomNumber;
                                worksheet.Cells[row, 2].Value = floorItemForExcelString.RoomName;
                                worksheet.Cells[row, 3].Value = floorItemForExcelString.ItemTypeDescription;
                                worksheet.Cells[row, 4].Value = floorItemForExcelString.ItemTypMark;
                                worksheet.Cells[row, 5].Value = floorItemForExcelString.ItemData;
                                worksheet.Cells[row, 6].Value = floorItemForExcelString.ItemUnits;
                                worksheet.Cells[row, 7].Value = floorItemForExcelString.ItemArea;
                                row++;
                            }

                            //Обработка стен
                            List<WallType> wallTypesList = new List<WallType>();
                            List<ElementId> wallTypesIdList = new List<ElementId>();
                            foreach (Wall wall in wallList)
                            {
                                if (!wallTypesIdList.Contains(wall.WallType.Id))
                                {
                                    wallTypesList.Add(wall.WallType);
                                    wallTypesIdList.Add(wall.WallType.Id);
                                }
                            }

                            wallTypesList = wallTypesList.OrderBy(wt => wt.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString()).ToList();
                            foreach (WallType wallType in wallTypesList)
                            {
                                ItemForEachRoomExcelString wallItemForExcelString = new ItemForEachRoomExcelString();
                                wallItemForExcelString.RoomNumber = room.Number;
                                wallItemForExcelString.RoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                                wallItemForExcelString.ItemTypeDescription = "Отделка\r\nстен";
                                wallItemForExcelString.ItemTypMark = wallType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString();
                                wallItemForExcelString.ItemData = wallType.get_Parameter(elemData).AsString();
                                wallItemForExcelString.ItemUnits = "м2";
                                double wallArea = 0;
                                List<Wall> tmpWallList = wallList.Where(w => w.WallType.Id == wallType.Id).ToList();
                                foreach (Wall wall in tmpWallList)
                                {
#if R2019 || R2020 || R2021
                                    wallArea += UnitUtils.ConvertFromInternalUnits(wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                                    wallArea += UnitUtils.ConvertFromInternalUnits(wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), UnitTypeId.SquareMeters);
#endif
                                }
                                wallItemForExcelString.ItemArea = Math.Round(wallArea, 2);

                                worksheet.Cells[row, 1].Value = wallItemForExcelString.RoomNumber;
                                worksheet.Cells[row, 2].Value = wallItemForExcelString.RoomName;
                                worksheet.Cells[row, 3].Value = wallItemForExcelString.ItemTypeDescription;
                                worksheet.Cells[row, 4].Value = wallItemForExcelString.ItemTypMark;
                                worksheet.Cells[row, 5].Value = wallItemForExcelString.ItemData;
                                worksheet.Cells[row, 6].Value = wallItemForExcelString.ItemUnits;
                                worksheet.Cells[row, 7].Value = wallItemForExcelString.ItemArea;
                                row++;
                            }

                            //Обработка потолков
                            List<CeilingType> ceilingTypesList = new List<CeilingType>();
                            List<ElementId> ceilingTypesIdList = new List<ElementId>();
                            foreach (Ceiling ceiling in ceilingList)
                            {
                                if (!ceilingTypesIdList.Contains(ceiling.GetTypeId()))
                                {
                                    ceilingTypesList.Add(doc.GetElement(ceiling.GetTypeId()) as CeilingType);
                                    ceilingTypesIdList.Add(ceiling.GetTypeId());
                                }
                            }

                            ceilingTypesList = ceilingTypesList.OrderBy(ct => ct.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString()).ToList();
                            foreach (CeilingType ceilingType in ceilingTypesList)
                            {
                                ItemForEachRoomExcelString ceilingItemForExcelString = new ItemForEachRoomExcelString();
                                ceilingItemForExcelString.RoomNumber = room.Number;
                                ceilingItemForExcelString.RoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                                ceilingItemForExcelString.ItemTypeDescription = "Потолок";
                                ceilingItemForExcelString.ItemTypMark = ceilingType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString();
                                ceilingItemForExcelString.ItemData = ceilingType.get_Parameter(elemData).AsString();
                                ceilingItemForExcelString.ItemUnits = "м2";
                                double ceilingArea = 0;
                                List<Ceiling> tmpCeilingList = ceilingList.Where(c => c.GetTypeId() == ceilingType.Id).ToList();
                                foreach (Ceiling ceiling in tmpCeilingList)
                                {
#if R2019 || R2020 || R2021
                                    ceilingArea += UnitUtils.ConvertFromInternalUnits(ceiling.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                                    ceilingArea += UnitUtils.ConvertFromInternalUnits(ceiling.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), UnitTypeId.SquareMeters);
#endif
                                }
                                ceilingItemForExcelString.ItemArea = Math.Round(ceilingArea, 2);

                                worksheet.Cells[row, 1].Value = ceilingItemForExcelString.RoomNumber;
                                worksheet.Cells[row, 2].Value = ceilingItemForExcelString.RoomName;
                                worksheet.Cells[row, 3].Value = ceilingItemForExcelString.ItemTypeDescription;
                                worksheet.Cells[row, 4].Value = ceilingItemForExcelString.ItemTypMark;
                                worksheet.Cells[row, 5].Value = ceilingItemForExcelString.ItemData;
                                worksheet.Cells[row, 6].Value = ceilingItemForExcelString.ItemUnits;
                                worksheet.Cells[row, 7].Value = ceilingItemForExcelString.ItemArea;
                                row++;
                            }
                            int endRow = row - 1;
                            worksheet.Cells[startRow, 1, endRow, 1].Merge = true;
                            worksheet.Cells[startRow, 2, endRow, 2].Merge = true;
                        }
                        roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());

                        worksheet.Cells[5, 1, row - 1, 8].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[5, 1, row - 1, 8].Style.WrapText = true;
                        worksheet.Cells[5, 1, row - 1, 8].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[5, 1, row - 1, 8].Style.Font.Size = 10;
                        worksheet.Cells[5, 1, row - 1, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[5, 5, row - 1, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        worksheet.Cells[5, 6, row - 1, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[5, 1, row - 1, 1].Style.Numberformat.Format = "0";

                        worksheet.Cells[4, 1, row - 1, 8].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[4, 1, row - 1, 8].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[4, 1, row - 1, 8].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[4, 1, row - 1, 8].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[4, 1, row - 1, 8].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);
                        worksheet.Cells[4, 1, 4, 8].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                        // сохраняем пакет Excel
                        System.Windows.Forms.SaveFileDialog saveDialog = new System.Windows.Forms.SaveFileDialog();
                        saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                        System.Windows.Forms.DialogResult result = saveDialog.ShowDialog();
                        string excelFilePath = "";
                        if (result == System.Windows.Forms.DialogResult.OK)
                        {
                            excelFilePath = saveDialog.FileName;
                            byte[] excelFile = package.GetAsByteArray();
                            File.WriteAllBytes(excelFilePath, excelFile);
                        }
                    }
                }
                catch (Exception theException)
                {
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);
                    TaskDialog.Show("Revit", errorMessage);
                    return Result.Cancelled;
                }
            }
            else if (exportOptionName == "rbt_FloorFinishByCombinationInRoom")
            {
                Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();
                Thread.Sleep(100);
                int step = 0;
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0);
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = roomList.Count);

                // создаем новый пакет Excel
                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (ExcelPackage package = new ExcelPackage())
                    {
                        List<ItemFloorFinishByRoom> itemFloorFinishByRoomList = new List<ItemFloorFinishByRoom>();
                        foreach (Room room in roomList)
                        {
                            step++;
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Сбор данных об отделке пола. Шаг {step} из {roomList.Count}");

                            ItemFloorFinishByRoom itemFloorFinishByRoom = new ItemFloorFinishByRoom();
                            itemFloorFinishByRoom.RoomNumber = room.Number;
                            itemFloorFinishByRoom.RoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();

                            //Стены в помещении
                            List<Floor> floorList = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Floors)
                                .OfClass(typeof(Floor))
                                .WhereElementIsNotElementType()
                                .Cast<Floor>()
                                .Where(w => w.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                                .Where(f => f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Пол"
                                || f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Полы")
                                .Where(w => w.get_Parameter(roombookRoomNumber) != null)
                                .Where(w => w.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                                .OrderBy(w => w.FloorType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
                                .ToList();

                            //Обработка стен
                            List<FloorType> floorTypesList = new List<FloorType>();
                            List<ElementId> floorTypesIdList = new List<ElementId>();
                            foreach (Floor floor in floorList)
                            {
                                if (!floorTypesIdList.Contains(floor.FloorType.Id))
                                {
                                    floorTypesList.Add(floor.FloorType);
                                    floorTypesIdList.Add(floor.FloorType.Id);
                                }
                            }

                            itemFloorFinishByRoom.FloorTypesList = floorTypesList.OrderBy(wt => wt.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString()).ToList();
                            itemFloorFinishByRoomList.Add(itemFloorFinishByRoom);

                        }

                        List<ItemFloorFinishByRoom> uniqueFloorFinishSet = itemFloorFinishByRoomList.Distinct(new ItemFloorFinishByRoomComparer()).ToList();
                        step = 0;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0);
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = uniqueFloorFinishSet.Count);

                        List<ItemFloorFinishByRoomExcelString> itemFloorFinishByRoomExcelStringList = new List<ItemFloorFinishByRoomExcelString>();
                        foreach (ItemFloorFinishByRoom uniqueFloorFinish in uniqueFloorFinishSet)
                        {
                            step++;
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Обработка сочетаний отделок. Шаг {step} из {uniqueFloorFinishSet.Count}");

                            ItemFloorFinishByRoomExcelString itemFloorFinishByRoomExcelString = new ItemFloorFinishByRoomExcelString();
                            itemFloorFinishByRoomExcelString.ItemData = new Dictionary<string, double>();
                            List<ItemFloorFinishByRoom> tmpItemFloorFinishList = itemFloorFinishByRoomList.Where(i => i.Equals(uniqueFloorFinish)).OrderBy(i => i.RoomNumber, new AlphanumComparatorFastString()).ToList();
                            List<string> roomNumbersList = new List<string>();
                            List<string> roomNamesList = new List<string>();

                            foreach (ItemFloorFinishByRoom tmpItemFloorFinish in tmpItemFloorFinishList)
                            {
                                if (!roomNumbersList.Contains(tmpItemFloorFinish.RoomNumber))
                                {
                                    roomNumbersList.Add(tmpItemFloorFinish.RoomNumber);
                                }
                                if (!roomNamesList.Contains(tmpItemFloorFinish.RoomName))
                                {
                                    roomNamesList.Add(tmpItemFloorFinish.RoomName);
                                }

                                foreach (FloorType floorType in tmpItemFloorFinish.FloorTypesList)
                                {
                                    List<Floor> tmpFloorList = new FilteredElementCollector(doc)
                                        .OfCategory(BuiltInCategory.OST_Floors)
                                        .OfClass(typeof(Floor))
                                        .WhereElementIsNotElementType()
                                        .Cast<Floor>()
                                        .Where(w => w.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                                        .Where(f => f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Пол"
                                        || f.FloorType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Полы")
                                        .Where(w => w.get_Parameter(roombookRoomNumber) != null)
                                        .Where(w => w.get_Parameter(roombookRoomNumber).AsString() == tmpItemFloorFinish.RoomNumber)
                                        .Where(w => w.FloorType.Id == floorType.Id)
                                        .ToList();

                                    double floorArea = 0;
                                    foreach (Floor floor in tmpFloorList)
                                    {
#if R2019 || R2020 || R2021
                                        floorArea += UnitUtils.ConvertFromInternalUnits(floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                                        floorArea += UnitUtils.ConvertFromInternalUnits(floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), UnitTypeId.SquareMeters);
#endif
                                    }

                                    string floorTypeDisc = floorType.get_Parameter(elemData).AsString();
                                    if (itemFloorFinishByRoomExcelString.ItemData.ContainsKey(floorTypeDisc))
                                    {
                                        itemFloorFinishByRoomExcelString.ItemData[floorTypeDisc] += Math.Round(floorArea, 2);
                                    }
                                    else
                                    {
                                        itemFloorFinishByRoomExcelString.ItemData.Add(floorTypeDisc, Math.Round(floorArea, 2));
                                    }
                                }
                            }
                            string roomNumbers = "";
                            roomNumbersList = roomNumbersList.OrderBy(n => n, new AlphanumComparatorFastString()).ToList();
                            foreach (string s in roomNumbersList)
                            {
                                if (roomNumbersList.IndexOf(s) != roomNumbersList.Count - 1)
                                {
                                    roomNumbers += $"{s}, ";
                                }
                                else
                                {
                                    roomNumbers += s;
                                }
                            }
                            itemFloorFinishByRoomExcelString.RoomNumber = roomNumbers;

                            string roomNames = "";
                            roomNamesList = roomNamesList.OrderBy(n => n, new AlphanumComparatorFastString()).ToList();
                            foreach (string s in roomNamesList)
                            {
                                if (roomNamesList.IndexOf(s) != roomNamesList.Count - 1)
                                {
                                    roomNames += $"{s}, ";
                                }
                                else
                                {
                                    roomNames += s;
                                }
                            }
                            itemFloorFinishByRoomExcelString.RoomName = roomNames;
                            itemFloorFinishByRoomExcelStringList.Add(itemFloorFinishByRoomExcelString);
                        }
                        roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                        int itemDataCnt = itemFloorFinishByRoomExcelStringList.Max(i => i.ItemData.Count) * 2;

                        // создаем новый лист
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("FloorFinish");

                        // задаем ширину столбцов
                        worksheet.Column(1).Width = 60;
                        worksheet.Column(2).Width = 120;
                        for (int i = 3; i <= itemDataCnt + 2; i += 2)
                        {
                            worksheet.Column(i).Width = 65;
                            worksheet.Column(i + 1).Width = 15;
                        }

                        worksheet.Column(itemDataCnt + 3).Width = 20;

                        // объединяем ячейки заголовка 1
                        worksheet.Cells[1, 1, 1, itemDataCnt + 3].Merge = true;
                        // вписываем текст
                        worksheet.Cells[1, 1].Value = "Ведомость отделки стен";
                        worksheet.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[1, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[1, 1].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[1, 1].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[1, 1].Style.Font.Size = 10;
                        worksheet.Cells[1, 1].Style.Font.Bold = true;

                        // объединяем ячейки заголовка 2
                        worksheet.Cells[2, 1, 3, 1].Merge = true;
                        // вписываем текст
                        worksheet.Cells[2, 1].Value = "Номера помещений";
                        worksheet.Cells[2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, 1].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[2, 1].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[2, 1].Style.Font.Size = 10;


                        // объединяем ячейки заголовка 3
                        worksheet.Cells[2, 2, 3, 2].Merge = true;
                        // вписываем текст
                        worksheet.Cells[2, 2].Value = "Наименования помещений";
                        worksheet.Cells[2, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, 2].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[2, 2].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[2, 2].Style.Font.Size = 10;

                        // объединяем ячейки заголовка 4
                        worksheet.Cells[2, 3, 2, itemDataCnt + 2].Merge = true;
                        // вписываем текст
                        worksheet.Cells[2, 3].Value = "Типы отделки помещений";
                        worksheet.Cells[2, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, 3].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[2, 3].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[2, 3].Style.Font.Size = 10;

                        //Заполняем заголовок отделки 6
                        int typeCnt = 1;
                        for (int i = 3; i <= itemDataCnt + 2; i += 2)
                        {
                            // вписываем текст
                            worksheet.Cells[3, i].Value = $"Отделка пола тип {typeCnt}";
                            worksheet.Cells[3, i + 1].Value = "Площ. м2";
                            typeCnt++;
                        }
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.Font.Size = 10;

                        // объединяем ячейки заголовка 7
                        worksheet.Cells[2, itemDataCnt + 3, 3, itemDataCnt + 3].Merge = true;
                        // вписываем текст
                        worksheet.Cells[2, itemDataCnt + 3].Value = "Примечание";
                        worksheet.Cells[2, itemDataCnt + 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, itemDataCnt + 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, itemDataCnt + 3].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[2, itemDataCnt + 3].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[2, itemDataCnt + 3].Style.Font.Size = 10;

                        for (int i = 4; i < itemFloorFinishByRoomExcelStringList.Count + 4; i++)
                        {
                            worksheet.Cells[i, 1].Value = itemFloorFinishByRoomExcelStringList[i - 4].RoomNumber;
                            worksheet.Cells[i, 2].Value = itemFloorFinishByRoomExcelStringList[i - 4].RoomName;
                            for (int j = 0; j < itemFloorFinishByRoomExcelStringList[i - 4].ItemData.Count; j++)
                            {
                                worksheet.Cells[i, j * 2 + 3].Value = itemFloorFinishByRoomExcelStringList[i - 4].ItemData.ElementAt(j).Key;
                                worksheet.Cells[i, j * 2 + 3].Style.WrapText = true;
                                worksheet.Cells[i, j * 2 + 3].Style.Font.Name = "ISOCPEUR";
                                worksheet.Cells[i, j * 2 + 3].Style.Font.Size = 10;
                                worksheet.Cells[i, j * 2 + 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                worksheet.Cells[i, j * 2 + 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                worksheet.Cells[i, j * 2 + 4].Value = itemFloorFinishByRoomExcelStringList[i - 4].ItemData.ElementAt(j).Value;
                                worksheet.Cells[i, j * 2 + 4].Style.WrapText = true;
                                worksheet.Cells[i, j * 2 + 4].Style.Font.Name = "ISOCPEUR";
                                worksheet.Cells[i, j * 2 + 4].Style.Font.Size = 10;
                                worksheet.Cells[i, j * 2 + 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet.Cells[i, j * 2 + 4].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            }
                        }

                        //форматирование
                        worksheet.Cells[4, 1, 4 + itemFloorFinishByRoomExcelStringList.Count, 2].Style.WrapText = true;
                        worksheet.Cells[4, 1, 4 + itemFloorFinishByRoomExcelStringList.Count, 2].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[4, 1, 4 + itemFloorFinishByRoomExcelStringList.Count, 2].Style.Font.Size = 10;
                        worksheet.Cells[4, 1, 4 + itemFloorFinishByRoomExcelStringList.Count, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[4, 1, 4 + itemFloorFinishByRoomExcelStringList.Count, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        //рамка
                        worksheet.Cells[2, 1, itemFloorFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 1, itemFloorFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, itemFloorFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, itemFloorFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 1, itemFloorFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                        worksheet.Cells[2, 1, 3, itemDataCnt + 3].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, 3, itemDataCnt + 3].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, 3, itemDataCnt + 3].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, 3, itemDataCnt + 3].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

                        // сохраняем пакет Excel
                        System.Windows.Forms.SaveFileDialog saveDialog = new System.Windows.Forms.SaveFileDialog();
                        saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                        System.Windows.Forms.DialogResult result = saveDialog.ShowDialog();
                        string excelFilePath = "";
                        if (result == System.Windows.Forms.DialogResult.OK)
                        {
                            excelFilePath = saveDialog.FileName;
                            byte[] excelFile = package.GetAsByteArray();
                            File.WriteAllBytes(excelFilePath, excelFile);
                        }
                    }

                }
                catch (Exception theException)
                {
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);
                    TaskDialog.Show("Revit", errorMessage);
                    return Result.Cancelled;
                }
            }
            else if(exportOptionName == "rbt_WallFinishByCombinationInRoom")
            {
                Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();
                Thread.Sleep(100);
                int step = 0;
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0);
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = roomList.Count);

                // создаем новый пакет Excel
                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (ExcelPackage package = new ExcelPackage())
                    {
                        List<ItemWallFinishByRoom> itemWallFinishByRoomList = new List<ItemWallFinishByRoom>();
                        foreach (Room room in roomList)
                        {
                            step++;
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Сбор данных об отделке стен. Шаг {step} из {roomList.Count}");

                            ItemWallFinishByRoom itemWallFinishByRoom = new ItemWallFinishByRoom();
                            itemWallFinishByRoom.RoomNumber = room.Number;
                            itemWallFinishByRoom.RoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();

                            //Стены в помещении
                            List<Wall> wallList = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Walls)
                                .OfClass(typeof(Wall))
                                .WhereElementIsNotElementType()
                                .Cast<Wall>()
                                .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                                .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Отделка стен")
                                .Where(w => w.get_Parameter(roombookRoomNumber) != null)
                                .Where(w => w.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                                .OrderBy(w => w.WallType.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
                                .ToList();

                            //Обработка стен
                            List<WallType> wallTypesList = new List<WallType>();
                            List<ElementId> wallTypesIdList = new List<ElementId>();
                            foreach (Wall wall in wallList)
                            {
                                if (!wallTypesIdList.Contains(wall.WallType.Id))
                                {
                                    wallTypesList.Add(wall.WallType);
                                    wallTypesIdList.Add(wall.WallType.Id);
                                }
                            }

                            itemWallFinishByRoom.WallTypesList = wallTypesList.OrderBy(wt => wt.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString()).ToList();
                            itemWallFinishByRoomList.Add(itemWallFinishByRoom);

                        }

                        List<ItemWallFinishByRoom> uniqueWallFinishSet = itemWallFinishByRoomList.Distinct(new ItemWallFinishByRoomComparer()).ToList();
                        step = 0;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0);
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = uniqueWallFinishSet.Count);

                        List<ItemWallFinishByRoomExcelString> itemWallFinishByRoomExcelStringList = new List<ItemWallFinishByRoomExcelString>();
                        foreach (ItemWallFinishByRoom uniqueWallFinish in uniqueWallFinishSet)
                        {
                            step++;
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Обработка сочетаний отделок. Шаг {step} из {uniqueWallFinishSet.Count}");

                            ItemWallFinishByRoomExcelString itemWallFinishByRoomExcelString = new ItemWallFinishByRoomExcelString();
                            itemWallFinishByRoomExcelString.ItemData = new Dictionary<string, double>();
                            List<ItemWallFinishByRoom> tmpItemWallFinishList = itemWallFinishByRoomList.Where(i => i.Equals(uniqueWallFinish)).OrderBy(i => i.RoomNumber, new AlphanumComparatorFastString()).ToList();
                            List<string> roomNumbersList = new List<string>();
                            List<string> roomNamesList = new List<string>();

                            foreach (ItemWallFinishByRoom tmpItemWallFinish in tmpItemWallFinishList)
                            {
                                if (!roomNumbersList.Contains(tmpItemWallFinish.RoomNumber))
                                {
                                    roomNumbersList.Add(tmpItemWallFinish.RoomNumber);
                                }
                                if (!roomNamesList.Contains(tmpItemWallFinish.RoomName))
                                {
                                    roomNamesList.Add(tmpItemWallFinish.RoomName);
                                }

                                foreach (WallType wallType in tmpItemWallFinish.WallTypesList)
                                {
                                    List<Wall> tmpWallList = new FilteredElementCollector(doc)
                                        .OfCategory(BuiltInCategory.OST_Walls)
                                        .OfClass(typeof(Wall))
                                        .WhereElementIsNotElementType()
                                        .Cast<Wall>()
                                        .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                                        .Where(w => w.WallType.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Отделка стен")
                                        .Where(w => w.get_Parameter(roombookRoomNumber) != null)
                                        .Where(w => w.get_Parameter(roombookRoomNumber).AsString() == tmpItemWallFinish.RoomNumber)
                                        .Where(w => w.WallType.Id == wallType.Id)
                                        .ToList();

                                    double wallArea = 0;
                                    foreach (Wall wall in tmpWallList)
                                    {
#if R2019 || R2020 || R2021
                                        wallArea += UnitUtils.ConvertFromInternalUnits(wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                                        wallArea += UnitUtils.ConvertFromInternalUnits(wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), UnitTypeId.SquareMeters);
#endif
                                    }

                                    string wallTypeDisc = wallType.get_Parameter(elemData).AsString();
                                    if (itemWallFinishByRoomExcelString.ItemData.ContainsKey(wallTypeDisc))
                                    {
                                        itemWallFinishByRoomExcelString.ItemData[wallTypeDisc] += Math.Round(wallArea, 2);
                                    }
                                    else
                                    {
                                        itemWallFinishByRoomExcelString.ItemData.Add(wallTypeDisc, Math.Round(wallArea, 2));
                                    }
                                }
                            }
                            string roomNumbers = "";
                            roomNumbersList = roomNumbersList.OrderBy(n => n, new AlphanumComparatorFastString()).ToList();
                            foreach (string s in roomNumbersList)
                            {
                                if (roomNumbersList.IndexOf(s) != roomNumbersList.Count - 1)
                                {
                                    roomNumbers += $"{s}, ";
                                }
                                else
                                {
                                    roomNumbers += s;
                                }
                            }
                            itemWallFinishByRoomExcelString.RoomNumber = roomNumbers;

                            string roomNames = "";
                            roomNamesList = roomNamesList.OrderBy(n => n, new AlphanumComparatorFastString()).ToList();
                            foreach (string s in roomNamesList)
                            {
                                if (roomNamesList.IndexOf(s) != roomNamesList.Count - 1)
                                {
                                    roomNames += $"{s}, ";
                                }
                                else
                                {
                                    roomNames += s;
                                }
                            }
                            itemWallFinishByRoomExcelString.RoomName = roomNames;
                            itemWallFinishByRoomExcelStringList.Add(itemWallFinishByRoomExcelString);
                        }
                        roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                        int itemDataCnt = itemWallFinishByRoomExcelStringList.Max(i => i.ItemData.Count) * 2;

                        // создаем новый лист
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("WallFinish");

                        // задаем ширину столбцов
                        worksheet.Column(1).Width = 60;
                        worksheet.Column(2).Width = 120;
                        for (int i = 3; i <= itemDataCnt + 2; i+=2)
                        {
                            worksheet.Column(i).Width = 65;
                            worksheet.Column(i+1).Width = 15;
                        }

                        worksheet.Column(itemDataCnt + 3).Width = 20;

                        // объединяем ячейки заголовка 1
                        worksheet.Cells[1, 1, 1, itemDataCnt + 3].Merge = true;
                        // вписываем текст
                        worksheet.Cells[1, 1].Value = "Ведомость отделки стен";
                        worksheet.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[1, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[1, 1].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[1, 1].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[1, 1].Style.Font.Size = 10;
                        worksheet.Cells[1, 1].Style.Font.Bold = true;

                        // объединяем ячейки заголовка 2
                        worksheet.Cells[2, 1, 3, 1].Merge = true;
                        // вписываем текст
                        worksheet.Cells[2, 1].Value = "Номера помещений";
                        worksheet.Cells[2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, 1].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[2, 1].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[2, 1].Style.Font.Size = 10;


                        // объединяем ячейки заголовка 3
                        worksheet.Cells[2, 2, 3, 2].Merge = true;
                        // вписываем текст
                        worksheet.Cells[2, 2].Value = "Наименования помещений";
                        worksheet.Cells[2, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, 2].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[2, 2].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[2, 2].Style.Font.Size = 10;

                        // объединяем ячейки заголовка 4
                        worksheet.Cells[2, 3, 2, itemDataCnt + 2].Merge = true;
                        // вписываем текст
                        worksheet.Cells[2, 3].Value = "Типы отделки помещений";
                        worksheet.Cells[2, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, 3].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[2, 3].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[2, 3].Style.Font.Size = 10;

                        //Заполняем заголовок отделки 6
                        int typeCnt = 1;
                        for (int i = 3; i <= itemDataCnt + 2; i += 2)
                        {
                            // вписываем текст
                            worksheet.Cells[3, i].Value = $"Отделка стен тип {typeCnt}";
                            worksheet.Cells[3, i + 1].Value = "Площ. м2";
                            typeCnt++;
                        }
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.Font.Size = 10;

                        // объединяем ячейки заголовка 7
                        worksheet.Cells[2, itemDataCnt + 3, 3, itemDataCnt + 3].Merge = true;
                        // вписываем текст
                        worksheet.Cells[2, itemDataCnt + 3].Value = "Примечание";
                        worksheet.Cells[2, itemDataCnt + 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, itemDataCnt + 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, itemDataCnt + 3].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[2, itemDataCnt + 3].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[2, itemDataCnt + 3].Style.Font.Size = 10;

                        for (int i = 4; i < itemWallFinishByRoomExcelStringList.Count + 4; i++)
                        {
                            worksheet.Cells[i, 1].Value = itemWallFinishByRoomExcelStringList[i - 4].RoomNumber;
                            worksheet.Cells[i, 2].Value = itemWallFinishByRoomExcelStringList[i - 4].RoomName;
                            for (int j = 0; j < itemWallFinishByRoomExcelStringList[i - 4].ItemData.Count; j++)
                            {
                                worksheet.Cells[i, j * 2 + 3].Value = itemWallFinishByRoomExcelStringList[i - 4].ItemData.ElementAt(j).Key;
                                worksheet.Cells[i, j * 2 + 3].Style.WrapText = true;
                                worksheet.Cells[i, j * 2 + 3].Style.Font.Name = "ISOCPEUR";
                                worksheet.Cells[i, j * 2 + 3].Style.Font.Size = 10;
                                worksheet.Cells[i, j * 2 + 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                worksheet.Cells[i, j * 2 + 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                worksheet.Cells[i, j * 2 + 4].Value = itemWallFinishByRoomExcelStringList[i - 4].ItemData.ElementAt(j).Value;
                                worksheet.Cells[i, j * 2 + 4].Style.WrapText = true;
                                worksheet.Cells[i, j * 2 + 4].Style.Font.Name = "ISOCPEUR";
                                worksheet.Cells[i, j * 2 + 4].Style.Font.Size = 10;
                                worksheet.Cells[i, j * 2 + 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet.Cells[i, j * 2 + 4].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            }
                        }

                        //форматирование
                        worksheet.Cells[4, 1, 4 + itemWallFinishByRoomExcelStringList.Count, 2].Style.WrapText = true;
                        worksheet.Cells[4, 1, 4 + itemWallFinishByRoomExcelStringList.Count, 2].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[4, 1, 4 + itemWallFinishByRoomExcelStringList.Count, 2].Style.Font.Size = 10;
                        worksheet.Cells[4, 1, 4 + itemWallFinishByRoomExcelStringList.Count, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[4, 1, 4 + itemWallFinishByRoomExcelStringList.Count, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        //рамка
                        worksheet.Cells[2, 1, itemWallFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 1, itemWallFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, itemWallFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, itemWallFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 1, itemWallFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                        worksheet.Cells[2, 1, 3, itemDataCnt + 3].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, 3, itemDataCnt + 3].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, 3, itemDataCnt + 3].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, 3, itemDataCnt + 3].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        
                        // сохраняем пакет Excel
                        System.Windows.Forms.SaveFileDialog saveDialog = new System.Windows.Forms.SaveFileDialog();
                        saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                        System.Windows.Forms.DialogResult result = saveDialog.ShowDialog();
                        string excelFilePath = "";
                        if (result == System.Windows.Forms.DialogResult.OK)
                        {
                            excelFilePath = saveDialog.FileName;
                            byte[] excelFile = package.GetAsByteArray();
                            File.WriteAllBytes(excelFilePath, excelFile);
                        }
                    }
  
                }
                catch (Exception theException)
                {
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);
                    TaskDialog.Show("Revit", errorMessage);
                    return Result.Cancelled;
                }
            }
            else
            {
                Thread newWindowThread = new Thread(new ThreadStart(ThreadStartingPoint));
                newWindowThread.SetApartmentState(ApartmentState.STA);
                newWindowThread.IsBackground = true;
                newWindowThread.Start();
                Thread.Sleep(100);
                int step = 0;
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0);
                roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = roomList.Count);

                // создаем новый пакет Excel
                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (ExcelPackage package = new ExcelPackage())
                    {
                        List<ItemCeilingFinishByRoom> itemCeilingFinishByRoomList = new List<ItemCeilingFinishByRoom>();
                        foreach (Room room in roomList)
                        {
                            step++;
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Сбор данных об отделке стен. Шаг {step} из {roomList.Count}");

                            ItemCeilingFinishByRoom itemCeilingFinishByRoom = new ItemCeilingFinishByRoom();
                            itemCeilingFinishByRoom.RoomNumber = room.Number;
                            itemCeilingFinishByRoom.RoomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();

                            //Стены в помещении
                            List<Ceiling> ceilingList = new FilteredElementCollector(doc)
                                .OfCategory(BuiltInCategory.OST_Ceilings)
                                .OfClass(typeof(Ceiling))
                                .WhereElementIsNotElementType()
                                .Cast<Ceiling>()
                                .Where(c => doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                                .Where(c => doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолок"
                                || doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолки")
                                .Where(c => c.get_Parameter(roombookRoomNumber) != null)
                                .Where(c => c.get_Parameter(roombookRoomNumber).AsString() == room.Number)
                                .OrderBy(c => doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString())
                                .ToList();

                            //Обработка стен
                            List<CeilingType> ceilingTypesList = new List<CeilingType>();
                            List<ElementId> ceilingTypesIdList = new List<ElementId>();
                            foreach (Ceiling ceiling in ceilingList)
                            {
                                if (!ceilingTypesIdList.Contains(ceiling.GetTypeId()))
                                {
                                    ceilingTypesList.Add(doc.GetElement(ceiling.GetTypeId()) as CeilingType);
                                    ceilingTypesIdList.Add(ceiling.GetTypeId());
                                }
                            }

                            itemCeilingFinishByRoom.CeilingTypesList = ceilingTypesList.OrderBy(wt => wt.get_Parameter(BuiltInParameter.WINDOW_TYPE_ID).AsString(), new AlphanumComparatorFastString()).ToList();
                            itemCeilingFinishByRoomList.Add(itemCeilingFinishByRoom);

                        }

                        List<ItemCeilingFinishByRoom> uniqueCeilingFinishSet = itemCeilingFinishByRoomList.Distinct(new ItemCeilingFinishByRoomComparer()).ToList();
                        step = 0;
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Minimum = 0);
                        roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Maximum = uniqueCeilingFinishSet.Count);

                        List<ItemCeilingFinishByRoomExcelString> itemCeilingFinishByRoomExcelStringList = new List<ItemCeilingFinishByRoomExcelString>();
                        foreach (ItemCeilingFinishByRoom uniqueCeilingFinish in uniqueCeilingFinishSet)
                        {
                            step++;
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Value = step);
                            roomBookToExcelProgressBarWPF.pb_RoomBookToExcelProgressBar.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.label_ItemName.Content = $"Обработка сочетаний отделок. Шаг {step} из {uniqueCeilingFinishSet.Count}");

                            ItemCeilingFinishByRoomExcelString itemCeilingFinishByRoomExcelString = new ItemCeilingFinishByRoomExcelString();
                            itemCeilingFinishByRoomExcelString.ItemData = new Dictionary<string, double>();
                            List<ItemCeilingFinishByRoom> tmpItemCeilingFinishList = itemCeilingFinishByRoomList.Where(i => i.Equals(uniqueCeilingFinish)).OrderBy(i => i.RoomNumber, new AlphanumComparatorFastString()).ToList();
                            List<string> roomNumbersList = new List<string>();
                            List<string> roomNamesList = new List<string>();

                            foreach (ItemCeilingFinishByRoom tmpItemCeilingFinish in tmpItemCeilingFinishList)
                            {
                                if (!roomNumbersList.Contains(tmpItemCeilingFinish.RoomNumber))
                                {
                                    roomNumbersList.Add(tmpItemCeilingFinish.RoomNumber);
                                }
                                if (!roomNamesList.Contains(tmpItemCeilingFinish.RoomName))
                                {
                                    roomNamesList.Add(tmpItemCeilingFinish.RoomName);
                                }

                                foreach (CeilingType ceilingType in tmpItemCeilingFinish.CeilingTypesList)
                                {
                                    List<Ceiling> tmpCeilingList = new FilteredElementCollector(doc)
                                        .OfCategory(BuiltInCategory.OST_Ceilings)
                                        .OfClass(typeof(Ceiling))
                                        .WhereElementIsNotElementType()
                                        .Cast<Ceiling>()
                                        .Where(c => doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL) != null)
                                        .Where(c => doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолок"
                                        || doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Потолки")
                                        .Where(c => c.get_Parameter(roombookRoomNumber) != null)
                                        .Where(c => c.get_Parameter(roombookRoomNumber).AsString() == tmpItemCeilingFinish.RoomNumber)
                                        .Where(c => doc.GetElement(c.GetTypeId()).Id == ceilingType.Id)
                                        .ToList();

                                    double ceilingArea = 0;
                                    foreach (Ceiling ceiling in tmpCeilingList)
                                    {
#if R2019 || R2020 || R2021
                                        ceilingArea += UnitUtils.ConvertFromInternalUnits(ceiling.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), DisplayUnitType.DUT_SQUARE_METERS);
#else
                                        ceilingArea += UnitUtils.ConvertFromInternalUnits(ceiling.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble(), UnitTypeId.SquareMeters);
#endif
                                    }

                                    string ceilingTypeDisc = ceilingType.get_Parameter(elemData).AsString();
                                    if (itemCeilingFinishByRoomExcelString.ItemData.ContainsKey(ceilingTypeDisc))
                                    {
                                        itemCeilingFinishByRoomExcelString.ItemData[ceilingTypeDisc] += Math.Round(ceilingArea, 2);
                                    }
                                    else
                                    {
                                        itemCeilingFinishByRoomExcelString.ItemData.Add(ceilingTypeDisc, Math.Round(ceilingArea, 2));
                                    }
                                }
                            }
                            string roomNumbers = "";
                            roomNumbersList = roomNumbersList.OrderBy(n => n, new AlphanumComparatorFastString()).ToList();
                            foreach (string s in roomNumbersList)
                            {
                                if (roomNumbersList.IndexOf(s) != roomNumbersList.Count - 1)
                                {
                                    roomNumbers += $"{s}, ";
                                }
                                else
                                {
                                    roomNumbers += s;
                                }
                            }
                            itemCeilingFinishByRoomExcelString.RoomNumber = roomNumbers;

                            string roomNames = "";
                            roomNamesList = roomNamesList.OrderBy(n => n, new AlphanumComparatorFastString()).ToList();
                            foreach (string s in roomNamesList)
                            {
                                if (roomNamesList.IndexOf(s) != roomNamesList.Count - 1)
                                {
                                    roomNames += $"{s}, ";
                                }
                                else
                                {
                                    roomNames += s;
                                }
                            }
                            itemCeilingFinishByRoomExcelString.RoomName = roomNames;
                            itemCeilingFinishByRoomExcelStringList.Add(itemCeilingFinishByRoomExcelString);
                        }
                        roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                        int itemDataCnt = itemCeilingFinishByRoomExcelStringList.Max(i => i.ItemData.Count) * 2;

                        // создаем новый лист
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("CeilingFinish");

                        // задаем ширину столбцов
                        worksheet.Column(1).Width = 60;
                        worksheet.Column(2).Width = 120;
                        for (int i = 3; i <= itemDataCnt + 2; i += 2)
                        {
                            worksheet.Column(i).Width = 65;
                            worksheet.Column(i + 1).Width = 15;
                        }

                        worksheet.Column(itemDataCnt + 3).Width = 20;

                        // объединяем ячейки заголовка 1
                        worksheet.Cells[1, 1, 1, itemDataCnt + 3].Merge = true;
                        // вписываем текст
                        worksheet.Cells[1, 1].Value = "Ведомость отделки стен";
                        worksheet.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[1, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[1, 1].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[1, 1].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[1, 1].Style.Font.Size = 10;
                        worksheet.Cells[1, 1].Style.Font.Bold = true;

                        // объединяем ячейки заголовка 2
                        worksheet.Cells[2, 1, 3, 1].Merge = true;
                        // вписываем текст
                        worksheet.Cells[2, 1].Value = "Номера помещений";
                        worksheet.Cells[2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, 1].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[2, 1].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[2, 1].Style.Font.Size = 10;


                        // объединяем ячейки заголовка 3
                        worksheet.Cells[2, 2, 3, 2].Merge = true;
                        // вписываем текст
                        worksheet.Cells[2, 2].Value = "Наименования помещений";
                        worksheet.Cells[2, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, 2].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[2, 2].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[2, 2].Style.Font.Size = 10;

                        // объединяем ячейки заголовка 4
                        worksheet.Cells[2, 3, 2, itemDataCnt + 2].Merge = true;
                        // вписываем текст
                        worksheet.Cells[2, 3].Value = "Типы отделки помещений";
                        worksheet.Cells[2, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, 3].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[2, 3].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[2, 3].Style.Font.Size = 10;

                        //Заполняем заголовок отделки 6
                        int typeCnt = 1;
                        for (int i = 3; i <= itemDataCnt + 2; i += 2)
                        {
                            // вписываем текст
                            worksheet.Cells[3, i].Value = $"Отделка потолка тип {typeCnt}";
                            worksheet.Cells[3, i + 1].Value = "Площ. м2";
                            typeCnt++;
                        }
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[3, 3, 3, itemDataCnt + 2].Style.Font.Size = 10;

                        // объединяем ячейки заголовка 7
                        worksheet.Cells[2, itemDataCnt + 3, 3, itemDataCnt + 3].Merge = true;
                        // вписываем текст
                        worksheet.Cells[2, itemDataCnt + 3].Value = "Примечание";
                        worksheet.Cells[2, itemDataCnt + 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[2, itemDataCnt + 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        worksheet.Cells[2, itemDataCnt + 3].Style.WrapText = true;
                        // устанавливаем шрифт
                        worksheet.Cells[2, itemDataCnt + 3].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[2, itemDataCnt + 3].Style.Font.Size = 10;

                        for (int i = 4; i < itemCeilingFinishByRoomExcelStringList.Count + 4; i++)
                        {
                            worksheet.Cells[i, 1].Value = itemCeilingFinishByRoomExcelStringList[i - 4].RoomNumber;
                            worksheet.Cells[i, 2].Value = itemCeilingFinishByRoomExcelStringList[i - 4].RoomName;
                            for (int j = 0; j < itemCeilingFinishByRoomExcelStringList[i - 4].ItemData.Count; j++)
                            {
                                worksheet.Cells[i, j * 2 + 3].Value = itemCeilingFinishByRoomExcelStringList[i - 4].ItemData.ElementAt(j).Key;
                                worksheet.Cells[i, j * 2 + 3].Style.WrapText = true;
                                worksheet.Cells[i, j * 2 + 3].Style.Font.Name = "ISOCPEUR";
                                worksheet.Cells[i, j * 2 + 3].Style.Font.Size = 10;
                                worksheet.Cells[i, j * 2 + 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                worksheet.Cells[i, j * 2 + 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                worksheet.Cells[i, j * 2 + 4].Value = itemCeilingFinishByRoomExcelStringList[i - 4].ItemData.ElementAt(j).Value;
                                worksheet.Cells[i, j * 2 + 4].Style.WrapText = true;
                                worksheet.Cells[i, j * 2 + 4].Style.Font.Name = "ISOCPEUR";
                                worksheet.Cells[i, j * 2 + 4].Style.Font.Size = 10;
                                worksheet.Cells[i, j * 2 + 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                worksheet.Cells[i, j * 2 + 4].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            }
                        }

                        //форматирование
                        worksheet.Cells[4, 1, 4 + itemCeilingFinishByRoomExcelStringList.Count, 2].Style.WrapText = true;
                        worksheet.Cells[4, 1, 4 + itemCeilingFinishByRoomExcelStringList.Count, 2].Style.Font.Name = "ISOCPEUR";
                        worksheet.Cells[4, 1, 4 + itemCeilingFinishByRoomExcelStringList.Count, 2].Style.Font.Size = 10;
                        worksheet.Cells[4, 1, 4 + itemCeilingFinishByRoomExcelStringList.Count, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        worksheet.Cells[4, 1, 4 + itemCeilingFinishByRoomExcelStringList.Count, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        //рамка
                        worksheet.Cells[2, 1, itemCeilingFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 1, itemCeilingFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, itemCeilingFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, itemCeilingFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells[2, 1, itemCeilingFinishByRoomExcelStringList.Count + 4, itemDataCnt + 3].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Medium);

                        worksheet.Cells[2, 1, 3, itemDataCnt + 3].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, 3, itemDataCnt + 3].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, 3, itemDataCnt + 3].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                        worksheet.Cells[2, 1, 3, itemDataCnt + 3].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

                        // сохраняем пакет Excel
                        System.Windows.Forms.SaveFileDialog saveDialog = new System.Windows.Forms.SaveFileDialog();
                        saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                        System.Windows.Forms.DialogResult result = saveDialog.ShowDialog();
                        string excelFilePath = "";
                        if (result == System.Windows.Forms.DialogResult.OK)
                        {
                            excelFilePath = saveDialog.FileName;
                            byte[] excelFile = package.GetAsByteArray();
                            File.WriteAllBytes(excelFilePath, excelFile);
                        }
                    }

                }
                catch (Exception theException)
                {
                    roomBookToExcelProgressBarWPF.Dispatcher.Invoke(() => roomBookToExcelProgressBarWPF.Close());
                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);
                    TaskDialog.Show("Revit", errorMessage);
                    return Result.Cancelled;
                }
            }
            return Result.Succeeded;
        }
        private void ThreadStartingPoint()
        {
            roomBookToExcelProgressBarWPF = new RoomBookToExcelProgressBarWPF();
            roomBookToExcelProgressBarWPF.Show();
            System.Windows.Threading.Dispatcher.Run();
        }

    }
}
