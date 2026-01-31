using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Acad = Autodesk.AutoCAD.ApplicationServices.Application;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

[assembly: CommandClass(typeof(MakeTextRegular.MakeTextRegularClass))]

namespace MakeTextRegular
{
    public class MakeTextRegularClass : IExtensionApplication
    {
        public void Initialize()
        {
            //throw new NotImplementedException();
        }

        [CommandMethod("MAKETEXTREGULAR")]
        public void MakeTextRegular()
        {
            // получаем БД и Editor текущего документа
            Document doc = Acad.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // создаем переменную, в которой будут содержаться данные для фильтра
            TypedValue[] filterlist = new TypedValue[1];

            // первый аргумент (0) указывает, что мы задаем тип объекта
            // второй аргумент ("MTEXT") - собственно тип
            filterlist[0] = new TypedValue(0, "MTEXT,TEXT,ACAD_TABLE");

            // создаем фильтр
            SelectionFilter filter = new(filterlist);

            // пытаемся получить ссылки на объекты с учетом фильтра
            // ВНИМАНИЕ! Нужно проверить работоспособность метода с замороженными и заблокированными слоями!
            PromptSelectionResult selRes = ed.SelectAll(filter);

            // если произошла ошибка - сообщаем о ней
            if (selRes.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nError!\n");
                return;
            }

            // получаем массив ID объектов
            ObjectId[] ids = selRes.Value.GetObjectIds();

            // начинаем транзакцию
            using Transaction tr = db.TransactionManager.StartTransaction();
            // "пробегаем" по всем полученным объектам
            foreach (ObjectId id in ids)
            {
                // приводим каждый из них к типу MText
                if (tr.GetObject(id, OpenMode.ForRead) is MText mtxt)
                {
                    // получаем значение текста
                    var text = mtxt.Text.Replace("\r\n", "\\P").Replace("\r", "\\P");
                    // открываем объект на запись
                    mtxt.UpgradeOpen();
                    // устанавливаем текст
                    mtxt.Contents = @"{\Q0;" + text.Replace("\r\n", "\\P") + @"}";
                }
                // приводим каждый из них к типу DBText
                else if (tr.GetObject(id, OpenMode.ForRead) is DBText txt)
                {
                    // получаем значение текста
                    var oblique = txt.Oblique;
                    // открываем объект на запись
                    txt.UpgradeOpen();
                    oblique = 0;
                    // устанавливаем наклон
                    txt.Oblique = oblique;
                }
                else if (tr.GetObject(id, OpenMode.ForRead) is Table tbl)
                {
                    string[,] arr = new string[tbl.Rows.Count, tbl.Columns.Count];
                    for (var i = 0; i < tbl.Rows.Count; i++)
                    {
                        for (var j = 0; j < tbl.Columns.Count; j++)
                        {
                            arr[i, j] = tbl.Cells[i, j].TextString;
                        }
                    }
                    // очистка форматов
                    for (var i = 0; i < arr.GetLength(0); i++)
                    {
                        for (var j = 0; j < arr.GetLength(1); j++)
                        {
                            var text = arr[i, j];
                            if (text.StartsWith('{'))
                            {
                                text = text.Trim('{', '}');
                                var vals = text.Split(';');
                                if (vals.Length == 2)
                                {
                                    text = vals[1];
                                }
                            }
                            arr[i, j] = text;
                        }
                    }
                    // принудительное отключение наклона
                    for (var i = 0; i < arr.GetLength(0); i++)
                    {
                        for (var j = 0; j < arr.GetLength(1); j++)
                        {
                            var text = arr[i, j].Replace("\r\n", "\\P").Replace("\r", "\\P");
                            arr[i, j] = @"{\Q0;" + text + @"}";
                        }
                    }
                    // открываем объект на запись
                    tbl.UpgradeOpen();
                    // запись содержимого в таблицу
                    for (var i = 0; i < tbl.Rows.Count; i++)
                    {
                        for (var j = 0; j < tbl.Columns.Count; j++)
                        {
                            try
                            {
                                tbl.Cells[i, j].TextString = arr[i, j];
                            }
                            catch { }
                        }
                    }
                }
            }
            tr.Commit();
        }

        public void Terminate()
        {
            //throw new NotImplementedException();
        }
    }
}
