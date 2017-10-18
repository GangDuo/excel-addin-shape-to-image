using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Excel.AddIn.Shape2Image
{
    class Workspace
    {
        public DirectoryInfo WorkingDirectory { get; private set; }
        public string FullName
        {
            get { return WorkingDirectory.FullName; }
        }

        public void Open()
        {
            var desktop = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string workbookName = Guid.NewGuid().ToString("N").Substring(0, 16).ToUpper();
            try
            {
                // 未保存エクセルのName取得で例外発生
                workbookName = Path.GetFileNameWithoutExtension(Globals.ThisAddIn.Application.ActiveWorkbook.Name);
            }
            catch (Exception) { }
            WorkingDirectory = Directory.CreateDirectory(Path.Combine(desktop, workbookName));
        }

        public void Close()
        {
            WorkingDirectory.Delete(true);
        }
    }
}
