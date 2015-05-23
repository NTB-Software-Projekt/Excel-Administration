using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MemberAdministration
{
    class BackupManager
    {
        private DateTime currentSavedDate = MemberAdministration.Properties.Settings.Default.currentSavedDate;
        private String dbPath;
        private String targetDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+@"\xlBackup";
        private String fileName = MemberAdministration.Properties.Settings.Default.fileName;

        public void saveFile()
        {
            if (!fileUpToDate() && !String.IsNullOrEmpty(fileName))
            {
                if (!System.IO.Directory.Exists(targetDirectory))
                {
                    System.IO.Directory.CreateDirectory(targetDirectory);
                }
                String pathToSave = targetDirectory + "\\"+fileName;
                System.IO.File.Copy(dbPath, pathToSave, true);
            }
        }

        private bool fileUpToDate()
        {
            
            dbPath = MemberAdministration.Properties.Settings.Default.dbPath;
            if (!String.IsNullOrEmpty(dbPath))
            {
                DateTime lastWriteTime = File.GetLastWriteTime(dbPath);
                if (currentSavedDate == null || currentSavedDate != lastWriteTime)
                {
                    currentSavedDate = lastWriteTime;
                    MemberAdministration.Properties.Settings.Default.currentSavedDate = currentSavedDate;
                    return false;
                }
            }
            return true;
        }

        public void setFileName(String fileName)
        {
            this.fileName = fileName;
            MemberAdministration.Properties.Settings.Default.fileName = fileName;

        }

    }
}
