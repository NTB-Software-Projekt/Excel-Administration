using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MemberAdministration
{
    /// <summary>
    /// This class handles all Backup related tasks such as creating, checking and updating the Backup file.
    /// 
    /// </summary>
    class BackupManager
    {
        private DateTime currentSavedDate = MemberAdministration.Properties.Settings.Default.currentSavedDate;
        private String dbPath;
        private String targetDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+@"\xlBackup";
        private String fileName = MemberAdministration.Properties.Settings.Default.fileName;

        /// <summary>
        /// If the file current Backup file is out of date or there is none,
        /// a new File will be copyied and the old file will be overwritten.
        /// </summary>
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

        /// <summary>
        /// A quick check if the current Backup is up to date.
        /// If not, the last write time will be saved in the application properties.
        /// </summary>
        /// <returns>True if the file is up to date and False if it isn't.</returns>
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

        /// <summary>
        /// Setting the Filename and saving it in the application properties.
        /// This is needed for specifying wich file to Backup.
        /// </summary>
        /// <param name="fileName">Name of the Excel File to Backup</param>
        public void setFileName(String fileName)
        {
            this.fileName = fileName;
            MemberAdministration.Properties.Settings.Default.fileName = fileName;
        }

    }
}
