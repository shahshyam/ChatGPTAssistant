using AssistantForWord.UI.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace AssistantForWord.SaveOption
{
    class ProcessData
    {
        private static List<PromptDetail> PromptDetailList { get; set; } = new List<PromptDetail>();
        private static string GetSaveDataFile()
        {
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WritingAssitant");
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return Path.Combine(folder, "Config.dll");
        }
        public static AssistantConfig GetData()
        {            
            var configuration = new AssistantConfig();
            string filePath = GetSaveDataFile();
            if (File.Exists(filePath))
            {
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    BinaryFormatter binaryFormatter = new BinaryFormatter();
                    configuration = binaryFormatter.Deserialize(fileStream) as AssistantConfig;
                    fileStream.Close();
                    fileStream.Dispose();
                }
            }
            if (PromptDetailList.Count == 0)
                PromptDetailList = configuration.PromptDetailList;
            return configuration;
        }
        public static void SaveData(AssistantConfig configuration)
        {
            PromptDetailList = configuration.PromptDetailList;
            string filePath = GetSaveDataFile();
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            {
                BinaryFormatter binaryFormatter = new BinaryFormatter();
                binaryFormatter.Serialize(fileStream, configuration);
                fileStream.Close();
                fileStream.Dispose();
            }
        }
        public static PromptDetail GetPromptDetailById(string guid)
        {
            Guid guId = new Guid(guid);
            return PromptDetailList.FirstOrDefault(x => x.Id == guId);
        }
    }
}
