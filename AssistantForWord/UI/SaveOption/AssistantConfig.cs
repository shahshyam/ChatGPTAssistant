using AssistantForWord.UI.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssistantForWord.SaveOption
{
    [Serializable]
    public class AssistantConfig
    {
        public AssistantConfig()
        {
            PromptDetailList = new List<PromptDetail>();           
        }
        public List<PromptDetail> PromptDetailList { get; set; }
        public string APIKEY { get; set; }        
        public string ModelName { get; set; } = OpenAI_API.Models.Model.ChatGPTTurbo;
        public int TokenSize { get; set; } = 500;
        public double Temperature { get; set; } = 0.1;
    }

   
}
