using AssistantForWord.SaveOption;
using OpenAI_API;
using OpenAI_API.Chat;
using OpenAI_API.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace AssistantForWord
{
    class OpenAIClient
    {
        internal static async Task<string> GetResponse(string input)
        {
            string response = string.Empty;
            var config = ProcessData.GetData();
            if (string.IsNullOrEmpty(config.APIKEY))
                return string.Empty;
            var client = new OpenAIAPI(config.APIKEY);            
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            try
            {               
                var result = await client.Chat.CreateChatCompletionAsync(new ChatRequest()
                {
                    Model = GetModel(config.ModelName),
                    Temperature = config.Temperature,
                    MaxTokens = config.TokenSize,
                    Messages = new ChatMessage[] {
                new ChatMessage(ChatMessageRole.User, input)
                }
                });
                response = result.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to get response" + ex.ToString());
            }
            var models = await client.Models.GetModelsAsync();
            return response;
        }
                
        private static Model GetModel(string modelName)
        {
            Model model = Model.ChatGPTTurbo;
            switch (modelName)
            {
                case "DefaultModel":
                    model = Model.DefaultModel;
                    break;
                case "GPT4_32k_Context":
                    model = Model.GPT4_32k_Context;
                    break;
                case "GPT4":
                    model = Model.GPT4;
                    break;
                case "ChatGPTTurbo0301":
                    model = Model.ChatGPTTurbo0301;
                    break;
                case "ChatGPTTurbo":
                    model = Model.ChatGPTTurbo;
                    break;
                case "AdaTextEmbedding":
                    model = Model.AdaTextEmbedding;
                    break;
                case "DavinciCode":
                    model = Model.DavinciCode;
                    break;
                case "CushmanCode":
                    model = Model.CushmanCode;
                    break;
                case "DavinciText":
                    model = Model.DavinciText;
                    break;
                case "CurieText":
                    model = Model.CurieText;
                    break;
                case "BabbageText":
                    model = Model.BabbageText;
                    break;
                case "AdaText":
                    model = Model.AdaText;
                    break;
                default:
                    break;
            }
            return model;
        }
    }

}
