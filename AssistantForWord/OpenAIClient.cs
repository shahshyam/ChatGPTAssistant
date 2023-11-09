using AssistantForWord.SaveOption;
using AssistantForWord.UI.Helpers;
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
                var model = await GetModelById(config.ModelName);
                var result = await client.Chat.CreateChatCompletionAsync(new ChatRequest()
                {
                    Model = model,
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
                MessageBox.Show("Failed to get response" + ex.ToString(), AppConstant.AppTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
            }           
            return response;
        }

        public static async Task<List<Model>> GetModels()
        {
            var models = new List<Model>();
            var config = ProcessData.GetData();
            if (string.IsNullOrEmpty(config.APIKEY))
                return models;
            var client = new OpenAIAPI(config.APIKEY);
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            try
            {
                models = await client.Models.GetModelsAsync();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to get Model" + ex.ToString(), AppConstant.AppTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            return models;
        }
        public static async Task<Model> GetModelById(string modelId )
        {            
            var config = ProcessData.GetData();
            if (string.IsNullOrEmpty(config.APIKEY))
                return null;
            var client = new OpenAIAPI(config.APIKEY);
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            try
            {
               return await client.Models.RetrieveModelDetailsAsync(modelId);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to get Model detail" + ex.ToString(), AppConstant.AppTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            return null;
        }
    }
}
