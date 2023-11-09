using AssistantForWord.SaveOption;
using AssistantForWord.UI.Commands;
using AssistantForWord.UI.Models;
using OpenAIModel = OpenAI_API.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Data;

namespace AssistantForWord.UI.ViewModels
{
   public class CustomPanelViewModel:ViewModelBase
    {
        public ObservableCollection<PromptViewModel> PromptList { get; set; }
        public ObservableCollection<string> ModelList { get; set; }
        private static object _lock = new object();
        public CustomPanelViewModel()
        {
            PromptList = new ObservableCollection<PromptViewModel>();
            ModelList = new ObservableCollection<string>();
            BindingOperations.EnableCollectionSynchronization(ModelList, _lock);
        }
        internal void LoadExistingData()
        {
            if (PromptList.Count > 0)
                PromptList.Clear();
            var config = ProcessData.GetData();
            APIKey = config.APIKEY;
            AllowPrompt = !string.IsNullOrEmpty(APIKey);
            IsConfigVisible= !string.IsNullOrEmpty(APIKey);
            Temperature = config.Temperature;
            TokenSize = config.TokenSize;
            if (!string.IsNullOrEmpty(config.APIKEY))
            {
                LoadModels();
                
            }
            
            SelectedModelName = config.ModelName;
            foreach (var prompt in config.PromptDetailList)
            {
                PromptList.Add(new PromptViewModel()
                {
                    Description = prompt.Description,
                    Title = prompt.Title,
                    Id = prompt.Id,
                    PromptName = prompt.Name
                }
                );
            }
        }
        #region API Setting
        private string apiKey;
        public string APIKey
        {
            get { return apiKey; }
            set { apiKey = value;
                OnPropertyChanged();
            }
        }

        private string selectedModelName;
        public string SelectedModelName
        {
            get { return selectedModelName; }
            set { selectedModelName = value;
                OnPropertyChanged();
            }
        }

        private int tokenSize;
        public int TokenSize
        {
            get { return tokenSize; }
            set { tokenSize = value;
                OnPropertyChanged();
            }
        }

        private double temperature;
        public double Temperature
        {
            get { return temperature; }
            set { temperature = value;
                OnPropertyChanged();
            }
        }
        private bool isConfigVisible;

        public bool IsConfigVisible
        {
            get { return isConfigVisible; }
            set { isConfigVisible = value;
                OnPropertyChanged();
            }
        }

       private ICommand saveKeyCommand;
        public ICommand SaveKeyCommand
        {
            get
            {
                return saveKeyCommand ?? (saveKeyCommand = new RelayCommand((x) =>
                {
                    SaveConfig();
                    IsConfigVisible = true;
                    LoadModels();
                }, y => { return !string.IsNullOrEmpty(APIKey); }
                ));
            }
        }
        private ICommand saveOpenAIConfigCommand;
        public ICommand SaveOpenAIConfigCommand
        {
            get
            {
                return saveOpenAIConfigCommand ?? (saveOpenAIConfigCommand = new RelayCommand((x) =>
                {
                    SaveConfig();
                    AllowPrompt = true;
                }, y => { return !string.IsNullOrEmpty(APIKey); }
                ));
            }
        }
        private bool allowPrompt;

        public bool AllowPrompt
        {
            get { return allowPrompt; }
            set { allowPrompt = value;
                OnPropertyChanged();
            }
        }        
        private async void LoadModels()
        {
            var models = await OpenAIClient.GetModels();
            if (ModelList.Count > 0)
                ModelList.Clear();
            foreach(var model in models)
            {
                ModelList.Add(model.ModelID);
            }
            //ModelList = new ObservableCollection<OpenAIModel.Model>(models);
        }
        private void SaveConfig()
        {
            var config = ProcessData.GetData();
            config.APIKEY = this.APIKey;
            config.Temperature = Temperature;
            config.ModelName = SelectedModelName;
            config.TokenSize = TokenSize;
            ProcessData.SaveData(config);
        }
        #endregion
        #region PROMPT
        
        private PromptViewModel newPromptDetail = new PromptViewModel();
        public PromptViewModel NewPromptDetail
        {
            get { return newPromptDetail; }
            set
            {
                newPromptDetail = value;
                OnPropertyChanged();
            }
        }

        private PromptViewModel selectedPromptDetail;
        public PromptViewModel SelectedPromptDetail
        {
            get { return selectedPromptDetail; }
            set { selectedPromptDetail = value;
                OnPropertyChanged();
            }
        }

        private bool isExpand = false;
        public bool IsExpand
        {
            get { return isExpand; }
            set
            {
                isExpand = value;
                OnPropertyChanged();
            }
        }

        public OperationMode Mode = OperationMode.Add;
        private ICommand editPromptCommand;
        public ICommand EditPromptCommand
        {
            get
            {
                return editPromptCommand ?? (editPromptCommand = new RelayCommand((x) =>
                {
                    NewPromptDetail = SelectedPromptDetail;
                    Mode = OperationMode.Edit;
                    IsExpand = true;
                }));
            }
        }

        private ICommand savePromptCommand;
        public ICommand SavePromptCommand
        {
            get
            {
                return savePromptCommand ?? (savePromptCommand = new RelayCommand((x) =>
                {
                    if (Mode == OperationMode.Edit)
                    {
                        Guid id = SelectedPromptDetail.Id;
                        var selectedItem = PromptList.FirstOrDefault(y => y.Id == id);

                        selectedItem = NewPromptDetail;
                        var config = ProcessData.GetData();
                        int index = config.PromptDetailList.FindIndex(z => z.Id == id);
                        if (index > -1)
                        {
                            config.PromptDetailList[index].Title = NewPromptDetail.Title;
                            config.PromptDetailList[index].Description = NewPromptDetail.Description;
                            config.PromptDetailList[index].Name = NewPromptDetail.PromptName;
                            ProcessData.SaveData(config);
                        }
                        IsExpand = false;
                        SelectedPromptDetail = null;
                    }
                    else
                    {
                        var proptdetail = new PromptDetail()
                        {
                            Description = NewPromptDetail.Description,
                            Title = NewPromptDetail.Title,
                            Id = NewPromptDetail.Id,
                            Name = NewPromptDetail.PromptName
                        };
                        PromptList.Add(NewPromptDetail);
                        var config = ProcessData.GetData();
                        config.PromptDetailList.Add(proptdetail);
                        ProcessData.SaveData(config);
                    }
                    NewPromptDetail = new PromptViewModel();
                    Mode = OperationMode.Add;
                }, y =>
                {
                    return (NewPromptDetail != null &&
              !string.IsNullOrEmpty(NewPromptDetail.Title)
              && !string.IsNullOrEmpty(NewPromptDetail.Description)
              && !string.IsNullOrEmpty(NewPromptDetail.PromptName)
              );
                }
                ));
            }
        }
        private ICommand deletePromptCommand;
        public ICommand DeletePromptCommand
        {
            get
            {
                return deletePromptCommand ?? (deletePromptCommand = new RelayCommand((param) =>
                {
                    if (Mode == OperationMode.Add && SelectedPromptDetail != null)
                    {
                        var config = ProcessData.GetData();
                        int index = config.PromptDetailList.FindIndex(x => x.Id == SelectedPromptDetail.Id);
                        if (index > -1)
                        {
                            config.PromptDetailList.RemoveAt(index);
                            ProcessData.SaveData(config);
                        }
                        PromptList.Remove(SelectedPromptDetail);
                    }
                }
                ));
            }
        }
        #endregion
    }
    public enum OperationMode
    {
        Edit,
        Add
    }
}
