using AssistantForWord.SaveOption;
using AssistantForWord.UI.Commands;
using AssistantForWord.UI.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace AssistantForWord.UI.ViewModels
{
   public class CustomPanelViewModel:ViewModelBase
    {
        public ObservableCollection<PromptViewModel> PromptList { get; set; }

        public CustomPanelViewModel()
        {
            PromptList = new ObservableCollection<PromptViewModel>();            
        }
        internal void LoadExistingData()
        {
            var config = ProcessData.GetData();
            APIKey = config.APIKEY;
            AllowPrompt = !string.IsNullOrEmpty(APIKey);
            Temperature = config.Temperature;
            TokenSize = config.TokenSize;
            SelectedModelName = config.ModelName;
            foreach (var prompt in config.PromptDetailList)
            {
                PromptList.Add(new PromptViewModel()
                {
                    Description = prompt.Description,
                    Title = prompt.Title,
                    Id = prompt.Id
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
        public List<string> ModelList
        {
            get => new List<string>() { "DefaultModel",
            "GPT4_32k_Context","GPT4" , "ChatGPTTurbo0301","ChatGPTTurbo",
            "AdaTextEmbedding","DavinciCode","CushmanCode","DavinciText","CurieText",
            "BabbageText","AdaText"};
        }

        private ICommand saveKeyCommand;
        public ICommand SaveKeyCommand
        {
            get
            {
                return saveKeyCommand ?? (saveKeyCommand = new RelayCommand((x) =>
                {
                    var config = ProcessData.GetData();
                    config.APIKEY = this.APIKey;
                    config.Temperature = Temperature;
                    config.ModelName = SelectedModelName;
                    config.TokenSize = TokenSize;
                    ProcessData.SaveData(config);
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

        #endregion
        #region PROMPT
        private string title;
        public string Title
        {
            get { return title; }
            set { title = value;
                OnPropertyChanged();
            }
        }
        private string description;
        public string Description
        {
            get { return description; }
            set { description = value;
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
                    Title = SelectedPromptDetail.Title;
                    Description = SelectedPromptDetail.Description;
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
                        selectedItem.Title = this.Title;
                        selectedItem.Description = this.Description;
                        var config = ProcessData.GetData();
                        int index = config.PromptDetailList.FindIndex(z => z.Id == id);
                        if (index > -1)
                        {
                            config.PromptDetailList[index].Title = Title;
                            config.PromptDetailList[index].Description = Description;
                            ProcessData.SaveData(config);
                        }
                        IsExpand = false;
                    }
                    else
                    {
                        var promptVM = new PromptViewModel()
                        {
                            Description = this.Description,
                            Title = this.Title
                        };
                        var proptdetail = new PromptDetail()
                        {
                            Description = this.Description,
                            Title = this.Title,
                            Id = promptVM.Id
                        };
                        PromptList.Add(promptVM);
                        Description = string.Empty;
                        Title = string.Empty;

                        var config = ProcessData.GetData();
                        config.PromptDetailList.Add(proptdetail);
                        ProcessData.SaveData(config);
                    }
                    Mode = OperationMode.Add;                   
                }, y => { return !string.IsNullOrEmpty(Title) && !string.IsNullOrEmpty(Description); }
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
