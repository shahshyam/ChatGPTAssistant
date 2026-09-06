using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssistantForWord.UI.ViewModels
{
   public class PromptViewModel : ViewModelBase
    {
        private string title;
        public string Title
        {
            get { return title; }
            set { title = value;
                OnPropertyChanged();
            }
        }

        private string desciption;
        public string Description
        {
            get { return desciption; }
            set
            {
                desciption = value;
                OnPropertyChanged();
            }
        }

        private Guid id = Guid.NewGuid();
        public Guid Id
        {
            get { return id; }
            set
            {
                id = value;
                OnPropertyChanged();
            }
        }

        private string promptName;
        public string PromptName
        {
            get { return promptName; }
            set { promptName = value;
                OnPropertyChanged();
            }
        }


    }
}
