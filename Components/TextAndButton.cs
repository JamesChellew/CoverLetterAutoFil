using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoverLetterAutoFil.Components
{
    public partial class TextAndButton : Component
    {
        public TextAndButton()
        {
            InitializeComponent();
        }

        public TextAndButton(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }
    }
}
