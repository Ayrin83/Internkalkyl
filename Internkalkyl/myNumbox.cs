using System.Windows.Forms;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
using System.Text.RegularExpressions;

namespace Internkalkyl
{
    public class myNumbox : TextBox
    {
        
        private Regex unacceptableChar;
        private string acceptableText;
        
        public myNumbox() : base()
        {
            unacceptableChar = new Regex("[^\\d.,-]");
        
        }


        protected override void OnTextChanged(System.EventArgs e)
        {
            if (unacceptableChar.IsMatch(this.Text))
            {
                Match match = unacceptableChar.Match(this.Text);

                //Logg.toLog("Found number of not num " + match.Value);
                        
                int selStart = this.SelectionStart - 1;

                if (selStart < 0)
                    selStart = 0;
                
                this.Text = acceptableText;
                this.SelectionStart = selStart;
            }
            else
            {
                acceptableText = this.Text;
            }
            
            base.OnTextChanged(e);
        }
    }
}
