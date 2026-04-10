using System.Windows.Forms;

namespace DrawioPpt.PowerPointAddIn.Services
{
    public class UserNotifier
    {
        public void ShowInfo(string message, string caption)
        {
            MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
