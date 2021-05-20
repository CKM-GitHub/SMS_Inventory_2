using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SMS_Entity;
using SMS_BL;

using System.Drawing;

namespace SMS
{
    public class CommonFunctions
    {
        public string DSP_MSG(string messageID, string messageParam1, string messageParam2, string messageParam3, string messageParam4, string messageParam5)
        {
            string message1, message2, message3, message = string.Empty;
            DialogResult dialog = DialogResult.None;

            try
            {
                bool result = GET_MSG(messageID, out message1, out message2, out message3);
                if (result)
                {
                    message = message1 + (string.IsNullOrEmpty(message2) ? string.Empty : "\n" + message2)
                        + (string.IsNullOrEmpty(message3) ? string.Empty : "\n" + message3);
                    message = message.Replace("%1", messageParam1);
                    message = message.Replace("%2", messageParam2);
                    message = message.Replace("%3", messageParam3);
                    message = message.Replace("%4", messageParam4);
                    message = message.Replace("%5", messageParam5);

                    switch (messageID.Substring(0, 1))
                    {
                        case "I":
                            dialog = MessageBox.Show(message + "","メッセージ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            break;
                        case "C":
                        case "Q":
                            dialog = MessageBox.Show(message + "", "メッセージ", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                            break;
                        case "E":
                            dialog = MessageBox.Show(message + "", "メッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            break;
                    }
                }
                else
                {
                    dialog = MessageBox.Show("メッセージが取得出来ませんでした。" + messageID, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception)
            {
            }

            switch (dialog)
            {
                case DialogResult.OK:
                    return "1";
                case DialogResult.Cancel:
                    return "2";
                default:
                    return "9";
            }
        }
    
        protected bool GET_MSG(string messageID, out string message1, out string message2, out string message3)
        {
            Message_BL message_bl = new Message_BL();
            SMSG_Entity smsg_data = new SMSG_Entity();
            smsg_data.MSGID = messageID.Substring(1);
            smsg_data.MSGSB = messageID.Substring(0, 1);
            return message_bl.SMSG_Select(smsg_data, out message1, out message2, out message3);
        }

        public IEnumerable<Control> GetAllControls(Control root)
        {
            foreach (Control control in root.Controls)
            {
                foreach (Control child in GetAllControls(control))
                {
                    yield return child;
                }
            }
            yield return root;
        }
    }
}
