using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop;
using System.Windows.Forms;

namespace JitsiMeetOutlook
{
    public class Utils
    {
        private static string escapeDomain(string domain)
        {
            string escapedDomain = Regex.Escape(domain);
            if (!escapedDomain.EndsWith("/"))
            {
                escapedDomain += "/";
            }
            return escapedDomain;
        }

        public static void HandleErrorWithUserNotification(Exception ex, string hint = null)
        {
            Log.Error(ex, hint ?? "Unexpected Error");
            MessageBox.Show(hint ?? ("An Error occured within JitsiOutlook: " + ex.Message));
        }

        public static string GetUrl(string oldBody, string domain)
        {
            return Regex.Match(oldBody, "http[s]*://" + escapeDomain(domain) + "[\\w\\/#%\\.=]+").Value;
        }

        public static bool SettingIsActive(string url, string setting)
        {
            return url.Contains("config." + setting + "=true");
        }


        public static string findRoomId(string body, string domain)
        {
            string roomId = Regex.Match(body, "(?<=" + escapeDomain(domain) + ")\\S+?(?=(#config|&config|\\s))").Value; // Match all non-blanks after jitsi url and before config or end
            return roomId;
        }

        public static string getNewRoomId()
        {
            if (Properties.Settings.Default.roomID.Length == 0)
            {
                return JitsiUrl.generateRandomId();
            }
            else
            {
                return Properties.Settings.Default.roomID;
            }
        }

        public static async System.Threading.Tasks.Task appendNewMeetingText(Microsoft.Office.Interop.Outlook.AppointmentItem appointmentItem, string roomId)
        {
            Microsoft.Office.Interop.Word.Document wordDocument = appointmentItem.GetInspector.WordEditor as Microsoft.Office.Interop.Word.Document;
            wordDocument.Select();
            var endSel = wordDocument.Application.Selection;
            endSel.Collapse(WdCollapseDirection.wdCollapseEnd);

            endSel.Font.Size = Constants.MainBodyTextSize;
            endSel.Font.Name = Constants.Font;

            var phoneNumbers = await Globals.ThisAddIn.JitsiApiService.getPhoneNumbers(roomId);
            var pinNumber = await Globals.ThisAddIn.JitsiApiService.getPIN(roomId);
            object missing = System.Reflection.Missing.Value;

            var link = JitsiUrl.getUrlBase() + roomId;
            if (Properties.Settings.Default.requireDisplayName)
            {
                link = ChangeSetting(link, Constants.JitsiConfig.RequireDisplayName);
            }
            if (Properties.Settings.Default.startWithAudioMuted)
            {
                link = ChangeSetting(link, Constants.JitsiConfig.AudioMuted);
            }
            if (Properties.Settings.Default.startWithVideoMuted)
            {
                link = ChangeSetting(link, Constants.JitsiConfig.VideoMuted);
            }

            endSel.InsertAfter("\n");
            endSel.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine);
            endSel.InsertAfter("\n");
            endSel.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine);
            endSel.InsertAfter("\n");
            endSel.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine);
            endSel.InsertAfter(Globals.ThisAddIn.getElementTranslation("appointmentItem", "textBodyMessage"));
            endSel.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine);
            var hyperlinkMeeting = wordDocument.Hyperlinks.Add(endSel.Range, link, ref missing, ref missing, link, ref missing);
            hyperlinkMeeting.Range.Font.Size = Constants.MainBodyTextSize;
            endSel.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine);
            endSel.InsertAfter("\n");
            endSel.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine);

            if (phoneNumbers.NumbersEnabled)
            {
                // Add Phone Number Text if they are enabled
                endSel.InsertAfter(Globals.ThisAddIn.getElementTranslation("appointmentItem", "textBodyMessagePhone"));
                endSel.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine);
                endSel.InsertAfter("\n");
                endSel.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine);
                foreach (var entry in phoneNumbers.Numbers)
                {
                    endSel.InsertAfter(entry.Key + ": ");
                    endSel.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine);
                    for (int i = 0; i < entry.Value.Count; i++)
                    {
                        var hyperlinkTel = wordDocument.Hyperlinks.Add(endSel.Range, "tel:" + entry.Value[i], ref missing, ref missing, entry.Value[i], ref missing);
                        hyperlinkTel.Range.Font.Size = Constants.MainBodyTextSize;

                        endSel.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine);
                        endSel.InsertAfter(" (");
                        endSel.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine);

                        var hyperlinkTelDirect = wordDocument.Hyperlinks.Add(endSel.Range, "tel:" + entry.Value[i] + ",,," + pinNumber + "%23", ref missing, ref missing, Globals.ThisAddIn.getElementTranslation("appointmentItem", "textBodyDirectCallString"), ref missing);
                        hyperlinkTelDirect.Range.Font.Size = Constants.MainBodyTextSize;
                        endSel.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine);
                        endSel.InsertAfter(")");
                        endSel.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine);

                        if (i < entry.Value.Count - 1)
                        {
                            endSel.InsertAfter(",");
                        }
                    }
                    endSel.InsertAfter("\n");
                    endSel.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine);
                }
                endSel.InsertAfter(Globals.ThisAddIn.getElementTranslation("appointmentItem", "textBodyPin") + pinNumber);
                endSel.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine);
            }
            endSel.InsertAfter("\n");
            endSel.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine);
            endSel.InsertAfter("\n");
            endSel.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine);

            endSel.Font.Size = Constants.DisclaimerTextSize;
            IEnumerable<KeyValuePair<bool, string>> disclaimer = Utils.SplitToTextAndHyperlinks(Globals.ThisAddIn.getElementTranslation("appointmentItem", "textBodyDisclaimer"));
            foreach (var textblock in disclaimer)
            {
                if (textblock.Key)
                {
                    // Textblock is a link
                    var hyperlink = wordDocument.Hyperlinks.Add(endSel.Range, textblock.Value, ref missing, ref missing, textblock.Value, ref missing);
                    hyperlink.Range.Font.Size = Constants.DisclaimerTextSize;
                    endSel.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine);
                }
                else
                {
                    // Textblock is no link
                    endSel.InsertAfter(textblock.Value);
                    endSel.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine);
                }
            }
            endSel.EndKey(Microsoft.Office.Interop.Word.WdUnits.wdLine);
            endSel.InsertAfter("\n");
            endSel.MoveDown(Microsoft.Office.Interop.Word.WdUnits.wdLine);

            wordDocument.Select();
            endSel.Collapse(Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart);
        }

        public static string ChangeSetting(string url, string setting)
        {
            var urlNew = string.Empty;
            if (Utils.SettingIsActive(url, setting))
            {
                urlNew = Regex.Replace(url, "(#|&)config\\." + setting + "=true", "");
            }
            else
            {
                // Otherwise add
                if (url.Contains("#config"))
                {
                    urlNew = url + "&config." + setting + "=true";
                }
                else
                {
                    urlNew = url + "#config." + setting + "=true";
                }
            }
            return urlNew;
        }


        public static void RunInThread(Action function)
        {
            try
            {
                ThreadStart s = new ThreadStart(() =>
                {
                    try
                    {
                        function();
                    }
                    catch (Exception ex)
                    {
                        HandleErrorWithUserNotification(ex);
                    }
                });
                Thread ss = new Thread(s);
                ss.Start();
            }
            catch (Exception ex)
            {
                HandleErrorWithUserNotification(ex);
            }
        }

        public static List<KeyValuePair<bool, string>> SplitToTextAndHyperlinks(string text)
        {
            var list = new List<KeyValuePair<bool, string>>();
            MatchCollection matches = Regex.Matches(text, "http[s]?:\\/\\/(?:[a-zA-Z]|[0-9]|[$-_@&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+(?<!\\.)");
            if (matches.Count == 0)
            {
                list.Add(new KeyValuePair<bool, string>(false, text));
            }
            var lastindex = 0;
            var index = 0;
            foreach (Match match in matches)
            {
                list.Add(new KeyValuePair<bool, string>(false, text.Substring(lastindex, match.Index - lastindex)));
                list.Add(new KeyValuePair<bool, string>(true, match.Value));
                lastindex = match.Index + match.Length;
                if (index == matches.Count - 1)
                {
                    list.Add(new KeyValuePair<bool, string>(false, text.Substring(lastindex, text.Length - lastindex)));
                }
                index++;
            }

            return list;
        }
    }
}
