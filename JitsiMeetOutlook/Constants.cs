using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JitsiMeetOutlook
{
    public static class Constants
    {
        public static readonly string MeetingLocationIdentifier = "Jitsi Meet";
        public static readonly string Font = "Calibri (Body)";
        public static readonly int MainBodyTextSize = 10;
        public static readonly int DisclaimerTextSize = 8;

        public static class JitsiConfig
        {
            public static readonly string AudioMuted = "startWithAudioMuted";
            public static readonly string VideoMuted = "startWithVideoMuted";
            public static readonly string RequireDisplayName = "requireDisplayName";
        }


    }
}
