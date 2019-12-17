using System;
using System.Collections.Generic;
using System.Text;

namespace _12122019_GraphCalendarReader_Test.Models
{
    public class EventItem
    {
        public string Subject { get; set; }
        public string UserName { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public IEnumerable<string> Categories { get; set; }
    }
}
