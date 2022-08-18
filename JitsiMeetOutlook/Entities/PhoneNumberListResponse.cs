using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace JitsiMeetOutlook.Entities
{
    /// <summary>
    /// PhoneNumberList Endpoint Response
    /// Derived from https://github.com/jitsi/jitsi-meet/blob/master/resources/cloud-api.swagger
    /// </summary>
    public class PhoneNumberListResponse : IValidatableObject
    {
        /// <summary>
        /// Message from the server.
        /// </summary>
        /// <value></value>
        public string Message { get; set; }

        /// <summary>
        /// Switch whether the numbers are enabled.
        /// </summary>
        /// <value></value>

        public bool NumbersEnabled { get; set; }

        /// <summary>
        /// Dictionary of Numbers.
        /// CountryCode - Number
        /// </summary>
        public Dictionary<string, List<string>> Numbers { get; set; }

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            var results = new List<ValidationResult>();
            if (Message == null)
            {
                results.Add(new ValidationResult("Message must be set"));
            };
            if (Numbers == null)
            {
                results.Add(new ValidationResult("Numbers must be set"));
            }
            return results;
        }
    }
}
