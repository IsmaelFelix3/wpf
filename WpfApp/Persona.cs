using System;
using System.Collections.Generic;
using System.Text;
using Persona;

namespace Persona
{

    public class Persona
    {
        public Datum[] Data { get; set; }
        public long Total { get; set; }
        public long Page { get; set; }
        public long Limit { get; set; }
        public long Offset { get; set; }
    }

    public class Datum
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public Uri Picture { get; set; }
    }
}
