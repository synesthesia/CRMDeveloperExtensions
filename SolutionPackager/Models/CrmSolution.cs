using System;

namespace SolutionPackager.Models
{
    public class CrmSolution
    {
        public Guid SolutionId { get; set; }
        public string Name { get; set; }
        public string Prefix { get; set; }
        public string UniqueName { get; set; }
        public Version Version { get; set; }
        public string BoundProject { get; set; }
        public bool DownloadManagedSolution { get; set; }
    }
}
