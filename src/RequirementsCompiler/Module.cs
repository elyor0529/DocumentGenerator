// -----------------------------------------------------------------------
// <copyright file="Module.cs" company="">
// TODO: Update copyright text.
// </copyright>
// -----------------------------------------------------------------------

namespace RequirementsCompiler
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    /// <summary>
    /// TODO: Update summary.
    /// </summary>
    public partial class Module
    {
        public int GetHoursFor(Priority priority)
        {
            return this.Features.First(fea => fea.Priority == priority).Remaining;
        }

        private Dictionary<Expertise, decimal> GetRelativeWorkloads()
        {
            var values = new Dictionary<Expertise, decimal>();

            foreach (var tech in this.Technologies)
            {
                if (values.ContainsKey(tech.Expertise))
                {
                    values[tech.Expertise] += this.GetWeight(tech.Workload, tech.Expertise);
                }
                else
                {
                    values.Add(tech.Expertise, this.GetWeight(tech.Workload, tech.Expertise));
                }
            }

            return values;
        }

        public decimal GetWeight(Workload workload, Expertise expertise)
        {
            var retValue = this.GetWorkloadValue(workload) * this.GetExpertiseWeight(expertise);
            return retValue;
        }

        public decimal GetHours(Priority priority, Expertise expertise)
        {
            if (this.PriorityHours().ContainsKey(priority))
            {
                var x = this.PriorityHours()[priority];

                if (this.GetRelativeWorkloads().ContainsKey(expertise))
                {
                     var weight = this.GetExpertiseSize(expertise);
                    return x * weight;
                }
                return 0;
            }
            return 0;
        }

        public IDictionary<Priority, int> PriorityHours()
        {
            IDictionary<Priority, int> values = new Dictionary<Priority, int>();

            foreach (var feature in this.Features)
            {
                if (values.ContainsKey(feature.Priority))
                {
                    values[feature.Priority] += feature.Remaining;
                }
                else
                {
                    values.Add(feature.Priority, feature.Remaining);
                }
            }

            return values;
        }

        public decimal GetWorkloadValue(Workload workload)
        {
            switch (workload)
            {
                case Workload.Heavy:
                    return 10;
                case Workload.Average:
                    return 4;
                case Workload.Light:
                    return 2;
                default:
                    return 1;
            }
        }

        public decimal GetExpertiseSize(Expertise expertise)
        {
            var x = this.GetRelativeWorkloads();

            decimal totalSize = 0;
            decimal totalSpecific = 0;
            foreach (var workload in x)
            {
                totalSize += workload.Value;
                if (workload.Key == expertise)
                {
                    totalSpecific = workload.Value;
                }
            }

            try
            {
                return totalSpecific / totalSize;
            }
            catch (Exception)
            {
                return 0;
            }            
        }

        public decimal GetExpertiseWeight(Expertise expertise)
        {
            switch (expertise)
            {
                case Expertise.Novice:
                    return 1;
                case Expertise.MidLevel:
                    return 2;
                case Expertise.Expert:
                    return 5;
                default:
                    return 0;
            }
        }

        public int OutstandingHours
        {
            get
            {
                return this.Features.Sum(x => x.Remaining);
            }
        }
    }
}
