﻿using System;
using Inedo.BuildMaster.Extensibility.Providers.IssueTracking;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace Inedo.BuildMasterExtensions.TFS2010
{
    [Serializable]
    internal sealed class Tfs2010Issue : IssueTrackerIssue
    {
        public static class DefaultStatusNames
        {
            public static string Active = "Active";
            public static string Resolved = "Resolved";
            public static string Closed = "Closed";
        }

        public Tfs2010Issue(WorkItem workItem, string customReleaseNumberFieldName)
            : base(workItem.Id.ToString(), workItem.State, workItem.Title, workItem.Description, GetReleaseNumber(workItem, customReleaseNumberFieldName))
        {
        }

        private static string GetReleaseNumber(WorkItem workItem, string customReleaseNumberFieldName)
        {
            return string.IsNullOrEmpty(customReleaseNumberFieldName)
                ? workItem.IterationPath.Substring(workItem.IterationPath.LastIndexOf('\\') + 1)
                : workItem.Fields[customReleaseNumberFieldName].Value.ToString().Trim();
        }
    }
}
