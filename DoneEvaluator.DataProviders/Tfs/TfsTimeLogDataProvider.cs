using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using PluggableService.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DoneEvaluator.DataProviders.Tfs
{
    public class TfsTimeLogDataProvider<T> : TimeLogDataProvider where T : TimeLog, new()
    {
        public string ConnectionString { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }

        public override List<TimeLog> LoadData(EvaluationServiceContext context)
        {
            string iterationPath = context.IterationPath;
            var wiStore = Connect(context).GetService<WorkItemStore>();

            //get linked work items
            //http://blogs.msdn.com/b/jsocha/archive/2012/02/22/retrieving-tfs-results-from-a-tree-query.aspx
            Query treeQuery = PrepareTreeQuery(wiStore, iterationPath, context.Project.Title);
            WorkItemLinkInfo[] links = treeQuery.RunLinkQuery();
            WorkItemCollection linkedResults = GetAssociatedWorkItems(wiStore, treeQuery, links);
            var linkedList = ConvertToTimeTrackingDetails(linkedResults);
            var relationMap = BuildRelationMap(linkedList, links);

            //get unlinked workitems
            WorkItemCollection results = GetAllWorkItems(wiStore, iterationPath, context.Project.Title);
            var list = ConvertToTimeTrackingDetails(results);
            var idList = relationMap.Select(q => q.WorkitemId).ToList();
            relationMap.AddRange(list.Where(p => !idList.Contains(p.WorkitemId)));

            PostProcessLoadedData(relationMap);

            return relationMap.Select(p => p as TimeLog).ToList();
        }

        private WorkItemCollection GetAllWorkItems(WorkItemStore wiStore, string iterationPath, string project)
        {
            var queryText = @"SELECT [System.Id], 
                                    [System.Title], 
                                    [Microsoft.VSTS.Common.BacklogPriority], 
                                    [System.AssignedTo], 
                                    [System.State], 
                                    [Microsoft.VSTS.Scheduling.RemainingWork], 
                                    [Microsoft.VSTS.CMMI.Blocked], 
                                    [System.WorkItemType] 
                                FROM WorkItems 
                                WHERE [System.TeamProject] = @project 
                                    AND [System.WorkItemType] IN ('Product Backlog Item', 'Bug')
                                    AND [System.State] <> 'Removed'         
                                    AND [System.IterationPath] Under @iterationPath
                                    ORDER BY [Microsoft.VSTS.Common.Priority], [System.Id] ";

            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("project", project);
            parameters.Add("iterationPath", iterationPath);

            var query = new Query(wiStore, queryText, parameters);

            return query.RunQuery();
        }

        private TfsTeamProjectCollection projectCollection;
        private TfsTeamProjectCollection Connect(ServiceContext context)
        {
            var token = new Microsoft.TeamFoundation.Client.SimpleWebTokenCredential(Username, Password);
            var clientCreds = new Microsoft.TeamFoundation.Client.TfsClientCredentials(token);
            projectCollection = new TfsTeamProjectCollection(new Uri(ConnectionString), clientCreds);
            projectCollection.EnsureAuthenticated();
            projectCollection.Connect(ConnectOptions.None);
            return projectCollection;
        }

        private List<T> BuildRelationMap(List<T> list, WorkItemLinkInfo[] links)
        {
            var pbiList = links.Where(p => p.SourceId == 0).Select(p => p.TargetId).ToList();
            var newList = list.Where(p => pbiList.Contains(p.WorkitemId)).ToList();
            foreach (var pbi in newList)
            {
                var taskList = links.Where(p => p.SourceId == pbi.WorkitemId).Select(p => p.TargetId);
                pbi.Tasks = list.Where(p => p.Type == "Task" && taskList.Contains(p.WorkitemId) && p.State != "Removed").Select(p => p as TimeLog).ToList();
            }

            return newList;
        }

        private static WorkItemCollection GetAssociatedWorkItems(WorkItemStore wiStore, Query treeQuery, WorkItemLinkInfo[] links)
        {
            int[] ids = links.Select(p => p.TargetId).Distinct().ToArray();

            var detailsWiql = new StringBuilder();
            detailsWiql.AppendLine("SELECT");
            bool first = true;
            foreach (FieldDefinition field in treeQuery.DisplayFieldList)
            {
                detailsWiql.Append("    ");
                if (!first)
                    detailsWiql.Append(",");
                detailsWiql.AppendLine("[" + field.ReferenceName + "]");
                first = false;
            }
            detailsWiql.AppendLine("FROM WorkItems");

            var flatQuery = new Query(wiStore, detailsWiql.ToString(), ids);
            WorkItemCollection results = flatQuery.RunQuery();
            return results;
        }

        private Query PrepareTreeQuery(WorkItemStore wiStore, string iterationPath, string project)
        {
            var queryText = @"SELECT [System.Id], 
                                    [System.Title], 
                                    [Microsoft.VSTS.Common.BacklogPriority], 
                                    [System.AssignedTo], 
                                    [System.State], 
                                    [Microsoft.VSTS.Scheduling.RemainingWork], 
                                    [Microsoft.VSTS.CMMI.Blocked], 
                                    [System.WorkItemType] 
                                FROM WorkItemLinks 
                                WHERE Source.[System.TeamProject] = @project 
                                    AND Source.[System.WorkItemType] IN ('Product Backlog Item', 'Bug')
                                    AND Source.[System.State] <> 'Removed'         
                                    AND Source.[System.IterationPath] Under @iterationPath
                                    ORDER BY [Microsoft.VSTS.Common.Priority], [System.Id] ";

            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add("project", project);
            parameters.Add("iterationPath", iterationPath);

            return new Query(wiStore, queryText, parameters);
        }

        protected const string WorkItemTypeField = "Work Item Type";
        protected const string IterationPathField = "Iteration Path";

        protected const string RemainWorkField = "Remaining Work";
        protected const string AssignedToField = "Assigned To";
        protected const string StateField = "State";
        protected const string ActivityTypeField = "Activity";

        private List<T> ConvertToTimeTrackingDetails(WorkItemCollection results)
        {
            List<T> workItems = new List<T>();
            if (results != null)
            {
                foreach (WorkItem wi in results)
                {
                    var workitem = new T();

                    workitem.WorkitemId = wi.Id;
                    workitem.Title = wi.Title;
                    workitem.Type = wi.Fields[WorkItemTypeField].Value.ToString();
                    workitem.IterationPath = wi.Fields[IterationPathField].Value.ToString();
                    workitem.AssignedTo = wi.Fields[AssignedToField].Value.ToString();
                    workitem.State = wi.Fields[StateField].Value.ToString();

                    workitem.Activity = string.Compare(wi.Fields[WorkItemTypeField].Value.ToString(), "Task", true) == 0
                            ? wi.Fields[ActivityTypeField].Value.ToString()
                            : string.Empty;

                    workitem.IsTaskMarkedAsDone = string.Compare(wi.Fields[WorkItemTypeField].Value.ToString(), "Task", true) == 0
                                            && string.Compare(wi.Fields[StateField].Value.ToString(), "Done", true) == 0;

                    workitem.TrackingDate = DateTime.Today.Date;

                    if (wi.Fields.Contains(RemainWorkField) && wi.Fields[RemainWorkField].Value != null)
                    {
                        workitem.RemainingWork = float.Parse(wi.Fields[RemainWorkField].Value.ToString());
                    }

                    PopulateCustomFields(wi, workitem);

                    workItems.Add(workitem);
                }
            }
            return workItems;
        }


        public virtual void PopulateCustomFields(WorkItem wi, T workitem) { }
        protected virtual void PostProcessLoadedData(List<T> relationMap) { }
    }

}
