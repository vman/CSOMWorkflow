using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Workflow;
using Microsoft.SharePoint.Client.WorkflowServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;

namespace CSOMWorkflow
{
    class Program
    {
        static void Main(string[] args)
        {
            //Replace the following values according to your configuration.

            //Site Details
            string siteUrl = "https://yoursite.sharepoint.com/sites/test";
            string userName = "login@yoursite.onmicrosoft.com";
            string password = "password";

            //Name of the SharePoint2013 List Workflow
            string workflowName = "SP2010DocWF";

            //Name of the List to which the Workflow is Associated
            string targetListName = "Documents";
            
            //Guid of the List to which the Workflow is Associated
            Guid targetListGUID = new Guid("b89e266c-5f20-4b83-9f95-10a42c629e84");
            
            //Guid of the ListItem on which to start the Workflow
            Guid targetItemId = new Guid("6ab8227b-66a3-4e12-8055-846dcfb53cab");

            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                SecureString securePassword = new SecureString();

                foreach (char c in password.ToCharArray()) securePassword.AppendChar(c);

                clientContext.Credentials = new SharePointOnlineCredentials(userName, securePassword);

                Web web = clientContext.Web;

                //Workflow Services Manager which will handle all the workflow interaction.
                WorkflowServicesManager wfServicesManager = new WorkflowServicesManager(clientContext, web);

                ////Deployment Service which holds all the Workflow Definitions deployed to the site
                //WorkflowDeploymentService wfDeploymentService = wfServicesManager.GetWorkflowDeploymentService();

                ////Get all the definitions from the Deployment Service, or get a specific definition using the GetDefinition method.
                //WorkflowDefinitionCollection wfDefinitions = wfDeploymentService.EnumerateDefinitions(false);

                //clientContext.Load(wfDefinitions, wfDefs => wfDefs.Where(wfd => wfd.DisplayName == workflowName));

                //clientContext.ExecuteQuery();

                //WorkflowDefinition wfDefinition = wfDefinitions.First();

                ////The Subscription service is used to get all the Associations currently on the SPSite
                //WorkflowSubscriptionService wfSubscriptionService = wfServicesManager.GetWorkflowSubscriptionService();

                ////The subscription (association)
                //WorkflowSubscription wfSubscription = new WorkflowSubscription(clientContext);
                //wfSubscription.DefinitionId = wfDefinition.Id;
                //wfSubscription.Enabled = true;
                //wfSubscription.Name = newSubscriptionName;

                //var startupOptions = new List<string>();
                
                //// manual start
                //startupOptions.Add("WorkflowStart");

                //// set the workflow start settings
                //wfSubscription.EventTypes = startupOptions;

                //// set the associated task and history lists
                //wfSubscription.SetProperty("HistoryListId", workflowHistoryListID);
                
                //wfSubscription.SetProperty("TaskListId", taskListID);

                //wfSubscriptionService.PublishSubscription(wfSubscription);

                //clientContext.ExecuteQuery();

                //WorkflowInstanceService wfInstanceService = manager.GetWorkflowInstanceService();

                //wfInstanceService.StartWorkflow(wfSubscription, null);

                //clientContext.ExecuteQuery();

                //WorkflowSubscriptionCollection wfSubscriptions = wfSubscriptionService.EnumerateSubscriptionsByDefinition(wfDefinition.Id);

                //clientContext.Load(wfSubscriptions);

                //clientContext.ExecuteQuery();

                //WorkflowSubscription wfSubscription = wfSubscriptions.First();

                ////The Subscription service is used to get all the Associations currently on the SPSite
                //WorkflowSubscriptionService wfSubscriptionService = wfServicesManager.GetWorkflowSubscriptionService();

                ////All the subscriptions (associations)
                //WorkflowSubscriptionCollection wfSubscriptions = wfSubscriptionService.EnumerateSubscriptions();

                ////Load only the subscription (association) which we want. You can also get a subscription by definition id.
                //clientContext.Load(wfSubscriptions, wfSubs => wfSubs.Where(wfSub => wfSub.Name == workflowName));

                //clientContext.ExecuteQuery();

                ////Get the subscription.
                //WorkflowSubscription wfSubscription = wfSubscriptions.First();

                ////The Instance Service is used to start workflows and create instances.
                //WorkflowInstanceService wfInstanceService = wfServicesManager.GetWorkflowInstanceService();

                ////Any custom parameters you want to send to the workflow.
                //var initiationData = new Dictionary<string, object>();

                //wfInstanceService.StartWorkflowOnListItem(wfSubscription, itemID, initiationData);

                //clientContext.ExecuteQuery();

                #region Interop
                //Will return all Workflow Associations which are running on the SharePoint 2010 Engine
                WorkflowAssociationCollection wfAssociations = web.Lists.GetByTitle(targetListName).WorkflowAssociations;

                //Get the required Workflow Association
                WorkflowAssociation wfAssociation = wfAssociations.GetByName(workflowName);

                clientContext.Load(wfAssociation);

                clientContext.ExecuteQuery();

                //Get the instance of the Interop Service which will be used to create an instance of the Workflow
                InteropService workflowInteropService = wfServicesManager.GetWorkflowInteropService();

                var initiationData = new Dictionary<string, object>();

                //Start the Workflow
                ClientResult<Guid> resultGuid = workflowInteropService.StartWorkflow(wfAssociation.Name, new Guid(), targetListGUID , targetItemId, initiationData);

                clientContext.ExecuteQuery(); 
                #endregion
            }
        }
    }
}
