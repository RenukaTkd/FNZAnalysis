using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace FNZAnalysis
{
    public class UpdateNotesCount : IPlugin
    {
        public object RetrrieveMultipleRequest { get; private set; }

        public void Execute(IServiceProvider serviceProvider)
    {
        //Extract the tracing service for use in debugging sandboxed plug-ins.
        ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

        // Obtain the execution context from the service provider.
        IPluginExecutionContext context = (IPluginExecutionContext)
            serviceProvider.GetService(typeof(IPluginExecutionContext));

        // The InputParameters collection contains all the data passed in the message request.
        if (context.InputParameters.Contains("Target"))
        {
                if (context.MessageName == "Delete")
                {
                    IOrganizationServiceFactory serviceFactory =
                      (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    // Obtain the target entity from the input parameters.
                 


                    if (context.PrimaryEntityName != "annotation")  //Need to replace with your entity
                        return;

                    try
                    {
                        
                        if (context.PreEntityImages != null && context.PreEntityImages.Contains("PreImage"))
                        {
                            Entity BeforeDeleteEntity = (Entity)context.PreEntityImages["PreImage"]; // Get Pre Image of the entity

                            if (BeforeDeleteEntity.Attributes.ContainsKey("objecttypecode") && BeforeDeleteEntity.Attributes["objecttypecode"].ToString() == "opportunity")
                            {
                                var oppId = ((EntityReference)BeforeDeleteEntity.Attributes["objectid"]).Id;

                                QueryExpression query = new QueryExpression();
                                query.EntityName = "annotation";
                                query.ColumnSet = new ColumnSet(true);
                                query.Criteria = new FilterExpression();
                                query.Criteria.AddCondition(new ConditionExpression("objectid", ConditionOperator.Equal, oppId));
                                query.Criteria.AddCondition(new ConditionExpression("annotationid", ConditionOperator.NotEqual, BeforeDeleteEntity.Id));


                                RetrieveMultipleRequest request = new RetrieveMultipleRequest();
                                request.Query = query;

                                IEnumerable<Entity> results = ((RetrieveMultipleResponse)service.Execute(request)).EntityCollection.Entities;

                                if (results.Any())
                                {
                                    var notesCount = results.Count();

                                    Entity updateOpportunity = new Entity("opportunity", oppId);

                                    updateOpportunity.Attributes["opportunityid"] = oppId;
                                    updateOpportunity.Attributes["red_numberofnotes"] = notesCount;
                                    service.Update(updateOpportunity);
                                }
                                else
                                {
                                    Entity updateOpportunity = new Entity("opportunity", oppId);

                                    updateOpportunity.Attributes["opportunityid"] = oppId;
                                    updateOpportunity.Attributes["red_numberofnotes"] = 0;
                                    service.Update(updateOpportunity);
                                }
                            }

                                if (BeforeDeleteEntity.Attributes.ContainsKey("objecttypecode") && BeforeDeleteEntity.Attributes["objecttypecode"].ToString() == "lead")
                                {
                                    var leadId = ((EntityReference)BeforeDeleteEntity.Attributes["objectid"]).Id;

                                    QueryExpression query1 = new QueryExpression();
                                    query1.EntityName = "annotation";
                                    query1.ColumnSet = new ColumnSet(true);
                                    query1.Criteria = new FilterExpression();
                                    query1.Criteria.AddCondition(new ConditionExpression("objectid", ConditionOperator.Equal, leadId));
                                    query1.Criteria.AddCondition(new ConditionExpression("annotationid", ConditionOperator.NotEqual, BeforeDeleteEntity.Id));


                                    RetrieveMultipleRequest request1 = new RetrieveMultipleRequest();
                                    request1.Query = query1;

                                    IEnumerable<Entity> results1 = ((RetrieveMultipleResponse)service.Execute(request1)).EntityCollection.Entities;

                                    if (results1.Any())
                                    {
                                        var notesCount1 = results1.Count();

                                        Entity updatelead = new Entity("lead", leadId);

                                        updatelead.Attributes["leadid"] = leadId;
                                        updatelead.Attributes["red_numberofnotes"] = notesCount1;
                                        service.Update(updatelead);
                                    }
                                    else
                                    {
                                        Entity updatelead = new Entity("lead", leadId);

                                        updatelead.Attributes["leadid"] = leadId;
                                        updatelead.Attributes["red_numberofnotes"] = 0;
                                        service.Update(updatelead);
                                    }

                                }
                                if (BeforeDeleteEntity.Attributes.ContainsKey("objecttypecode") && BeforeDeleteEntity.Attributes["objecttypecode"].ToString() == "account")
                                {
                                    var accountId = ((EntityReference)BeforeDeleteEntity.Attributes["objectid"]).Id;

                                    QueryExpression query2 = new QueryExpression();
                                    query2.EntityName = "annotation";
                                    query2.ColumnSet = new ColumnSet(true);
                                    query2.Criteria = new FilterExpression();
                                    query2.Criteria.AddCondition(new ConditionExpression("objectid", ConditionOperator.Equal, accountId));
                                    query2.Criteria.AddCondition(new ConditionExpression("annotationid", ConditionOperator.NotEqual, BeforeDeleteEntity.Id));


                                    RetrieveMultipleRequest request2 = new RetrieveMultipleRequest();
                                    request2.Query = query2;

                                    IEnumerable<Entity> results2 = ((RetrieveMultipleResponse)service.Execute(request2)).EntityCollection.Entities;

                                    if (results2.Any())
                                    {
                                        var notesCount2 = results2.Count();

                                        Entity updateaccount = new Entity("account", accountId);

                                        updateaccount.Attributes["accountid"] = accountId;
                                        updateaccount.Attributes["red_numberofnotes"] = notesCount2;
                                        service.Update(updateaccount);
                                    }
                                    else
                                    {
                                        Entity updateaccount = new Entity("account", accountId);

                                        updateaccount.Attributes["accountid"] = accountId;
                                        updateaccount.Attributes["red_numberofnotes"] = 0;
                                        service.Update(updateaccount);
                                    }

                                }

                            }
                        

                    }
                    catch (FaultException<OrganizationServiceFault> ex)
                    {
                        throw new InvalidPluginExecutionException("An error occurred in the FollowupPlugin plug-in.", ex);
                    }

                    catch (Exception ex)
                    {
                        tracingService.Trace("FollowupPlugin: {0}", ex.ToString());
                        throw;
                    }
                }
        }
    }
}
}