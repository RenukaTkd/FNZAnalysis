using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.Runtime.Serialization.Json;
using System.IO;
using System.Web.Script.Serialization;
using FNZAnalysis.Model;
using Microsoft.Xrm.Sdk.Query;

namespace FNZAnalysis
{
    public class OpportunityAnalysis : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.  
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.  
                Entity entity = (Entity)context.InputParameters["Target"];

                // Obtain the organization service reference which you will need for  
                // web service calls.  
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                
                Dictionary<string, OpportunityChanges> ocold = new Dictionary<string, OpportunityChanges>();

                try
                {
                    tracingService.Trace("Entered Plugin.... OpportunityAnalysis");
                    //get the current record guid from context
                    Guid opportunityid = entity.Id;

                    string trackedOpportunityFields = null;
                    string updatedFields = null;


                    //Define variables to store Preimage and Postimage
                    string pre_red__asset_value = null;
                    object AssestValuePre;


                    string pre_red_stage = null;

                    string pre_red_sltmember = null;
                    string postSltmember = null;

                    string pre_red_ifyes = null;

                    string pre_name = null;

                    string pre_red_opptype = null;

                    string pre_estimatedclosedate = null;

                    string pre_description = null;

                    string pre_red_multisolution = null;

                    string pre_red_probability = null;

                    string pre_red_regulatorypermissions = null;

                    string pre_red__basepoint = null;

                    string pre_red_basispointrange = null;

                    string pre_red_includeintherevenueforecast = null;

                    string pre_opportunityratingcode = null;
                    string preopportunitycode = null;

                    string pre_StatusCode = null;

                    string pre_statusreason = null;

                    string pre_exchangerate = null;

                    EntityReference pre_parentaccountid = null;

                    EntityReference pre_parentcontactid = null;

                    EntityReference pre_modifiedby = null;

                    EntityReference pre_transactioncurrencyid = null;

                    EntityReference pre_red_territory = null;

                    EntityReference pre_red_division = null;

                    EntityReference pre_red_regions = null;

                    EntityReference pre_originatingleadid = null;

                    EntityReference pre_red_primarysolution = null;

                    string pre_red__consultancy_fee = null;
                    Object ConsultancyValuePre;

                    string pre_red__consultancy_fee_base = null;
                    Object ConsultancybaseValuePre;

                    string pre_red_implamentationfee = null;
                    Object implementationValuePre;
                    string pre_red_implamentationfee_base = null;
                    Object implementationbaseValuePre;

                    string pre_red__bps_cal = null;
                    Object bpsValuePre;

                    string pre_red_deliveryfees = null;
                    Object deliveryFeeValuePre;

                    string pre_red__asset_value_Base = null;
                    Object AssestValueBasePre;

                    string pre_red__bps_cal_base = null;
                    Object bpsValueBasePre;

                    string pre_red_deliveryfees_Base = null;
                    Object DeliveryFeePre;


                    // Post Image
                    string post_red__Asset_value = null;
                    object AssestValuePost;

                    string post_name = null;

                    string post_uniqueid = null;

                    string post_red_stage = null;

                    string post_red_sltmember = null;

                    string post_red_ifyes = null;

                    string post_red_opptype = null;

                    string post_estimatedclosedate = null;
                    string Post_modifiedon = null;

                    string post_description = null;

                    string post_opportunityratingcode = null;
                    string postopportunitycode = null;

                    string post_StatusCode = null;

                    string post_statusreason = null;


                    string post_red_multisolution = null;

                    string post_red_probability = null;

                    string post_red_regulatorypermissions = null;

                    string post_red__basepoint = null;

                    string post_red_basispointrange = null;

                    string post_red_includeintherevenueforecast = null;

                    string post_exchangerate = null;

                    EntityReference post_parentaccountid = null;

                    EntityReference post_parentcontactid = null;

                   EntityReference post_modifiedby = null;

                   EntityReference post_dealowner = null;

                    EntityReference post_transactioncurrencyid = null;

                    EntityReference post_red_territory = null;

                    EntityReference post_red_division = null;

                    EntityReference post_red_regions = null;

                    EntityReference post_originatingleadid = null;

                    EntityReference post_red_primarysolution = null;

                    string post_red__consultancy_fee = null;
                    Object ConsultancyValuePost;


                    string post_red__consultancy_fee_base = null;
                    Object ConsultancybaseValuePost;

                    string post_red_implamentationfee = null;
                    Object implementationValuePost;
                    string post_red_implamentationfee_base = null;
                    Object implementationbaseValuePost;

                    string post_red__bps_cal = null;
                    Object bpsValuePost;

                    string post_red_deliveryfees = null;
                    Object deliveryFeeValuePost;

                    string post_red__asset_value_Base = null;
                    Object AssestValueBasePost;

                    string post_red__bps_cal_base = null;
                    Object bpsValueBasePost;

                    string post_red_deliveryfees_Base = null;
                    Object DeliveryFeePost;

                    string postStageValue = null;
                    string postProbabilityValue = null;
                    string postincludeinrevenue = null;

                    string preStageValue = null;
                    string preProbabilityValue = null;
                    string preincludeinrevenue = null;
                    // Updated values collection

                    Dictionary<string, string> changeValues = new Dictionary<string, string>();
                    //Dictionary<string, string> unchangeValues = new Dictionary<string, string>();
                    OpportunityCounting oc = new OpportunityCounting();

                    // In below leadimage has been added in plugin registration tool
                    //get PreImage from Context
                    
                    
                    if (context.PreEntityImages.Contains("OpportunityAnalysisImage") && context.PreEntityImages["OpportunityAnalysisImage"] is Entity)
                    {
                        Entity preMessageImage = (Entity)context.PreEntityImages["OpportunityAnalysisImage"];
                        //get topic field value beforedatabase update perform 

                        //opportunity details
                        if (preMessageImage.Attributes.Contains("red__asset_value"))
                        {

                            if (preMessageImage.Attributes.TryGetValue("red__asset_value", out AssestValuePre))
                                pre_red__asset_value = ((Money)AssestValuePre).Value.ToString();
                        }

                        if (preMessageImage.Attributes.Contains("red__asset_value_Base"))
                        {

                            if (preMessageImage.Attributes.TryGetValue("red__asset_value_Base", out AssestValueBasePre))
                                pre_red__asset_value_Base = ((Money)AssestValueBasePre).Value.ToString();
                        }
                        if (preMessageImage.Attributes.Contains("red_deliveryfees"))
                        {

                            if (preMessageImage.Attributes.TryGetValue("red_deliveryfees", out deliveryFeeValuePre))
                                pre_red_deliveryfees = ((Money)deliveryFeeValuePre).Value.ToString();
                        }

                        if (preMessageImage.Attributes.Contains("red_deliveryfees_Base"))
                        {

                            if (preMessageImage.Attributes.TryGetValue("red_deliveryfees_Base", out DeliveryFeePre))
                                pre_red_deliveryfees = ((Money)DeliveryFeePre).Value.ToString();
                        }

                        if (preMessageImage.Attributes.Contains("red__consultancy_fee"))
                        {

                            if (preMessageImage.Attributes.TryGetValue("red__consultancy_fee", out ConsultancyValuePre))
                                pre_red__consultancy_fee = ((Money)ConsultancyValuePre).Value.ToString();
                        }

                        if (preMessageImage.Attributes.Contains("red__consultancy_fee_base"))
                        {

                            if (preMessageImage.Attributes.TryGetValue("red__consultancy_fee_base", out ConsultancybaseValuePre))
                                pre_red__consultancy_fee_base = ((Money)ConsultancybaseValuePre).Value.ToString();
                        }



                        if (preMessageImage.Attributes.Contains("red_implamentationfee"))
                        {

                            if (preMessageImage.Attributes.TryGetValue("red_implamentationfee", out implementationValuePre))
                                pre_red_implamentationfee = ((Money)implementationValuePre).Value.ToString();
                        }
                        if (preMessageImage.Attributes.Contains("red_implamentationfee_base"))
                        {

                            if (preMessageImage.Attributes.TryGetValue("red_implamentationfee_base", out implementationbaseValuePre))
                                pre_red_implamentationfee_base = ((Money)implementationbaseValuePre).Value.ToString();
                        }

                        if (preMessageImage.Attributes.Contains("red__bps_cal"))
                        {

                            if (preMessageImage.Attributes.TryGetValue("red__bps_cal", out bpsValuePre))
                                pre_red__bps_cal = ((Money)bpsValuePre).Value.ToString();
                        }


                        if (preMessageImage.Attributes.Contains("red__basepoint"))
                            pre_red__basepoint = ((decimal)preMessageImage.Attributes["red__basepoint"]).ToString();

                        if (preMessageImage.Attributes.Contains("exchangerate"))
                            pre_exchangerate = ((decimal)preMessageImage.Attributes["exchangerate"]).ToString();

                        if (preMessageImage.Attributes.Contains("red_opptype"))
                            pre_red_opptype = ((OptionSetValue)preMessageImage.Attributes["red_opptype"]).Value.ToString();

                        if (preMessageImage.Attributes.Contains("opportunityratingcode"))
                        {
                            pre_opportunityratingcode = ((OptionSetValue)preMessageImage.Attributes["opportunityratingcode"]).Value.ToString();
                            preopportunitycode = preMessageImage.FormattedValues.Where(x => x.Key == "opportunityratingcode").Select(y => y.Value).First().ToString();
                        }

                        if (preMessageImage.Attributes.Contains("statecode"))
                            pre_StatusCode = ((OptionSetValue)preMessageImage.Attributes["statecode"]).Value.ToString();

                        if (preMessageImage.Attributes.Contains("statuscode"))
                            pre_statusreason = ((OptionSetValue)preMessageImage.Attributes["statuscode"]).Value.ToString();

                        if (preMessageImage.Attributes.Contains("red_probability"))
                        {
                            pre_red_probability = ((OptionSetValue)preMessageImage.Attributes["red_probability"]).Value.ToString();
                            preProbabilityValue = preMessageImage.FormattedValues.Where(x => x.Key == "red_probability").Select(y => y.Value).First().ToString();
                        }

                        if (preMessageImage.Attributes.Contains("red_regulatorypermissions"))
                            pre_red_regulatorypermissions = ((OptionSetValue)preMessageImage.Attributes["red_regulatorypermissions"]).Value.ToString();

                        if (preMessageImage.Attributes.Contains("red_stage"))
                        {
                            pre_red_stage = ((OptionSetValue)preMessageImage.Attributes["red_stage"]).Value.ToString();
                            preStageValue = preMessageImage.FormattedValues.Where(x => x.Key == "red_stage").Select(y => y.Value).First().ToString();
                        }
                        if (preMessageImage.Attributes.Contains("red_sltmember"))
                            pre_red_sltmember = ((OptionSetValue)preMessageImage.Attributes["red_sltmember"]).Value.ToString();


                        // if (preMessageImage.Attributes.Contains("red_multisolution"))
                        //     pre_red_multisolution = preMessageImage.GetAttributeValue < OptionSetValueCollection > ("red_multisolution").ToString();

                        if (preMessageImage.Attributes.Contains("red_includeintherevenueforecast"))
                        {
                            pre_red_includeintherevenueforecast = ((OptionSetValue)preMessageImage.Attributes["red_includeintherevenueforecast"]).Value.ToString();
                            preincludeinrevenue = preMessageImage.FormattedValues.Where(x => x.Key == "red_includeintherevenueforecast").Select(y => y.Value).First().ToString();
                        }

                        if (preMessageImage.Attributes.Contains("parentaccountid"))
                            pre_parentaccountid = (EntityReference)preMessageImage.Attributes["parentaccountid"];

                        if (preMessageImage.Attributes.Contains("parentcontactid"))
                            pre_parentcontactid = (EntityReference)preMessageImage.Attributes["parentcontactid"];

                        if (preMessageImage.Attributes.Contains("red_territory"))
                            pre_red_territory = (EntityReference)preMessageImage.Attributes["red_territory"];
                        pre_red_territory = (EntityReference)preMessageImage.Attributes["red_territory"];

                        if (preMessageImage.Attributes.Contains("modifiedby"))
                        pre_modifiedby = (EntityReference)preMessageImage.Attributes["modifiedby"];

                        if (preMessageImage.Attributes.Contains("red_division"))
                            pre_red_division = (EntityReference)preMessageImage.Attributes["red_division"];

                        if (preMessageImage.Attributes.Contains("red_regions"))
                            pre_red_regions = (EntityReference)preMessageImage.Attributes["red_regions"];

                        if (preMessageImage.Attributes.Contains("originatingleadid"))
                            pre_originatingleadid = (EntityReference)preMessageImage.Attributes["originatingleadid"];

                        if (preMessageImage.Attributes.Contains("red_primarysolution"))
                            pre_red_primarysolution = (EntityReference)preMessageImage.Attributes["red_primarysolution"];


                        if (preMessageImage.Attributes.Contains("name"))
                            pre_name = (String)preMessageImage.Attributes["name"];

                        if (preMessageImage.Attributes.Contains("red_ifyes"))
                            pre_red_ifyes = (String)preMessageImage.Attributes["red_ifyes"];

                        if (preMessageImage.Attributes.Contains("description"))
                            pre_description = (String)preMessageImage.Attributes["description"];

                        if (preMessageImage.Attributes.Contains("red_basispointrange"))
                            pre_red_basispointrange = (String)preMessageImage.Attributes["red_basispointrange"];

                        if (preMessageImage.Attributes.Contains("estimatedclosedate"))
                            pre_estimatedclosedate = preMessageImage.Attributes["estimatedclosedate"].ToString();

                    }
                    // pre image

                    // post image

                    if (context.PostEntityImages.Contains("OpportunityAnalysisImage") && context.PostEntityImages["OpportunityAnalysisImage"] is Entity)
                    {
                        Entity postMessageImage = (Entity)context.PostEntityImages["OpportunityAnalysisImage"];
                        //get topic field value beforedatabase update perform 

                        //opportunity details
                        if (postMessageImage.Attributes.Contains("red__asset_value"))
                        {

                            if (postMessageImage.Attributes.TryGetValue("red__asset_value", out AssestValuePost))
                                post_red__Asset_value = ((Money)AssestValuePost).Value.ToString();
                        }
                        if (postMessageImage.Attributes.Contains("red__asset_value_Base"))
                        {

                            if (postMessageImage.Attributes.TryGetValue("red__asset_value_Base", out AssestValueBasePost))
                                post_red__asset_value_Base = ((Money)AssestValueBasePost).Value.ToString();
                        }
                        //if (postMessageImage.Attributes.Contains("red_deliveryfees"))
                        //{

                        //    if (postMessageImage.Attributes.TryGetValue("red_deliveryfees", out deliveryFeeValuePost))
                        //        post_red_deliveryfees = ((Money)deliveryFeeValuePost).Value.ToString();
                        //}

                        if (postMessageImage.Attributes.Contains("red_deliveryfees_Base"))
                        {

                            if (postMessageImage.Attributes.TryGetValue("red_deliveryfees_Base", out DeliveryFeePost))
                                post_red_deliveryfees = ((Money)DeliveryFeePost).Value.ToString();
                        }
                        if (postMessageImage.Attributes.Contains("red__consultancy_fee"))
                        {

                            if (postMessageImage.Attributes.TryGetValue("red__consultancy_fee", out ConsultancyValuePost))
                                post_red__consultancy_fee = ((Money)ConsultancyValuePost).Value.ToString();
                        }

                        if (postMessageImage.Attributes.Contains("red__consultancy_fee_base"))
                        {

                            if (postMessageImage.Attributes.TryGetValue("red__consultancy_fee_base", out ConsultancybaseValuePost))
                                post_red__consultancy_fee_base = ((Money)ConsultancybaseValuePost).Value.ToString();
                        }
                        if (postMessageImage.Attributes.Contains("red_implamentationfee"))
                        {

                            if (postMessageImage.Attributes.TryGetValue("red_implamentationfee", out implementationValuePost))
                                post_red_implamentationfee = ((Money)implementationValuePost).Value.ToString();

                        }
                        if (postMessageImage.Attributes.Contains("red_implamentationfee_base"))
                        {

                            if (postMessageImage.Attributes.TryGetValue("red_implamentationfee_base", out implementationbaseValuePost))
                                post_red_implamentationfee_base = ((Money)implementationbaseValuePost).Value.ToString();

                        }


                        if (postMessageImage.Attributes.Contains("red__bps_cal"))
                        {

                            if (postMessageImage.Attributes.TryGetValue("red__bps_cal", out bpsValuePost))
                                post_red__bps_cal = ((Money)bpsValuePost).Value.ToString();
                        }



                        if (postMessageImage.Attributes.Contains("red__basepoint"))
                            post_red__basepoint = ((decimal)postMessageImage.Attributes["red__basepoint"]).ToString();

                        if (postMessageImage.Attributes.Contains("exchangerate"))
                            post_exchangerate = ((decimal)postMessageImage.Attributes["exchangerate"]).ToString();

                        if (postMessageImage.Attributes.Contains("red_opptype"))
                            post_red_opptype = ((OptionSetValue)postMessageImage.Attributes["red_opptype"]).Value.ToString();

                        if (postMessageImage.Attributes.Contains("red_probability"))
                        {
                            post_red_probability = ((OptionSetValue)postMessageImage.Attributes["red_probability"]).Value.ToString();

                            postProbabilityValue = postMessageImage.FormattedValues.Where(x => x.Key == "red_probability").Select(y => y.Value).First().ToString();
                        }
                        if (postMessageImage.Attributes.Contains("red_regulatorypermissions"))
                            post_red_regulatorypermissions = ((OptionSetValue)postMessageImage.Attributes["red_regulatorypermissions"]).Value.ToString();

                        if (postMessageImage.Attributes.Contains("red_stage"))
                        {

                            post_red_stage = ((OptionSetValue)postMessageImage.Attributes["red_stage"]).Value.ToString();
                            postStageValue = postMessageImage.FormattedValues.Where(x => x.Key == "red_stage").Select(y => y.Value).First().ToString();
                        }

                        if (postMessageImage.Attributes.Contains("opportunityratingcode"))
                        {
                            post_opportunityratingcode = ((OptionSetValue)postMessageImage.Attributes["opportunityratingcode"]).Value.ToString();
                            postopportunitycode = postMessageImage.FormattedValues.Where(x => x.Key == "opportunityratingcode").Select(y => y.Value).First().ToString();
                        }

                        if (postMessageImage.Attributes.Contains("statecode"))
                            post_StatusCode = ((OptionSetValue)postMessageImage.Attributes["statecode"]).Value.ToString();

                        if (postMessageImage.Attributes.Contains("statuscode"))
                            post_statusreason = ((OptionSetValue)postMessageImage.Attributes["statuscode"]).Value.ToString();

                        if (postMessageImage.Attributes.Contains("red_sltmember"))
                            post_red_sltmember = ((OptionSetValue)postMessageImage.Attributes["red_sltmember"]).Value.ToString();
                        postSltmember = postMessageImage.FormattedValues.Where(x => x.Key == "red_sltmember").Select(y => y.Value).First().ToString();

                        //if (postMessageImage.Attributes.Contains("red_multisolution"))
                        //     post_red_multisolution = ((OptionSetValue)postMessageImage.Attributes["red_multisolution"]).Value.ToString();

                        if (postMessageImage.Attributes.Contains("red_includeintherevenueforecast"))
                        {
                            post_red_includeintherevenueforecast = ((OptionSetValue)postMessageImage.Attributes["red_includeintherevenueforecast"]).Value.ToString();
                            postincludeinrevenue = postMessageImage.FormattedValues.Where(x => x.Key == "red_includeintherevenueforecast").Select(y => y.Value).First().ToString();
                        }

                        if (postMessageImage.Attributes.Contains("parentaccountid"))
                            post_parentaccountid = (EntityReference)postMessageImage.Attributes["parentaccountid"];

                        if (postMessageImage.Attributes.Contains("parentcontactid"))
                            post_parentcontactid = (EntityReference)postMessageImage.Attributes["parentcontactid"];

                        if (postMessageImage.Attributes.Contains("red_territory"))
                            post_red_territory = (EntityReference)postMessageImage.Attributes["red_territory"];

                        if (postMessageImage.Attributes.Contains("red_division"))
                            post_red_division = (EntityReference)postMessageImage.Attributes["red_division"];

                        if (postMessageImage.Attributes.Contains("modifiedby"))
                          post_modifiedby = (EntityReference)postMessageImage.Attributes["modifiedby"];


                        if (postMessageImage.Attributes.Contains("ownerid"))
                           post_dealowner = (EntityReference)postMessageImage.Attributes["ownerid"];

                        if (postMessageImage.Attributes.Contains("red_regions"))
                            post_red_regions = (EntityReference)postMessageImage.Attributes["red_regions"];

                        if (postMessageImage.Attributes.Contains("originatingleadid"))
                            post_originatingleadid = (EntityReference)postMessageImage.Attributes["originatingleadid"];

                        if (postMessageImage.Attributes.Contains("red_primarysolution"))
                            post_red_primarysolution = (EntityReference)postMessageImage.Attributes["red_primarysolution"];

                        if (postMessageImage.Attributes.Contains("name"))
                            post_name = (String)postMessageImage.Attributes["name"];

                        if (postMessageImage.Attributes.Contains("red_opportunitycode"))
                            post_uniqueid = (String)postMessageImage.Attributes["red_opportunitycode"];

                        if (postMessageImage.Attributes.Contains("red_ifyes"))
                            post_red_ifyes = (String)postMessageImage.Attributes["red_ifyes"];

                        if (postMessageImage.Attributes.Contains("description"))
                            post_description = (String)postMessageImage.Attributes["description"];

                        if (postMessageImage.Attributes.Contains("red_basispointrange"))
                            post_red_basispointrange = (String)postMessageImage.Attributes["red_basispointrange"];

                        if (postMessageImage.Attributes.Contains("estimatedclosedate"))
                            post_estimatedclosedate = postMessageImage.Attributes["estimatedclosedate"].ToString();

                        //   if (postMessageImage.Attributes.Contains("modifiedon"))
                        Post_modifiedon = (DateTime.UtcNow).ToString();

                        //update the old and new values of topic field in description field 

                        Entity opportunity_analysis = new Entity("red_opportunityanalysis");
                        // var changed_values = "";


                        if (pre_estimatedclosedate != null && pre_estimatedclosedate != post_estimatedclosedate)
                        {

                            opportunity_analysis["red_targetsigningdate"] = DateTime.Parse(post_estimatedclosedate.ToString());
                            opportunity_analysis["red_estimatedclosedate"] = DateTime.Parse(pre_estimatedclosedate.ToString());
                            changeValues.Add("Target Signing Date:", post_estimatedclosedate.ToString());
                            OpportunityChanges changes = new OpportunityChanges();
                            changes.SummaryDescription = pre_name;
                            changes.opportunitycode = post_uniqueid;
                            changes.Division = post_red_division.Name;
                            changes.ElementChanged = "Target Signing Date";
                            changes.PreviousInput = pre_estimatedclosedate;
                            changes.PresentInput = post_estimatedclosedate;
                            changes.CustomerName = pre_parentaccountid.Name;
                           // changes.Modifiedby = post_modifiedby.Name;
                            changes.SLTMember = postSltmember;
                            changes.Modifiedon = Post_modifiedon;
                          //  changes.DealOwner = post_dealowner.Name;



                            // ocold.Add("Target Signing Date:", pre_estimatedclosedate.ToString());

                            ocold.Add("TargetSigningDate", changes);

                            oc.red_targetsigningdate = 1;
                        }


                        if (pre_red_includeintherevenueforecast != null && pre_red_includeintherevenueforecast != post_red_includeintherevenueforecast)
                        {
                            opportunity_analysis["red_includeintherevenueforecast"] = new OptionSetValue(Convert.ToInt32(pre_red_includeintherevenueforecast));
                            opportunity_analysis["red_newincludeinrevenueforecast"] = new OptionSetValue(Convert.ToInt32(post_red_includeintherevenueforecast));
                            changeValues.Add("2022 Commitment:", postincludeinrevenue.ToString());
                            OpportunityChanges revenuechanges = new OpportunityChanges();
                            revenuechanges.SummaryDescription = pre_name;
                            revenuechanges.opportunitycode = post_uniqueid;
                            revenuechanges.Division = post_red_division.Name;
                            revenuechanges.ElementChanged = "2022 Commitment";
                            revenuechanges.PreviousInput = preincludeinrevenue;
                            revenuechanges.PresentInput = postincludeinrevenue;
                            revenuechanges.CustomerName = pre_parentaccountid.Name;
//      revenuechanges.Modifiedby = post_modifiedby.Name;
                            revenuechanges.SLTMember = postSltmember;
                            revenuechanges.Modifiedon = Post_modifiedon;
                         //  revenuechanges.DealOwner = post_dealowner.Name;


                            // ocold.Add("Target Signing Date:", pre_estimatedclosedate.ToString());

                            ocold.Add("2022 Commitment", revenuechanges);
                            // unchangeValues.Add("Include in the Revenue Forecast :", preincludeinrevenue.ToString());
                            oc.red_includeinrevenueforecast = 1;

                        }


                        //if (pre_red__asset_value != null)



                        if (pre_red__asset_value != null && pre_red__asset_value != post_red__Asset_value)
                        {
                            opportunity_analysis["red_asset_value"] = new Money(Decimal.Parse(pre_red__asset_value.ToString()));
                            opportunity_analysis["red_new_aua"] = new Money(Decimal.Parse(post_red__Asset_value.ToString()));
                            changeValues.Add("AUA (billion) :", post_red__Asset_value.ToString());

                            OpportunityChanges AUAchanges = new OpportunityChanges();
                            AUAchanges.SummaryDescription = pre_name;
                            AUAchanges.opportunitycode = post_uniqueid;
                            AUAchanges.Division = post_red_division.Name;
                            AUAchanges.ElementChanged = "AUA (billion)";
                            AUAchanges.PreviousInput = pre_red__asset_value;
                            AUAchanges.PresentInput = post_red__Asset_value;
                            AUAchanges.CustomerName = pre_parentaccountid.Name;
                           // AUAchanges.Modifiedby = post_modifiedby.Name;
                            AUAchanges.SLTMember = postSltmember;
                            AUAchanges.Modifiedon = Post_modifiedon;
                        //   AUAchanges.DealOwner = post_dealowner.Name;


                            ocold.Add("AUA", AUAchanges);

                            //unchangeValues.Add("AUA (billion) :", pre_red__asset_value.ToString());
                            oc.red_aua = 1;

                        }
                        if (pre_red__bps_cal != null && pre_red__bps_cal != post_red__bps_cal)
                        {
                            opportunity_analysis["red_bps_cal"] = new Money(Decimal.Parse(pre_red__bps_cal.ToString()));
                            opportunity_analysis["red_new_bps_cal"] = new Money(Decimal.Parse(post_red__bps_cal.ToString()));
                        }
                        //opportunity_analysis["red_new_bps_cal_base"] = new Money(Decimal.Parse(post_red__bps_cal_base.ToString())); }

                        /*  if (pre_red__asset_value_Base != null && pre_red__asset_value_Base != post_red__asset_value_Base)
                           {
                               opportunity_analysis["red_asset_value_base"] = new Money(Decimal.Parse(pre_red__asset_value_Base.ToString()));

                              opportunity_analysis["red_new_aua_base"] = new Money(Decimal.Parse(post_red__asset_value_Base.ToString()));
                          }*/
                    /*    if (pre_modifiedby != null && pre_modifiedby != post_modifiedby)
                        {
                            opportunity_analysis["modifiedby"] = pre_modifiedby;
                            opportunity_analysis["red_newdealowner"] = post_modifiedby;
                        }*/

                        if (pre_StatusCode != null && pre_StatusCode != post_StatusCode)
                        {
                            opportunity_analysis["red_status"] = new OptionSetValue(Convert.ToInt32(pre_StatusCode.ToString()));
                            opportunity_analysis["red_newstatus"] = new OptionSetValue(Convert.ToInt32(post_StatusCode.ToString()));
                        }


                        if (pre_statusreason != null && pre_statusreason != post_statusreason)
                        {
                            opportunity_analysis["red_statusreason"] = new OptionSetValue(Convert.ToInt32(pre_statusreason.ToString()));
                            opportunity_analysis["red_newstatusreason"] = new OptionSetValue(Convert.ToInt32(post_statusreason.ToString()));
                        }
                        if (pre_opportunityratingcode != null && pre_opportunityratingcode != post_opportunityratingcode)
                        {
                            opportunity_analysis["red_prioritymatrix"] = new OptionSetValue(Convert.ToInt32(pre_opportunityratingcode.ToString()));
                            opportunity_analysis["red_newprioritymatrix"] = new OptionSetValue(Convert.ToInt32(post_opportunityratingcode.ToString()));
                          

                            changeValues.Add("Priority:", post_opportunityratingcode);

                            OpportunityChanges prioritychanges = new OpportunityChanges();
                            prioritychanges.SummaryDescription = pre_name;
                            prioritychanges.opportunitycode = post_uniqueid;
                            prioritychanges.Division = post_red_division.Name;
                            prioritychanges.ElementChanged = "Priority";
                            prioritychanges.PreviousInput = preopportunitycode;
                            prioritychanges.PresentInput = postopportunitycode;
                            prioritychanges.CustomerName = pre_parentaccountid.Name;
                           // prioritychanges.Modifiedby = post_modifiedby.Name;
                            prioritychanges.SLTMember = postSltmember;
                            prioritychanges.Modifiedon = Post_modifiedon;
                         //   prioritychanges.DealOwner = post_dealowner.Name;



                            ocold.Add("Priority", prioritychanges);
                            oc.red_priorityaspersalesprioritisationmatrix = 1;

                        }
                        if (pre_red__consultancy_fee_base != null && pre_red__consultancy_fee_base != post_red__consultancy_fee_base)
                        {
                            opportunity_analysis["red_consultancy_fee_base"] = new Money(Decimal.Parse(pre_red__consultancy_fee_base.ToString()));
                            opportunity_analysis["red_newconsultancy_fee_base"] = new Money(Decimal.Parse(post_red__consultancy_fee_base.ToString()));
                        }
                        if (pre_red__consultancy_fee != null && pre_red__consultancy_fee != post_red__consultancy_fee)
                        {
                            opportunity_analysis["red_consultancy_fee"] = new Money(Decimal.Parse(pre_red__consultancy_fee.ToString()));
                            opportunity_analysis["red_newconsultancy_fee"] = new Money(Decimal.Parse(post_red__consultancy_fee.ToString()));
                            changeValues.Add("Total Scoping/Consultancy Fee (million) :", post_red__consultancy_fee.ToString());

                            OpportunityChanges scopechanges = new OpportunityChanges();
                            scopechanges.SummaryDescription = pre_name;
                            scopechanges.Division = post_red_division.Name;
                            scopechanges.opportunitycode = post_uniqueid;
                            scopechanges.ElementChanged = "Total Scoping/Consultancy Fee (million)";
                            scopechanges.PreviousInput = pre_red__consultancy_fee;
                            scopechanges.PresentInput = post_red__consultancy_fee;
                            scopechanges.CustomerName = pre_parentaccountid.Name;
                            //scopechanges.Modifiedby = post_modifiedby.Name;
                            scopechanges.SLTMember = postSltmember;
                            scopechanges.Modifiedon = Post_modifiedon;
                           // scopechanges.DealOwner = post_dealowner.Name;

                            ocold.Add("TotalScoping", scopechanges);
                            // unchangeValues.Add("Total Scoping/Consultancy Fee (million) :", pre_red__consultancy_fee.ToString());
                            oc.red_totalscopingconsultancyfeemillion = 1;

                        }

                        if (pre_red_deliveryfees != null && pre_red_deliveryfees != post_red_deliveryfees)
                        {
                            opportunity_analysis["red_deliveryfees"] = new Money(Decimal.Parse(pre_red_deliveryfees.ToString()));
                            // opportunity_analysis["red_new_deliveryfees"] = new Money(Decimal.Parse(post_red_deliveryfees.ToString()));
                        }
                        /* if (pre_red_implamentationfee_base != null && pre_red_implamentationfee_base != post_red_implamentationfee_base)
                         {
                             opportunity_analysis["red_implamentationfee_base"] = new Money(Decimal.Parse(pre_red_implamentationfee_base.ToString()));
                             opportunity_analysis["red_newimplamentationfee_base"] = new Money(Decimal.Parse(post_red_implamentationfee_base.ToString()));
                         }*/

                        if (pre_red_implamentationfee != null && pre_red_implamentationfee != post_red_implamentationfee)
                        {
                            opportunity_analysis["red_implamentationfee"] = new Money(Decimal.Parse(pre_red_implamentationfee.ToString()));
                            opportunity_analysis["red_newimplamentationfee"] = new Money(Decimal.Parse(post_red_implamentationfee.ToString()));
                            changeValues.Add("Total D&E Fee :", post_red_implamentationfee.ToString());

                            OpportunityChanges DEchanges = new OpportunityChanges();
                            DEchanges.SummaryDescription = pre_name;
                            DEchanges.Division = post_red_division.Name;
                            DEchanges.ElementChanged = "Total D&E Fee :";
                            DEchanges.PreviousInput = pre_red_implamentationfee;
                            DEchanges.PresentInput = post_red_implamentationfee;
                            DEchanges.CustomerName = pre_parentaccountid.Name;
                          // DEchanges.Modifiedby = post_modifiedby.Name;
                            DEchanges.SLTMember = postSltmember;
                            DEchanges.Modifiedon = Post_modifiedon;
                         //  DEchanges.DealOwner = post_dealowner.Name;

                            ocold.Add("TotalDandE", DEchanges);

                            // unchangeValues.Add("Total D&E Fee :", pre_red_implamentationfee.ToString());
                            oc.red_totaldefeemillion = 1;

                        }
                        /*  if (pre_red_implamentationfee_base != null && pre_red_implamentationfee != post_red_implamentationfee)
                          {
                              opportunity_analysis["red_implamentationfee"] = new Money(Decimal.Parse(pre_red_implamentationfee.ToString()));
                              opportunity_analysis["red_newimplamentationfee"] = new Money(Decimal.Parse(post_red_implamentationfee.ToString()));
                              changeValues.Add("Total D&E Fee :", post_red_implamentationfee.ToString());

                              OpportunityChanges DEchanges = new OpportunityChanges();
                              DEchanges.OpportunityName = pre_name;
                              DEchanges.Division = post_red_division.Name;
                              DEchanges.ElementChanged = "Total D&E Fee :";
                              DEchanges.PreviousInput = pre_red_implamentationfee;
                              DEchanges.PresentInput = post_red_implamentationfee;

                              ocold.Add("TotalDandE", DEchanges);

                              // unchangeValues.Add("Total D&E Fee :", pre_red_implamentationfee.ToString());
                              oc.red_totaldefeemillion = 1;

                          } */


                        /* if (pre_red__bps_cal != null )
                         opportunity_analysis["red_bps_cal"] = new Money(Decimal.Parse(pre_red__bps_cal.ToString()));
                         if (pre_red__bps_cal != null && pre_red__bps_cal != post_red__bps_cal)
                             changed_values += " , " + pre_red__bps_cal;
                         if (pre_red__bps_cal_base != null)
                             opportunity_analysis["red_bps_cal_base"] = new Money(Decimal.Parse(pre_red__bps_cal_base.ToString()));
                         if (pre_red__bps_cal_base != null && pre_red__bps_cal_base != post_red__bps_cal_base)
                             changed_values += " , " + pre_red__bps_cal_base;



                         if (pre_red_deliveryfees != null )
                         {
                             opportunity_analysis["red_deliveryfees"] = new Money(Decimal.Parse(pre_red_deliveryfees.ToString()));
                         }

                         #/                    if (pre_red_deliveryfees != null && pre_red_deliveryfees != post_red_deliveryfees)
                         if (pre_red_deliveryfees != null && pre_red_deliveryfees != post_red_deliveryfees)
                             changed_values += " , " + pre_red_deliveryfees;


                         if (pre_red_deliveryfees_Base != null)
                         {
                             opportunity_analysis["red_deliveryfees_base"] = new Money(Decimal.Parse(pre_red_deliveryfees_Base.ToString()));
                         }
                         if (pre_red_deliveryfees_Base != null && pre_red_deliveryfees_Base != post_red_deliveryfees_Base)
                             changed_values += " , " + pre_red_deliveryfees_Base;


                         if (pre_red__basepoint != null)
                         {
                             opportunity_analysis["red_basepoint"] = Decimal.Parse(pre_red__basepoint.ToString());

                         }
                         */

                        if (pre_red__basepoint != null && pre_red__basepoint != post_red__basepoint)
                        {
                            opportunity_analysis["red_newbasepoint"] = Decimal.Parse(post_red__basepoint.ToString());
                            opportunity_analysis["red_basepoint"] = Decimal.Parse(pre_red__basepoint.ToString());
                            changeValues.Add("Average Basis Point :", post_red__basepoint.ToString());
                            OpportunityChanges ABPchanges = new OpportunityChanges();
                            ABPchanges.SummaryDescription = pre_name;
                            ABPchanges.opportunitycode = post_uniqueid;
                            ABPchanges.Division = post_red_division.Name;
                            ABPchanges.ElementChanged = "Average Basis Point";
                            ABPchanges.PreviousInput = pre_red__basepoint;
                            ABPchanges.PresentInput = post_red__basepoint;
                            ABPchanges.CustomerName = pre_parentaccountid.Name;
                          // ABPchanges.Modifiedby = post_modifiedby.Name;
                            ABPchanges.SLTMember = postSltmember;
                            ABPchanges.Modifiedon = Post_modifiedon;
                         //  ABPchanges.DealOwner = post_dealowner.Name;

                            ocold.Add("AverageBasisPoint", ABPchanges);
                            // unchangeValues.Add("Average Basis Point :", pre_red__basepoint.ToString());
                            oc.red_averagebasispoint = 1;

                        }


                        if (pre_exchangerate != null)
                        {
                            opportunity_analysis["exchangerate"] = pre_exchangerate;
                        }
                        if (pre_parentaccountid != null)
                        {
                            opportunity_analysis["red_parentaccountid"] = pre_parentaccountid;

                        }

                        if (pre_parentcontactid != null)
                        {
                            opportunity_analysis["red_parentcontactid"] = pre_parentcontactid;

                        }

                        if (pre_transactioncurrencyid != null)
                        {
                            opportunity_analysis["transactioncurrencyid"] = pre_transactioncurrencyid;


                        }
                        if (pre_red_territory != null)
                        {
                            opportunity_analysis["red_territory"] = pre_red_territory;

                        }

                        if (pre_red_division != null)
                        {
                            opportunity_analysis["red_division"] = pre_red_division;

                            oc.red_division = pre_red_division;

                        }

                        if (pre_red_regions != null)
                        {
                            opportunity_analysis["red_regions"] = pre_red_regions;


                        }

                        if (pre_originatingleadid != null)
                        {
                            opportunity_analysis["red_originatingleadid"] = pre_originatingleadid;


                        }



                        if (pre_red_opptype != null)
                        {
                            opportunity_analysis["red_opptype"] = new OptionSetValue(Convert.ToInt32(pre_red_opptype));

                        }


                        //    if (pre_red_multisolution != null)
                        //  if (pre_red_multisolution != null)

                        //      opportunity_analysis["red_multisolution"] = new OptionSetValue(Convert.ToInt32(pre_red_multisolution));



                        if (pre_red_probability != null)
                        {
                            opportunity_analysis["red_probability"] = new OptionSetValue(Convert.ToInt32(pre_red_probability));
                        }

                        if (pre_red_primarysolution != null && pre_red_primarysolution != post_red_primarysolution)
                        {

                            opportunity_analysis["red_primarysolution"] = pre_red_primarysolution;
                            opportunity_analysis["red_newprimarysolution"] = post_red_primarysolution;
                        }


                        if (pre_red_probability != null && pre_red_probability != post_red_probability)
                        {
                            opportunity_analysis["red_probability"] = new OptionSetValue(Convert.ToInt32(pre_red_probability.ToString()));
                            opportunity_analysis["red_newprobability"] = new OptionSetValue(Convert.ToInt32(post_red_probability.ToString()));

                            changeValues.Add("Probability :", postProbabilityValue.ToString());
                            OpportunityChanges prochanges = new OpportunityChanges();
                            prochanges.SummaryDescription = pre_name;
                            prochanges.opportunitycode = post_uniqueid;
                            prochanges.Division = post_red_division.Name;
                            prochanges.ElementChanged = "Probability";
                            prochanges.PreviousInput = preProbabilityValue;
                            prochanges.PresentInput = postProbabilityValue;
                            prochanges.CustomerName = pre_parentaccountid.Name;
                        //  prochanges.Modifiedby = post_modifiedby.Name;
                            prochanges.SLTMember = postSltmember;
                            prochanges.Modifiedon = Post_modifiedon;
                        //   prochanges.DealOwner = post_dealowner.Name;



                            ocold.Add("Probability", prochanges);

                            //   unchangeValues.Add("Probability :", preProbabilityValue.ToString());xws
                            oc.red_probability = 1;

                        }

                        if (pre_red_regulatorypermissions != null)
                        {
                            opportunity_analysis["red_regulatorypermissions"] = new OptionSetValue(Convert.ToInt32(pre_red_regulatorypermissions));
                        }

                        if (pre_red_stage != null && pre_red_stage != post_red_stage)
                        {
                            if (Convert.ToInt32(post_red_stage) == 283390005 || Convert.ToInt32(post_red_stage) == 283390007)
                            {
                                opportunity_analysis["red_newstage"] = new OptionSetValue(Convert.ToInt32(post_red_stage));
                                opportunity_analysis["red_stage"] = new OptionSetValue(Convert.ToInt32(pre_red_stage));
                                changeValues.Add("Stage:", postStageValue);
                                OpportunityChanges stagechanges = new OpportunityChanges();
                                stagechanges.SummaryDescription = pre_name;
                                stagechanges.opportunitycode = post_uniqueid;
                                stagechanges.Division = post_red_division.Name;
                                stagechanges.ElementChanged = "Stage";
                                stagechanges.PreviousInput = preStageValue;
                                stagechanges.PresentInput = postStageValue;
                                stagechanges.CustomerName = pre_parentaccountid.Name;
                            //    stagechanges.Modifiedby = post_modifiedby.Name;
                                stagechanges.SLTMember = postSltmember;
                                stagechanges.Modifiedon = Post_modifiedon;
                             //  stagechanges.DealOwner = post_dealowner.Name;

                                ocold.Add("Stage", stagechanges);

                                //unchangeValues.Add("Probability :", preStageValue.ToString());
                                oc.red_stage = 1;
                            }


                        }
                        if (pre_red_sltmember != null)
                            opportunity_analysis["red_sltmember"] = new OptionSetValue(Convert.ToInt32(pre_red_sltmember));
                        if (pre_name != null)
                            opportunity_analysis["red_name"] = pre_name;

                        if (pre_red_ifyes != null)
                            opportunity_analysis["red_ifyes"] = pre_red_ifyes;
                        if (pre_description != null)
                            opportunity_analysis["red_description"] = pre_description;

                        if (pre_red_basispointrange != null)
                            opportunity_analysis["red_basispointrange"] = pre_red_basispointrange;

                        /*    var ExistingOppAnalystByOppId = service.Retrieve("")




                          if (changed_values != " " || changed_values != null)
                            {
                                if (post_estimatedclosedate != null)

                                if (post_red_includeintherevenueforecast != null)

                                if (post_red__Asset_value != null)

                                if (post_red__basepoint != null)
                                    opportunity_analysis["red_new_bps_cal_base"] = new Money(Decimal.Parse(post_red__bps_cal_base.ToString()));
                              //  if (post_red_primarysolution opportunity_analysis["red_new_bps_cal_base"] = new Money(Decimal.Parse(post_red__bps_cal_base.ToString())); != null)

                                if (post_red__bps_cal != null)
                                    opportunity_analysis["red_new_bps_cal"] = new Money(Decimal.Parse(post_red__bps_cal.ToString()));
                          opportunity_analysis["red_new_bps_cal_base"] = new Money(Decimal.Parse(post_red__bps_cal_base.ToString()));
         opportunity_analysis["red_new_bps_cal_base"] = new Money(Decimal.Parse(post_red__bps_cal_base.ToString()));
                                if (post_red__bps_cal_base != null)
                                    opportunity_analysis["red_new_bps_cal_base"] = new Money(Decimal.Parse(post_red__bps_cal_base.ToString()));

         opportunity_analysis["red_new_bps_cal_base"] = new Money(Decimal.Parse(post_red__bps_cal_base.ToString()));
                                if (post_red_deliveryfees != null)

                                if (post_red_deliveryfees_Base != null)
                                    opportunity_analysis["red_new_deliveryfees_base"] = new Money(Decimal.Parse(post_red_deliveryfees_Base.ToString()));
                               // if (post_red_multisolution != null)
                                 //   opportunity_analysis["red_new_multisolution"] = new OptionSetValue(Convert.ToInt32(post_red_multisolution.ToString()));
                                if (post_red_probability != null)


                                if (post_red_implamentationfee != null)



                                if (pre_estimatedclosedate != null)

                                if(pre_red_includeintherevenueforecast != null)

                                //if (pre_red_primarysolution != null)
                                  //  opportunity_analysis["red_primarysolution"] = new OptionSetValue(Convert.ToInt32(pre_red_primarysolution));
                                if (pre_red__basepoint != null)
                                    opportunity_analysis["red_basepoint"] = Decimal.Parse(pre_red__basepoint.ToString());
                                if (pre_red__asset_value != null)
                                    opportunity_analysis["red_asset_value"] = new Money(Decimal.Parse(pre_red__asset_value.ToString()));
                                if (pre_red__bps_cal != null)
                                    opportunity_analysis["red_bps_cal"] = new Money(Decimal.Parse(pre_red__bps_cal.ToString()));
                                if (pre_red__bps_cal_base != null)
                                    opportunity_analysis["red_bps_cal_base"] = new Money(Decimal.Parse(pre_red__bps_cal_base.ToString()));
                                if (pre_red_deliveryfees != null)
                                    opportunity_analysis["red_deliveryfees"] = new Money(Decimal.Parse(pre_red_deliveryfees.ToString()));
                                if (post_red_deliveryfees_Base != null)
                                    opportunity_analysis["red_deliveryfees_base"] = new Money(Decimal.Parse(post_red_deliveryfees_Base.ToString()));

                               //     opportunity_analysis["red_multisolution"] = new OptionSetValue(Convert.ToInt32(pre_red_multisolution.ToString()));
                                if (pre_red_probability != null)

                                )); */

                        QueryExpression query = new QueryExpression();

                        query.EntityName = "red_opportunityanalysis";
                        query.ColumnSet = new ColumnSet(true);

                        query.LinkEntities.Add(new LinkEntity("red_opportunityanalysis", "opportunity", "red_opportunityid", "opportunityid", JoinOperator.Inner));
                        query.LinkEntities[0].Columns.AddColumns("opportunityid", "name");
                        query.LinkEntities[0].EntityAlias = "temp";

                        query.Criteria = new FilterExpression();
                        query.Criteria.AddCondition("red_opportunityid", ConditionOperator.Equal, opportunityid);
                        query.Orders.Add(new OrderExpression("modifiedon", OrderType.Descending));

                        query.Distinct = true;
                        query.TopCount = 1;

                        EntityCollection opportunityAnalysisCollection = service.RetrieveMultiple(query);

                        if (opportunityAnalysisCollection.Entities.Count > 0)
                        {
                            opportunity_analysis["red_opportunityanalysisid"] = opportunityAnalysisCollection.Entities[0].Id;
                            service.Update(opportunity_analysis);
                        }
                        else
                        {
                            opportunity_analysis["red_opportunityid"] = new EntityReference("opportunity", opportunityid);

                            service.Create(opportunity_analysis);
                        }



                        var jsonValues = "";
                        var previousChangedValues = "";
                        var newTextValues = "";

                        Dictionary<string, string> existingValues = new Dictionary<string, string>();


                        Entity opp = service.Retrieve(entity.LogicalName, entity.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));
                        JavaScriptSerializer deSerializer = new JavaScriptSerializer();
                        JavaScriptSerializer serializer = new JavaScriptSerializer();


                        if (changeValues.Count > 0)
                        {

                            if (opp.Attributes.ContainsKey("red_json") && opp.Attributes["red_json"].ToString() != "")
                            {

                                existingValues = deSerializer.Deserialize<Dictionary<string, string>>(opp.Attributes["red_json"].ToString());
                            }

                            if (opp.Attributes.ContainsKey("red_texttoadd") && opp.Attributes["red_texttoadd"].ToString() != "")
                            {

                                previousChangedValues = opp.Attributes["red_texttoadd"].ToString();
                                newTextValues = previousChangedValues + "\n";
                            }
                            Dictionary<string, string> newValues = new Dictionary<string, string>();
                            Dictionary<string, string> oldValues = new Dictionary<string, string>();



                            var totalDE = "";
                            var stage = "";
                            var probability = "";
                            var aua = "";
                            var basePoint = "";
                            var estimateddate = "";
                            var scopingfee = "";
                            var includeForecast = "";
                            var prioritymatrix = "";
                            var pretotalDE = "";
                            var prestage = "";
                            var preprobability = "";
                            var preaua = "";
                            var prebasePoint = "";
                            var preestimateddate = "";
                            var prescopingfee = "";
                            var preincludeForecast = "";
                            var preprioritymatrix = "";

                            totalDE = existingValues.ContainsKey("Total D&E Fee :") && changeValues.ContainsKey("Total D&E Fee :") ? changeValues.FirstOrDefault(x => x.Key == "Total D&E Fee :").Value : existingValues.ContainsKey("Total D&E Fee :") ? existingValues.FirstOrDefault(x => x.Key == "Total D&E Fee :").Value : changeValues.ContainsKey("Total D&E Fee :") ? changeValues.FirstOrDefault(x => x.Key == "Total D&E Fee :").Value : "";
                            stage = existingValues.ContainsKey("Stage:") && changeValues.ContainsKey("Stage:") ? changeValues.FirstOrDefault(x => x.Key == "Stage:").Value : existingValues.ContainsKey("Stage:") ? existingValues.FirstOrDefault(x => x.Key == "Stage:").Value : changeValues.ContainsKey("Stage:") ? changeValues.FirstOrDefault(x => x.Key == "Stage:").Value : "";
                            probability = existingValues.ContainsKey("Probability :") && changeValues.ContainsKey("Probability :") ? changeValues.FirstOrDefault(x => x.Key == "Probability :").Value : existingValues.ContainsKey("Probability :") ? existingValues.FirstOrDefault(x => x.Key == "Probability :").Value : changeValues.ContainsKey("Probability :") ? changeValues.FirstOrDefault(x => x.Key == "Probability :").Value : "";
                            aua = existingValues.ContainsKey("AUA (billion) :") && changeValues.ContainsKey("AUA (billion) :") ? changeValues.FirstOrDefault(x => x.Key == "AUA (billion) :").Value : existingValues.ContainsKey("AUA (billion) :") ? existingValues.FirstOrDefault(x => x.Key == "AUA (billion) :").Value : changeValues.ContainsKey("AUA (billion) :") ? changeValues.FirstOrDefault(x => x.Key == "AUA (billion) :").Value : "";
                            basePoint = existingValues.ContainsKey("Average Basis Point :") && changeValues.ContainsKey("Average Basis Point :") ? changeValues.FirstOrDefault(x => x.Key == "Average Basis Point :").Value : existingValues.ContainsKey("Average Basis Point :") ? existingValues.FirstOrDefault(x => x.Key == "Average Basis Point :").Value : changeValues.ContainsKey("Average Basis Point :") ? changeValues.FirstOrDefault(x => x.Key == "Average Basis Point :").Value : "";
                            estimateddate = existingValues.ContainsKey("Target Signing Date:") && changeValues.ContainsKey("Target Signing Date:") ? changeValues.FirstOrDefault(x => x.Key == "Target Signing Date:").Value : existingValues.ContainsKey("Target Signing Date:") ? existingValues.FirstOrDefault(x => x.Key == "Target Signing Date:").Value : changeValues.ContainsKey("Target Signing Date:") ? changeValues.FirstOrDefault(x => x.Key == "Target Signing Date:").Value : "";
                            includeForecast = existingValues.ContainsKey("2022 Commitment:") && changeValues.ContainsKey("2022 Commitment:") ? changeValues.FirstOrDefault(x => x.Key == "2022 Commitment:").Value : existingValues.ContainsKey("2022 Commitment:") ? existingValues.FirstOrDefault(x => x.Key == "2022 Commitment:").Value : changeValues.ContainsKey("2022 Commitment:") ? changeValues.FirstOrDefault(x => x.Key == "2022 Commitment:").Value : "";
                            scopingfee = existingValues.ContainsKey("Total Scoping/Consultancy Fee (million) :") && changeValues.ContainsKey("Total Scoping/Consultancy Fee (million) :") ? changeValues.FirstOrDefault(x => x.Key == "Total Scoping/Consultancy Fee (million) :").Value : existingValues.ContainsKey("Total Scoping/Consultancy Fee (million) :") ? existingValues.FirstOrDefault(x => x.Key == "Total Scoping/Consultancy Fee (million) :").Value : changeValues.ContainsKey("Total Scoping/Consultancy Fee (million) :") ? changeValues.FirstOrDefault(x => x.Key == "Total Scoping/Consultancy Fee (million) :").Value : "";
                            prioritymatrix = existingValues.ContainsKey("Priority:") && changeValues.ContainsKey("Priority:") ? changeValues.FirstOrDefault(x => x.Key == "Priority:").Value : existingValues.ContainsKey("Priority:") ? existingValues.FirstOrDefault(x => x.Key == "Priority:").Value : changeValues.ContainsKey("Priority:") ? changeValues.FirstOrDefault(x => x.Key == "Priority:").Value : "";
                            //                        pretotalDE = existingValues.ContainsKey("Total D&E Fee :") && unchangeValues.ContainsKey("Total D&E Fee :") ? unchangeValues.FirstOrDefault(x => x.Key == "Total D&E Fee :").Value : existingValues.ContainsKey("Total D&E Fee :") ? existingValues.FirstOrDefault(x => x.Key == "Total D&E Fee :").Value : unchangeValues.ContainsKey("Total D&E Fee :") ? unchangeValues.FirstOrDefault(x => x.Key == "Total D&E Fee :").Value : "";
                            //  prestage = existingValues.ContainsKey("Stage:") && unchangeValues.ContainsKey("Stage:") ? unchangeValues.FirstOrDefault(x => x.Key == "Stage:").Value : existingValues.ContainsKey("Stage:") ? existingValues.FirstOrDefault(x => x.Key == "Stage:").Value : unchangeValues.ContainsKey("Stage:") ? unchangeValues.FirstOrDefault(x => x.Key == "Stage:").Value : "";
                            //preprobability = existingValues.ContainsKey("Probability :") && unchangeValues.ContainsKey("Probability :") ? unchangeValues.FirstOrDefault(x => x.Key == "Probability :").Value : existingValues.ContainsKey("Probability :") ? existingValues.FirstOrDefault(x => x.Key == "Probability :").Value : unchangeValues.ContainsKey("Probability :") ? unchangeValues.FirstOrDefault(x => x.Key == "Probability :").Value : "";
                            // preaua = existingValues.ContainsKey("AUA (billion) :") && unchangeValues.ContainsKey("AUA (billion) :") ? unchangeValues.FirstOrDefault(x => x.Key == "AUA (billion) :").Value : existingValues.ContainsKey("AUA (billion) :") ? existingValues.FirstOrDefault(x => x.Key == "AUA (billion) :").Value : unchangeValues.ContainsKey("AUA (billion) :") ? unchangeValues.FirstOrDefault(x => x.Key == "AUA (billion) :").Value : "";
                            //prebasePoint = existingValues.ContainsKey("Average Basis Point :") && unchangeValues.ContainsKey("Average Basis Point :") ? unchangeValues.FirstOrDefault(x => x.Key == "Average Basis Point :").Value : existingValues.ContainsKey("Average Basis Point :") ? existingValues.FirstOrDefault(x => x.Key == "Average Basis Point :").Value : unchangeValues.ContainsKey("Average Basis Point :") ? unchangeValues.FirstOrDefault(x => x.Key == "Average Basis Point :").Value : "";
                            //preestimateddate = existingValues.ContainsKey("Target Signing Date:") && unchangeValues.ContainsKey("Target Signing Date:") ? unchangeValues.FirstOrDefault(x => x.Key == "Target Signing Date:").Value : existingValues.ContainsKey("Target Signing Date:") ? existingValues.FirstOrDefault(x => x.Key == "Target Signing Date:").Value : unchangeValues.ContainsKey("Target Signing Date:") ? unchangeValues.FirstOrDefault(x => x.Key == "Target Signing Date:").Value : "";
                            //preincludeForecast = existingValues.ContainsKey("Include in the Revenue Forecast :") && unchangeValues.ContainsKey("Include in the Revenue Forecast :") ? unchangeValues.FirstOrDefault(x => x.Key == "Include in the Revenue Forecast :").Value : existingValues.ContainsKey("Include in the Revenue Forecast :") ? existingValues.FirstOrDefault(x => x.Key == "Include in the Revenue Forecast :").Value : unchangeValues.ContainsKey("Include in the Revenue Forecast :") ? unchangeValues.FirstOrDefault(x => x.Key == "Include in the Revenue Forecast :").Value : "";
                            //prescopingfee = existingValues.ContainsKey("Total Scoping/Consultancy Fee (million) :") && unchangeValues.ContainsKey("Total Scoping/Consultancy Fee (million) :") ? unchangeValues.FirstOrDefault(x => x.Key == "Total Scoping/Consultancy Fee (million) :").Value : existingValues.ContainsKey("Total Scoping/Consultancy Fee (million) :") ? existingValues.FirstOrDefault(x => x.Key == "Total Scoping/Consultancy Fee (million) :").Value : unchangeValues.ContainsKey("Total Scoping/Consultancy Fee (million) :") ? unchangeValues.FirstOrDefault(x => x.Key == "Total Scoping/Consultancy Fee (million) :").Value : "";



                            newTextValues = "<pre>";

                            // newTextValues = newTextValues + " Opportunity Name : " + post_name + "\n";
                            if (totalDE != "")
                            {
                                newValues.Add("Total D&E Fee :", totalDE);
                                oldValues.Add("Total D&E Fee :", pretotalDE);
                                newTextValues = newTextValues + "Total D&E Fee: " + totalDE + "\n";
                            }
                            if (stage != "")
                            {
                                newValues.Add("Stage:", stage);
                                oldValues.Add("Stage:", prestage);
                                newTextValues = newTextValues + "Stage: " + stage + "\n";
                            }
                            if (probability != "")
                            {
                                newValues.Add("Probability :", probability);
                                oldValues.Add("Probability :", preprobability);
                                newTextValues = newTextValues + "Probability: " + probability + "\n";

                            }
                            if (aua != "")
                            {
                                newValues.Add("AUA (billion) :", aua);
                                oldValues.Add("AUA (billion) :", preaua);

                                newTextValues = newTextValues + "AUA (billion) : " + aua + "\n";

                            }
                            if (scopingfee != "")
                            {
                                newValues.Add("Total Scoping/Consultancy Fee (million) :", scopingfee);
                                oldValues.Add("Total Scoping/Consultancy Fee (million) :", prescopingfee);
                                newTextValues = newTextValues + "Total Scoping/Consultancy Fee (million) : " + scopingfee + "\n";

                            }
                            if (basePoint != "")
                            {
                                newValues.Add("Average Basis Point :", basePoint);
                                oldValues.Add("Average Basis Point :", prebasePoint);
                                newTextValues = newTextValues + "Average Basis Point : " + basePoint + "\n";

                            }
                            if (estimateddate != "")
                            {
                                newValues.Add("Target Signing Date:", estimateddate);
                                oldValues.Add("Target Signing Date:", preestimateddate);
                                newTextValues = newTextValues + "Target Signing Date: " + estimateddate + "\n";

                            }
                            if (includeForecast != "")
                            {
                                newValues.Add("2022 Commitment:", includeForecast);
                                oldValues.Add("2022 Commitment:", preincludeForecast);
                                newTextValues = newTextValues + "2022 Commitment:" + includeForecast + "\n";

                            }
                            if (prioritymatrix != "")
                            {
                                newValues.Add("Priority:", prioritymatrix);
                                //  oldValues.Add("Priority as per sales prioritisation matrix :", preincludeForecast);
                                newTextValues = newTextValues + "Priority: " + prioritymatrix + "\n";

                            }
                            newTextValues = newTextValues + "</pre>";

                            //tracingService.Trace(" Existing Changed values : " + opp.Attributes["red_texttoadd"].ToString());

                            EntityReference entityReference = new EntityReference(entity.LogicalName, entity.Id);

                            trackedOpportunityFields = newTextValues;

                            jsonValues = serializer.Serialize(newValues);

                            var oldJsonValues = "";

                            if (opp.Attributes.ContainsKey("red_oldjson"))
                            {
                                var existingOppChanges = opp.Attributes["red_oldjson"].ToString();

                                Dictionary<string, OpportunityChanges> extChanges = serializer.Deserialize<Dictionary<string, OpportunityChanges>>(existingOppChanges);
                                Dictionary<string, OpportunityChanges> extwithChanges = new Dictionary<string, OpportunityChanges>();

                                foreach (var extVal in extChanges)
                                {
                                    if (ocold.ContainsKey(extVal.Key))
                                    {

                                        OpportunityChanges changes = ocold.Where(x => x.Key == extVal.Key).Select(y => y.Value).FirstOrDefault();

                                        extwithChanges.Add(extVal.Key, changes);
                                        //  extwithChanges.Add(extVal.Key, ocold.Where(x => x.Key == extVal.Key).Select(y => y.Value));
                                    }
                                    else
                                    {

                                        OpportunityChanges changes = extVal.Value;

                                        extwithChanges.Add(extVal.Key, changes);
                                        // extwithChanges.Add(extVal.Key, extVal.Value);

                                    }
                                }

                                foreach (var newVal in ocold)
                                {
                                    if (!extwithChanges.ContainsKey(newVal.Key))
                                    {
                                        //  extwithChanges.Add(newVal.Key, ocold.Where(x => x.Key == newVal.Key).Select(y => y.Value));

                                        // OpportunityChanges changes = serializer.Deserialize<OpportunityChanges>(ocold.Where(x => x.Key == newVal.Key).Select(y => y.Value).ToString());
                                        OpportunityChanges changes = ocold.Where(x => x.Key == newVal.Key).Select(y => y.Value).FirstOrDefault();

                                        extwithChanges.Add(newVal.Key, changes);
                                    }

                                }
                                oldJsonValues = serializer.Serialize(extwithChanges);

                            }
                            else
                            {
                                oldJsonValues = serializer.Serialize(ocold);
                            }



                            TrackChangedOpportunityFields(entityReference, service, trackedOpportunityFields, jsonValues, oldJsonValues);

                            CreateUpdateOpportunityCounting(oc, service);
                        }
                        //else
                        //{
                        //    if (changeValues.Count > 0)
                        //    {


                        //        JavaScriptSerializer serializer = new JavaScriptSerializer();

                        //        jsonValues = serializer.Serialize(changeValues);

                        //        tracingService.Trace(" New Changed values : " + trackedOpportunityFields);

                        //        EntityReference entityReference = new EntityReference(entity.LogicalName, entity.Id);

                        //        TrackChangedOpportunityFields(entityReference, service, trackedOpportunityFields, jsonValues);

                        //    }
                        //}


                        tracingService.Trace("Exited Plugin.... OpportunityAnalysis");



                    }
                }

                catch (FaultException<OrganizationServiceFault> ex)
                {

                    throw new NotImplementedException();
                }
                catch (Exception ex)
                {

                    tracingService.Trace("FollowUpPlugin: {0}", ex.ToString());
                    throw;

                }
                

            }
        }

        public void TrackChangedOpportunityFields(EntityReference opportunityReference, IOrganizationService service, string changedFields, string jsonValues, string oldJsonValues)
        {
            Entity opportunity = new Entity(opportunityReference.LogicalName, opportunityReference.Id);
            opportunity["red_texttoadd"] = changedFields;
            opportunity["red_json"] = jsonValues;
            opportunity["red_oldjson"] = oldJsonValues;
            opportunity["red_emailsent"] = false;
            service.Update(opportunity);
        }

        public void CreateUpdateOpportunityCounting(OpportunityCounting opportunityCounting, IOrganizationService service)
        {
            QueryExpression query = new QueryExpression();

            query.EntityName = "red_opportunitycounting";
            query.ColumnSet = new ColumnSet(true);

            EntityCollection opportunityCountingCollection = service.RetrieveMultiple(query);

            if(opportunityCountingCollection.Entities.Count > 0)
            {
                var sss = opportunityCountingCollection.Entities.Where(x=> x.Attributes.Keys.ToString() == "red_division" && x.Attributes.Values.Equals(opportunityCounting.red_division)).Select(x => x.Attributes.Values);

                bool division = false;

                Guid recordId;

                foreach(var oc in opportunityCountingCollection.Entities)
                {
                    if(((EntityReference)oc.Attributes["red_division"]).Id == opportunityCounting.red_division.Id)
                    {
                        division = true;
                        Entity ocEntity = new Entity("red_opportunitycounting");

                        ocEntity.Attributes.Add("red_aua", Convert.ToInt32(oc.Attributes["red_aua"]) + Convert.ToInt32(opportunityCounting.red_aua));
                        ocEntity.Attributes.Add("red_averagebasispoint", Convert.ToInt32(oc.Attributes["red_averagebasispoint"]) + Convert.ToInt32(opportunityCounting.red_averagebasispoint));
                        ocEntity.Attributes.Add("red_includeinrevenueforecast", Convert.ToInt32(oc.Attributes["red_includeinrevenueforecast"]) + Convert.ToInt32(opportunityCounting.red_includeinrevenueforecast));
                        ocEntity.Attributes.Add("red_probability", Convert.ToInt32(oc.Attributes["red_probability"]) + Convert.ToInt32(opportunityCounting.red_probability));
                        ocEntity.Attributes.Add("red_stage", Convert.ToInt32(oc.Attributes["red_stage"]) + Convert.ToInt32(opportunityCounting.red_stage));
                        ocEntity.Attributes.Add("red_targetsigningdate", Convert.ToInt32(oc.Attributes["red_targetsigningdate"]) + Convert.ToInt32(opportunityCounting.red_targetsigningdate));
                        ocEntity.Attributes.Add("red_totalscopingconsultancyfeemillion", Convert.ToInt32(oc.Attributes["red_totalscopingconsultancyfeemillion"]) + Convert.ToInt32(opportunityCounting.red_totalscopingconsultancyfeemillion));
                        ocEntity.Attributes.Add("red_totaldefeemillion", Convert.ToInt32(oc.Attributes["red_totaldefeemillion"]) + Convert.ToInt32(opportunityCounting.red_totaldefeemillion));
                        // ocEntity.Attributes.Add("red_priorityaspersalesprioritisationmatrix", oc.Attributes.ContainsKey("red_priorityaspersalesprioritisationmatrix") ? Convert.ToInt32(oc.Attributes["red_priorityaspersalesprioritisationmatrix"]) : 0 + Convert.ToInt32(opportunityCounting.red_priorityaspersalesprioritisationmatrix));
                        ocEntity.Attributes.Add("red_priorityaspersalesprioritisationmatrix", Convert.ToInt32(oc.Attributes["red_priorityaspersalesprioritisationmatrix"]) + Convert.ToInt32(opportunityCounting.red_priorityaspersalesprioritisationmatrix));
                        var sumofaua = Convert.ToInt32(ocEntity.Attributes["red_aua"].ToString());
                        var sumofabp = Convert.ToInt32(ocEntity.Attributes["red_averagebasispoint"].ToString());
                        var sumofirf = Convert.ToInt32(ocEntity.Attributes["red_includeinrevenueforecast"].ToString());
                        var sumofprobability = Convert.ToInt32(ocEntity.Attributes["red_probability"].ToString());
                        var sumofstage = Convert.ToInt32(ocEntity.Attributes["red_stage"].ToString());
                        var sumofTSD = Convert.ToInt32(ocEntity.Attributes["red_targetsigningdate"].ToString());
                        var sumofTSCF = Convert.ToInt32(ocEntity.Attributes["red_totalscopingconsultancyfeemillion"].ToString());
                        var sumofTDEfee = Convert.ToInt32(ocEntity.Attributes["red_totaldefeemillion"].ToString());
                        var sumofpriority = Convert.ToInt32(ocEntity.Attributes["red_priorityaspersalesprioritisationmatrix"].ToString());
                        var totalchanges = (sumofabp + sumofaua + sumofirf + sumofprobability + sumofstage + sumofTSD + sumofTSCF + sumofTDEfee + sumofpriority);

                        ocEntity.Attributes.Add("red_totalchanges", totalchanges);
                        


                        ocEntity.Attributes.Add("red_opportunitycountingid", oc.Id);
                        service.Update(ocEntity);
                    }
                    
                }

                if(division == false)
                {
                    OpportunityCounting oc = new OpportunityCounting();
                    Entity ocEntity = new Entity("red_opportunitycounting");

                    ocEntity.Attributes.Add("red_aua", Convert.ToInt32(opportunityCounting.red_aua));
                    ocEntity.Attributes.Add("red_averagebasispoint", Convert.ToInt32(opportunityCounting.red_averagebasispoint));
                    ocEntity.Attributes.Add("red_division", opportunityCounting.red_division);
                    ocEntity.Attributes.Add("red_includeinrevenueforecast", Convert.ToInt32(opportunityCounting.red_includeinrevenueforecast));
                    ocEntity.Attributes.Add("red_probability", Convert.ToInt32(opportunityCounting.red_probability));
                    ocEntity.Attributes.Add("red_stage", Convert.ToInt32(opportunityCounting.red_stage));
                    ocEntity.Attributes.Add("red_targetsigningdate", Convert.ToInt32(opportunityCounting.red_targetsigningdate));
                    ocEntity.Attributes.Add("red_totalscopingconsultancyfeemillion", Convert.ToInt32(opportunityCounting.red_totalscopingconsultancyfeemillion));
                    ocEntity.Attributes.Add("red_totaldefeemillion", Convert.ToInt32(opportunityCounting.red_totaldefeemillion));
                    ocEntity.Attributes.Add("red_priorityaspersalesprioritisationmatrix", Convert.ToInt32(opportunityCounting.red_priorityaspersalesprioritisationmatrix));
                    ocEntity.Attributes.Add("red_name", opportunityCounting.red_division.Name);

                    service.Create(ocEntity);
                }

               
                    
                
            }
            else
            {
                OpportunityCounting oc = new OpportunityCounting();
                Entity ocEntity = new Entity("red_opportunitycounting");

                ocEntity.Attributes.Add("red_aua", Convert.ToInt32(opportunityCounting.red_aua));
                ocEntity.Attributes.Add("red_averagebasispoint", Convert.ToInt32(opportunityCounting.red_averagebasispoint));
                ocEntity.Attributes.Add("red_division", opportunityCounting.red_division);
                ocEntity.Attributes.Add("red_includeinrevenueforecast", Convert.ToInt32(opportunityCounting.red_includeinrevenueforecast));
                ocEntity.Attributes.Add("red_probability", Convert.ToInt32(opportunityCounting.red_probability));
                ocEntity.Attributes.Add("red_stage", Convert.ToInt32(opportunityCounting.red_stage));
                ocEntity.Attributes.Add("red_targetsigningdate", Convert.ToInt32(opportunityCounting.red_targetsigningdate));
                ocEntity.Attributes.Add("red_totalscopingconsultancyfeemillion", Convert.ToInt32(opportunityCounting.red_totalscopingconsultancyfeemillion));
                ocEntity.Attributes.Add("red_totaldefeemillion", Convert.ToInt32(opportunityCounting.red_totaldefeemillion));
                                
                ocEntity.Attributes.Add("red_name", opportunityCounting.red_division.Name);

                service.Create(ocEntity);

            }
        }

    }
}
      