using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace msdyncrmWorkflowTools
{
    public class Common
    {
        public ITracingService tracingService;
        public IWorkflowContext context;
        public IOrganizationServiceFactory serviceFactory;
        public IOrganizationService service;
        public CodeActivityContext codeActivityContext;

        public Common(CodeActivityContext executionContext)
        {
            codeActivityContext = executionContext;
            tracingService = executionContext.GetExtension<ITracingService>();
            context = executionContext.GetExtension<IWorkflowContext>();
            serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            service = serviceFactory.CreateOrganizationService(context.UserId);
        }

        /// <summary>
        /// Query the Metadata to get the Entity Schema Name from the Object Type Code
        /// </summary>
        /// <param name="ObjectTypeCode"></param>
        /// <param name="service"></param>
        /// <returns>Entity Schema Name</returns>
        public string sGetEntityNameFromCode(string ObjectTypeCode, IOrganizationService service)
        {
            MetadataFilterExpression entityFilter = new MetadataFilterExpression(LogicalOperator.And);
            entityFilter.Conditions.Add(new MetadataConditionExpression("ObjectTypeCode", MetadataConditionOperator.Equals, Convert.ToInt32(ObjectTypeCode)));
            EntityQueryExpression entityQueryExpression = new EntityQueryExpression()
            {
                Criteria = entityFilter
            };
            RetrieveMetadataChangesRequest retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest()
            {
                Query = entityQueryExpression,
                ClientVersionStamp = null
            };
            RetrieveMetadataChangesResponse response = (RetrieveMetadataChangesResponse)service.Execute(retrieveMetadataChangesRequest);

            EntityMetadata entityMetadata = (EntityMetadata)response.EntityMetadata[0];
            return entityMetadata.SchemaName.ToLower();
        }

        /// <summary>
        /// Returns entity object type code for specified entity name
        /// </summary>
        /// <param name="EntityName"></param>
        /// <param name="service"></param>
        /// <returns></returns>
        public int sGetEntityCodeForName(string EntityName, IOrganizationService service)
        {
            MetadataFilterExpression entityFilter = new MetadataFilterExpression(LogicalOperator.And);
            entityFilter.Conditions.Add(new MetadataConditionExpression("SchemaName", MetadataConditionOperator.Equals, EntityName));
            EntityQueryExpression entityQueryExpression = new EntityQueryExpression()
            {
                Criteria = entityFilter
            };
            RetrieveMetadataChangesRequest retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest()
            {
                Query = entityQueryExpression,
                ClientVersionStamp = null
            };
            RetrieveMetadataChangesResponse response = (RetrieveMetadataChangesResponse)service.Execute(retrieveMetadataChangesRequest);

            EntityMetadata entityMetadata = (EntityMetadata)response.EntityMetadata[0];
            return entityMetadata.ObjectTypeCode ?? 0;
        }

        public EntityCollection getAssociations(string PrimaryEntityName, Guid PrimaryEntityId, string _relationshipName, string entityName, string ParentId)
        {
            //
            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                      <entity name='" + PrimaryEntityName + @"'>
                                        <link-entity name='" + _relationshipName + @"' from='" + PrimaryEntityName + @"id' to='" + PrimaryEntityName + @"id' visible='false' intersect='true'>
                                        
                                            <filter type='and'>
                                            <condition attribute='" + PrimaryEntityName + @"id' operator='eq' value='" + PrimaryEntityId.ToString() + @"' />
                                            </filter>
                                       
                                        <link-entity name='" + entityName + @"' from='" + entityName + @"id' to='" + entityName + @"id' alias='ac'>
                                                <filter type='and'>
                                                  <condition attribute='" + entityName + @"id' operator='eq' value='" + ParentId + @"' />
                                                </filter>
                                              </link-entity>
                                        </link-entity>
                                      </entity>
                                    </fetch>";
            tracingService.Trace(String.Format("FetchXML: {0} ", fetchXML));
            EntityCollection relations = service.RetrieveMultiple(new FetchExpression(fetchXML));

            return relations;
        }

        public List<string> getEntityAttributesToClone(string entityName, IOrganizationService service,
            ref string PrimaryIdAttribute, ref string PrimaryNameAttribute)
        {


            List<string> atts = new List<string>();
            RetrieveEntityRequest req = new RetrieveEntityRequest()
            {
                EntityFilters = EntityFilters.Attributes,
                LogicalName = entityName
            };

            RetrieveEntityResponse res = (RetrieveEntityResponse)service.Execute(req);
            PrimaryIdAttribute = res.EntityMetadata.PrimaryIdAttribute;

            foreach (AttributeMetadata attMetadata in res.EntityMetadata.Attributes)
            {
                if (attMetadata.IsPrimaryName.Value)
                {
                    PrimaryNameAttribute = attMetadata.LogicalName;
                }
                if ((attMetadata.IsValidForCreate.Value || attMetadata.IsValidForUpdate.Value)
                    && !attMetadata.IsPrimaryId.Value)
                {
                    //tracingService.Trace("Tipo:{0}", attMetadata.AttributeTypeName.Value.ToLower());
                    if (attMetadata.AttributeTypeName.Value.ToLower() == "partylisttype")
                    {
                        atts.Add("partylist-" + attMetadata.LogicalName);
                        //atts.Add(attMetadata.LogicalName);
                    }
                    else
                    {
                        atts.Add(attMetadata.LogicalName);
                    }
                }
            }

            return (atts);
        }

        /// <summary>
        /// Clone record
        /// If <param name="cloneChildRecord"/> is set to True -> then other input parameters will be used in cloning process.
        /// In this case we should update parent field. 30.11.2021. Before this fix, CloneChilg WF would completely clone child recod, and then in the WF
        /// it would call Update to set parent field to new parent value. This doesn't work for i.e. cloning active Quote and related Quote products.
        /// </summary>
        /// <param name="entityName"></param>
        /// <param name="objectId"></param>
        /// <param name="fieldstoIgnore"></param>
        /// <param name="prefix"></param>
        /// <param name="parentEntityId">Used in clone child proccess</param>
        /// <param name="cloneChildRecord">True indicates that child record is cloned.</param>
        /// <param name="parentFieldName">Used in clone child proccess</param>
        /// <param name="parentEntityName">Used in clone child proccess</param>
        /// <returns></returns>
        public Guid CloneRecord(string entityName, string objectId, string fieldstoIgnore, string prefix,
            Guid parentEntityId, bool cloneChildRecord = false, string parentFieldName = "", string parentEntityName = "")
        {
            tracingService.Trace("entering CloneRecord");
            if (fieldstoIgnore == null) fieldstoIgnore = "";
            fieldstoIgnore = fieldstoIgnore.ToLower();
            tracingService.Trace("fieldstoIgnore="+ fieldstoIgnore);
            Entity retrievedObject = service.Retrieve(entityName, new Guid(objectId), new ColumnSet(allColumns: true));
            tracingService.Trace("retrieved object OK");

            Entity newEntity = new Entity(entityName);
            string PrimaryIdAttribute = "";
            string PrimaryNameAttribute = "";
            List<string> atts = getEntityAttributesToClone(entityName, service, ref PrimaryIdAttribute, ref PrimaryNameAttribute);



            foreach (string att in atts)
            {
                if (fieldstoIgnore != null && fieldstoIgnore != "")
                {
                    if (Array.IndexOf(fieldstoIgnore.Split(';'), att) >= 0 || Array.IndexOf(fieldstoIgnore.Split(','), att) >= 0)
                    {
                        continue;
                    }
                }

                // Ignore parent entity field when cloning child - it will be set later
                if(cloneChildRecord && att == parentFieldName)
                {
                    tracingService.Trace("CloneChild process - skip parent field in this loop:{0}", att);
                    continue;
                }


                if (retrievedObject.Attributes.Contains(att) && att != "statuscode" && att != "statecode"
                    || att.StartsWith("partylist-"))
                {
                    if (att.StartsWith("partylist-"))
                    {
                        string att2 = att.Replace("partylist-", "");

                        string fetchParty = @"<fetch version='1.0' output-format='xml - platform' mapping='logical' distinct='true'>
                                                <entity name='activityparty'>
                                                    <attribute name = 'partyid'/>
                                                        <filter type = 'and' >
                                                            <condition attribute = 'activityid' operator= 'eq' value = '" + objectId + @"' />
                                                            <condition attribute = 'participationtypemask' operator= 'eq' value = '" + getParticipation(att2) + @"' />
                                                         </filter>
                                                </entity>
                                            </fetch> ";

                        RetrieveMultipleRequest fetchRequest1 = new RetrieveMultipleRequest
                        {
                            Query = new FetchExpression(fetchParty)
                        };
                        tracingService.Trace(fetchParty);
                        EntityCollection returnCollection = ((RetrieveMultipleResponse)service.Execute(fetchRequest1)).EntityCollection;


                        EntityCollection arrPartiesNew = new EntityCollection();
                        tracingService.Trace("attribute:{0}", att2);

                        foreach (Entity ent in returnCollection.Entities)
                        {
                            Entity party = new Entity("activityparty");
                            EntityReference partyid = (EntityReference)ent.Attributes["partyid"];


                            party.Attributes.Add("partyid", new EntityReference(partyid.LogicalName, partyid.Id));
                            tracingService.Trace("attribute:{0}:{1}:{2}", att2, partyid.LogicalName, partyid.Id.ToString());
                            arrPartiesNew.Entities.Add(party);
                        }

                        newEntity.Attributes.Add(att2, arrPartiesNew);
                        continue;

                    }

                    tracingService.Trace("attribute:{0}", att);
                    if (att == PrimaryNameAttribute && prefix != null)
                    {
                        retrievedObject.Attributes[att] = prefix + retrievedObject.Attributes[att];
                    }
                    newEntity.Attributes.Add(att, retrievedObject.Attributes[att]);
                }
            }

            // Ignore parent entity field when cloning child - it will be set later
            if (cloneChildRecord)
            {
                tracingService.Trace($"CloneChild process - set parent field: Name:{parentFieldName} Id:{parentEntityId}");
                newEntity.Attributes.Add(parentFieldName, new EntityReference(parentEntityName, parentEntityId));
            }

            tracingService.Trace("creating cloned object...");
            Guid createdGUID = service.Create(newEntity);
            tracingService.Trace($"created cloned object (id={createdGUID}) OK");

            if (newEntity.Attributes.Contains("statuscode") && newEntity.Attributes.Contains("statecode"))
            {
                Entity record = service.Retrieve(entityName, createdGUID, new ColumnSet("statuscode", "statecode"));


                if (retrievedObject.Attributes["statuscode"] != record.Attributes["statuscode"] ||
                    retrievedObject.Attributes["statecode"] != record.Attributes["statecode"])
                {
                    tracingService.Trace("set statuscode and statecode for cloned object same as original object...");

                    Entity setStatusEnt = new Entity(entityName, createdGUID);
                    setStatusEnt.Attributes.Add("statuscode", retrievedObject.Attributes["statuscode"]);
                    setStatusEnt.Attributes.Add("statecode", retrievedObject.Attributes["statecode"]);

                    service.Update(setStatusEnt);

                    tracingService.Trace("set statuscode and statecode for cloned object same as original object OK");
                }
            }

            tracingService.Trace($"cloned object id={createdGUID} OK");
            return createdGUID;
        }

        protected string getParticipation(string attributeName)
        {
            string sReturn = "";
            switch (attributeName)
            {
                case "from":
                    sReturn = "1";
                    break;
                case "to":
                    sReturn = "2";
                    break;
                case "cc":
                    sReturn = "3";
                    break;
                case "bcc":
                    sReturn = "4";
                    break;

                case "organizer":
                    sReturn = "7";
                    break;
                case "requiredattendees":
                    sReturn = "5";
                    break;
                case "optionalattendees":
                    sReturn = "6";
                    break;
                case "customer":
                    sReturn = "11";
                    break;
                case "resources":
                    sReturn = "10";
                    break;
            }
            return sReturn;
            /*Sender  1
                Specifies the sender.

                ToRecipient
                2
                Specifies the recipient in the To field.

                CCRecipient
                3
                Specifies the recipient in the Cc field.

                BccRecipient
                4
                Specifies the recipient in the Bcc field.

                RequiredAttendee
                5
                Specifies a required attendee.

                OptionalAttendee
                6
                Specifies an optional attendee.

                Organizer
                7
                Specifies the activity organizer.

                Regarding
                8
                Specifies the regarding item.

                Owner
                9
                Specifies the activity owner.

                Resource
                10
                Specifies a resource.

                Customer
                11

            */
        }



    }
}
