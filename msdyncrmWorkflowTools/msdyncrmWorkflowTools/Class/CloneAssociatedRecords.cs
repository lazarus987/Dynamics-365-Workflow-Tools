using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace msdyncrmWorkflowTools.Class
{
    /// <summary>
    /// Clones Associated records (N:N)
    /// 
    /// 1. Takes all the associated record from "Source Record URL" using "Relationship Name (N:N entity)"
    /// 2. Creates new associations to record from "Source Record URL"
    ///     
    /// </summary>
    public class CloneAssociatedRecords : CodeActivity
    {
        #region "Parameter Definition"

        [RequiredArgument]
        [Input("Source Record URL")]
        [ReferenceTarget("")]
        public InArgument<String> SourceRecordUrl { get; set; }

        [RequiredArgument]
        [Input("Target Record ID")]
        [ReferenceTarget("")]
        public InArgument<String> TargetRecordId { get; set; }

        [RequiredArgument]
        [Input("Relationship Name")]
        [ReferenceTarget("")]
        public InArgument<String> RelationshipName { get; set; }

        [RequiredArgument]
        [Input("Source Field Id Name")]
        [ReferenceTarget("")]
        public InArgument<String> SourceEntityFieldIdName { get; set; }

        [RequiredArgument]
        [Input("Related Field Id Name")]
        [ReferenceTarget("")]
        public InArgument<String> RelatedEntityFieldIdName { get; set; }

        [RequiredArgument]
        [Input("Related Entity name")]
        [ReferenceTarget("")]
        public InArgument<String> RelatedEntityName { get; set; }

        #endregion

        protected override void Execute(CodeActivityContext executionContext)
        {
            #region "Load CRM Service from context"

            Common objCommon = new Common(executionContext);
            objCommon.tracingService.Trace("Load CRM Service from context --- OK");
            #endregion

            #region "Read Parameters"

            String _relationshipName = this.RelationshipName.Get(executionContext);
            if (_relationshipName == null || _relationshipName == "")
            {
                throw new ArgumentNullException("RelationshipName");
            }

            String _sourceEntityFieldIdName = this.SourceEntityFieldIdName.Get(executionContext);
            if (_sourceEntityFieldIdName == null || _sourceEntityFieldIdName == "")
            {
                throw new ArgumentNullException("SourceEntityFieldIdName");
            }

            String _relatedEntityFieldIdName = this.RelatedEntityFieldIdName.Get(executionContext);
            if (_relatedEntityFieldIdName == null || _relatedEntityFieldIdName == "")
            {
                throw new ArgumentNullException("RelatedEntityFieldIdName");
            }

            String _relatedEntityName = this.RelatedEntityName.Get(executionContext);
            if (_relatedEntityName == null || _relatedEntityName == "")
            {
                throw new ArgumentNullException("RelatedEntityName");
            }

            String _source = this.SourceRecordUrl.Get(executionContext);
            if (_source == null || _source == "")
            {
                throw new ArgumentNullException("SourceRecordUrl");
            }

            String _targetRecordId = this.TargetRecordId.Get(executionContext);
            if (_targetRecordId == null || _targetRecordId == "")
            {
                throw new ArgumentNullException("TargetRecordId");
            }

            string[] urlParts = _source.Split("?".ToArray());
            string[] urlParams = urlParts[1].Split("&".ToCharArray());
            string parentObjectTypeCode = urlParams[0].Replace("etc=", "");
            string parentEntityName = objCommon.sGetEntityNameFromCode(parentObjectTypeCode, objCommon.service);
            string parentId = urlParams[1].Replace("id=", "");
            objCommon.tracingService.Trace("ObjectTypeCode=" + parentObjectTypeCode + "--ParentId=" + parentId);

            #endregion

            #region "Clone Execution"

            // Get N:N records base on input parameters:
            // - relationship name
            // - SourceRelatedEntityFieldIdName

            List<Entity> entitiesForNewAssociation = GetAssociatedRecords(objCommon, _relationshipName,
                _sourceEntityFieldIdName, sourceId: Guid.Parse(parentId), _relatedEntityFieldIdName);

            foreach(Entity e in entitiesForNewAssociation)
            {
                // Associate entity with Target Record
                objCommon.tracingService.Trace($"Start associate target(cloned) record with [Entity={_relatedEntityName}; ID={e.GetAttributeValue<Guid>(_relatedEntityFieldIdName)}]");
                
                AssociateEntitiesRequest fooToBar = new AssociateEntitiesRequest
                {
                    Moniker1 = new EntityReference(parentEntityName, Guid.Parse(_targetRecordId)), // target entity
                    Moniker2 = new EntityReference(_relatedEntityName, e.GetAttributeValue<Guid>(_relatedEntityFieldIdName)), // related entity
                    RelationshipName = _relationshipName,  // name of the relationship
                };

                objCommon.service.Execute(fooToBar);
                
                objCommon.tracingService.Trace($"Associated source record with [Entity={_relatedEntityName}; ID={e.GetAttributeValue<Guid>(_relatedEntityFieldIdName)}] - DONE");
            }

            objCommon.tracingService.Trace("cloned object OK");

            #endregion

        }

        /// <summary>
        /// Returns list of N:N related records.
        /// </summary>
        /// <param name="objCommon"></param>
        /// <param name="relationshipName">N:N Relationship name</param>
        /// <param name="sourceIdName">Source ID - based on which we are filtering</param>
        /// <param name="sourceId">id</param>
        /// <returns></returns>
        private List<Entity> GetAssociatedRecords(Common objCommon, string relationshipName, 
            string sourceIdName, Guid sourceId, string relatedEntityIdName)
        {
            objCommon.tracingService.Trace($"GetAssociatedRecords(relationshipName={relationshipName}, sourceIdName={sourceIdName}, " +
                $"sourceId={sourceId}, relatedEntityIdName={relatedEntityIdName}) called...");

            var fetchXml = $@"
                <fetch>
                  <entity name='{relationshipName}'>
                    <attribute name='{relatedEntityIdName}' />
                    <filter type='and'>
                      <condition attribute='{sourceIdName}' operator='eq' value='{sourceId}'/>
                    </filter>
                  </entity>
                </fetch>";

            EntityCollection ec = objCommon.service.RetrieveMultiple(new FetchExpression(fetchXml));

            objCommon.tracingService.Trace($"GetAssociatedRecords returns {ec.Entities.Count} records...");

            return ec.Entities.ToList();
        }
    }
}
