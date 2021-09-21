using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

namespace msdyncrmWorkflowTools.Class
{
    public class GetEntityObjectTypeCode : CodeActivity
    {
        #region "Parameter Definition"

        [RequiredArgument]
        [Input("Entity name")]        
        public InArgument<String> EntityName { get; set; }

        [Output("Entity object type code")]
        public OutArgument<int> EntityObjectTypeCode { get; set; }

        #endregion

        protected override void Execute(CodeActivityContext executionContext)
        {
            #region "Load CRM Service from context"
            Common objCommon = new Common(executionContext);
            objCommon.tracingService.Trace("Load CRM Service from context --- OK");
            #endregion

            string entityName = this.EntityName.Get(executionContext);

            int entityObjectCode = objCommon.sGetEntityCodeForName(entityName, objCommon.service);
            objCommon.tracingService.Trace("ObjectTypeCode=" + entityObjectCode + "--EntitySchemaName=" + entityName);

            EntityObjectTypeCode.Set(executionContext, entityObjectCode);

            objCommon.tracingService.Trace("Retrieval of object type code OK");

        }

    }
}
