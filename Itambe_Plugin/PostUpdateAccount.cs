using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Itambe_Plugin
{
    public class PostUpdateAccount : Plugin
    {
        private IOrganizationService service;

        public PostUpdateAccount() : base(typeof(PostUpdateAccount))
        {
            RegisteredEvents.Add(new Tuple<Int32, String, String, Action<LocalPluginContext>>(40, "Update", "account", Execute));
        }
        protected void Execute(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new ArgumentNullException("localContext");
            }
            IPluginExecutionContext context = localContext.PluginExecutionContext;
            if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity)) return;

            service = localContext.OrganizationService;
            OrganizationServiceContext ctx = new OrganizationServiceContext(service);

            if (context.Depth > 1)
            {
                return;
            }

            try
            {
                Entity entidadeTarget = localContext.PluginExecutionContext.InputParameters["Target"] as Entity;

                int total = entidadeTarget.Contains("results_diassemfaturamento") ?
                                           entidadeTarget.GetAttributeValue<int>("results_diassemfaturamento") :
                                           0;
                if (total == 30 || total == 45 || total == 60 || total == 83 || total == 90)
                {
                    ColumnSet attributes = new ColumnSet(new string[] { "results_diassemfaturamento", "results_qtdedias" });
                    Entity empresa = service.Retrieve("account", entidadeTarget.Id, attributes);
                    empresa["results_qtdedias"] = total;

                    service.Update(empresa);
                }
            }
            catch (Exception ex) { throw new InvalidPluginExecutionException(ex.Message); }
        }
    }
}
