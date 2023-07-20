using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Runtime.Serialization;
using System;
using System.Configuration;
using System.Net;
using System.ServiceModel.Description;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;

namespace InativarClientes
{
    public static class Util
    {
        public static IOrganizationService service = Conexao();

        public static EntityCollection checkOneCondition(string _entityName, string _attributeName, string _attributeValue, IOrganizationService _service)
        {
            ConditionExpression condition = new ConditionExpression();
            condition.AttributeName = _attributeName;
            condition.Operator = ConditionOperator.Equal;
            condition.Values.Add(_attributeValue);

            FilterExpression filter = new FilterExpression();
            filter.FilterOperator = LogicalOperator.And;
            filter.Conditions.Add(condition);

            QueryExpression query = new QueryExpression();
            query.EntityName = _entityName;
            query.ColumnSet = new ColumnSet(true);
            query.Criteria.AddFilter(filter);

            return _service.RetrieveMultiple(query);
        }

        public static IOrganizationService Conexao()
        {
            ClientCredentials Credentials = new ClientCredentials();
            Credentials.UserName.UserName = ConfigurationManager.AppSettings["usrCRM"].ToString();
            Credentials.UserName.Password = ConfigurationManager.AppSettings["pwdCRM"].ToString();
            string crm = ConfigurationManager.AppSettings["crmUrl"].ToString();
            Uri OrganizationUrl = new Uri(crm);
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUrl, null, Credentials, null))
            {
                serviceProxy.EnableProxyTypes();
                service = (IOrganizationService)serviceProxy;
            }
            return service;
        }

        public static EntityCollection BuscaCliente()
        {
            QueryExpression qry = new QueryExpression();
            qry.EntityName = "account";
            qry.ColumnSet = new ColumnSet("name");

            EntityCollection lista = RetrieveMultiplePlus(qry);
            return lista;
        }

        public static void SetState(this Entity entity, IOrganizationService service, int state, int status)
        {
            SetStateRequest setStateRequest = new SetStateRequest();
            setStateRequest.EntityMoniker = new EntityReference(entity.LogicalName, entity.Id);
            setStateRequest.State = new OptionSetValue(state);
            setStateRequest.Status = new OptionSetValue(status);
            service.Execute(setStateRequest);
        }

        public static EntityCollection RetrieveMultiplePlus(QueryExpression query)
        {
            EntityCollection collectionReturn;
            EntityCollection collectionAux;
            query.PageInfo.Count = 5000;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            collectionReturn = service.RetrieveMultiple(query);
            if (collectionReturn.MoreRecords)
            {
                do
                {
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = collectionReturn.PagingCookie;
                    collectionAux = service.RetrieveMultiple(query);
                    collectionReturn.Entities.AddRange(collectionAux.Entities);
                } while (collectionAux.MoreRecords);
            }
            return collectionReturn;
        }
    }
}