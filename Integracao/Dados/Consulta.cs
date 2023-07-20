using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

namespace Integracao
{
   public static class Consulta
    {
        public static IOrganizationService service = Conexao();

        public static EntityReference Busca(string entityName, string condition, string cargo)
        {
            QueryExpression query = new QueryExpression();
            query.EntityName = entityName;
            query.Criteria.AddCondition(condition, ConditionOperator.Equal, cargo);

            EntityCollection lista = service.RetrieveMultiple(query);
            return lista.Entities.Count > 0 ? new EntityReference(entityName, lista[0].Id) : null;
        }

        public static EntityReference BuscaUsuario(string entityName, string condition, object cargo)
        {
            string nome = cargo.ToString();

            while (nome.IndexOf("  ") >= 0)
                nome = nome.Replace("  ", " ");

            QueryExpression query = new QueryExpression();
            query.EntityName = entityName;
            query.Criteria.AddCondition(condition, ConditionOperator.Equal, nome);

            EntityCollection lista = service.RetrieveMultiple(query);
            return lista.Entities.Count > 0 ? new EntityReference(entityName, lista[0].Id) : null;
        }

        public static string Busca(string entityName, string[] columnset, string condition, string valor)
        {
            QueryExpression query = new QueryExpression();
            query.EntityName = entityName;
            query.Criteria.AddCondition(condition, ConditionOperator.Equal, valor);

            EntityCollection lista = service.RetrieveMultiple(query);
            return lista.Entities.Count > 0 ? lista[0].GetAttributeValue<string>(columnset[0]) : null;
        }

        public static EntityCollection RetornaColecao(string entityName, string[] columnset, string condition, string valor)
        {
            QueryExpression query = new QueryExpression();
            query.EntityName = entityName;
            //query.ColumnSet = new ColumnSet(columnset);
            query.Criteria.AddCondition(condition, ConditionOperator.Equal, valor);

            EntityCollection lista = service.RetrieveMultiple(query);

      
            return lista;
        }


        public static bool VerificaExistencia(string nomeEntidade, string nomeCondicao, string valorCondicao)
        {

            QueryExpression query = new QueryExpression();
            query.EntityName = nomeEntidade;
            query.Criteria.AddCondition(nomeCondicao, ConditionOperator.Equal, valorCondicao);

            EntityCollection lista = service.RetrieveMultiple(query);

            return lista.Entities.Count == 0;
        }

        public static bool VerificaExistencia(string nomeEntidade, string nomeCondicao, string valorCondicao, int? estrutura)
        {

            QueryExpression query = new QueryExpression();
            query.EntityName = nomeEntidade;
            query.Criteria.AddCondition(nomeCondicao, ConditionOperator.Equal, valorCondicao);
            query.Criteria.AddCondition("productstructure", ConditionOperator.Equal, estrutura);
            EntityCollection lista = service.RetrieveMultiple(query);

            return lista.Entities.Count > 0;
        }

        public static EntityReference RetornaListaPrecoPadrao(object listaPreco)
        {

            QueryExpression query = new QueryExpression();
            query.EntityName = "pricelevel";
            query.ColumnSet = new ColumnSet("pricelevelid");
            if (String.IsNullOrEmpty(listaPreco.ToString()))

                query.Criteria.AddCondition("name", ConditionOperator.Equal, "Padrão");
            else
                query.Criteria.AddCondition("name", ConditionOperator.Equal, listaPreco.ToString());
            EntityCollection lista = service.RetrieveMultiple(query);

            return new EntityReference("pricelevel", lista[0].Id);

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
    }
}
