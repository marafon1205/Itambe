
using Integracao.Objetos;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Integracao
{
     static class Insercoes
    {
        public static void InsereGrupoProdutos(GrupoDeProdutos grp)
        {
            if (Consulta.VerificaExistencia("results_grupodoproduto", "results_codigo", grp.results_codigo.ToString()))
            {
                try
                {
                    Entity entidade = new Entity("results_grupodoproduto");
                    entidade["results_codigo"] = grp.results_codigo.ToString();
                    entidade["results_name"] = grp.results_name.ToString();
                   // entidade["results_integrado"] = true;
                    Consulta.service.Create(entidade);
                    InsereLog("IntegracaoGrupoProdutos", true, "Grupo Produto", String.Format("Código do Produto {0} ", grp.results_codigo.ToString()));
                }
                catch (Exception e)
                {
                    InsereLog("IntegracaoGrupoProdutos", false, "Grupo Produto", e.Message);
                }
            }
        }

        public static void InsereProdutos(Produtos produto)
        {
            bool naoExiste = Consulta.VerificaExistencia("product", "productnumber", produto.productnumber.ToString());

            try
            {
                String[] columnset = new string[]{ "name","productnumber", "results_grupodoprodutoid", "results_embalagem", "defaultuomid", "parentproductid", "pricelevelid" };

                EntityCollection lista = Consulta.RetornaColecao("product",  columnset, "productnumber", produto.productnumber.ToString());
                Entity entidade = naoExiste? new Entity("product"):new Entity("product", lista[0].Id);
                if (lista[0].GetAttributeValue<string>("name") != null ? !lista[0].GetAttributeValue<string>("name").Equals(produto.name) : true)
                    entidade["name"] = produto.name.ToString();
                if (naoExiste)
                    entidade["productnumber"] = produto.productnumber.ToString();

                if (lista[0].GetAttributeValue<string>("results_embalagem") != null ? !lista[0].GetAttributeValue<string>("results_embalagem").Equals(produto.results_embalagem) : true)
                   entidade["results_embalagem"] = produto.results_embalagem;
                EntityReference grupo = Consulta.Busca("results_grupodoproduto", "results_codigo", produto.results_grupodoprodutoid.ToString());
               
                if (lista[0].Contains("results_grupodoprodutoid") ? !lista[0].GetAttributeValue<EntityReference>("results_grupodoprodutoid").Id.Equals(grupo.Id) : true)
                   entidade["results_grupodoprodutoid"] = grupo;

                entidade["quantitydecimal"] = 2;

                entidade["defaultuomid"] = Consulta.Busca("uom", "name", produto.defaultuomid.ToString()); // Unidade Medida
                entidade["defaultuomscheduleid"] = Consulta.Busca("uomschedule", "name", produto.defaultuomscheduleid.ToString()); // grupo de unidades Unidade Padrão
                entidade["statecode"] = Comum.Status(produto.statecode);
                if (Consulta.VerificaExistencia("product", "name", produto.parentproductid.ToString(), 2))
                    entidade["parentproductid"] = Consulta.Busca("product", "name", produto.parentproductid.ToString());

                entidade["pricelevelid"] = Consulta.RetornaListaPrecoPadrao("");
                // entidade["results_dataintegracao"] = DateTime.Now;
                if (naoExiste)
                {
                    
                   
                    Guid product =  Consulta.service.Create(entidade);

                   bool existe =  Consulta.VerificaExistencia("productpricelevel", "productid", product.ToString());
                    if (existe)
                    {
                        // Verifica se o produto está na tabela de item da tabela de preço
                        Entity itemListaPreco = new Entity("productpricelevel");
                        itemListaPreco["pricelevelid"] = new EntityReference("pricelevel", Consulta.RetornaListaPrecoPadrao("").Id);
                        itemListaPreco["productid"] = new EntityReference("product", product);
                        itemListaPreco["amount"] = new Money(1);
                        itemListaPreco["uomid"] = Consulta.Busca("uom", "name", produto.defaultuomid.ToString());
                        Consulta.service.Create(itemListaPreco);

                        InsereLog("IntegracaoProdutos", true, "Item da lista de Preço", String.Format("Código do Produto {0} ", produto.productnumber.ToString()));
                    }

                    // Faz o vinculo produto tabela de preço
                    EntityReference item = new EntityReference("pricelevel", Consulta.RetornaListaPrecoPadrao("").Id);

                    // Cria a EntityReferenceCollection para Linha do Contrato
                    EntityReferenceCollection relatedEntities = new EntityReferenceCollection();

                    // Adiciona o relacionamento com a entidade
                    relatedEntities.Add(item);

                    // Adiciona o nome do relacionamento entra as entidades
                    Relationship relationship = new Relationship("price_level_products");

                    // Associate the contact record to Account
                    Consulta.service.Associate("product", product, relationship, relatedEntities);
                                      
                    
                }
                else Consulta.service.Update(entidade);

                InsereLog("IntegracaoProdutos", true, "Produto", String.Format("Código do Produto {0} ", produto.productnumber.ToString()));
            }
            catch (Exception e)
            {
                InsereLog("IntegracaoProdutos", false, "Produto", e.Message);
            }

        }

        public static void InsereClientes(Cliente cliente)
        {
            try
            {
                if (Consulta.VerificaExistencia("account", String.IsNullOrEmpty(cliente.accountnumber.ToString()) ? "mps_cod_erp":  "accountnumber" ,  String.IsNullOrEmpty(cliente.accountnumber.ToString()) ?  cliente.mps_cod_erp.ToString() :  Comum.FormatCNPJ(cliente.accountnumber.ToString())))
                {

                    Entity entidade = new Entity("account");

                    entidade["accountnumber"] = Comum.FormatCNPJ(cliente.accountnumber.ToString());
                    entidade["name"] = cliente.name;
                    entidade["mps_fantasia"] = cliente.mps_fantasia;
                    entidade["mps_inscricaoestadual"] = cliente.mps_inscricaoestadual;
                    entidade["emailaddress1"] = cliente.emailaddress1;
                    // entidade["mps_tipologradouro"] = ;
                    entidade["mps_lkppais"] = Consulta.Busca("mps_pais", "mps_name", "BRASIL");
                    entidade["mps_lkpuf"] = Consulta.Busca("mps_uf", "mps_name", cliente.mps_lkpuf.ToString());
                    entidade["address1_city"] = cliente.address1_city.ToString();
                    entidade["address1_line1"] = cliente.address1_line1;
                    entidade["address1_line2"] = cliente.address1_line2;
                    entidade["address1_line3"] = cliente.address1_line3;
                    entidade["address1_postalcode"] = cliente.address1_postalcode;
                    entidade["mps_cod_erp"] = cliente.mps_cod_erp;
                    entidade["mps_atendente"] = Consulta.BuscaUsuario("systemuser", "fullname", cliente.mps_atendente);
                    entidade["ownerid"] = Consulta.BuscaUsuario("systemuser", "fullname", cliente.ownerid.ToString());
                    entidade["mps_ramos"] = Consulta.Busca("mps_ramo", "mps_descricao", cliente.mps_ramos.ToString());
                    entidade["telephone1"] = cliente.telephone1;
                    entidade["telephone2"] = cliente.telephone2;
                    entidade["mps_complemento"] = cliente.mps_complemento;
                    // entidade["mps_tipodeorigemdoprospect"] = Comum.Status(produto.statecode);
                    // entidade["mps_listtipoderelacao"] =;
                    // entidade["results_integrado"] = true;
                    // entidade["results_dataintegracao"] = DateTime.Now;
                    Consulta.service.Create(entidade);
                    // else Consulta.service.Update(entidade);

                    InsereLog("IntegracaoClientes", true, "Empresa", String.Format("Cliente com o CNPJ {0} foi adicionado!", Comum.FormatCNPJ(cliente.accountnumber.ToString())));
                }
            }
            catch (Exception e)
            {

                InsereLog("IntegracaoClientes", false, "Empresa", e.Message);
            }
        }

        public static void InsereFaturamentos(Faturamento faturamento)
        {
            string mensagem = String.Empty;

            try
            {
                if (Consulta.VerificaExistencia("invoice", "invoicenumber", faturamento.invoicenumber.ToString()))
                {
                    mensagem = String.Format("Fatura {0} -  Cliente {1}", faturamento.invoicenumber, String.IsNullOrEmpty(faturamento.customerid.ToString()) ? faturamento.codigoerp.ToString(): faturamento.customerid.ToString());
                    Entity entidade = new Entity("invoice");
                    entidade["customerid"] = String.IsNullOrEmpty(faturamento.customerid.ToString())? Consulta.Busca("account", "mps_cod_erp", faturamento.codigoerp.ToString()) : Consulta.Busca("account", "accountnumber", Comum.FormatCNPJ(faturamento.customerid.ToString()));
                    entidade["invoicenumber"] = faturamento.invoicenumber;
                    entidade["pricelevelid"] = Consulta.RetornaListaPrecoPadrao(faturamento.pricelevelid);
                    entidade["ispricelocked"] =  false; 
                    entidade["name"] = faturamento.name;
                   
                    entidade["description"] = faturamento.description;
                    entidade["results_numeropedido"] = faturamento.results_numeropedido;
                    if (!String.IsNullOrEmpty(faturamento.ownerid.ToString()))
                    {
                        EntityReference userReference = Consulta.Busca("systemuser", "fullname", faturamento.ownerid.ToString());
                        if(userReference != null)
                            entidade["ownerid"] = userReference;
                        else
                        {
                            EntityReference teamReference = Consulta.Busca("team", "name", faturamento.ownerid.ToString());
                            entidade["ownerid"] = teamReference;
                        }
                    }
                    if (!String.IsNullOrEmpty(faturamento.results_dataemissao.ToString()))
                    {
                        int ano = Convert.ToInt16(faturamento.results_dataemissao.ToString().Substring(0, 4));
                        int mes = Convert.ToInt16(faturamento.results_dataemissao.ToString().Substring(4, 2));
                        int dia = Convert.ToInt16(faturamento.results_dataemissao.ToString().Substring(6));

                        entidade["results_dataemissao"] = new DateTime(ano, mes, dia);
                    }
                    if (!String.IsNullOrEmpty(faturamento.datedelivered.ToString()))
                    {
                        int ano = Convert.ToInt16(faturamento.datedelivered.ToString().Substring(0, 4));
                        int mes = Convert.ToInt16(faturamento.datedelivered.ToString().Substring(4, 2));
                        int dia = Convert.ToInt16(faturamento.datedelivered.ToString().Substring(6));
                        entidade["datedelivered"] = new DateTime(ano, mes, dia);
                    }

                    entidade["billto_city"] = faturamento.billto_city;
                    entidade["billto_line1"] = faturamento.billto_line1;
                    entidade["billto_line2"] = faturamento.billto_line2;
                    entidade["billto_line3"] = faturamento.billto_line3;
                    entidade["billto_postalcode"] = faturamento.billto_postalcode;
                    entidade["billto_stateorprovince"] = faturamento.billto_stateorprovince.ToString();
                    entidade["billto_telephone"] = faturamento.billto_telephone;
                    entidade["results_inscricaoestadual"] = faturamento.results_inscricaoestadual;
                    entidade["results_nometransportadora"] = faturamento.results_nometransportadora;
                    entidade["shipto_city"] = faturamento.shipto_city;
                    entidade["shipto_line1"] = faturamento.shipto_line1;
                    entidade["shipto_line2"] = faturamento.shipto_line2;
                    entidade["shipto_line3"] = faturamento.shipto_line3;
                    entidade["shipto_postalcode"] = faturamento.shipto_postalcode;
                    entidade["shipto_stateorprovince"] = faturamento.shipto_stateorprovince;
                    entidade["shipto_telephone"] = faturamento.shipto_telephone;
                    if (!String.IsNullOrEmpty(faturamento.shippingmethodcode.ToString()))
                      entidade["shippingmethodcode"] = Comum.TipoTransporte(faturamento.shippingmethodcode);

                    // Transportadora
                    entidade["results_inscestadualtransportadora"] = faturamento.results_inscestadualtransportadora;
                    entidade["results_cnpjcpf"] = faturamento.results_cnpjcpf;
                    entidade["results_placaveiculo"] = faturamento.results_placaveiculo;
                    entidade["results_codigoantt"] = faturamento.results_codigoantt;
                    entidade["results_nomemotorista"] = faturamento.results_nomemotorista;
                    entidade["results_enderecoentrega"] = faturamento.results_enderecoentrega;

                    if (!String.IsNullOrEmpty(faturamento.totalamount.ToString()))
                        entidade["totalamount"] = new Money(Convert.ToDecimal(faturamento.totalamount));
                    if (!String.IsNullOrEmpty(faturamento.results_totalprodutos.ToString()))
                        entidade["results_totalprodutos"] = new Money(Convert.ToDecimal(faturamento.results_totalprodutos));
                    if (!String.IsNullOrEmpty(faturamento.freightamount.ToString()))
                        entidade["freightamount"] = new Money(Convert.ToDecimal(faturamento.freightamount));
                    if (!String.IsNullOrEmpty(faturamento.results_valorseguro.ToString()))
                        entidade["results_valorseguro"] = new Money(Convert.ToDecimal(faturamento.results_valorseguro));
                    if (!String.IsNullOrEmpty(faturamento.totaldiscountamount.ToString()))
                        entidade["totaldiscountamount"] = new Money(Convert.ToDecimal(faturamento.totaldiscountamount));
                    if (!String.IsNullOrEmpty(faturamento.results_valoricms.ToString()))
                        entidade["results_valoricms"] = new Money(Convert.ToDecimal(faturamento.results_valoricms));
                    if (!String.IsNullOrEmpty(faturamento.results_quantidade.ToString()))
                        entidade["results_quantidade"] = Convert.ToDecimal(faturamento.results_quantidade);
                    
                    Guid invoiceid=  Consulta.service.Create(entidade);
                    InsereLog("IntegracaoFaturamentos", true, "Faturamento", mensagem);


                    try
                    {
                        Entity produtoFatura = new Entity("invoicedetail");
                        produtoFatura["invoiceid"] = new EntityReference("invoice", invoiceid);
                        produtoFatura["ispriceoverridden"] = true;

                        produtoFatura["results_notafiscal"] = faturamento.results_notafiscal;
                        produtoFatura["productid"] = Consulta.Busca("product", "productnumber", faturamento.productnumber.ToString());
                        if (!String.IsNullOrEmpty(faturamento.quantity.ToString()))
                            produtoFatura["quantity"] = Convert.ToDecimal(faturamento.quantity.ToString().Replace(".", ",")); //.Replace(",", "").Replace(".", ","));
                        if (!String.IsNullOrEmpty(faturamento.priceperunit.ToString()))
                            produtoFatura["priceperunit"] = new Money(Convert.ToDecimal(faturamento.priceperunit.ToString().Replace(".", ","))); //.Replace(".", ",")));
                        if (!String.IsNullOrEmpty(faturamento.results_valoripi.ToString()))
                            produtoFatura["results_valoripi"] = Convert.ToDecimal(faturamento.results_valoripi);
                        if (!String.IsNullOrEmpty(faturamento.results_valoricms.ToString()))
                            produtoFatura["results_valoricms"] = Convert.ToDecimal(faturamento.results_valoricms);
                        if (!String.IsNullOrEmpty(faturamento.extendedamount.ToString()))
                            produtoFatura["extendedamount"] = Convert.ToDecimal(faturamento.extendedamount);
                        if (!String.IsNullOrEmpty(faturamento.tax.ToString()))
                            produtoFatura["tax"] = new Money(Convert.ToDecimal(faturamento.tax.ToString().Replace(".", ",")));
                        produtoFatura["results_ncm"] = faturamento.results_ncm;
                        produtoFatura["results_cst"] = faturamento.Results_cst;
                        produtoFatura["results_cfop"] = faturamento.results_cfop;
                        produtoFatura["uomid"] = Consulta.Busca("uom", "name", faturamento.uomi.ToString());
                        produtoFatura["salesrepid"] = Consulta.Busca("systemuser", "fullname", faturamento.ownerid.ToString());
                        Consulta.service.Create(produtoFatura);
                        InsereLog("IntegracaoFaturamentosProduto", true, "Faturamento Produto", String.Format("ID da Fatura {0} - Produto {1} ", faturamento.invoicenumber.ToString(), faturamento.productnumber));


                      /*  Entity fatura = new Entity("invoice", invoiceid);

                         if (!String.IsNullOrEmpty(faturamento.totalamount.ToString()))
                            fatura["totalamount"] = Convert.ToDecimal(faturamento.totalamount.ToString());
                        if (!String.IsNullOrEmpty(faturamento.results_totalprodutos.ToString()))
                            fatura["results_totalprodutos"] = Convert.ToDecimal(faturamento.results_totalprodutos.ToString());
                        if (!String.IsNullOrEmpty(faturamento.freightamount.ToString()))
                            fatura["freightamount"] = Convert.ToDecimal(faturamento.freightamount);
                        if (!String.IsNullOrEmpty(faturamento.results_valorseguro.ToString()))
                            fatura["results_valorseguro"] = Convert.ToDecimal(faturamento.results_valorseguro);
                        if (!String.IsNullOrEmpty(faturamento.totaldiscountamount.ToString()))
                            fatura["totaldiscountamount"] = Convert.ToDecimal(faturamento.totaldiscountamount);
                        if (!String.IsNullOrEmpty(faturamento.results_valoricms.ToString()))
                            fatura["results_valoricms"] = Convert.ToDecimal(faturamento.results_valoricms);
                        if (!String.IsNullOrEmpty(faturamento.results_quantidade.ToString()))
                            fatura["results_quantidade"] = Convert.ToDecimal(faturamento.results_quantidade);*/


                      //  Consulta.service.Update(fatura);
                    }
                    catch (Exception e)
                    {

                        InsereLog("IntegracaoFaturamentosProduto", false, "Faturamento Produto", mensagem +": "+ e.Message);
                    }
                }
            }
            catch (Exception e)
            {

                InsereLog("IntegracaoFaturamentos", false, "Faturamento", mensagem + ": " +  e.Message);
            }
        }
        public static void InsereLog(string mensagem, bool resultado, string entidade, string descricao)
        {
            Entity log = new Entity("results_logintegracao");

            log["results_descricao"] = descricao.Length > 5000 ? descricao.Substring(0, 4999) : descricao;

            log["results_resultado"] = resultado;
            log["results_entidade"] = entidade;
            log["results_name"] = mensagem + "-Hora:" + DateTime.Now.Hour;
            Consulta.service.Create(log);
        }

        public static void AtualizaDevolucao(Faturamento faturamento)
        {
            string mensagem = string.Empty;
            try
            {
                EntityCollection lista = Consulta.VerificaExistenciaDevolucao("invoice", "invoicenumber", faturamento.invoicenumber.ToString());
                mensagem = string.Format("Fatura {0} -  Cliente {1}", faturamento.invoicenumber, (string.IsNullOrEmpty(faturamento.customerid.ToString()) ? faturamento.codigoerp.ToString() : faturamento.customerid.ToString()));
                Entity entidade = new Entity("invoice", lista[0].Id);
                EntityCollection lista1 = Consulta.VerificarProdutoFatura(entidade);
                if (!string.IsNullOrEmpty(faturamento.ownerid.ToString()))
                {
                    EntityReference teamReference = Consulta.Busca("team", "name", faturamento.ownerid.ToString());
                    entidade["ownerid"] = teamReference;
                }
                Consulta.service.Update(entidade);
                Insercoes.InsereLog("IntegracaoFaturamentos", true, "Faturamento", mensagem);
                Guid invoiceid = lista1[0].Id;
                try
                {
                    Entity produtoFatura = new Entity("invoicedetail", invoiceid);
                    produtoFatura["ispriceoverridden"] = true;
                    if (!string.IsNullOrEmpty(faturamento.quantity.ToString()))
                    {
                        produtoFatura["quantity"] = Convert.ToDecimal(faturamento.quantity.ToString());
                    }
                    Consulta.service.Update(produtoFatura);
                    Insercoes.InsereLog("IntegracaoFaturamentosProduto", true, "Faturamento Produto", string.Format("ID da Fatura {0} - Produto {1} ", faturamento.invoicenumber.ToString(), faturamento.productnumber));
                }
                catch (Exception exception)
                {
                    Exception e = exception;
                    Insercoes.InsereLog("IntegracaoFaturamentosProduto", false, "Faturamento Produto", string.Concat(mensagem, ": ", e.Message));
                }
            }
            catch (Exception exception1)
            {
                Exception e = exception1;
                Insercoes.InsereLog("IntegracaoFaturamentos", false, "Faturamento", string.Concat(mensagem, ": ", e.Message));
            }
        }
    }
}
