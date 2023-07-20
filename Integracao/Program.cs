using Integracao.Objetos;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace Integracao
{
    static class Program
    {

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            //GrupoProdutos();

            //Produtos();

            //Clientes();

            //Faturamentos();

            //  AtualizarCodContato();

            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(@"C:\Importar\XML");
            foreach (System.IO.FileInfo f in di.GetFiles())
            {
                if (f.Extension.ToLower().Equals(".xlsx") || f.Extension.ToLower().Equals(".csv"))
                {
                    FileInfo newFile = new FileInfo(f.FullName);
                       if (f.FullName.Contains("clientes"))
                     {
                         ClientesExcel(newFile);
                     }

                    if (f.FullName.Contains("faturamento"))
                    {
                        Faturamentos(newFile);
                    }
                }
            }
        }

        static void Produtos()
        {
            try {
                string xmlData = @"C:\Importar\XML\produtos.xml";
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlData); //Carregando o arquivo*/

                //Pegando elemento pelo nome da TAG
                XmlNodeList xnList = xmlDoc.SelectNodes("produtos");
                foreach (XmlNode xn in xnList)
                {
                    foreach (XmlNode childNode in xn.ChildNodes)
                    {
                        Produtos prod = new Produtos();
                        prod.productnumber = childNode["b1_cod"].InnerText;
                        prod.name = childNode["b1_desc"].InnerText;
                        prod.results_grupodoprodutoid = childNode["b1_grupo"].InnerText;
                        prod.parentproductid =  childNode["b1_ativid"].InnerText;
                        prod.results_embalagem = childNode["b1_embalagem"].InnerText;
                        prod.defaultuomid = childNode["b1_um"].InnerText;
                        prod.defaultuomscheduleid = "Unidade Padrão";
                        prod.statecode = childNode["b1_msblql"].InnerText;

                        Insercoes.InsereProdutos(prod);
                    }
                }
                File.Delete(xmlData);
            }
            catch
            {

            }
        }

        static void GrupoProdutos()
        {
            try
            {
                string xmlData = @"C:\Importar\XML\grupoprodutos.xml";
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlData); //Carregando o arquivo*/

                //Pegando elemento pelo nome da TAG
                XmlNodeList xnList = xmlDoc.SelectNodes("grupoprodutos");
                foreach (XmlNode xn in xnList)
                {
                    foreach (XmlNode childNode in xn.ChildNodes)
                    {
                        GrupoDeProdutos prod = new GrupoDeProdutos();
                        prod.results_codigo = childNode["bm_grupo"].InnerText;
                        prod.results_name = childNode["bm_desc"].InnerText;

                        Insercoes.InsereGrupoProdutos(prod);
                    }
                }
                File.Delete(xmlData);
            }
            catch ( Exception e )
            {
                Insercoes.InsereLog(e.Message, false, "grupoProduto", "");
            }
        }

        static void Clientes()
        {
            try
            {
                string xmlData = @"C:\Importar\XML\clientes.xml";
            XmlDocument xmlDoc = new XmlDocument();
          

            xmlDoc.Load(xmlData); //Carregando o arquivo*/

            //Pegando elemento pelo nome da TAG
            XmlNodeList xnList = xmlDoc.SelectNodes("clientes");
            foreach (XmlNode xn in xnList)
            {
                    foreach (XmlNode childNode in xn.ChildNodes)
                    {
                        Cliente prod = new Cliente();
                        prod.accountnumber = childNode["a1_cgc"].InnerText;
                        prod.name = childNode["a1_nome"].InnerText;
                        prod.mps_fantasia = childNode["a1_nreduz"].InnerText;
                        prod.mps_inscricaoestadual = childNode["a1_inscr"].InnerText;
                        prod.emailaddress1 = childNode["a1_email"].InnerText;
                        prod.mps_lkpuf = childNode["a1_est"].InnerText;
                        prod.address1_city = childNode["a1_mun"].InnerText;
                        prod.address1_line1 = childNode["a1_end"].InnerText;
                        prod.address1_line2 = childNode["a1_bairro"].InnerText;
                        prod.address1_line3 = childNode["a1_nr_end"].InnerText;
                        prod.address1_postalcode = childNode["a1_cep"].InnerText;
                        //prod.mps_complemento = childNode["a1_complem].InnerText;;
                        prod.mps_cod_erp = childNode["a1_cod"].InnerText;
                        prod.mps_atendente = childNode["a3_nome"].InnerText;
                        prod.mps_ramos = childNode["a1_grpven"].InnerText;
                        prod.telephone1 = childNode["a1_tel"].InnerText;
                        prod.telephone2 = childNode["a1_telcont"].InnerText;
                        prod.ownerid = childNode["x_propri"].InnerText;

                        Insercoes.InsereClientes(prod);
                    }
                }
                File.Delete(xmlData);
            }
            catch (Exception e)
            {
                Insercoes.InsereLog("IntegracaoClientes", false, "Empresa", e.Message);
            }
        }

        static void ClientesExcel(FileInfo newFile)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                

                    int row = 2;
                    while (row <= sheet.Dimension.Rows)
                    {
                        Cliente prod = new Cliente();
                        prod.accountnumber = sheet.Cells[row, 1].Value.ToString().Trim(); //childNode["a1_cgc"].InnerText;
                        prod.address1_city = sheet.Cells[row, 2].Value.ToString().Trim(); //childNode["a1_mun"].InnerText;
                        prod.address1_line2 = sheet.Cells[row, 3].Value.ToString().Trim(); //a1_bairro
                        prod.address1_line1 = sheet.Cells[row, 4].Value.ToString().Trim(); //a1_end
                        prod.address1_line3 = sheet.Cells[row, 5].Value.ToString().Trim(); //a1_nr_end
                        prod.address1_postalcode = sheet.Cells[row, 6].Value.ToString().Trim(); //a1_cep
                        prod.emailaddress1 = sheet.Cells[row, 7].Value.ToString().Trim(); //childNode["a1_email"].InnerText;
                        prod.mps_atendente = sheet.Cells[row, 8].Value.ToString().Trim(); // childNode["a3_nome"].InnerText;
                        prod.mps_cod_erp = sheet.Cells[row, 9].Value; //childNode["a1_cod"].InnerText;
                        prod.mps_complemento = sheet.Cells[row, 10].Value.ToString().Trim(); //childNode["a1_complem].InnerText;

                        prod.mps_fantasia = sheet.Cells[row, 14].Value.ToString().Trim(); // childNode["a1_nreduz"].InnerText;

                        prod.mps_inscricaoestadual = sheet.Cells[row, 16].Value.ToString().Trim(); //childNode["a1_inscr"].InnerText;

                        prod.mps_lkpuf = sheet.Cells[row, 18].Value.ToString().Trim(); //childNode["a1_est"].InnerText;

                        prod.mps_ramos = sheet.Cells[row, 20].Value.ToString().Trim(); //childNode["a1_grpven"].InnerText;

                        prod.name = sheet.Cells[row, 22].Value.ToString().Trim(); //childNode["a1_nome"].InnerText;
                        prod.ownerid = sheet.Cells[row, 23].Value.ToString().Trim();//childNode["x_propri"].InnerText;
                        prod.telephone1 = sheet.Cells[row, 24].Value.ToString().Trim(); //childNode["a1_tel"].InnerText;
                        prod.telephone2 = sheet.Cells[row, 25].Value.ToString().Trim(); //childNode["a1_telcont"].InnerText;


                        Insercoes.InsereClientes(prod);
                        row++;
                    }
                }
               //File.Delete(xmlData);
            }
            catch (Exception e)
            {
                Insercoes.InsereLog("IntegracaoClientes", false, "Empresa", e.Message);
            }
        }

        /*static void Faturamentos()
        {
            try {
                 string xmlData = @"C:\Importar\XML\faturamento.xml";
                 XmlDocument xmlDoc = new XmlDocument();
                 xmlDoc.Load(xmlData); //Carregando o arquivo

              

                        //Pegando elemento pelo nome da TAG
                        XmlNodeList xnList = xmlDoc.SelectNodes("faturamento");
                foreach (XmlNode xn in xnList)
                {
                    foreach (XmlNode childNode in xn.ChildNodes)
                    {
                        Faturamento fat = new Faturamento();
                        fat.customerid = childNode["a1_cgc"].InnerText;
                        fat.codigoerp = childNode["a1_cliente"].InnerText;
                        fat.invoicenumber = childNode["f2_doc"].InnerText;
                        fat.pricelevelid = childNode["x_lista"].InnerText;
                        fat.ispricelocked = childNode["x_precbloq"].InnerText;
                        fat.name = childNode["a1_nome"].InnerText;
                        fat.totalamount = childNode["f2_valbrut"].InnerText;
                        fat.results_totalprodutos = childNode["f2_valmerc"].InnerText;
                        fat.freightamount = childNode["f2_frete"].InnerText;
                        fat.results_valorseguro = childNode["f2_seguro"].InnerText;
                        fat.totaldiscountamount = childNode["d2_desc"].InnerText;
                        fat.results_valoricms = childNode["f2_valicm"].InnerText;
                        fat.results_quantidade = childNode["d2_quant"].InnerText;
                        fat.description = childNode["x_descr"].InnerText;
                        fat.results_numeropedido = childNode["d2_pedido"].InnerText;
                        fat.ownerid = childNode["x_repres"].InnerText;
                        fat.salesrepid = childNode["x_repres"].InnerText;
                        fat.results_dataemissao = childNode["f2_emissao"].InnerText;

                        fat.datedelivered = childNode["x_dtentrega"].InnerText;
                        fat.billto_city = childNode["a1_mun"].InnerText;
                        fat.billto_line1 = childNode["a1_end"].InnerText;
                        fat.billto_line2 = childNode["a1_bairro"].InnerText;
                        fat.billto_line3 = childNode["a1_complem"].InnerText;
                        fat.billto_postalcode = childNode["a1_cep"].InnerText;
                        fat.billto_stateorprovince = childNode["a1_est"].InnerText;
                        fat.billto_telephone = childNode["a1_tel"].InnerText;
                        fat.results_inscricaoestadual = childNode["a1_inscr"].InnerText;
                        //ALIMENTAR OS DADOS DA TRASNPORTADORA

                        fat.results_notafiscal = childNode["d2_doc"].InnerText;
                        fat.productnumber = childNode["d2_cod"].InnerText;
                        fat.quantity = childNode["d2_quant"].InnerText;
                        fat.priceperunit = childNode["d2_prcven"].InnerText;
                        fat.results_valoripi = childNode["d2_ipi"].InnerText;
                        fat.results_valoricms = childNode["d2_icmsret"].InnerText;
                        fat.extendedamount = childNode["d2_total"].InnerText;
                        fat.tax = childNode["x_impostos"].InnerText;
                        fat.results_ncm = childNode["b1_posipi"].InnerText;
                        fat.Results_cst = childNode["d2_clasfis"].InnerText;
                        fat.results_cfop = childNode["d2_cf"].InnerText;
                        fat.uomi = childNode["d2_um"].InnerText;
                        fat.shippingmethodcode = childNode["c5_transp"].InnerText;
                        

                        Insercoes.InsereFaturamentos(fat);
                    }
                }
                File.Delete(xmlData);
            }
            catch
            {

            }
        }*/

        static void Faturamentos(FileInfo newFile)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets[1];


                    Int32 row = 1;
                    while (row <= sheet.Dimension.Rows)
                    {
                        row++;
                        Faturamento fat = new Faturamento();
                        fat.codigoerp = sheet.Cells[row, 1].Value.ToString();// childNode["a1_cliente"].InnerText;
                        fat.customerid = sheet.Cells[row, 2].Value.ToString().Trim(); //childNode["a1_cgc"].InnerText;
                        fat.invoicenumber = sheet.Cells[row, 3].Value.ToString().Trim();// childNode["f2_doc"].InnerText;
                        fat.pricelevelid = sheet.Cells[row, 4].Value != null ? sheet.Cells[row, 4].Value.ToString().Trim(): ""; //childNode["x_lista"].InnerText;
                        fat.ispricelocked = sheet.Cells[row, 5].Value != null ? sheet.Cells[row, 5].Value.ToString().Trim() : ""; //childNode["x_precbloq"].InnerText;
                        fat.name = sheet.Cells[row, 6].Value.ToString().Trim(); //childNode["a1_nome"].InnerText;
                        fat.totalamount = sheet.Cells[row, 7].Value.ToString().Trim(); //childNode["f2_valbrut"].InnerText;
                        fat.results_totalprodutos = sheet.Cells[row, 8].Value.ToString().Trim(); //childNode["f2_valmerc"].InnerText;
                        fat.freightamount = sheet.Cells[row, 9].Value.ToString().Trim(); //childNode["f2_frete"].InnerText;
                        fat.results_valorseguro = sheet.Cells[row, 10].Value.ToString().Trim(); //childNode["f2_seguro"].InnerText;
                        fat.totaldiscountamount = sheet.Cells[row, 11].Value.ToString().Trim(); //childNode["d2_desc"].InnerText;
                        fat.results_valoricms = sheet.Cells[row, 12].Value.ToString().Trim(); //childNode["f2_valicm"].InnerText;
                        fat.results_quantidade = sheet.Cells[row, 13].Value.ToString().Trim(); //childNode["d2_quant"].InnerText;
                        fat.description = sheet.Cells[row, 14].Value.ToString().Trim(); //childNode["x_descr"].InnerText;
                        fat.results_numeropedido = sheet.Cells[row, 15].Value.ToString().Trim(); //childNode["d2_pedido"].InnerText;
                        fat.ownerid = sheet.Cells[row, 16].Value.ToString().Trim(); // childNode["x_repres"].InnerText;
                        fat.results_dataemissao = sheet.Cells[row, 17].Value.ToString().Trim(); //childNode["f2_emissao"].InnerText;

                        fat.datedelivered = sheet.Cells[row, 18].Value.ToString().Trim(); //childNode["x_dtentrega"].InnerText;
                        fat.billto_city = sheet.Cells[row, 19].Value.ToString().Trim(); //childNode["a1_mun"].InnerText;
                        fat.billto_line1 = sheet.Cells[row, 20].Value.ToString().Trim(); //childNode["a1_end"].InnerText;
                        fat.billto_line2 = sheet.Cells[row, 21].Value.ToString().Trim(); //hildNode["a1_bairro"].InnerText;
                        fat.billto_line3 = sheet.Cells[row, 22].Value.ToString().Trim(); //childNode["a1_complem"].InnerText;
                        fat.billto_postalcode = sheet.Cells[row, 23].Value.ToString().Trim(); //childNode["a1_cep"].InnerText;
                        fat.billto_stateorprovince = sheet.Cells[row, 24].Value.ToString().Trim(); //childNode["a1_est"].InnerText;
                        fat.billto_telephone = sheet.Cells[row, 25].Value.ToString().Trim(); //childNode["a1_tel"].InnerText;
                        fat.results_inscricaoestadual = sheet.Cells[row, 26].Value.ToString().Trim(); //childNode["a1_inscr"].InnerText;
                                                                                                      //ALIMENTAR OS DADOS DA TRASNPORTADORA

                        fat.results_notafiscal = sheet.Cells[row, 27].Value.ToString().Trim(); //childNode["d2_doc"].InnerText;
                        fat.productnumber = sheet.Cells[row, 28].Value.ToString().Trim(); //childNode["d2_cod"].InnerText;
                        fat.quantity = sheet.Cells[row, 29].Value.ToString().Trim(); //childNode["d2_quant"].InnerText;
                        fat.priceperunit = sheet.Cells[row, 30].Value.ToString().Trim(); //childNode["d2_prcven"].InnerText;
                        fat.results_valoripi = sheet.Cells[row, 31].Value.ToString().Trim(); //childNode["d2_ipi"].InnerText;
                        fat.results_valoricms = sheet.Cells[row, 32].Value.ToString().Trim(); //childNode["d2_icmsret"].InnerText;
                        fat.extendedamount = sheet.Cells[row, 33].Value.ToString().Trim(); //childNode["d2_total"].InnerText;
                        fat.tax = sheet.Cells[row, 34].Value != null ? sheet.Cells[row, 34].Value.ToString().Trim() : ""; //childNode["x_impostos"].InnerText;
                        fat.results_ncm = sheet.Cells[row, 35].Value.ToString().Trim(); //childNode["b1_posipi"].InnerText;
                        fat.Results_cst = sheet.Cells[row, 36].Value.ToString().Trim();// childNode["d2_clasfis"].InnerText;
                        fat.results_cfop = sheet.Cells[row, 37].Value.ToString().Trim();// childNode["d2_cf"].InnerText;
                        fat.uomi = sheet.Cells[row, 38].Value.ToString().Trim();// childNode["d2_um"].InnerText;
                        fat.shippingmethodcode = sheet.Cells[row, 39].Value.ToString().Trim(); //childNode["c5_transp"].InnerText;

                        if (row > 41000)
                           Insercoes.InsereFaturamentos(fat);
                       
                    }
                }
                //  File.Delete(xmlData);
            }
            catch
            {

            }
        }


        static void AtualizarCodContato()
        {
            EntityCollection GetAll;
            QueryExpression qry = new QueryExpression();
            qry.EntityName = "contact";
            qry.ColumnSet = new ColumnSet("fullname");
            qry.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
            qry.Criteria.AddCondition("results_codcontato", ConditionOperator.Null);
            qry.AddOrder("fullname", OrderType.Ascending);

            EntityCollection list = Consulta.service.RetrieveMultiple(qry);

            GetAll = list;
            while (list.MoreRecords)
            {
                qry.PageInfo.PageNumber += 1;
                qry.PageInfo.PagingCookie = list.PagingCookie;
                list = Consulta.service.RetrieveMultiple(qry);

                foreach (Entity e in list.Entities)
                    GetAll.Entities.Add(e);
            }

            if (GetAll.Entities.Count > 0)
            {
                Int32 i = 19999;
                foreach (var item in GetAll.Entities)
                {
                    i++;
                    Entity contato = new Entity("contact", item.Id);
                    contato["results_codcontato"] = i.ToString();

                    Consulta.service.Update(contato);

                }
            }
        }
        
    }
}
