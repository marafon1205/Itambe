using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace InativaClientes
{
    static class Program
    {
        public static void Main()
        {
            EntityCollection listaCliente = InativarClientes.Util.BuscaCliente();

                foreach (var clientes in listaCliente.Entities)
                {
                    EntityCollection faturaCliente = InativarClientes.Util.checkOneCondition("invoice", "customerid", clientes.Id.ToString(), InativarClientes.Util.service);
                    var details1 = faturaCliente.Entities.AsEnumerable().ToList();
                    var dataAuxMenos90dias = DateTime.Now.AddDays(-90);

                    if (details1.Count > 0)
                    {
                        var data = details1.Max(x => x.GetAttributeValue<DateTime>("results_dataemissao"));

                    if (data < dataAuxMenos90dias)
                    {
                        Entity entidade = new Entity("account", clientes.Id);
                        entidade.Attributes["mps_listtipoderelacao"] = new OptionSetValue(100000001);
                        InativarClientes.Util.service.Update(entidade);

                        EntityCollection contatos = InativarClientes.Util.checkOneCondition("contact", "parentcustomerid", clientes.Id.ToString(), InativarClientes.Util.service);
                        foreach (var cont in contatos.Entities)
                        {
                            Entity entidade2 = new Entity("contact", cont.Id);
                            entidade2.Attributes["mps_listtipoderelacao"] = new OptionSetValue(100000001);
                            InativarClientes.Util.service.Update(entidade2);
                        }

                        InativarClientes.Util.SetState(clientes, InativarClientes.Util.service, 1, 2);
                    }
                }
            }
        }
    }
}

