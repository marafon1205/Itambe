using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Integracao
{
    public static class Comum
    {
        public static OptionSetValue Status(object tipo)
        {
            OptionSetValue valor = null;
            switch (Convert.ToInt16(tipo))
            {
                case 1: valor = new OptionSetValue(0); break;
                case 2: valor = new OptionSetValue(1); break;
               // case "RASCUNHO": valor = new OptionSetValue(2); break;
              //  case "EM REVISÃO": valor = new OptionSetValue(3); break;
            }
            return valor;
        }
        public static OptionSetValue TipoTransporte(object tipo)
        {
            int valor=0;
            switch (tipo.ToString().ToUpper())
            {

                case "FOB":
                    valor = 2; break;
                case "FOBD":
                    valor = 3; break;
                case "CIF":
                    valor = 4; break;
                case "Correio Comum":
                    valor = 5; break;
                case "Cliente Retira no Local":
                    valor = 7; break;
                case "Transportadora":
                    valor = 17; break;
                case "SEDEX":
                    valor = 23; break;
            }
            return new OptionSetValue(valor);
        }

        public static string FormatCNPJ(string CNPJ)
        {
            return Convert.ToUInt64(CNPJ).ToString(@"00\.000\.000\/0000\-00");
        }
    }
}
