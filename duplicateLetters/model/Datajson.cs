using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace duplicateLetters.model
{
    public class DataJson
    {
        public string Nomedapessoa { get; set; }
        //[JsonProperty("nomedapessoa")]
        public string Datadeadmissao { get; set; }
        public string Gestor { get; set; }
        public List<Periodo> Periodo { get; set; }
    }

    public class Periodo
    {
        public string Nome { get; set; }
        public string Data { get; set; }
    }
}
