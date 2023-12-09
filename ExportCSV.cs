using System.Globalization;
using System.Reflection;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;

/*
    DEPENDE DE: dotnet add package CsvHelper
*/

namespace ExpCSV
{
    public class ExportCSV
    {
        private string[] ObterNomesDasPropriedades(object instancia)
        {
            try
            {
                Type tipo = instancia.GetType();
                PropertyInfo[] propriedades = tipo.GetProperties(BindingFlags.Instance | BindingFlags.Public);

                string[] nomesDasPropriedades = new string[propriedades.Length];

                for (int i = 0; i < propriedades.Length; i++)
                {
                    nomesDasPropriedades[i] = propriedades[i].Name;
                }

                return nomesDasPropriedades;
            }
            catch (Exception)
            {

                throw;
            }

        }
        private async IAsyncEnumerable<string[]> ObterArraysDeValores<T>(List<T> listaDeObjetos)
        {
            foreach (var objeto in listaDeObjetos)
            {
                Type tipo = objeto.GetType();
                PropertyInfo[] propriedades = tipo.GetProperties(BindingFlags.Instance | BindingFlags.Public);

                string[] valoresDosCampos = new string[propriedades.Length];

                for (int i = 0; i < propriedades.Length; i++)
                {
                    object valor = propriedades[i].GetValue(objeto);

                    if (valor == null)
                    {
                        valoresDosCampos[i] = string.Empty;
                    }
                    else
                    {
                        if (propriedades[i].PropertyType == typeof(decimal))
                        {
                            valoresDosCampos[i] = ((decimal)valor).ToString(CultureInfo.CurrentCulture);
                        }
                        else if (propriedades[i].PropertyType == typeof(double))
                        {
                            valoresDosCampos[i] = ((double)valor).ToString("F2", CultureInfo.CurrentCulture);
                        }
                        else if (propriedades[i].PropertyType == typeof(DateTime))
                        {
                            valoresDosCampos[i] = ((DateTime)valor).ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.CurrentCulture);
                        }
                        else if (propriedades[i].PropertyType == typeof(int))
                        {
                            valoresDosCampos[i] = valor.ToString();
                        }
                        else if (propriedades[i].PropertyType == typeof(bool))
                        {
                            valoresDosCampos[i] = ((bool)valor) ? "Sim" : "Não";
                        }
                        else
                        {
                            valoresDosCampos[i] = valor.ToString();
                        }
                    }
                }

                yield return valoresDosCampos;
            }
        }

        public async Task ExportarCSV<T>(List<T> list, string localExport = "local")
        {
            try
            {
                var memory = await GeraCSV(list);
                localExport = localExport == "local" ? Guid.NewGuid().ToString() + ".csv" : localExport;
                await File.WriteAllBytesAsync(localExport, memory.ToArray());
            }
            catch (System.Exception)
            {
                throw;
            }

        }

        public async Task<byte[]> GeraCSV<T>(List<T> list)
        {
            using var memoryStream = new MemoryStream();
            using var write = new StreamWriter(memoryStream, System.Text.Encoding.UTF8);
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ";"
            };
            using var csv = new CsvWriter(write, config);

            try
            {
                await csv.WriteRecordsAsync(list);
                await write.FlushAsync();
                memoryStream.Position = 0;
                return memoryStream.ToArray();
            }
            catch (Exception)
            {

                throw new Exception("Erro ao criar arquivo.");
            }
        }

    }
}