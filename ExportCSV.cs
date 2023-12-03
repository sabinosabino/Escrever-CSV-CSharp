using System.Globalization;
using System.Reflection;
using System.Text;

/*
    DEPENDE DE: dotnet add package Csv
*/

namespace ExpCSV
{
    public class ExportCSV
    {
        private string[] ObterNomesDasPropriedades(object instancia)
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

                    if (valor != null)
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
                var memory = await ExportCsvMemory(list);
                localExport = localExport == "local" ? Guid.NewGuid().ToString() + ".csv" : localExport;
                await File.WriteAllBytesAsync(localExport, memory.ToArray());
            }
            catch (System.Exception)
            {
                throw;
            }

        }
        public async Task<MemoryStream> ExportCsvMemory<T>(List<T> list)
        {
            try
            {
                var columnNames = ObterNomesDasPropriedades(list.FirstOrDefault());
                var rows = ObterArraysDeValores(list);
                var csv = await Csv.CsvWriter.WriteToTextAsync(columnNames, rows, ';');
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (StreamWriter streamWriter = new StreamWriter(memoryStream, Encoding.UTF8))
                    {
                        await streamWriter.WriteAsync(csv);
                        await streamWriter.FlushAsync();
                    }
                    return memoryStream;
                }
            }
            catch (Exception)
            {
                throw;
            }

        }
    }
}