using System.IO;

namespace AzureFunctionsDemo.Arquivos
{
  public static class ArquivosUtil
  {
    public static byte[] RetornarArquivo(string caminhoArquivo) => File.ReadAllBytes(caminhoArquivo);
  }
}
