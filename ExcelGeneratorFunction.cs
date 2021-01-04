using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using AzureFunctionsDemo.Excel;
using System.Net.Http;
using System.Text;
using Microsoft.WindowsAzure.Storage.Auth;
using Microsoft.WindowsAzure.Storage;

namespace AzureFunctionsDemo
{
  public static class ExcelGeneratorFunction
  {
    [FunctionName("ExcelGeneratorFunction")]
    public static async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Function, "get", Route = null)] HttpRequest req,
        ILogger log)
    {
      try
      {
        log.LogInformation("C# HTTP trigger function processed a request.");

        string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
        string idUsuario = req.Headers["IdUsuario"];
        string indicador = req.Headers["Indicador"];
        dynamic data = JsonConvert.DeserializeObject(requestBody);

        var repository = new CommandRepository();
        IEnumerable<Command> result = await repository.GetCommandsThroughFunction(log);
        string[] nomeAtributosExibir = new string[] { "Id", "HowTo", "Line", "Plataform" };
        MemoryStream planilha = ExcelUtil.ObterPlanilhaPorLista(result, "Commands", "How to", nomeAtributosExibir, 0);
        var arquivoEnviado = await Upload(planilha, "Commands.xls", "application/vnd.ms-excel", idUsuario, indicador);
        return new OkObjectResult(arquivoEnviado);
      }
      catch (System.Exception e)
      {
        log.LogInformation(e.ToString());
        return new BadRequestResult();
      }

    }

    private static async Task<string> Upload(MemoryStream file, string fileName, string fileType, string idUsuario, string indicador)
    {
      var accountName = "devstoreaccount1";
      var accountKey = "Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsuFq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw==";
      var containerName = "samples-workitems";
      var blobEndpoint = new Uri("http://127.0.0.1:10000/devstoreaccount1");
      var storageCredentials = new StorageCredentials(accountName, accountKey);
      var storageAccount = new CloudStorageAccount(storageCredentials, blobEndpoint, null, null, null);
      var blobAzure = storageAccount.CreateCloudBlobClient();
      var container = blobAzure.GetContainerReference(containerName);
      var blob = container.GetBlockBlobReference(fileName);

      blob.Metadata["IdUsuario"] = idUsuario;
      blob.Metadata["Indicador"] = indicador;
      blob.Properties.ContentType = fileType;

      await blob.UploadFromByteArrayAsync(file.ToArray(), 0, file.ToArray().Length);
      return blob.SnapshotQualifiedStorageUri.PrimaryUri.ToString();
    }
  }



  public class CommandRepository
  {

    public async Task<IEnumerable<Command>> GetCommandsThroughFunction(ILogger log)
    {
      var connectionString = "Server=DPCRAMONPRATA\\SQLEXPRESS;Initial Catalog=CommanderDB;User ID=commander;Password=admin";
      var queryString = "SELECT * FROM dbo.Commands";
      var commands = new List<Command>();

      using (var connection = new SqlConnection(connectionString))
      {
        await connection.OpenAsync();
        SqlCommand command = new SqlCommand(queryString, connection);
        SqlDataReader reader = await command.ExecuteReaderAsync();
        while (reader.Read())
        {
          var newCommand = ReadSingleRow((IDataRecord)reader);
          commands.Add(newCommand);
        }
        reader.Close();
        connection.Close();
        return commands;
      }

    }

    private Command ReadSingleRow(IDataRecord record)
    {
      Command command = new Command();
      command.Id = (int)record["Id"];
      command.HowTo = (string)record["HowTo"];
      command.Line = (string)record["Line"];
      command.Plataform = (string)record["Plataform"];
      return command;
    }

  }

  public class Command
  {

    public int Id { get; set; }
    public string HowTo { get; set; }
    public string Line { get; set; }
    public string Plataform { get; set; }
  }

}
