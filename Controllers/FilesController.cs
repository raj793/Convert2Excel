using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Web;
using System.Web.Http;
using Aspose.Cells;

namespace convert2excel
{
	[RoutePrefix("files")]
	public class FilesController : ApiController
	{
		[HttpPost]
		[Route("HtmlToExcel")]
		public HttpResponseMessage HtmltoExcel()
		{
			var httpRequest = HttpContext.Current.Request;
			//Set license
			//License license = new License();
			//FileStream myStream = new FileStream("Aspose.Words.lic", FileMode.Open, FileAccess.Read);
			//license.SetLicense(myStream);
			foreach (string file in httpRequest.Files)
			{
				var postFile = httpRequest.Files[file];
				if (postFile != null && postFile.ContentLength > 0)
				{
					Workbook book = new Workbook(postFile.InputStream);
					var originalFilename = postFile.FileName;
					var indexOfFileSeparator = originalFilename.IndexOf('.');
					var fileName = originalFilename.Substring(0, indexOfFileSeparator);
					MemoryStream outputStream = new MemoryStream();
					try
					{
						book.Save(outputStream, SaveFormat.Excel97To2003);
					}
					catch (Exception ex)
					{
						throw new Exception("Error converting file to html format : " + ex.Message);
					}
					outputStream.Position = 0;
					HttpResponseMessage response = new HttpResponseMessage();
					response.StatusCode = HttpStatusCode.OK;

					if (outputStream != null)
					{
						//Write the memory stream to HttpResponseMessage content
						response.Content = new StreamContent(outputStream);
						string contentDisposition = string.Concat("attachment; filename=", fileName + ".xlsx");
						response.Content.Headers.ContentDisposition =
						ContentDispositionHeaderValue.Parse(contentDisposition);
						response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
						response.Content.Headers.ContentLength = outputStream.Length;
						return response;
					}
				}
			}

			HttpResponseMessage response1 = new HttpResponseMessage();
			response1.StatusCode = HttpStatusCode.BadRequest;
			return response1;

		}
	}
}
