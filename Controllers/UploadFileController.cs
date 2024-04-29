using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Upload__ExcelAndCSV_Project.CommonLayer.Model;
using Upload__ExcelAndCSV_Project.DataAccessLayer;

namespace Upload__ExcelAndCSV_Project.Controllers
{
    [Route("api/[controller]/[Action]")]
    [ApiController]
    public class UploadFileController : ControllerBase
    {
        public readonly IUploadFileDL _uploadFileDL;
       public UploadFileController(IUploadFileDL uploadFileDL) 
        {
            _uploadFileDL = uploadFileDL;   
        }
        [HttpPost]
        public async Task<IActionResult> UploadExcelFile(UploadExcelFileRequest request)
        {
            UploadExcelFileResponse response = new UploadExcelFileResponse();
            string Path = "UploadFileFolder/" + request.File.FileName; 
           
            try
            {
                using (FileStream stream = new FileStream(Path ,FileMode.CreateNew))
                {
                    await request.File.CopyToAsync(stream);
                }
                response = await _uploadFileDL.UploadExcelFile(request, Path); 
            }
            catch(Exception ex)
            {
              response.IsSuccess = false;   
              response.Message = ex.Message;    
            }
            return Ok(response);    
        }
    }
}
