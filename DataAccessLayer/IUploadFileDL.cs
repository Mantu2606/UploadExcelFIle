using Upload__ExcelAndCSV_Project.CommonLayer.Model;

namespace Upload__ExcelAndCSV_Project.DataAccessLayer
{
    public interface IUploadFileDL
    {
        public Task<UploadExcelFileResponse> UploadExcelFile(UploadExcelFileRequest request , string Path);
    }
}
