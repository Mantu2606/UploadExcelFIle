using ExcelDataReader;
using MySql.Data.MySqlClient;
using System.Data;
using Upload__ExcelAndCSV_Project.CommonLayer.Model;

namespace Upload__ExcelAndCSV_Project.DataAccessLayer
{
    public class UploadFileDL : IUploadFileDL
    {
        public readonly IConfiguration _configuration;
        public readonly MySqlConnection _mySqlConnection;
       
        public UploadFileDL(IConfiguration configuration)
        {
           _configuration = configuration;
            _mySqlConnection = new MySqlConnection(_configuration[key: "ConnectionStrings : MySqlDBConnectionString"]);
        }
        public async Task<UploadExcelFileResponse> UploadExcelFile(UploadExcelFileRequest request, string Path)
        { 
            UploadExcelFileResponse response = new UploadExcelFileResponse();
            List<ExcelBulkUploadParameter> Parameters = new List<ExcelBulkUploadParameter>();
            response.IsSuccess = true;
            response.Message = "Successful";
            try
            {
                if(_mySqlConnection.State != System.Data.ConnectionState.Open) 
                {
                    await _mySqlConnection.OpenAsync(); 
                }
                if (request.File.FileName.ToLower().Contains(value: ".xlsx"))
                 { 
                    FileStream stream = new FileStream(Path,FileMode.Open, FileAccess.Read, FileShare.ReadWrite );
                    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                    IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                    DataSet dataset = reader.AsDataSet(
                     configuration: new ExcelDataSetConfiguration()
                     {
                         UseColumnDataType = false,
                         ConfigureDataTable = (IExcelDataReader tableReader) => new ExcelDataTableConfiguration()
                         {  // UseHeaderRow => it is 1st heading row 
                             UseHeaderRow = true
                         }
                     });

                    for(int i = 0; i < dataset.Tables[index: 0].Rows.Count; i++)
                    {
                        ExcelBulkUploadParameter rows = new ExcelBulkUploadParameter();
                        rows.UserName = dataset.Tables[index: 0].Rows[i].ItemArray[0] != null ? Convert.ToString(dataset.Tables[index: 0].Rows[i].ItemArray[0]) : "-1";
                        rows.EmailID = dataset.Tables[index: 0].Rows[i].ItemArray[1] != null ? Convert.ToString(dataset.Tables[index: 0].Rows[i].ItemArray[1]) : "-1";
                        rows.MobileNumber = dataset.Tables[index: 0].Rows[i].ItemArray[2] != null ? Convert.ToString(dataset.Tables[index: 0].Rows[i].ItemArray[2]) : "-1";
                        rows.Age = dataset.Tables[index: 0].Rows[i].ItemArray[3] != null ? Convert.ToInt32(dataset.Tables[index: 0].Rows[i].ItemArray[3]) : -1;
                        rows.Salary = dataset.Tables[index: 0].Rows[i].ItemArray[4] != null ? Convert.ToInt32(dataset.Tables[index: 0].Rows[i].ItemArray[4]) : -1;
                        rows.Gender = dataset.Tables[index: 0].Rows[i].ItemArray[5] != null ? Convert.ToString(dataset.Tables[index: 0].Rows[i].ItemArray[5]) : "-1";
                        Parameters.Add(rows);  
                    }
                    stream.Close();
                    if(Parameters.Count > 0 )
                     {
                        string SqlQuery = @"INSERT INTO crudoperation.bulkuploadtable(UserName, EmailID, MobileNumber,Age, Salary, Gender) VALUES
                        (@UserName, @EmailID, @MobileNumber , @Age, @Salary, @Gender)"; 
                       
                        foreach(ExcelBulkUploadParameter rows in Parameters)
                        {
                         using (MySqlCommand sqlCommand = new MySqlCommand(SqlQuery,_mySqlConnection))
                            {
                                sqlCommand.CommandType = CommandType.Text; 
                                sqlCommand.CommandTimeout = 180;
                                sqlCommand.Parameters.AddWithValue(parameterName: "@UserName", rows.UserName);
                                sqlCommand.Parameters.AddWithValue(parameterName: "@EmailID", rows.EmailID);
                                sqlCommand.Parameters.AddWithValue(parameterName: "@MobileNumber", rows.MobileNumber);
                                sqlCommand.Parameters.AddWithValue(parameterName: "@Age", rows.Age);
                                sqlCommand.Parameters.AddWithValue(parameterName: "@Salary", rows.Salary);
                                sqlCommand.Parameters.AddWithValue(parameterName: "@Gender", rows.Gender);
                                int Status = await sqlCommand.ExecuteNonQueryAsync();                        
                                if(Status <= 0 ) 
                                { 
                                  response.IsSuccess = false;
                                    response.Message = "Query Not Executed"; 
                                    return response;
                                }
                            }
                        }   
                    }
                } 
                else
                {
                    response.IsSuccess = false;
                    response.Message = "InCurrect File";
                    return response; 
                }
            }
            catch(Exception ex)
            {
                response.IsSuccess = false; 
                response.Message = ex.Message;  
            }
            finally 
            {
                await _mySqlConnection.CloseAsync();   
                await _mySqlConnection.DisposeAsync();    
            }
            return response;    
        }
    }
}
