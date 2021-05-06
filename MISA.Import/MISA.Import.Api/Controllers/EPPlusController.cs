using Dapper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using MISA.Import.Api.Model;
using MySqlConnector;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace MISA.Import.Api.Controllers
{
    [Route("api/[controller]s")]
    [ApiController]
    public class EPPlusController : ControllerBase
    {
        /// <summary>
        /// Import dữ liệu từ file excel
        /// </summary>
        /// <param name="formFile">định dạng file</param>
        /// <param name="cancellationToken">hủy bỏ hoạt động</param>
        /// <returns>List danh sách dữ liệu trong file</returns>
        /// CreatedBy: LMDuc (06/05/2021)
        [HttpPost("import")]
        public async Task<CustomerResponse<List<Customer>>> Import(IFormFile formFile,
            CancellationToken cancellationToken)
        {
            //kiểm tra file null
            if (formFile == null || formFile.Length <= 0)
            {
                return CustomerResponse<List<Customer>>.GetResult(-1, "formfile is empty");
            }
            // kiểm tra định dạng file
            if (!Path.GetExtension(formFile.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return CustomerResponse<List<Customer>>.GetResult(-1, "Not support file extension");
            }
            //thực hiện get dữ liệu từ file
            var list = new List<Customer>();
            using (var stream = new MemoryStream())
            {
                await formFile.CopyToAsync(stream, cancellationToken);
                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;
                    for (int row = 3; row < rowCount; row++)
                    {
                        //get ngày tháng
                        //khởi tạo biến dateTime = ngày hiện tại
                        DateTime dateTime = DateTime.Today;
                        // kiểm tra null
                        if (worksheet.Cells[row, 6].Value == null)
                        {
                            // nếu null gán = ngày hiện tại
                            dateTime = DateTime.UtcNow;
                        }
                        else
                        {   // lấy dữ liệu dateTime trong file excel và parse về kiểu Datetime
                            string sDate = worksheet.Cells[row, 6].Value.ToString();
                            //dateTime = DateTime.ParseExact(sDate, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                            //validate datetime
                            //- Nếu chỉ nhập năm thì ngày sinh hiển thị sẽ là 01/01/[năm] 
                            //- Nếu nhập tháng/ năm thì sẽ lấy ngày là ngày 01: 01/[tháng]/[năm]
                            dateTime = (DateTime)validDateOfBirth(sDate);
                        }

                        // get email
                        //khởi tạo email = ""
                        var email = "";
                        // kiểm tra giá trị null khi lấy từ trong file excel
                        if (worksheet.Cells[row, 9].Value == null)
                        { // gán email = null nếu giá trị = null
                            email = null;
                        }
                        else
                        {// gán giá trị email = giá trị trong file excel
                            email = worksheet.Cells[row, 9].Value.ToString().Trim();
                        }
                        //thêm dữ liệu và list
                        list.Add(new Customer
                        {
                            CustomerCode = (worksheet.Cells[row, 1].Value == null) ? "" : worksheet.Cells[row, 1].Value.ToString().Trim(),
                            FullName = (worksheet.Cells[row, 2].Value == null) ? "" : worksheet.Cells[row, 2].Value.ToString().Trim(),
                            MemberCardCode = (worksheet.Cells[row, 3].Value == null) ? "" : worksheet.Cells[row, 3].Value.ToString().Trim(),
                            CustomerGroupName = (worksheet.Cells[row, 4].Value == null) ? "" : worksheet.Cells[row, 4].Value.ToString().Trim(),
                            PhoneNumber = (worksheet.Cells[row, 5].Value == null) ? "" : worksheet.Cells[row, 5].Value.ToString().Trim(),
                            //DateOfBirth = DateTime.Parse( worksheet.Cells[row, 6].Value.ToString()),
                            DateOfBirth = dateTime,
                            CompanyName = (worksheet.Cells[row, 7].Value == null) ? "" : worksheet.Cells[row, 7].Value.ToString().Trim(),
                            CompanyTaxCode = (worksheet.Cells[row, 8].Value == null) ? "" : worksheet.Cells[row, 8].Value.ToString().Trim(),
                            Email = email,
                            Address = (worksheet.Cells[row, 10].Value == null) ? "" : worksheet.Cells[row, 10].Value.ToString().Trim(),
                            Note = (worksheet.Cells[row, 11].Value == null) ? "" : worksheet.Cells[row, 11].Value.ToString().Trim()
                        });
                    }
                }
            }
            //kiểm tra các trường mã khách hàng, tên nhóm khách hàng, số điện thoại có bị trùng lặp trong file hoặc đã tồn tại trên hệ thống
            for (var k = 0; k < list.Count; k++)
            {
                list[k].Status = null;
                //kiểm tra trùng lặp trong file
                // khởi tạo biến check Mã khách hàng:
                // false - không có trong file
                // true - có trong file
                bool checkCustomerCode = false;
                // khởi tạo biến check số điện thoại:
                // false - không có trong file
                // true - có trong file
                bool checkPhoneNumber = false;
                for (int i = k + 1; i < list.Count; i++)
                {
                    // kiểm tra trùng mã khách hàng
                    if (list[k].CustomerCode == list[i].CustomerCode)
                    {
                        checkCustomerCode = true;
                    }
                    //Kiểm tra trùng số điện thoại
                    if (list[k].PhoneNumber == list[i].PhoneNumber)
                    {
                        checkPhoneNumber = true;
                    }
                }
                //add status
                if (checkCustomerCode)
                {
                    list[k].Status += "Mã khách hàng đã trùng với mã khách hàng khách trong tệp nhập khẩu. \n";
                }
                if (checkPhoneNumber)
                    list[k].Status += "Số điện thoại " + list[k].PhoneNumber + " đã trùng với số điện thoại của khách hàng khác trong tệp nhập khẩu.\n";
                //kiểm tra trùng trong cơ sở dữ liệu.
                check(list[k]);
            }
            // thêm vào db
            var rowAffect = Insert(list);
            return CustomerResponse<List<Customer>>.GetResult(0, "OK", list);

        }

        /// <summary>
        /// validate Ngày tháng
        /// </summary>
        /// <param name="sdate">chuỗi truyền vào</param>
        /// <returns>Định dạng ngày tháng năm</returns>
        private DateTime? validDateOfBirth(string sdate)
        {
            try
            {
                var araayDate = sdate.Split("/");
                if (araayDate.Length == 1)
                {
                    return new DateTime(int.Parse(araayDate[0]), 1, 1);
                }
                else if (araayDate.Length == 2)
                {
                    return new DateTime(int.Parse(araayDate[1]), int.Parse(araayDate[0]), 1);
                }
                else
                {
                    return new DateTime(int.Parse(araayDate[2]), int.Parse(araayDate[1]), int.Parse(araayDate[0]));
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        string connectionString = "Host = 47.241.69.179; " +
            "Port = 3306;" +
            "Database = MF819_Import_LMDuc;" +
            "User Id= dev;" +
            "Password = 12345678;Convert Zero Datetime=true;";
        protected IDbConnection dbConnection;

        /// <summary>
        /// Thêm dữ liệu vào database
        /// </summary>
        /// <param name="customers">list thông tin khách hàng</param>
        /// <returns>số bản ghi được thêm vào</returns>
        /// CreatedBy: LMDuc(06/05/2021)
        private int Insert(List<Customer> customers)
        {
            // kết nối db
            using (dbConnection = new MySqlConnection(connectionString))
            {
                var rowAffect = 0;
                //Thực hiện thêm
                foreach (var customer in customers)
                {
                    if (customer.Status == null)
                    {
                        dbConnection.Execute("Proc_InsertCustomer", param: customer, commandType: CommandType.StoredProcedure);
                        rowAffect++;
                    }
                }
                return rowAffect;
            }
        }
        /// <summary>
        /// Lấy dữ liệu  trong bảng
        /// </summary>
        /// <returns>danh sách bản ghi có db</returns>
        /// CreatedBy: LMDuc(06/05/2021)
        private IEnumerable<Customer> GetAll()
        {
            using (dbConnection = new MySqlConnection(connectionString))
            {
                var listCustomers = dbConnection.Query<Customer>("Proc_GetCustomer", commandType: CommandType.StoredProcedure);
                return listCustomers;
            }
        }

        private Customer check(Customer customer)
        {
            var listCustomerDB = GetAll();
            // khởi tạo biến check tên nhóm khách hàng:
            // false - không có trong hệ thống
            // true - có trong hệ thống
            bool checkCustomerGroupName = false;
            // khởi tạo biến check mã khách hàng:
            // false - không có trong hệ thống
            // true - có trong hệ thống
            bool checkCustomerCode = false;
            // khởi tạo biến check số điện thoại:
            // false - không có trong hệ thống
            // true - có trong hệ thống
            bool checkPhoneNumber = false;
            foreach (var item in listCustomerDB)
            {
                if (customer.CustomerGroupName == item.CustomerGroupName)
                {
                    checkCustomerGroupName = true;
                }
                if (customer.PhoneNumber == item.PhoneNumber)
                {
                    checkPhoneNumber = true;
                }
                if (customer.CustomerCode == item.CustomerCode)
                {
                    checkCustomerCode = true;
                }
            }
            //kiểm tra mã khách hàng
            if (checkCustomerCode)
            {
                customer.Status += "Mã khách hàng đã tồn tại trong hệ thống. \n";
            }
            // kiểm tra tên nhóm khách hàng
            if (checkCustomerGroupName == false)
            {
                customer.Status += "Nhóm khách hàng không có trong hệ thống. \n";
            }
            // kiểm tra số điện thoại
            if (checkPhoneNumber)
            {
                customer.Status += "Số điện thoại đã có trong hệ thống. \n";
            }
            return customer;
        }

    }
}
