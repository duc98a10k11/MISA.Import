using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MISA.Import.Api.Model
{
    public class CustomerResponse<T>
    {
        public int Code { get; set; }
        public string Msg { get; set; }
        public T Data { get; set; }
        public static CustomerResponse<T> GetResult(int code, string msg, T data = default(T))
        {
            return new CustomerResponse<T>
            {
                Code = code,
                Msg = msg,
                Data = data,
                
            };
        }
    }
}
