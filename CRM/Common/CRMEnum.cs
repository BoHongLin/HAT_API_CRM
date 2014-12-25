using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRM.Common
{
    /***** fix time 2014/12/24  *****/


    public enum TransactionType
    {
        None = 0,
        Insert = 100000000,
        Update = 100000001
    }

    /***** Allen syncDetail 2014/12/24 start *****/
    public enum TransactionStatus
    {
        None = 0,
        Fail = 100000000,
        Success = 100000001,
        Partially = 100000002
    }
    /***** Allen syncDetail 2014/12/24 end *****/

    public enum ErrorType
    {
        None = 0,
        INI = 1,
        CRM = 2,
        DATASYNC = 3,
        DB = 4,
        DATASYNCDETAIL = 5
    }
}
