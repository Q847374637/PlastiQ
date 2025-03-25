using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlastiQ
{
    public static class SQLCOMMANDS
    {

        public static string DeleteOSTFinal = "DELETE * FROM ost_final";
        
        public static string DeleteOSTFinal_VBAL = "DELETE * FROM ost_final_vbal";
        
        public static string DeleteOSTMonitoring = "DELETE * FROM ost_monitoring";
        
        public static string DeleteOSTMonitoring_VBAL = "DELETE * FROM ost_monitoring_vbal";
        
        public static string DeleteOSTNew = "DELETE * FROM ost_new";
        
        public static string DeleteOSTNew_VBAL1 = "DELETE * FROM ost_new_vbal";
        
        public static string DeleteOSTNew_VBAL2 = "DELETE * FROM ost_new_vbal2";

        public static string clearBadClients = "Delete * from [bad_clients]";
        
        public static string clearBadAccounts = "Delete * from [bad_accounts]";
    }
}
