﻿using System;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace ConnectSharePointOnline
{
    class AzureFunctionSettings
    {
        public string TenantId { get; set; }
        public string ClientId { get; set; }
        public StoreName CertificateStoreName { get; set; }
        public StoreLocation CertificateStoreLocation { get; set; }
        public string CertificateThumbprint { get; set; }
    }
}
