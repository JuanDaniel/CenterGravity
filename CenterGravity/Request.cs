using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BBI.JD
{
    public enum RequestId : int
    {
        None = 0,
        CenterGravityFamily = 1,
        Select = 2,
        Update = 3,
        VisualizeCenterGravity = 4,
        RemoveCenterGravity = 5
    }

    public class Request
    {
        private int request = (int)RequestId.None;

        public RequestId Take()
        {
            return (RequestId)Interlocked.Exchange(ref request, (int)RequestId.None);
        }

        public void Make(RequestId r)
        {
            Interlocked.Exchange(ref request, (int)r);
        }
    }
}
