using System.Threading;

namespace BBI.JD
{
    public enum RequestId : int
    {
        None = 0,
        CenterGravityFamily = 1,
        Select = 2,
        PlaceCenterGravity = 3,
        RemoveCenterGravity = 4,
        CreateSchedule = 5
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
