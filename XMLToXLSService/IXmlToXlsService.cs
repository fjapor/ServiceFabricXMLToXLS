using Microsoft.ServiceFabric.Services.Remoting;
using System.Threading.Tasks;

namespace XMLToXLSService
{
    interface IXmlToXlsService: IService
    {
        Task ExecuteXMLToXLS(string xlsFileName);
    }
}
