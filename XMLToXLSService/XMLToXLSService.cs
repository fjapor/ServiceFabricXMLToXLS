using System;
using System.Collections.Generic;
using System.Fabric;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.ServiceFabric.Services.Communication.Runtime;
using Microsoft.ServiceFabric.Services.Runtime;
using Microsoft.ServiceFabric.Services.Remoting.Runtime;
using XMLToXLSService.XMLToXLSHandlers;

namespace XMLToXLSService
{
    internal sealed class XMLToXLSService : StatelessService, IXmlToXlsService
    {
        public XMLToXLSService(StatelessServiceContext context)
            : base(context)
        { }

        public Task ExecuteXMLToXLS(string xlsFileName)
        {
            if (!IsValidXLM(xlsFileName))
                LogError("Invalid XML File");

            var xmlhandler = GetXmlHandler();
            return Task.Run(() => xmlhandler.ExecuteXMLToXls(xlsFileName));
        }

        private void LogError(string error)
        {
            //Error logs come here
        }

        private bool IsValidXLM(string xlsFileName)
        {
            if (xlsFileName.ToLower().EndsWith("xml"))
                return true;

            return false;
        }

        protected override IEnumerable<ServiceInstanceListener> CreateServiceInstanceListeners()
        {
            //return new ServiceInstanceListener[0];

            return new[] { new ServiceInstanceListener(context =>
            this.CreateServiceRemotingListener(context)) };
        }

        protected override async Task RunAsync(CancellationToken cancellationToken)
        {
            long iterations = 0;

            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();

                ServiceEventSource.Current.ServiceMessage(this.Context, "Working-{0}", ++iterations);

                await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken);
            }
        }

        private IXmlToXls GetXmlHandler()
        {
            return new MinimalMemoryLoad();
        }
    }


}
