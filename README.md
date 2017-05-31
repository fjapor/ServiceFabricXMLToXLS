# ServiceFabricXMLToXLS
Sample Azure Service Fabric Service to convert an XML to XLS

1- The windows service was implemented using Azure Service Fabric SDK (I'll call it ASF in the notes below)

2- There is a simple factory with two approaches (XMLToXLSService Class Line 65): The used one utilizes the minimum amount of memory, and takes some time to open/close the xls writing each xml entry. The second approach uses a bit of memory setting up an entity collection, that creates the xml faster.

3- The listener is ready to receive a filename using .net remoting. It seems ASF allows listening several protocols, however i'm not sure how this handler will work with service fabric, and i have not created the client service.

4- There is no type of log (i'm not sure how ASF handle filesystem or application logs), just the stubs inside the application

5- I haven't created detailed validators to the XML (created the stubs inside the code), but i'm considering receiving the xml in the right format.