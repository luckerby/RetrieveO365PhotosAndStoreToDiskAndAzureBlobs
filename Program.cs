using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Azure.Storage.Blobs;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;

namespace RetrieveO365PhotosAndInjectToStorage
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // The full path to where this process' image started from will be
            //  used to store the photos when writing to disk
            string pathToCurrentProcess = System.IO.Directory.GetCurrentDirectory();

            // Build a client application
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create("c190be44-4ed1-4106-b310-81dcb0e47d1e")
                .WithTenantId("022d92de-141e-4cb1-8578-e9af93f8ea31")
                .WithClientSecret("K1b.4B-~YxNwy5I3Y-x5RnEUh4.Jl7YRxW")
                .Build();
                
            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);
            // Create a new instance of GraphServiceClient with the authentication provider.
            // Using beta endpoint for now since it's only here where the photos can be retrieved successfully both
            //  for on-prem as well as cloud users
            GraphServiceClient graphClient = new GraphServiceClient("https://graph.microsoft.com/beta", authProvider);

            // == Retrieve the list of users from Azure AD ==
            // Build the request that will retrieve the users with enabled accounts
            var userRequest = graphClient.Users
                .Request()
                .Filter("accountEnabled eq true");

            // Setup the hash that will contain the pairs of user id and its corresponding employeeNumber
            Dictionary<string, string> userIdsAndEmployeeNumber = new Dictionary<string, string>();

            // Process each page of the user result set
            //do
            {
                var users = await userRequest.GetAsync();
                foreach (var user in users)
                {
                    // the Graph API will return this as a string
                    string employeeNumber = null;

                    // There are enabled accounts that don't have this extension (as
                    //  there's no backing employee number), so use try/catch
                    try
                    {
                        employeeNumber = (string)user.AdditionalData["extension_4922b266f3ce4f7ea0ba403a0bca8fc0_employeeNumber"];
                    }
                    catch
                    {
                    }

                    if (employeeNumber != null)
                        userIdsAndEmployeeNumber.Add(user.Id, employeeNumber);
                    Console.WriteLine($"{user.DisplayName} [no:{employeeNumber}]");
                }

                // Switch to the next page
                userRequest = users.NextPageRequest;
            } 
            //while (userRequest != null);


            // == Azure setup ==
            var connectionString = "DefaultEndpointsProtocol=https;AccountName=office365photosstorage;AccountKey=8uv+JdD6aBHWZH92hs1lMISOnqMkYdQ1PAJ4Dm5mM391x3hhZeaXTH3lEyEEEJ+oGsKkpJiNvL2X6kLxYCS2Fg==;EndpointSuffix=core.windows.net";
            // Build the object that will connect to the container within our target storage account
            BlobContainerClient blobContainerClient = new BlobContainerClient(connectionString, "photos");

            // == Pull the photos using Graph ==
            // Retrieve the photos for the list of users built previously
            foreach (string userId in userIdsAndEmployeeNumber.Keys)
            {
                Stream photoStream = default;
                try
                {
                    Task<Stream> photoTask = graphClient.Users[userId].Photo.Content.Request().GetAsync();
                    photoStream = await photoTask;
                }
                catch (Microsoft.Graph.ServiceException)
                {
                    // We'll end up here if the photo isn't found stamped on the user,
                    //  but we do nothing since photoStream will be left at its default (null)
                    //  value
                }

                if (photoStream != null)
                {
                    string filenameToWrite = userIdsAndEmployeeNumber[userId] + ".jpg";
                    string pathToFilenameToWrite = pathToCurrentProcess + '\\' + filenameToWrite;

                    Console.WriteLine(filenameToWrite);

                    // Write to local disk
                    await using (System.IO.FileStream fileStream =
                        new FileStream(pathToFilenameToWrite, System.IO.FileMode.Create))
                    {
                        await photoStream.CopyToAsync(fileStream);
                    }

                    // Rewind the position within the stream, otherwise we'll deadlock next
                    photoStream.Position = 0;

                    // Write to Azure Blob Storage
                    BlobClient blobClient = blobContainerClient.GetBlobClient(filenameToWrite);
                    await blobClient.UploadAsync(photoStream, true);
                }
            }
        }
    }
}
