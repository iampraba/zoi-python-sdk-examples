import os

from officeintegrator.src.com.zoho.officeintegrator.exception.sdk_exception import SDKException
from officeintegrator.src.com.zoho.officeintegrator.dc import APIServer
from officeintegrator.src.com.zoho.api.authenticator import Auth
from officeintegrator.src.com.zoho.officeintegrator.util import StreamWrapper
from officeintegrator.src.com.zoho.officeintegrator.logger import Logger
from officeintegrator.src.com.zoho.officeintegrator import Initializer
from officeintegrator.src.com.zoho.officeintegrator.v1 import InvalidConfigurationException, \
    MergeAndDownloadDocumentParameters, FileBodyWrapper, Authentication
from officeintegrator.src.com.zoho.officeintegrator.v1.v1_operations import V1Operations

class MergeAndDownload:

    # Refer Preview API documentation - https://www.zoho.com/officeintegrator/api/v1/merge-document.html
    @staticmethod
    def execute():
        MergeAndDownload.init_sdk()
        parameter = MergeAndDownloadDocumentParameters()

        # Either use url as document source or attach the document in request body use below methods
        parameter.set_file_url('https://demo.office-integrator.com/zdocs/OfferLetter.zdoc')
        ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
        # filePath = ROOT_DIR + "/sample_documents/OfferLetter.zdoc"
        # print('Source document file path : ' + filePath)
        # parameter.set_file_content(StreamWrapper(file_path=filePath))

        parameter.set_merge_data_json_url("https://demo.office-integrator.com/data/candidates.json")
        # jsonFilePath = ROOT_DIR + "/sample_documents/candidates.json"
        # print('Data Source Json file to be path : ' + jsonFilePath)
        # parameter.set_merge_data_json_content(StreamWrapper(file_path=jsonFilePath))

        # parameter.set_merge_data_csv_url("https://demo.office-integrator.com/data/csv_data_source.csv")
        # filePath = ROOT_DIR + "/sample_documents/csv_data_source.csv"
        # print('Source document file path : ' + filePath)
        # parameter.set_merge_data_csv_content(StreamWrapper(file_path=filePath))

        parameter.set_output_format('pdf')
        parameter.set_password('***')

        v1Operations = V1Operations()
        response = v1Operations.merge_and_download_document(parameter)

        if response is not None:
            print('Status Code: ' + str(response.get_status_code()))
            responseObject = response.get_object()

            if responseObject is not None:
                if isinstance(responseObject, FileBodyWrapper):
                    convertedDocument = responseObject.get_file()

                    if isinstance(convertedDocument, StreamWrapper):
                        outputFileStream = convertedDocument.get_stream()
                        outputFilePath = ROOT_DIR + "/sample_documents/merge_output.pdf"

                        with open(outputFilePath, 'wb') as outputFileObj:
                            # while True:
                            #     # Read a chunk of data from the input stream
                            #     chunk = outputFileStream.read(1024)  # You can adjust the chunk size as needed
                            #
                            #     # If no more data is read, break the loop
                            #     if not chunk:
                            #         break
                            #
                            #     # Write the chunk of data to the file
                            #     outputFileObj.write(chunk)
                            outputFileObj.write(outputFileStream.content)

                        print("\nCheck converted output file in file path - " + outputFilePath)
                elif isinstance(responseObject, InvalidConfigurationException):
                    print('Invalid configuration exception.')
                    print('Error Code  : ' + str(responseObject.get_code()))
                    print("Error Message : " + str(responseObject.get_message()))
                    if responseObject.get_parameter_name() is not None:
                        print("Error Parameter Name : " + str(responseObject.get_parameter_name()))
                    if responseObject.get_key_name() is not None:
                        print("Error Key Name : " + str(responseObject.get_key_name()))
                else:
                    print('Document Conversion Request Failed')

    @staticmethod
    def init_sdk():
        try:
            #Sdk application log configuration
            logger = Logger.get_instance(Logger.Levels.INFO, "./logs.txt")
            #Update this apikey with your own apikey signed up in office integrator service
            auth = Auth.Builder().add_param("apikey", "2ae438cf864488657cc9754a27daa480").set_authentication_schema(Authentication.TokenFlow()).build()
            tokens = [ auth ]
            # Refer this help page for api end point domain details -  https://www.zoho.com/officeintegrator/api/v1/getting-started.html
            environment = APIServer.Production("https://api.office-integrator.com")

            Initializer.initialize(environment, tokens,None, None, logger, None)
        except SDKException as ex:
            print(ex.code)

MergeAndDownload.execute()